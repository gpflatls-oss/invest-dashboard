<#
  지표 이력 백필 (history.csv)
  ----------------------------------------------------------------
  refresh.ps1 은 "지금 값"만 쌓기 때문에, 처음 돌린 날 이전 이력이 없습니다.
  이 스크립트로 과거를 한 번 채워 넣습니다. 평소에는 돌릴 필요가 없습니다.

  사용법:
    powershell -ExecutionPolicy Bypass -File backfill_history.ps1
    powershell -ExecutionPolicy Bypass -File backfill_history.ps1 -From 2023-01-01
    powershell -ExecutionPolicy Bypass -File backfill_history.ps1 -Code USDKRW,KOSPI
    powershell -ExecutionPolicy Bypass -File backfill_history.ps1 -DryRun
    powershell -ExecutionPolicy Bypass -File backfill_history.ps1 -DiscoverEcos 721Y001

  이미 있는 (date,code) 는 건드리지 않습니다. 여러 번 돌려도 안전합니다.
  결측일(휴장·공휴일)은 채우지 않습니다 — 행이 없다는 것 자체가 정보입니다.
#>
[CmdletBinding()]
param(
  [string]   $From         = "",        # 기본: 3년 전
  [string]   $To           = "",        # 기본: 오늘 (KST)
  [string[]] $Code         = @(),       # 특정 지표만
  [string]   $EcosKey      = "",        # 없으면 환경변수 ECOS_API_KEY
  [string]   $DiscoverEcos = "",        # 통계표코드를 주면 항목 목록만 출력하고 종료
  [switch]   $DryRun
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol =
  [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$UA       = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36"
$HistFile = Join-Path $PSScriptRoot "history.csv"
$EPOCH    = [datetime]::SpecifyKind((Get-Date "1970-01-01"), [DateTimeKind]::Utc)

function Log([string]$m) { Write-Host $m }
function Warn([string]$m) { Write-Host ("  ! " + $m) -ForegroundColor Yellow }

function To-Number([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $d = 0.0
  if ([double]::TryParse(($s -replace "[,\s%]", ""), [Globalization.NumberStyles]::Float,
                         [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
  return $null
}

function To-Epoch([datetime]$d) { return [int64](($d.ToUniversalTime() - $EPOCH).TotalSeconds) }

function Get-Json([string]$Url, [hashtable]$Headers, [int]$Retries = 2) {
  $h = @{ "User-Agent" = $UA; "Accept" = "application/json" }
  if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
  for ($i = 0; $i -le $Retries; $i++) {
    try {
      return (Invoke-WebRequest -Uri $Url -Headers $h -UseBasicParsing -TimeoutSec 30).Content | ConvertFrom-Json
    } catch {
      if ($i -eq $Retries) { throw }
      Start-Sleep -Milliseconds 800
    }
  }
}

# ────────────────────────────────────────────────────────────────
# 백필 대상
# ────────────────────────────────────────────────────────────────
#
# code 는 refresh.ps1 의 $HISTORY_CODES / -Code 와 같은 값이어야 한다.
#
# scale : 소스 단위가 대시보드 표기와 다를 때 곱한다.
#         야후 JPYKRW=X 는 1엔당 원이고 대시보드는 100엔당 원이라 100 을 곱한다.
#
# 소스별 커버리지
#   yahoo : 키 불필요. 지수·환율·원자재.
#   naver : 키 불필요. 국고채. m.stock 채권 화면이 쓰는 일별 시세 API.
#   ecos  : 한국은행 키 필요. 네이버가 이력 API 를 주지 않는 국내 금리들.
#   (없음): 휘발유(오피넷)·국내 금은 과거 API 가 없어 백필하지 못한다.
#           refresh.ps1 이 도는 날부터 앞으로만 쌓인다.

$BACKFILL = @(
  # ── 주요 지수 (야후)
  @{ code = "KOSPI";       source = "yahoo"; symbol = "^KS11";     scale = 1 },
  @{ code = "KOSDAQ";      source = "yahoo"; symbol = "^KQ11";     scale = 1 },
  @{ code = "DJI";         source = "yahoo"; symbol = "^DJI";      scale = 1 },
  @{ code = "IXIC";        source = "yahoo"; symbol = "^IXIC";     scale = 1 },
  @{ code = "SPX";         source = "yahoo"; symbol = "^GSPC";     scale = 1 },
  @{ code = "SOX";         source = "yahoo"; symbol = "^SOX";      scale = 1 },
  @{ code = "N225";        source = "yahoo"; symbol = "^N225";     scale = 1 },
  @{ code = "HSI";         source = "yahoo"; symbol = "^HSI";      scale = 1 },

  # ── 환율 (야후 시장환율. 대시보드의 하나은행 고시와는 미세하게 다르다)
  @{ code = "USDKRW";      source = "yahoo"; symbol = "KRW=X";     scale = 1 },
  @{ code = "JPYKRW100";   source = "yahoo"; symbol = "JPYKRW=X";  scale = 100 },
  @{ code = "EURKRW";      source = "yahoo"; symbol = "EURKRW=X";  scale = 1 },
  # 야후의 CNYKRW=X 는 사실상 죽은 심볼이라(전 구간 1건) 원/달러 ÷ 달러/위안 으로 만든다
  @{ code = "CNYKRW";      source = "derived"; numer = "KRW=X"; denom = "CNY=X" },

  # ── 국제 시장 환율 (야후)
  @{ code = "USDJPY";      source = "yahoo"; symbol = "JPY=X";     scale = 1 },
  @{ code = "EURUSD";      source = "yahoo"; symbol = "EURUSD=X";  scale = 1 },
  @{ code = "GBPUSD";      source = "yahoo"; symbol = "GBPUSD=X";  scale = 1 },
  @{ code = "DXY";         source = "yahoo"; symbol = "DX-Y.NYB";  scale = 1 },

  # ── 유가·금 (야후 선물 근월물)
  @{ code = "WTI";         source = "yahoo"; symbol = "CL=F";      scale = 1 },
  @{ code = "GOLD_INTL";   source = "yahoo"; symbol = "GC=F";      scale = 1 },

  # ── 국고채 (네이버 채권 일별 시세)
  @{ code = "KTB3Y";       source = "naver"; symbol = "KR3YT=RR" },
  @{ code = "KTB5Y";       source = "naver"; symbol = "KR5YT=RR" },
  @{ code = "KTB10Y";      source = "naver"; symbol = "KR10YT=RR" },

  # ── 국내 금리 (ECOS 817Y002 = 시장금리(일별). 키가 없으면 통째로 건너뛴다)
  #
  #    통계표·항목 코드는 -DiscoverEcos 817Y002 로 확인한 값이다.
  #    721Y001(시장금리)에도 같은 이름의 항목이 있지만 그쪽은 연/분기/월만 있어
  #    일별 이력에는 쓸 수 없다.
  #
  #    COFIX(잔액/신규취급액)는 ECOS 에 없다. 은행연합회 고시라 과거 API 가
  #    없어 백필하지 못한다 — $NO_BACKFILL 참고.
  @{ code = "CALL";        source = "ecos"; stat = "817Y002"; item = "010101000"; cycle = "D" },
  @{ code = "CD91";        source = "ecos"; stat = "817Y002"; item = "010502000"; cycle = "D" },
  @{ code = "CORPAA3Y";    source = "ecos"; stat = "817Y002"; item = "010300000"; cycle = "D" }
)

# 백필 소스가 없는 지표 — 안내용
$NO_BACKFILL = @{
  "GASOLINE_KR" = "휘발유 (오피넷 · 과거 API 없음)"
  "GOLD_KR"     = "국내 금 (네이버 시세 · 과거 API 없음)"
  "COFIX_BAL"   = "COFIX 잔액 (은행연합회 고시 · ECOS 에 없음)"
  "COFIX_NEW"   = "COFIX 신규취급액 (은행연합회 고시 · ECOS 에 없음)"
}

# ────────────────────────────────────────────────────────────────
# 소스별 수집기 — 전부 @{ "yyyy-MM-dd" = 값 } 해시테이블을 돌려준다
# ────────────────────────────────────────────────────────────────

# 야후 봉 타임스탬프는 UTC 가 아니라 '거래소 현지 시각'이다. 예를 들어 KRW=X 는
# Europe/London 기준이라 서머타임 구간에는 현지 자정이 UTC 23:00(전날)로 찍힌다.
# 그대로 UTC 날짜로 읽으면 3~10월 데이터가 통째로 하루씩 밀린다(월요일 종가가
# 일요일로 들어간다). meta.gmtoffset 을 더한 뒤 가장 가까운 자정으로 반올림한다.
# gmtoffset 은 응답 시점의 값 하나뿐이라 과거 서머타임 구간에는 1시간이
# 어긋나는데, 장 시작이 현지 정오 이전이면 그 오차는 반올림이 흡수한다.
function Get-YahooSeries([string]$symbol, [datetime]$from, [datetime]$to) {
  $p1 = To-Epoch $from.AddDays(-4)
  $p2 = To-Epoch $to.AddDays(1)
  $url = "https://query1.finance.yahoo.com/v8/finance/chart/" +
         [Uri]::EscapeDataString($symbol) + "?period1=$p1&period2=$p2&interval=1d"

  $j = Get-Json $url
  $res = $j.chart.result
  if (-not $res -or $res.Count -eq 0) { return @{} }
  $r = $res[0]

  $stamps = @($r.timestamp)
  $closes = @($r.indicators.quote[0].close)
  $off = 0
  if ($r.meta -and $null -ne $r.meta.gmtoffset) { $off = [int]$r.meta.gmtoffset }

  $out = @{}
  for ($i = 0; $i -lt $stamps.Count; $i++) {
    if ($i -ge $closes.Count) { break }
    $c = $closes[$i]
    if ($null -eq $c) { continue }          # 휴장일. 직전값으로 메우지 않는다.
    $local = [double]($stamps[$i] + $off)
    $dayEpoch = [Math]::Round($local / 86400.0, 0, [MidpointRounding]::AwayFromZero) * 86400.0
    $out[$EPOCH.AddSeconds($dayEpoch).ToString("yyyy-MM-dd")] = [double]$c
  }
  return $out
}

# 네이버 채권 일별 시세. pageSize 는 60 을 넘기면 400 이 온다.
function Get-NaverBondSeries([string]$symbol, [datetime]$from, [datetime]$to) {
  $h = @{ "Referer" = "https://m.stock.naver.com/" }
  $out = @{}
  $fromText = $from.ToString("yyyy-MM-dd")

  for ($page = 1; $page -le 80; $page++) {
    $url = "https://api.stock.naver.com/marketindex/bond/" +
           [Uri]::EscapeDataString($symbol) + "/prices?page=$page&pageSize=60"
    # PowerShell 5.1 의 ConvertFrom-Json 은 JSON 배열을 원소별로 내보내지 않고
    # 배열 하나를 통째로 내보낸다. @() 로 감싸면 "원소 1개(=배열)" 가 되므로
    # foreach 로 한 겹 펼쳐야 한다.
    $resp = Get-Json $url $h
    $rows = @(); foreach ($r in $resp) { $rows += $r }
    if ($rows.Count -eq 0) { break }

    $oldest = ""
    foreach ($r in $rows) {
      if (-not $r.localTradedAt) { continue }
      $d = ([string]$r.localTradedAt).Substring(0, 10)
      $oldest = $d
      $v = To-Number ([string]$r.closePrice)
      if ($null -ne $v) { $out[$d] = $v }
    }

    # 요청 구간보다 과거로 넘어갔으면 더 볼 필요가 없다
    if ($oldest -and ($oldest -lt $fromText)) { break }
    Start-Sleep -Milliseconds 300
  }
  return $out
}

# 두 야후 시리즈를 나눠 만든 파생 시계열. 양쪽에 다 값이 있는 날만 남는다.
function Get-DerivedSeries($spec, [datetime]$from, [datetime]$to) {
  $a = Get-YahooSeries $spec.numer $from $to
  Start-Sleep -Milliseconds 400
  $b = Get-YahooSeries $spec.denom $from $to

  $out = @{}
  foreach ($d in $a.Keys) {
    if (-not $b.ContainsKey($d)) { continue }
    $den = [double]$b[$d]
    if ($den -eq 0) { continue }
    $out[$d] = [double]$a[$d] / $den
  }
  return $out
}

function Get-EcosKey {
  if ($EcosKey) { return $EcosKey }
  $v = [Environment]::GetEnvironmentVariable("ECOS_API_KEY")
  if ($v) { return $v.Trim() }
  return ""
}

function Format-EcosDate([datetime]$d, [string]$cycle) {
  if ($cycle -eq "M") { return $d.ToString("yyyyMM") }
  if ($cycle -eq "A") { return $d.ToString("yyyy") }
  if ($cycle -eq "Q") { return ("" + $d.Year + "Q" + [Math]::Floor(($d.Month - 1) / 3 + 1)) }
  return $d.ToString("yyyyMMdd")
}

function Parse-EcosTime([string]$t, [string]$cycle) {
  if ($cycle -eq "D" -and $t.Length -eq 8) { return $t.Substring(0,4) + "-" + $t.Substring(4,2) + "-" + $t.Substring(6,2) }
  if ($cycle -eq "M" -and $t.Length -eq 6) { return $t.Substring(0,4) + "-" + $t.Substring(4,2) + "-01" }
  if ($cycle -eq "A" -and $t.Length -eq 4) { return $t + "-01-01" }
  return ""
}

function Get-EcosSeries($spec, [datetime]$from, [datetime]$to) {
  $key = Get-EcosKey
  if (-not $key) { throw "ECOS_API_KEY 가 없습니다" }

  $cycle = $spec.cycle
  $s = Format-EcosDate $from $cycle
  $e = Format-EcosDate $to   $cycle
  $out = @{}
  $first = 1
  $itemName = ""

  for ($page = 0; $page -lt 40; $page++) {
    $url = "https://ecos.bok.or.kr/api/StatisticSearch/$key/json/kr/$first/" + ($first + 999) +
           "/" + $spec.stat + "/$cycle/$s/$e/" + $spec.item
    $j = Get-Json $url

    # 오류/자료없음은 RESULT 로 온다
    if ($j.RESULT) {
      if ($j.RESULT.CODE -eq "INFO-200") { break }
      throw ("" + $j.RESULT.CODE + " " + $j.RESULT.MESSAGE)
    }

    $rows = @($j.StatisticSearch.row)
    if ($rows.Count -eq 0) { break }

    foreach ($r in $rows) {
      if (-not $itemName -and $r.ITEM_NAME1) { $itemName = [string]$r.ITEM_NAME1 }
      $d = Parse-EcosTime ([string]$r.TIME) $cycle
      $v = To-Number ([string]$r.DATA_VALUE)
      if ($d -and $null -ne $v) { $out[$d] = $v }
    }

    $first += 1000
    if ($rows.Count -lt 1000) { break }
    Start-Sleep -Milliseconds 400
  }

  # 항목 코드가 맞는지 사람이 확인할 수 있게 실제 항목명을 남긴다
  if ($itemName) { Log ("      ECOS 항목명: " + $itemName) }
  return $out
}

# ECOS 통계표의 항목 코드를 훑어본다 (item 코드가 맞는지 확인용)
function Show-EcosItems([string]$stat) {
  $key = Get-EcosKey
  if (-not $key) { Warn "ECOS_API_KEY 가 없습니다."; return }
  $j = Get-Json "https://ecos.bok.or.kr/api/StatisticItemList/$key/json/kr/1/500/$stat"
  if ($j.RESULT) { Warn ("" + $j.RESULT.CODE + " " + $j.RESULT.MESSAGE); return }
  Log ("통계표 " + $stat + " 항목 목록")
  foreach ($r in @($j.StatisticItemList.row)) {
    Log ("  " + ([string]$r.ITEM_CODE).PadRight(14) + " " + $r.ITEM_NAME + "   [" + $r.CYCLE + "] " + $r.UNIT_NAME)
  }
}

# ────────────────────────────────────────────────────────────────
# history.csv 병합
# ────────────────────────────────────────────────────────────────

function Read-History {
  $map = New-Object 'System.Collections.Generic.Dictionary[string,string]'
  if (-not (Test-Path $HistFile)) { return $map }
  foreach ($line in [IO.File]::ReadAllLines($HistFile, [Text.Encoding]::UTF8)) {
    if (-not $line -or $line.StartsWith("date,")) { continue }
    $p = $line.Split(",")
    if ($p.Count -lt 3) { continue }
    $map[$p[0] + "|" + $p[1]] = $p[2]
  }
  return $map
}

function Save-History($map) {
  $keys = New-Object 'string[]' $map.Count
  $map.Keys.CopyTo($keys, 0)
  [Array]::Sort($keys, [StringComparer]::Ordinal)

  $sb = New-Object Text.StringBuilder
  [void]$sb.AppendLine("date,code,value")
  foreach ($k in $keys) {
    $i = $k.IndexOf("|")
    [void]$sb.AppendLine($k.Substring(0, $i) + "," + $k.Substring($i + 1) + "," + $map[$k])
  }
  [IO.File]::WriteAllText($HistFile, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
}

# ────────────────────────────────────────────────────────────────
# 실행
# ────────────────────────────────────────────────────────────────

if ($DiscoverEcos) { Show-EcosItems $DiscoverEcos; return }

$today = (Get-Date).ToUniversalTime().AddHours(9).Date
if ($From) { $fromDate = [datetime]::ParseExact($From, "yyyy-MM-dd", $null) } else { $fromDate = $today.AddYears(-3) }
if ($To)   { $toDate   = [datetime]::ParseExact($To,   "yyyy-MM-dd", $null) } else { $toDate   = $today }
if ($fromDate -gt $toDate) { throw "-From 이 -To 보다 뒤입니다." }

# powershell -File 로 부르면 -Code A,B,C 가 배열이 아니라 문자열 하나로 들어온다.
# 여기서 쉼표를 풀어 준다 (-Command 로 부르면 이미 배열이라 그대로 통과).
$Code = @($Code | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } |
          Where-Object { $_ })

$targets = $BACKFILL
if ($Code.Count -gt 0) {
  $wanted = @($Code | ForEach-Object { $_.ToUpperInvariant() })
  $targets = @($BACKFILL | Where-Object { $wanted -contains $_.code.ToUpperInvariant() })
  if ($targets.Count -eq 0) { throw ("해당하는 code 가 없습니다: " + ($Code -join ", ")) }
}

$hasEcos = [bool](Get-EcosKey)
if (-not $hasEcos) {
  $skipped = @($targets | Where-Object { $_.source -eq "ecos" })
  if ($skipped.Count -gt 0) {
    Warn ("ECOS_API_KEY 가 없어 국내 금리 " + $skipped.Count + "종을 건너뜁니다: " +
          (($skipped | ForEach-Object { $_.code }) -join ", "))
    Warn "  키는 https://ecos.bok.or.kr/api/ 에서 무료 발급됩니다."
    $targets = @($targets | Where-Object { $_.source -ne "ecos" })
  }
}

Log ""
Log ("지표 이력을 백필합니다 · " + $fromDate.ToString("yyyy-MM-dd") + " ~ " + $toDate.ToString("yyyy-MM-dd") +
     " · 대상 " + $targets.Count + "종")
if ($DryRun) { Log "(dry-run: history.csv 를 쓰지 않습니다)" }
Log ""

$map = Read-History
$before = $map.Count
$results = @()
$i = 0

foreach ($t in $targets) {
  $i++
  Write-Host ("  [{0,2}/{1}] {2,-12} " -f $i, $targets.Count, $t.code) -NoNewline
  $added = 0; $exist = 0; $err = ""
  $series = @{}

  try {
    switch ($t.source) {
      "yahoo"   { $series = Get-YahooSeries $t.symbol $fromDate $toDate }
      "naver"   { $series = Get-NaverBondSeries $t.symbol $fromDate $toDate }
      "ecos"    { $series = Get-EcosSeries $t $fromDate $toDate }
      "derived" { $series = Get-DerivedSeries $t $fromDate $toDate }
    }
  } catch {
    $err = $_.Exception.Message
  }

  if ($err) {
    Write-Host ("실패 · " + $err.Substring(0, [Math]::Min(60, $err.Length))) -ForegroundColor Yellow
    $results += [pscustomobject]@{ code=$t.code; source=$t.source; fetched=0; added=0; first=""; last=""; error=$err }
    Start-Sleep -Milliseconds 400
    continue
  }

  $scale = 1.0
  if ($t.ContainsKey("scale")) { $scale = [double]$t.scale }

  $dates = @($series.Keys | Sort-Object)
  foreach ($d in $dates) {
    if ($d -lt $fromDate.ToString("yyyy-MM-dd") -or $d -gt $toDate.ToString("yyyy-MM-dd")) { continue }
    $key = $d + "|" + $t.code
    if ($map.ContainsKey($key)) { $exist++; continue }   # 기존 값은 건드리지 않는다
    $v = [double]$series[$d] * $scale
    $map[$key] = $v.ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
    $added++
  }

  $kept = @($dates | Where-Object { $_ -ge $fromDate.ToString("yyyy-MM-dd") -and $_ -le $toDate.ToString("yyyy-MM-dd") })
  $f = ""; $l = ""
  if ($kept.Count -gt 0) { $f = $kept[0]; $l = $kept[$kept.Count - 1] }

  Write-Host ("수집 {0,5} → 신규 {1,5} (기존 {2})" -f $kept.Count, $added, $exist)
  $results += [pscustomobject]@{ code=$t.code; source=$t.source; fetched=$kept.Count; added=$added; first=$f; last=$l; error="" }
  Start-Sleep -Milliseconds 400
}

if (-not $DryRun) { Save-History $map }

Log ""
Log ("=" * 78)
Log "백필 결과"
Log ("=" * 78)
Log ("{0,-12} {1,-6} {2,7} {3,7}  {4}" -f "지표", "소스", "수집", "신규", "기간")
Log ("-" * 78)
foreach ($r in $results) {
  if ($r.error) {
    Log ("{0,-12} {1,-6} {2}" -f $r.code, $r.source, ("!! " + $r.error.Substring(0, [Math]::Min(50, $r.error.Length))))
  } else {
    Log ("{0,-12} {1,-6} {2,7} {3,7}  {4} ~ {5}" -f $r.code, $r.source, $r.fetched, $r.added, $r.first, $r.last)
  }
}
Log ("-" * 78)
Log ("{0,-12} {1,-6} {2,7} {3,7}" -f "합계", "", (($results | Measure-Object fetched -Sum).Sum), (($results | Measure-Object added -Sum).Sum))
Log ""
Log ("history.csv · " + $before + "행 → " + $map.Count + "행")

if ($NO_BACKFILL.Count -gt 0) {
  Log ""
  Log "과거 API 가 없어 백필하지 못하는 지표 (refresh.ps1 이 도는 날부터 앞으로만 쌓입니다):"
  foreach ($k in $NO_BACKFILL.Keys) { Log ("  " + $k.PadRight(14) + $NO_BACKFILL[$k]) }
}
Log ""
