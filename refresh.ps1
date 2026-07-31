<#
  투자관리 대시보드 데이터 수집기
  ----------------------------------------------------------------
  네이버 금융 시장지표 / 지수 API / 구글 뉴스 RSS 를 읽어 data.js 를 만듭니다.
  Windows 기본 PowerShell 5.1 에서 그대로 돌아가며, 별도 설치가 필요 없습니다.

  사용법:  powershell -ExecutionPolicy Bypass -File refresh.ps1
           또는 같은 폴더의 갱신.bat 더블클릭
#>
[CmdletBinding()]
param(
  [int] $NewsDays     = 30,   # 뉴스 수집 기간(일)
  [int] $NewsPerClass = 12    # 자산군당 최대 기사 수
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol =
  [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$UA      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36"
$OutFile = Join-Path $PSScriptRoot "data.js"
$errors  = New-Object System.Collections.Generic.List[string]

function Log([string]$msg) { Write-Host $msg }
function Fail([string]$msg) { $errors.Add($msg) | Out-Null; Write-Host ("  ! " + $msg) -ForegroundColor Yellow }

# ────────────────────────────────────────────────────────────────
# 공통 유틸
# ────────────────────────────────────────────────────────────────

function Get-Web {
  param([string]$Url, [hashtable]$Headers, [int]$Retries = 2)
  $h = @{ "User-Agent" = $UA }
  if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
  for ($i = 0; $i -le $Retries; $i++) {
    try {
      return (Invoke-WebRequest -Uri $Url -Headers $h -UseBasicParsing -TimeoutSec 25).Content
    } catch {
      if ($i -eq $Retries) { throw }
      Start-Sleep -Milliseconds 700
    }
  }
}

function Get-Group([string]$Html, [string]$Pattern, [int]$Index = 1) {
  $m = [regex]::Match($Html, $Pattern, "Singleline")
  if ($m.Success) { return $m.Groups[$Index].Value }
  return ""
}

# "1,433.30" → 1433.30 (실패하면 $null)
function To-Number([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $t = $s -replace "[,\s]", ""
  $d = 0.0
  if ([double]::TryParse($t, [ref]$d)) { return $d }
  return $null
}

# 등락률을 직접 계산한다 (시장지표 페이지는 절대 등락폭만 제공)
function Get-Ratio([double]$value, [double]$change, [string]$dir) {
  $signed = [Math]::Abs($change)
  if ($dir -eq "down") { $signed = -$signed }
  $prev = $value - $signed
  if ($prev -eq 0) { return $null }
  return [Math]::Round(($signed / $prev) * 100.0, 2)
}

function New-Metric {
  param(
    [string]$Name, [string]$Value, [string]$Unit,
    [string]$Change, [string]$ChangeUnit, $Ratio,
    [string]$Dir, [string]$Note, [string]$AsOf
  )
  return [ordered]@{
    name        = $Name
    value       = $Value
    unit        = $Unit
    change      = $Change
    change_unit = $ChangeUnit
    ratio       = $Ratio
    dir         = $Dir
    note        = $Note
    asof        = $AsOf
  }
}

# ────────────────────────────────────────────────────────────────
# 1. 네이버 금융 시장지표 (환율 / 국제환율 / 유가·금 / 국내금리)
# ────────────────────────────────────────────────────────────────

# <li> 카드 형태(환율·국제환율·유가금)를 공통 파싱
function Parse-Cards([string]$section) {
  $rows = @()
  if (-not $section) { return $rows }

  $blocks = [regex]::Split($section, '<li[^>]*>')
  foreach ($b in $blocks) {
    $name = Get-Group $b '<h3 class="h_lst"><span class="blind">([^<]+)</span>'
    if (-not $name) { continue }

    $headClass = Get-Group $b '<div class="head_info([^"]*)">'
    $inner     = Get-Group $b '<div class="head_info[^"]*">(.*?)</div>'
    if (-not $inner) { continue }

    $valueTxt  = (Get-Group $inner '<span class="value">([^<]*)</span>').Trim()
    $unit      = (Get-Group $inner '<span class="txt_[a-z]+"><span class="blind">([^<]*)</span>').Trim()
    $changeTxt = (Get-Group $inner '<span class="change">([^<]*)</span>').Trim()

    $dir = "flat"
    if ($headClass -match "point_up") { $dir = "up" }
    elseif ($headClass -match "point_dn") { $dir = "down" }

    $ratio = $null
    $v = To-Number $valueTxt
    $c = To-Number $changeTxt
    if ($null -ne $v -and $null -ne $c -and $dir -ne "flat") { $ratio = Get-Ratio $v $c $dir }

    $asof   = (Get-Group $b '<span class="time">([^<]*)</span>').Trim()
    $source = (Get-Group $b '<span class="source">([^<]*)</span>').Trim()

    $rows += (New-Metric -Name $name.Trim() -Value $valueTxt -Unit $unit `
                         -Change $changeTxt -ChangeUnit "" -Ratio $ratio `
                         -Dir $dir -Note $source -AsOf $asof)
  }
  return $rows
}

function Parse-Rates([string]$html) {
  $rows = @()
  $tbody = Get-Group $html '<h3 class="h_interest"><span>국내시장금리</span></h3>.*?<tbody>(.*?)</tbody>'
  if (-not $tbody) { return $rows }

  $trs = [regex]::Matches($tbody, '<tr[^>]*>(.*?)</tr>', "Singleline")
  foreach ($tr in $trs) {
    $row  = $tr.Groups[1].Value
    $name = (Get-Group $row '<th[^>]*>.*?<span>([^<]+)</span>').Trim()
    if (-not $name) { continue }

    $tds = [regex]::Matches($row, '<td[^>]*>(.*?)</td>', "Singleline")
    if ($tds.Count -lt 2) { continue }

    $valueTxt = ($tds[0].Groups[1].Value -replace '<[^>]+>', '').Trim()
    $chgCell  = $tds[1].Groups[1].Value
    $dirWord  = (Get-Group $chgCell 'alt="([^"]*)"').Trim()

    $dir = "flat"
    if ($dirWord -eq "상승") { $dir = "up" }
    elseif ($dirWord -eq "하락") { $dir = "down" }

    $changeTxt = ($chgCell -replace '<[^>]+>', '').Trim()
    $ratio = $null   # 금리는 %p 변화가 의미 있으므로 비율은 표시하지 않는다

    $rows += (New-Metric -Name $name -Value $valueTxt -Unit "%" `
                         -Change $changeTxt -ChangeUnit "%p" -Ratio $ratio `
                         -Dir $dir -Note "" -AsOf "")
  }
  return $rows
}

function Get-MarketIndex {
  Log "· 네이버 금융 시장지표 …"
  $html = Get-Web "https://finance.naver.com/marketindex/"

  $fxSec    = Get-Group $html 'id="exchangeList">(.*?)</ul>'
  $worldSec = Get-Group $html 'id="worldExchangeList">(.*?)</ul>'
  $oilSec   = Get-Group $html 'id="oilGoldList">(.*?)</ul>'

  return [ordered]@{
    fx     = Parse-Cards $fxSec
    world  = Parse-Cards $worldSec
    oil    = Parse-Cards $oilSec
    rates  = Parse-Rates $html
  }
}

# ────────────────────────────────────────────────────────────────
# 2. 주가지수
# ────────────────────────────────────────────────────────────────

$DOMESTIC_INDEX = @(
  @{ code = "KOSPI";  label = "코스피"  },
  @{ code = "KOSDAQ"; label = "코스닥"  }
)
$WORLD_INDEX = @(
  @{ code = ".DJI";  label = "다우존스"          },
  @{ code = ".IXIC"; label = "나스닥 종합"       },
  @{ code = ".INX";  label = "S&P 500"           },
  @{ code = ".SOX";  label = "필라델피아 반도체" },
  @{ code = ".N225"; label = "니케이 225"        },
  @{ code = ".HSI";  label = "항셍"              }
)

function Convert-IndexPayload($j, [string]$label, [string]$note) {
  $dir = "flat"
  $t = $j.compareToPreviousPrice.text
  if ($t -eq "상승") { $dir = "up" }
  elseif ($t -eq "하락") { $dir = "down" }

  $ratio = To-Number $j.fluctuationsRatio
  if ($null -ne $ratio -and $dir -eq "down" -and $ratio -gt 0) { $ratio = -$ratio }

  $asof = ""
  if ($j.localTradedAt) {
    try { $asof = ([datetimeoffset]$j.localTradedAt).ToString("yyyy.MM.dd HH:mm") } catch { $asof = "" }
  }

  return (New-Metric -Name $label -Value $j.closePrice -Unit "" `
                     -Change $j.compareToPreviousClosePrice -ChangeUnit "" -Ratio $ratio `
                     -Dir $dir -Note $note -AsOf $asof)
}

function Get-Indices {
  Log "· 주가지수 …"
  $rows = @()

  foreach ($ix in $DOMESTIC_INDEX) {
    try {
      $j = (Get-Web ("https://polling.finance.naver.com/api/realtime/domestic/index/" + $ix.code)) | ConvertFrom-Json
      $rows += (Convert-IndexPayload $j.datas[0] $ix.label "한국거래소")
    } catch { Fail ("지수 " + $ix.label + " 수집 실패") }
  }

  $h = @{ "Referer" = "https://m.stock.naver.com/"; "Accept" = "application/json" }
  foreach ($ix in $WORLD_INDEX) {
    try {
      $j = (Get-Web ("https://api.stock.naver.com/index/" + $ix.code + "/basic") $h) | ConvertFrom-Json
      $name = $ix.label
      if ($j.indexName) { $name = $j.indexName }
      $rows += (Convert-IndexPayload $j $name "해외 지수 · 종가 기준")
    } catch { Fail ("지수 " + $ix.label + " 수집 실패") }
  }
  return $rows
}

# ────────────────────────────────────────────────────────────────
# 3. 자산군별 뉴스 (구글 뉴스 RSS)
# ────────────────────────────────────────────────────────────────

$NEWS_CLASSES = @(
  @{ key="fx";       slot=1; label="환율·외환";   sublabel="원/달러, 외환시장";     query="환율 OR 원달러 OR 외환시장" },
  @{ key="kr";       slot=2; label="국내주식";     sublabel="코스피·코스닥";         query="코스피 OR 코스닥 OR 국내증시" },
  @{ key="us";       slot=3; label="해외주식";     sublabel="미국·글로벌 증시";      query='미국증시 OR 나스닥 OR "S&P500"' },
  @{ key="rate";     slot=4; label="금리·채권";    sublabel="기준금리, 국고채";      query="기준금리 OR 국고채 OR 채권시장" },
  @{ key="cmdt";     slot=5; label="원자재·금";    sublabel="유가, 금 시세";         query="국제유가 OR 금값 OR 원자재" },
  @{ key="realty";   slot=6; label="부동산";       sublabel="주택·상업용";           query="부동산시장 OR 아파트값 OR 주택시장" },
  @{ key="crypto";   slot=7; label="가상자산";     sublabel="비트코인 등";           query="비트코인 OR 가상자산 OR 이더리움" }
)

function Get-News([hashtable]$cls, [datetime]$cutoff) {
  $q   = [uri]::EscapeDataString($cls.query + " when:" + $NewsDays + "d")
  $url = "https://news.google.com/rss/search?q=$q&hl=ko&gl=KR&ceid=KR:ko"

  $xmlText = Get-Web $url
  $doc = New-Object System.Xml.XmlDocument
  $doc.LoadXml($xmlText)

  $items = @()
  $seen  = New-Object 'System.Collections.Generic.HashSet[string]'

  foreach ($node in $doc.SelectNodes("/rss/channel/item")) {
    $title = $node.SelectSingleNode("title").InnerText
    $link  = $node.SelectSingleNode("link").InnerText

    $source = ""
    $srcNode = $node.SelectSingleNode("source")
    if ($srcNode) { $source = $srcNode.InnerText }

    # 구글은 제목 끝에 " - 매체명"을 붙인다
    if ($source -and $title.EndsWith(" - " + $source)) {
      $title = $title.Substring(0, $title.Length - ($source.Length + 3))
    }
    $title = $title.Trim()
    if (-not $title) { continue }
    if (-not $seen.Add($title)) { continue }

    $published = ""
    $pubNode = $node.SelectSingleNode("pubDate")
    if ($pubNode) {
      try {
        $dt = [datetimeoffset]::Parse($pubNode.InnerText, [Globalization.CultureInfo]::InvariantCulture)
        if ($dt.UtcDateTime -lt $cutoff) { continue }
        $published = $dt.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
      } catch { }
    }

    $items += [ordered]@{
      title     = $title
      url       = $link
      source    = $source
      published = $published
    }
    if ($items.Count -ge $NewsPerClass) { break }
  }
  return $items
}

# ────────────────────────────────────────────────────────────────
# 4. 이전 수집값 (실패한 구간을 메우는 용도)
# ────────────────────────────────────────────────────────────────

function Get-Previous {
  if (-not (Test-Path $OutFile)) { return $null }
  try {
    $raw = [IO.File]::ReadAllText($OutFile, [Text.Encoding]::UTF8)
    $json = Get-Group $raw 'window\.DASHBOARD_DATA\s*=\s*(\{.*\})\s*;'
    if (-not $json) { return $null }
    return ($json | ConvertFrom-Json)
  } catch { return $null }
}

function Get-PrevGroup($prev, [string]$label) {
  if (-not $prev -or -not $prev.market) { return $null }
  foreach ($g in $prev.market) { if ($g.label -eq $label) { return $g } }
  return $null
}

function Get-PrevNews($prev, [string]$key) {
  if (-not $prev -or -not $prev.news) { return $null }
  foreach ($c in $prev.news) { if ($c.key -eq $key) { return $c } }
  return $null
}

# ────────────────────────────────────────────────────────────────
# 실행
# ────────────────────────────────────────────────────────────────

Log ""
Log "투자관리 대시보드 데이터를 수집합니다."
Log ""

$prev   = Get-Previous
$market = @()

# 지수
$indices = @()
try { $indices = Get-Indices } catch { Fail "주가지수 수집 실패" }
if ($indices.Count -gt 0) {
  $market += [ordered]@{ label="주요 지수"; note="국내 실시간 · 해외 최근 종가"; stale=$false; items=@($indices) }
} else {
  $p = Get-PrevGroup $prev "주요 지수"
  if ($p) { $market += [ordered]@{ label=$p.label; note=$p.note; stale=$true; items=@($p.items) } }
}

# 시장지표
$mi = $null
try { $mi = Get-MarketIndex } catch { Fail "네이버 금융 시장지표 수집 실패" }

$miGroups = @(
  @{ label="환율";           note="하나은행 고시 기준";       key="fx"    },
  @{ label="국내시장금리";   note="최종 고시치";              key="rates" },
  @{ label="유가·금 시세";   note="국제·국내 시세";           key="oil"   },
  @{ label="국제 시장 환율"; note="주요 통화쌍";              key="world" }
)

foreach ($g in $miGroups) {
  $items = @()
  if ($mi) { $items = @($mi[$g.key]) }
  if ($items.Count -gt 0) {
    $market += [ordered]@{ label=$g.label; note=$g.note; stale=$false; items=@($items) }
  } else {
    if ($mi) { Fail ($g.label + " 항목을 찾지 못했습니다 (페이지 구조 변경 가능성)") }
    $p = Get-PrevGroup $prev $g.label
    if ($p) { $market += [ordered]@{ label=$p.label; note=$p.note; stale=$true; items=@($p.items) } }
  }
}

# 뉴스
Log "· 자산군별 뉴스 …"
$cutoff = (Get-Date).ToUniversalTime().AddDays(-$NewsDays)
$news = @()
foreach ($cls in $NEWS_CLASSES) {
  $items = @()
  $stale = $false
  try {
    $items = @(Get-News $cls $cutoff)
  } catch {
    Fail ($cls.label + " 뉴스 수집 실패")
    $p = Get-PrevNews $prev $cls.key
    if ($p) { $items = @($p.items); $stale = $true }
  }
  Log ("    " + $cls.label + " " + $items.Count + "건")
  $news += [ordered]@{
    key      = $cls.key
    slot     = $cls.slot
    label    = $cls.label
    sublabel = $cls.sublabel
    stale    = $stale
    items    = @($items)
  }
  Start-Sleep -Milliseconds 250
}

# 결과 기록
# GitHub Actions 러너는 UTC 로 돌기 때문에 표시용 시각은 KST 로 직접 환산한다.
$nowUtc = (Get-Date).ToUniversalTime()
$nowKst = $nowUtc.AddHours(9)
$data = [ordered]@{
  generated_at         = $nowUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
  generated_at_display = $nowKst.ToString("yyyy년 M월 d일 HH:mm") + " KST"
  news_window_days     = $NewsDays
  market               = @($market)
  news                 = @($news)
  errors               = @($errors)
}

$json = $data | ConvertTo-Json -Depth 12
$body = @"
/* 이 파일은 refresh.ps1 이 자동으로 만듭니다. 직접 고치지 마세요. */
window.DASHBOARD_DATA = $json;
"@

[IO.File]::WriteAllText($OutFile, $body, (New-Object Text.UTF8Encoding($false)))

$metricCount = 0
foreach ($g in $market) { $metricCount += @($g.items).Count }
$newsCount = 0
foreach ($c in $news) { $newsCount += @($c.items).Count }

Log ""
Log ("완료 · 지표 " + $metricCount + "개, 뉴스 " + $newsCount + "건 → data.js")
if ($errors.Count -gt 0) {
  Log ("경고 " + $errors.Count + "건:")
  foreach ($e in $errors) { Log ("  - " + $e) }
}
Log "index.html 을 열고 새로고침하면 반영됩니다."
Log ""
