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
  [int] $NewsDays         = 30,   # 뉴스 수집 기간(일)
  [int] $NewsPerClass     = 12,   # 자산군당 국내 매체 기사 수
  [int] $NewsIntlPerClass = 6     # 자산군당 해외 매체 기사 수 (한국어로 번역해 싣는다)
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol =
  [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$UA      = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0 Safari/537.36"
$OutFile  = Join-Path $PSScriptRoot "data.js"
$HistFile = Join-Path $PSScriptRoot "history.csv"
$errors   = New-Object System.Collections.Generic.List[string]

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
    [string]$Dir, [string]$Note, [string]$AsOf,
    [string]$Code = ""
  )
  # code 는 history.csv 의 시계열 식별자다. 표시명(name)은 네이버 표기가 바뀌면
  # 같이 바뀌지만 code 는 고정이라, 이력이 끊기지 않는다.
  # 이름으로 긁어오는 항목은 비워 두고 Resolve-MetricCode 가 나중에 채운다.
  return [ordered]@{
    name        = $Name
    code        = $Code
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

# 표시명 → 시계열 코드. 네이버에서 이름으로 긁어오는 항목들이 대상이다.
# (지수·국고채는 수집 루프에서 -Code 로 직접 넣으므로 여기 없다.)
#
# 이름이 바뀌면 이력이 끊기므로, 네이버 표기가 바뀌면 이 표만 고치면 된다.
# 여기에 없는 항목은 이력을 쌓지 않고 넘어간다 (수집은 정상, 차트만 없음).
$HISTORY_CODES = @{
  # 환율 (하나은행 고시)
  "미국 USD"          = "USDKRW"
  "일본 JPY(100엔)"   = "JPYKRW100"
  "유럽연합 EUR"      = "EURKRW"
  "중국 CNY"          = "CNYKRW"
  # 국제 시장 환율
  "달러/일본 엔"      = "USDJPY"
  "유로/달러"         = "EURUSD"
  "영국 파운드/달러"  = "GBPUSD"
  "달러인덱스"        = "DXY"
  # 유가·금
  "WTI"               = "WTI"
  "휘발유"            = "GASOLINE_KR"
  "국제 금"           = "GOLD_INTL"
  "국내 금"           = "GOLD_KR"
  # 국내시장금리
  "CD금리(91일)"      = "CD91"
  "콜 금리"           = "CALL"
  "회사채 AA- (3년)"  = "CORPAA3Y"
  "COFIX 잔액"        = "COFIX_BAL"
  "COFIX 신규취급액"  = "COFIX_NEW"
}

# 수집 루프에서 코드를 못 넣은 항목을 이름으로 채운다.
function Resolve-MetricCode($metric) {
  if ($metric.code) { return $metric.code }
  if ($HISTORY_CODES.ContainsKey($metric.name)) { return $HISTORY_CODES[$metric.name] }
  return ""
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

# 네이버는 회사채 등급을 표기하지 않는다. 국내시장금리의 회사채(3년)는
# 금융투자협회가 고시하는 무보증 3년 AA- 최종호가수익률이다.
# 실제 값으로도 확인된다 — 국고채 3년 대비 스프레드가 70bp 안팎으로,
# BBB- 였다면 수백 bp 벌어져 있어야 한다.
$RATE_LABELS = @{
  "회사채 (3년)" = @{ name = "회사채 AA- (3년)"; note = "무보증 3년 · AA- 등급" }
  "국고채 (3년)" = @{ name = "국고채 3년";       note = "" }
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

    $note = ""
    if ($RATE_LABELS.ContainsKey($name)) {
      $note = $RATE_LABELS[$name].note
      $name = $RATE_LABELS[$name].name
    }

    $rows += (New-Metric -Name $name -Value $valueTxt -Unit "%" `
                         -Change $changeTxt -ChangeUnit "%p" -Ratio $ratio `
                         -Dir $dir -Note $note -AsOf "")
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
# 1-2. 한국은행 ECOS — 네이버에 없는 만기·등급
# ────────────────────────────────────────────────────────────────

# 네이버 시장지표 표에는 국고채 3년밖에 없다. 장기물은 모바일 채권 화면이 쓰는
# API 에서 받는다. 코드는 로이터 형식이다 (KR5YT=RR).
# 1·2·20·30년도 같은 방식으로 받을 수 있다.
$BOND_TENORS = @(
  @{ code = "KR3YT=RR";  label = "국고채 3년";  hist = "KTB3Y"  },
  @{ code = "KR5YT=RR";  label = "국고채 5년";  hist = "KTB5Y"  },
  @{ code = "KR10YT=RR"; label = "국고채 10년"; hist = "KTB10Y" }
)

function Get-BondYields {
  $rows = @()
  Log "· 국고채 수익률 …"

  $h = @{ "Referer" = "https://m.stock.naver.com/"; "Accept" = "application/json" }
  foreach ($t in $BOND_TENORS) {
    try {
      $j = (Get-Web ("https://api.stock.naver.com/marketindex/bond/" + $t.code) $h) | ConvertFrom-Json

      $dir = "flat"
      if ($j.fluctuationsType.text -eq "상승") { $dir = "up" }
      elseif ($j.fluctuationsType.text -eq "하락") { $dir = "down" }

      $val = To-Number $j.closePrice
      $chg = To-Number $j.fluctuations
      if ($null -eq $val) { Fail ($t.label + " 값을 읽지 못했습니다"); continue }

      $valTxt = $val.ToString("0.000")
      $chgTxt = "0.000"
      if ($null -ne $chg) { $chgTxt = [Math]::Abs($chg).ToString("0.000") }

      $asof = ""
      if ($j.localTradedAt) {
        try { $asof = ([datetimeoffset]$j.localTradedAt).ToString("yyyy.MM.dd HH:mm") } catch { }
      }

      # 시장지표 표의 국고채는 최종호가수익률이고 이쪽은 체결 기준이라 값이 조금 다르다
      $rows += (New-Metric -Name $t.label -Value $valTxt -Unit "%" `
                           -Change $chgTxt -ChangeUnit "%p" -Ratio $null `
                           -Dir $dir -Note "체결 기준" -AsOf $asof -Code $t.hist)
    } catch {
      Fail ($t.label + " 수집 실패")
    }
  }
  return $rows
}

# 만기가 뒤섞이면 읽기 나쁘다. 짧은 것부터 긴 것 순으로 세운다.
$RATE_ORDER = @(
  "CD금리(91일)", "콜 금리",
  "국고채 3년", "국고채 5년", "국고채 10년",
  "회사채 AA- (3년)",
  "COFIX 잔액", "COFIX 신규취급액"
)

function Sort-Rates($rates) {
  return @($rates | Sort-Object -Property @{ Expression = {
    $i = $RATE_ORDER.IndexOf($_.name)
    if ($i -lt 0) { 99 } else { $i }
  } })
}

# ────────────────────────────────────────────────────────────────
# 2. 주가지수
# ────────────────────────────────────────────────────────────────

$DOMESTIC_INDEX = @(
  @{ code = "KOSPI";  label = "코스피";  hist = "KOSPI"  },
  @{ code = "KOSDAQ"; label = "코스닥";  hist = "KOSDAQ" }
)
$WORLD_INDEX = @(
  @{ code = ".DJI";  label = "다우존스";          hist = "DJI"  },
  @{ code = ".IXIC"; label = "나스닥 종합";       hist = "IXIC" },
  @{ code = ".INX";  label = "S&P 500";           hist = "SPX"  },
  @{ code = ".SOX";  label = "필라델피아 반도체"; hist = "SOX"  },
  @{ code = ".N225"; label = "니케이 225";        hist = "N225" },
  @{ code = ".HSI";  label = "항셍";              hist = "HSI"  }
)

function Convert-IndexPayload($j, [string]$label, [string]$note, [string]$hist = "") {
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
                     -Dir $dir -Note $note -AsOf $asof -Code $hist)
}

function Get-Indices {
  Log "· 주가지수 …"
  $rows = @()

  foreach ($ix in $DOMESTIC_INDEX) {
    try {
      $j = (Get-Web ("https://polling.finance.naver.com/api/realtime/domestic/index/" + $ix.code)) | ConvertFrom-Json
      $rows += (Convert-IndexPayload $j.datas[0] $ix.label "한국거래소" $ix.hist)
    } catch { Fail ("지수 " + $ix.label + " 수집 실패") }
  }

  $h = @{ "Referer" = "https://m.stock.naver.com/"; "Accept" = "application/json" }
  foreach ($ix in $WORLD_INDEX) {
    try {
      $j = (Get-Web ("https://api.stock.naver.com/index/" + $ix.code + "/basic") $h) | ConvertFrom-Json
      $name = $ix.label
      if ($j.indexName) { $name = $j.indexName }
      $rows += (Convert-IndexPayload $j $name "해외 지수 · 종가 기준" $ix.hist)
    } catch { Fail ("지수 " + $ix.label + " 수집 실패") }
  }
  return $rows
}

# ────────────────────────────────────────────────────────────────
# 3. 자산군별 뉴스 (구글 뉴스 RSS)
# ────────────────────────────────────────────────────────────────

# 자산군마다 쿼리를 여러 개 던진다. 하나만 쓰면 그날 큰 기사 하나에 목록이
# 통째로 잠식되거나(예: 잠실 MICE), 반대로 너무 좁아 몇 건밖에 안 잡힌다.
$NEWS_CLASSES = @(
  @{ key="pe";        slot=1; label="Private Equity";  sublabel="사모펀드·바이아웃"
     queries = @(
       '"사모펀드" OR "PEF" OR "프라이빗에쿼티"',
       '"바이아웃 펀드" OR "경영권 인수" OR "블라인드 펀드 결성"',
       '"MBK파트너스" OR "한앤컴퍼니" OR "IMM PE" OR "스틱인베스트먼트"'
     )
     filter = @('사모|PEF|프라이빗|바이아웃|경영권|MBK|한앤|IMM|스틱')
     queriesEn = @(
       '"private equity" OR "buyout fund"',
       '"Blackstone" OR "KKR" OR "Carlyle" OR "Apollo Global"',
       '"PE firm acquires" OR "take-private deal" OR "GP stakes"'
     )
     filterEn = @('(?i)private equity|buyout|blackstone|kkr|carlyle|apollo|take-private|pe firm') },

  @{ key="pd";        slot=2; label="Private Debt";    sublabel="사모대출·프라이빗 크레딧"
     queries = @(
       '"사모대출" OR "프라이빗 크레딧" OR "private credit"',
       '"메자닌 투자" OR "인수금융" OR "사모사채"',
       '"대출채권 투자" OR "직접대출 펀드" OR "private debt"'
     )
     filter = @('사모대출|프라이빗|크레딧|메자닌|인수금융|사모사채|대출채권|직접대출|private')
     queriesEn = @(
       '"private credit" OR "private debt"',
       '"direct lending" OR "leveraged loan market" OR "BDC"',
       '"credit fund" OR "mezzanine financing" OR "CLO market"'
     )
     filterEn = @('(?i)private credit|private debt|direct lending|leveraged loan|bdc|mezzanine|clo|credit fund') },

  @{ key="vc";        slot=3; label="Venture Capital"; sublabel="벤처투자·스타트업"
     queries = @(
       '"벤처캐피탈" OR "벤처캐피털"',
       '"시리즈A 투자유치" OR "시리즈B 투자유치" OR "시리즈C 투자유치"',
       '"벤처펀드 결성" OR "모태펀드" OR "신기술투자조합"'
     )
     filter = @('벤처|스타트업|시리즈\s?[ABC]|모태펀드|투자조합|VC')
     queriesEn = @(
       '"venture capital" OR "VC funding"',
       '"Series A funding" OR "Series B funding" OR "seed round"',
       # 그냥 "Sequoia"로 두면 세쿼이아 나무와 동명의 항암제 임상시험이 딸려온다
       '"Sequoia Capital" OR "Andreessen Horowitz" OR "startup raises"'
     )
     filterEn = @('(?i)venture capital|vc fund|series [abc]|seed round|sequoia capital|andreessen|startup rais') },

  # 국내 기사와 섞이지 않도록 "부동산 낱말"과 "해외 낱말"을 둘 다 요구한다.
  @{ key="re_intl";   slot=4; label="해외 부동산";      sublabel="글로벌 상업용·오피스"
     queries = @(
       '"해외 부동산 투자" OR "해외부동산 펀드"',
       '"글로벌 상업용 부동산" OR "해외 오피스 빌딩" OR "미국 상업용 부동산"',
       '"유럽 부동산 시장" OR "일본 부동산 투자" OR "해외 물류센터 투자"',
       '"미국 리츠" OR "일본 리츠"',   # "글로벌 리츠"는 국내 상장사 이름에 걸린다
       '"뉴욕 오피스" OR "런던 오피스" OR "도쿄 부동산" OR "해외 오피스 공실률"',
       '"글로벌 부동산 시장" OR "해외 부동산 펀드 손실" OR "해외 대체투자"',
       '"싱가포르 부동산" OR "베트남 부동산" OR "호주 부동산 투자"'
     )
     filter = @(
       '부동산|리츠|오피스|빌딩|물류센터|주택|호텔|상업용',
       '해외|글로벌|미국|美|유럽|일본|日|뉴욕|런던|도쿄|중국|호주|싱가포르|베트남|아시아'
     )
     queriesEn = @(
       '"commercial real estate" OR "office market"',
       '"real estate fund" OR "property investment" OR "REIT market"',
       '"office vacancy" OR "logistics property" OR "data center real estate"'
     )
     filterEn = @('(?i)real estate|office|property|reit|logistics|vacancy|landlord') },

  # "인프라 펀드" 처럼 넓은 말은 국내 BTL 기사까지 끌어와 국내 인프라와 겹친다.
  @{ key="infra_intl";slot=5; label="해외 인프라";      sublabel="글로벌 인프라 자산"
     queries = @(
       '"해외 인프라 투자" OR "글로벌 인프라 펀드"',
       '"브룩필드" OR "해외 발전소 인수" OR "글로벌 인프라 자산"',
       '"해외 신재생 발전 투자" OR "해외 데이터센터 투자" OR "해외 공항 민영화"',
       '"해외 도로 사업" OR "해외 항만 투자" OR "글로벌 에너지 인프라"',
       '"해외 전력망 투자" OR "글로벌 인프라 M&A" OR "해외 수처리 사업"'
     )
     filter = @(
       '인프라|발전|데이터센터|공항|항만|철도|도로|파이프라인|신재생|전력|ESS',
       '해외|글로벌|미국|美|유럽|일본|日|브룩필드|맥쿼리|블랙스톤|아시아|중동'
     )
     queriesEn = @(
       # "infrastructure investment" 은 미국 법률명(IIJA)에 걸려 엉뚱한 기사를 부른다
       '"infrastructure fund" OR "infrastructure investor" OR "infrastructure deal"',
       '"Brookfield" OR "Macquarie Asset Management" OR "toll road concession"',
       '"data center investment" OR "renewable energy project finance" OR "airport privatisation"'
     )
     filterEn = @('(?i)infrastructure|data cent|renewable|power plant|grid|airport|port|toll road|pipeline|brookfield|macquarie') },

  # "리츠" 만 쓰면 미국 리츠 기사가 섞여 들어온다.
  @{ key="re_kr";     slot=6; label="국내 부동산";      sublabel="상업용·리츠·개발"
     queries = @(
       '"국내 상업용 부동산" OR "오피스 빌딩 매각"',
       '"상장리츠" OR "리츠 배당"',
       '"부동산 PF" OR "물류센터 거래" OR "부동산 개발사업"',
       '"이지스자산운용" OR "코람코자산신탁" OR "마스턴투자운용" OR "캡스톤자산운용"',
       '"서울 오피스 공실률" OR "지식산업센터" OR "데이터센터 부지"'
     )
     filter = @('부동산|리츠|오피스|빌딩|물류센터|PF|개발사업|상업용')
     # 국내 자산군이라 해외 매체는 한국 시장을 다룬 기사만 의미가 있다
     queriesEn = @(
       '"Korea real estate" OR "Seoul office market"',
       '"Korean property market" OR "Korea REIT"'
     )
     filterEn = @('(?i)korea|korean|seoul', '(?i)real estate|office|property|reit|housing') },

  @{ key="infra_kr";  slot=7; label="국내 인프라";      sublabel="민자사업·SOC"
     queries = @(
       '"민자사업" OR "BTL 사업" OR "BTO 사업"',
       '"사회기반시설 투자" OR "국내 인프라 펀드"',
       '"도로 민자" OR "철도 민자사업" OR "환경기초시설 민자"',
       '"국가철도망" OR "GTX 사업" OR "고속도로 민자"',
       '"신재생 프로젝트 파이낸싱" OR "국내 발전소 매각" OR "인프라 자산 인수"'
     )
     filter = @('민자|BTL|BTO|사회기반시설|인프라|SOC|민투심|실시협약')
     queriesEn = @(
       '"Korea infrastructure project" OR "Korea PPP project"',
       '"Korea data center" OR "Korea power plant project"'
     )
     filterEn = @('(?i)korea|korean|seoul', '(?i)infrastructure|ppp|power plant|data cent|rail|highway|port|grid') }
)

# 쿼리에 쓴 낱말은 그 자산군 기사 대부분에 나타나므로 중복 판정에서 제외한다.
# (예: 국내 인프라에서 "민자사업"은 변별력이 없다)
function Get-QueryStopwords([string[]]$queries) {
  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($q in $queries) {
    foreach ($w in (($q -replace '"', ' ') -split '\s+')) {
      $w = $w.Trim().ToLower()
      if ($w.Length -ge 2 -and $w -ne "or") { $set.Add($w) | Out-Null }
    }
  }
  return ,(@($set))
}

# 한국어는 조사가 붙어 "네이버"와 "네이버는"이 다른 낱말이 된다.
# 그래서 낱말이 아니라 두 글자 단위로 쪼개 비교한다.
function Get-TitleShingles([string]$title, $stopwords) {
  $t = $title.ToLower()
  foreach ($w in $stopwords) { $t = $t.Replace($w, " ") }   # 쿼리 낱말은 변별력이 없다
  $t = $t -replace '[^0-9a-z가-힣]', ''

  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  for ($i = 0; $i -lt $t.Length - 1; $i++) { $set.Add($t.Substring($i, 2)) | Out-Null }
  # 쉼표를 붙이지 않으면 PowerShell 이 반환하면서 원소를 낱개로 펼쳐 버린다
  return ,(@($set))
}

# 같은 사건을 다룬 기사가 제목만 바꿔 여러 건 올라온다.
# 실측해 보면 같은 사건 기사끼리 겹침이 0.4 안팎이라 0.35 를 경계로 잡았다.
# (예: "잠실 스포츠·MICE 민자사업 민투심 통과" 와 "잠실 MICE·청주 … 민투심 통과")
function Test-NearDuplicate($shingles, $accepted) {
  if ($shingles.Count -lt 4) { return $false }
  foreach ($prev in $accepted) {
    if ($prev.Count -lt 4) { continue }
    $inter = 0
    foreach ($s in $shingles) { if ($prev -contains $s) { $inter++ } }
    $minLen = [Math]::Min($shingles.Count, $prev.Count)
    if ((([double]$inter) / $minLen) -ge 0.35) { return $true }
  }
  return $false
}

# 구글 뉴스는 결과가 모자라면 관련성이 뚝 떨어지는 기사까지 채워 넣는다.
# 자산군마다 반드시 걸려야 할 낱말을 정해 걸러낸다 (여러 개면 모두 만족해야 함).
function Test-TitleFilter([string]$title, $patterns) {
  if (-not $patterns) { return $true }
  foreach ($p in $patterns) { if ($title -notmatch $p) { return $false } }
  return $true
}

# 영어판 구글 뉴스에는 국내 언론의 영문판도 함께 잡힌다. 이걸 번역해 실으면
# 한국어 → 영어 → 한국어 왕복이 되어 고유명사가 망가진다
# (실제로 "한국부동산원"이 "한국부동산위원회"로 돌아왔다). 아예 제외한다.
$KOREAN_OUTLETS = '(?i)chosun|korea herald|joongang|korea times|yonhap|pulse|maeil|헤럴드|매일경제|조선|중앙|한국경제|서울경제|아시아경제|asiae|seoul economic|ked global|businesskorea|business korea|aju|thelec|koreabizwire|kbs|mbc|sbs|newsis|hankyung|donga|edaily|mk\.co\.kr|korea economic'

# 기계번역이 업계 용어를 어색하게 옮기는 것만 몇 개 바로잡는다.
$GLOSSARY = @(
  @{ from = '민간 신용';   to = '사모대출'     },
  @{ from = '사모 신용';   to = '사모대출'     },
  @{ from = '개인 신용';   to = '사모대출'     },
  @{ from = '직접 대출';   to = '직접대출'     },
  @{ from = '사모 펀드';   to = '사모펀드'     },
  @{ from = '벤처 캐피탈'; to = '벤처캐피탈'   },
  @{ from = '벤처 캐피털'; to = '벤처캐피탈'   },
  @{ from = '비공개 거래'; to = '상장폐지 인수' },
  @{ from = '데이터 센터'; to = '데이터센터'   }
)

# 해외 매체 기사는 원문이 영어라 제목을 한국어로 옮겨 싣는다.
# 구글 번역의 비공식 엔드포인트라 키는 필요 없지만 언제든 막히거나 조여질 수 있다.
# 번역에 실패한 기사는 아예 싣지 않는다 (영어 제목이 섞이면 안 되므로).
function Convert-ToKorean([string]$text) {
  $u = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=ko&dt=t&q=" +
       [uri]::EscapeDataString($text)
  try {
    $json = (Get-Web $u $null 1) | ConvertFrom-Json
  } catch {
    return $null
  }
  $out = ""
  foreach ($seg in $json[0]) { $out += [string]$seg[0] }
  $out = $out.Trim()
  if (-not $out) { return $null }
  # 번역이 사실상 그대로면 (한글이 거의 없으면) 실패로 본다
  if (($out -replace '[^가-힣]', '').Length -lt 2) { return $null }
  foreach ($g in $GLOSSARY) { $out = $out.Replace($g.from, $g.to) }
  return $out
}

# 해외 매체 기사 수집. 영어판 구글 뉴스에서 받아 거른 뒤 번역한다.
function Get-IntlNews([hashtable]$cls, [datetime]$cutoff, $seenTitles, $accepted) {
  if (-not $cls.queriesEn) { return @() }

  $stopwords = Get-QueryStopwords $cls.queriesEn
  $buckets   = @()
  $seenEn    = @()

  foreach ($query in $cls.queriesEn) {
    $bucket = @()
    $q   = [uri]::EscapeDataString($query + " when:" + $NewsDays + "d")
    $url = "https://news.google.com/rss/search?q=$q&hl=en-US&gl=US&ceid=US:en"

    $doc = New-Object System.Xml.XmlDocument
    try { $doc.LoadXml((Get-Web $url)) } catch { continue }

    foreach ($node in $doc.SelectNodes("/rss/channel/item")) {
      $title = $node.SelectSingleNode("title").InnerText

      $source = ""
      $srcNode = $node.SelectSingleNode("source")
      if ($srcNode) { $source = $srcNode.InnerText }
      if ($source -and $title.EndsWith(" - " + $source)) {
        $title = $title.Substring(0, $title.Length - ($source.Length + 3))
      }
      $title = $title.Trim()
      if (-not $title) { continue }
      if ($source -match $KOREAN_OUTLETS) { continue }   # 국내 언론 영문판은 제외
      if (-not (Test-TitleFilter $title $cls.filterEn)) { continue }

      # 번역은 비싸므로 영어 제목 단계에서 먼저 중복을 걸러낸다
      $enShingles = Get-TitleShingles $title $stopwords
      if (Test-NearDuplicate $enShingles $seenEn) { continue }
      $seenEn += ,$enShingles

      $published = ""
      $pubNode = $node.SelectSingleNode("pubDate")
      if ($pubNode) {
        try {
          $dt = [datetimeoffset]::Parse($pubNode.InnerText, [Globalization.CultureInfo]::InvariantCulture)
          if ($dt.UtcDateTime -lt $cutoff) { continue }
          $published = $dt.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } catch { }
      }

      $bucket += [ordered]@{
        title_en  = $title
        url       = $node.SelectSingleNode("link").InnerText
        source    = $source
        published = $published
      }
      if ($bucket.Count -ge $NewsIntlPerClass) { break }
    }
    $buckets += ,$bucket
    Start-Sleep -Milliseconds 200
  }

  # 국내 기사와 같은 방식으로 쿼리를 번갈아 뽑는다
  $picked = @()
  $depth  = 0
  while ($picked.Count -lt ($NewsIntlPerClass * 2)) {
    $tookAny = $false
    foreach ($b in $buckets) {
      if ($depth -lt $b.Count) {
        $picked += $b[$depth]
        $tookAny = $true
      }
    }
    if (-not $tookAny) { break }
    $depth++
  }

  $items    = @()
  $failed   = 0
  foreach ($cand in $picked) {
    if ($items.Count -ge $NewsIntlPerClass) { break }

    $ko = Convert-ToKorean $cand.title_en
    if (-not $ko) { $failed++; continue }
    if (-not $seenTitles.Add($ko)) { continue }

    # 번역된 제목이 국내 기사와 같은 사건이면 뺀다
    $shingles = Get-TitleShingles $ko @()
    if (Test-NearDuplicate $shingles $accepted) { continue }

    $items += [ordered]@{
      title     = $ko
      title_en  = $cand.title_en
      url       = $cand.url
      source    = $cand.source
      published = $cand.published
      intl      = $true
    }
    Start-Sleep -Milliseconds 120
  }

  if ($failed -gt 0) { Fail ($cls.label + " 해외 기사 " + $failed + "건 번역 실패") }
  return $items
}

function Get-News([hashtable]$cls, [datetime]$cutoff, $seenTitles) {
  $stopwords = Get-QueryStopwords $cls.queries
  $accepted  = @()
  $buckets   = @()   # 쿼리마다 후보를 따로 담는다

  foreach ($query in $cls.queries) {
    $bucket = @()

    $q   = [uri]::EscapeDataString($query + " when:" + $NewsDays + "d")
    $url = "https://news.google.com/rss/search?q=$q&hl=ko&gl=KR&ceid=KR:ko"

    $doc = New-Object System.Xml.XmlDocument
    try {
      $doc.LoadXml((Get-Web $url))
    } catch {
      Fail ($cls.label + " 뉴스 쿼리 하나를 받지 못했습니다: " + $query)
      continue
    }

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
      if (-not (Test-TitleFilter $title $cls.filter)) { continue }
      if (-not $seenTitles.Add($title)) { continue }   # 다른 자산군에 이미 실린 기사

      $shingles = Get-TitleShingles $title $stopwords
      if (Test-NearDuplicate $shingles $accepted) { continue }

      $published = ""
      $pubNode = $node.SelectSingleNode("pubDate")
      if ($pubNode) {
        try {
          $dt = [datetimeoffset]::Parse($pubNode.InnerText, [Globalization.CultureInfo]::InvariantCulture)
          if ($dt.UtcDateTime -lt $cutoff) { continue }
          $published = $dt.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
        } catch { }
      }

      $bucket   += [ordered]@{
        title     = $title
        url       = $link
        source    = $source
        published = $published
      }
      $accepted += ,$shingles
      if ($bucket.Count -ge $NewsPerClass) { break }
    }

    $buckets += ,$bucket
    Start-Sleep -Milliseconds 200
  }

  # 쿼리 하나가 목록을 독식하지 않도록 버킷을 돌아가며 한 건씩 뽑는다.
  # (예: "글로벌 리츠" 하나가 특정 종목 기사로 12칸을 다 채우는 상황을 막는다)
  $items = @()
  $depth = 0
  while ($items.Count -lt $NewsPerClass) {
    $tookAny = $false
    foreach ($b in $buckets) {
      if ($depth -lt $b.Count) {
        $items += $b[$depth]
        $tookAny = $true
        if ($items.Count -ge $NewsPerClass) { break }
      }
    }
    if (-not $tookAny) { break }
    $depth++
  }

  return @($items | Sort-Object -Property @{ Expression = { $_.published }; Descending = $true })
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
# 5. 지표 이력 (history.csv)
# ────────────────────────────────────────────────────────────────
#
#   date,code,value
#   2026-08-05,USDKRW,1421.88
#
# data.js 는 "지금 값"만 담는 스냅샷이라 어제와 비교할 수 없다. 이력은 따로 쌓는다.
#
# SQLite 를 쓰지 않는 이유: 이 저장소는 GitHub Actions(러너는 매번 새로 뜬다)에서
# 돌기 때문에 이력이 남으려면 파일이 리포에 커밋돼야 한다. 바이너리 DB 를 평일
# 24회 커밋하면 리포가 급격히 부풀고 diff 도 볼 수 없다. CSV 는 추가된 줄만
# diff 에 잡힌다.
#
# 하루에 24번 도는데 한 날짜에 한 행만 남긴다 — 그날 마지막 값이 이긴다.
# 그래서 append 가 아니라 (date,code) 로 덮어쓰고 정렬해 다시 쓴다. 정렬 순서가
# 고정이라 실제로 바뀐 줄만 diff 에 나온다.

# "2026.07.30 16:56" / "2026.07.30" → "2026-07-30".
# asof 가 비면(금리처럼 고시 시각이 없는 항목) 수집 시점(KST)으로 둔다.
#
# 수집일이 아니라 asof 를 쓰는 게 중요하다. 미국 지수는 KST 오전에 받아도
# 전날 종가라, 수집일로 찍으면 하루씩 밀린다.
function Get-MetricDate([string]$asof) {
  $m = [regex]::Match($asof, '^(\d{4})\.(\d{2})\.(\d{2})')
  if ($m.Success) {
    return ($m.Groups[1].Value + "-" + $m.Groups[2].Value + "-" + $m.Groups[3].Value)
  }
  return (Get-Date).ToUniversalTime().AddHours(9).ToString("yyyy-MM-dd")
}

# "date|code" → value 사전으로 읽는다. 파일이 없으면 빈 사전.
function Read-History {
  $map = New-Object 'System.Collections.Generic.Dictionary[string,string]'
  if (-not (Test-Path $HistFile)) { return $map }
  try {
    $lines = [IO.File]::ReadAllLines($HistFile, [Text.Encoding]::UTF8)
  } catch {
    Fail "history.csv 를 읽지 못했습니다 (이번 회차 이력은 건너뜁니다)"
    return $null
  }
  foreach ($line in $lines) {
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
  # 날짜(고정 10자리) → 코드 순. 순서가 고정이라야 diff 가 바뀐 줄만 잡는다.
  [Array]::Sort($keys, [StringComparer]::Ordinal)

  $sb = New-Object Text.StringBuilder
  [void]$sb.AppendLine("date,code,value")
  foreach ($k in $keys) {
    $i = $k.IndexOf("|")
    [void]$sb.AppendLine($k.Substring(0, $i) + "," + $k.Substring($i + 1) + "," + $map[$k])
  }
  [IO.File]::WriteAllText($HistFile, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
}

function Update-History($market) {
  $map = Read-History
  if ($null -eq $map) { return }

  $added = 0; $changed = 0; $unknown = @()
  foreach ($g in $market) {
    # 수집이 실패해 이전 값을 재사용 중인 그룹은 건너뛴다.
    # 그대로 쌓으면 어제 값이 오늘 값으로 둔갑한다.
    if ($g.stale) { continue }

    foreach ($m in @($g.items)) {
      $code = Resolve-MetricCode $m
      if (-not $code) { $unknown += $m.name; continue }

      $v = To-Number $m.value
      if ($null -eq $v) { continue }

      $key = (Get-MetricDate $m.asof) + "|" + $code
      $txt = $v.ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)

      if ($map.ContainsKey($key)) {
        if ($map[$key] -ne $txt) { $map[$key] = $txt; $changed++ }
      } else {
        $map[$key] = $txt; $added++
      }
    }
  }

  try {
    Save-History $map
  } catch {
    Fail "history.csv 를 쓰지 못했습니다"
    return
  }

  Log ("· 지표 이력 · 신규 " + $added + " / 갱신 " + $changed + " / 누적 " + $map.Count + "행")
  if ($unknown.Count -gt 0) {
    # 새 지표가 늘었거나 네이버 표기가 바뀐 것이다. 수집 자체는 정상이라 경고만.
    Log ("  (코드 미지정 " + $unknown.Count + "건: " + (($unknown | Select-Object -Unique) -join ", ") + ")")
  }
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

# 네이버 시장지표 표에 없는 만기를 채권 API 로 채운다.
# 표의 국고채(3년)는 빼고 3·5·10년을 같은 기준(체결)으로 통일한다 — 기준이 다른
# 3년물이 둘 나란히 있으면 값이 어긋나 보인다.
if ($mi -and $mi.rates.Count) {
  $bonds = @()
  try { $bonds = @(Get-BondYields) } catch { Fail "국고채 수익률 수집 실패" }

  if ($bonds.Count) {
    $mi.rates = @(@($mi.rates | Where-Object { $_.name -ne "국고채 3년" }) + $bonds)
  }
  $mi.rates = Sort-Rates $mi.rates
}

$miGroups = @(
  @{ label="환율";           note="하나은행 고시 기준";       key="fx"    },
  @{ label="국내시장금리";   note="최종 고시치 · 등락은 %p";  key="rates" },
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

# 지표 이력. 뉴스보다 먼저 쌓는다 — 뉴스 수집이 길고 실패도 잦은데,
# 거기서 죽더라도 이번 회차 지표는 이력에 남아야 한다.
try { Update-History $market } catch { Fail "지표 이력 기록 실패" }

# 뉴스
Log "· 자산군별 뉴스 …"
$cutoff = (Get-Date).ToUniversalTime().AddDays(-$NewsDays)
$seenTitles = New-Object 'System.Collections.Generic.HashSet[string]'
$news = @()
foreach ($cls in $NEWS_CLASSES) {
  $items = @()
  $stale = $false
  try {
    $items = @(Get-News $cls $cutoff $seenTitles)
  } catch {
    Fail ($cls.label + " 뉴스 수집 실패")
    $p = Get-PrevNews $prev $cls.key
    if ($p) { $items = @($p.items); $stale = $true }
  }

  # 해외 매체 기사를 덧붙인다. 국내 수집이 실패해 이전 값을 쓰는 중이면 건드리지 않는다.
  $intlCount = 0
  if (-not $stale) {
    $accepted = @()
    foreach ($it in $items) { $accepted += ,(Get-TitleShingles $it.title @()) }
    try {
      $intl = @(Get-IntlNews $cls $cutoff $seenTitles $accepted)
      $intlCount = $intl.Count
      if ($intlCount -gt 0) {
        $items = @(@($items) + @($intl) |
                   Sort-Object -Property @{ Expression = { $_.published }; Descending = $true })
      }
    } catch { Fail ($cls.label + " 해외 뉴스 수집 실패") }
  }

  Log ("    " + $cls.label + " " + $items.Count + "건 (해외 " + $intlCount + "건)")
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
