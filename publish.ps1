# 로컬에서 갱신한 data.js / history.csv 를 GitHub 에 올려 온라인 대시보드까지 반영한다.
# 갱신.bat 이 refresh.ps1 다음에 자동으로 호출한다. 단독 실행해도 된다.
$ErrorActionPreference = 'Continue'
Set-Location $PSScriptRoot

git add data.js history.csv
$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "변경된 데이터가 없어 온라인 반영을 건너뜁니다."
    exit 0
}

$stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm')
git commit -m "data: refreshed at $stamp KST (local)" | Out-Null

git push
if ($LASTEXITCODE -ne 0) {
    # 자동 배치(bot)가 그 사이 커밋을 올린 경우: 원격을 받아 얹은 뒤 다시 푸시.
    # 충돌 구간은 방금 수집한 로컬 값(-X theirs)을 우선한다.
    Write-Host "원격에 새 커밋이 있어 병합 후 다시 푸시합니다..."
    git pull --rebase --autostash -X theirs
    if ($LASTEXITCODE -ne 0) {
        git rebase --abort 2>$null
        Write-Host "자동 병합에 실패했습니다. Claude Code 에서 'git 충돌 해결하고 푸시해줘' 라고 요청하세요."
        exit 1
    }
    git push
    if ($LASTEXITCODE -ne 0) {
        Write-Host "푸시에 실패했습니다. 네트워크 상태를 확인하거나 Claude Code 에 요청하세요."
        exit 1
    }
}

Write-Host ""
Write-Host "온라인 반영 완료. 1~2분 뒤 사이트에서 확인할 수 있습니다:"
Write-Host "https://gpflatls-oss.github.io/invest-dashboard/"
