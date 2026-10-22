# Собирает «Установить-MCP.bat» для сотрудников: батник скачивает установщик с GitHub
# и передаёт ему ключ Google-клиента. Ключ (client_secret) в GitHub не кладём: GitHub
# сообщает о таких ключах в Google, и Google отключает клиента. Батник раздаётся внутри компании.
param(
    [Parameter(Mandatory = $true)][string]$Repo,
    [Parameter(Mandatory = $true)][string]$GoogleClientPath,
    [string]$Branch = 'main',
    [string]$Output = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Установить-MCP.bat')
)
$ErrorActionPreference = 'Stop'
if ($Repo -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') { throw 'Repo: владелец/репозиторий, например ivanov/zakupki-mcp' }
if ($Branch -notmatch '^[A-Za-z0-9_./-]+$') { throw 'Некорректная ветка' }
$client = Get-Content -LiteralPath $GoogleClientPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $client.installed.client_id) { throw 'Нужен OAuth client JSON типа Desktop app (ключ "installed")' }
$encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($GoogleClientPath))
$raw = "https://raw.githubusercontent.com/$Repo/$Branch/installer/local-install.ps1"
$command = "`$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; " +
    "`$r=Invoke-WebRequest -UseBasicParsing -Uri '$raw'; `$s=[Text.Encoding]::UTF8.GetString(`$r.RawContentStream.ToArray()); " +
    "& ([scriptblock]::Create(`$s)) -Repo '$Repo' -Branch '$Branch' -GoogleClient '$encoded'"
$bat = @(
    '@echo off', 'chcp 65001 >nul', 'title Zakupki MCP setup',
    "powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command `"$command`"",
    'set "RC=%ERRORLEVEL%"', 'echo.', 'pause', 'exit /b %RC%', ''
) -join "`r`n"
[IO.File]::WriteAllText($Output, $bat, [Text.Encoding]::ASCII)
Write-Host "Готово: $Output" -ForegroundColor Green
Write-Host 'Раздайте файл сотрудникам внутри компании (не выкладывайте в интернет): в нём ключ Google-клиента.'
