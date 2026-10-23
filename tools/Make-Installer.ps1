# Собирает «Установить-MCP.bat» для сотрудников: батник-полиглот, содержащий полный
# текст installer/local-install.ps1 в виде обычного текста после метки ::ZAKUPKI-PS1::.
# При запуске батник извлекает скрипт во %TEMP% и выполняет его как файл.
# Так сделано из-за ложных срабатываний антивируса (Kaspersky Endpoint Security):
# прежняя схема «батник качает ps1 и исполняет из памяти через scriptblock::Create»
# — это типовой шаблон download-cradle, который эвристика режет. Здесь весь код
# лежит в файле открытым текстом и доступен антивирусу для проверки.
# Ключ (client_secret) в GitHub не кладём: GitHub сообщает о таких ключах в Google,
# и Google отключает клиента. Батник раздаётся внутри компании.
param(
    [Parameter(Mandatory = $true)][string]$Repo,
    [Parameter(Mandatory = $true)][string]$GoogleClientPath,
    [Parameter(Mandatory = $true)][string]$SpreadsheetId,
    [string]$SheetName = 'незакр 2025/26',
    [string]$Branch = 'main',
    [string]$Output = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Установить-MCP.bat')
)
$ErrorActionPreference = 'Stop'
if ($Repo -notmatch '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') { throw 'Repo: владелец/репозиторий, например ivanov/zakupki-mcp' }
if ($Branch -notmatch '^[A-Za-z0-9_./-]+$') { throw 'Некорректная ветка' }
if ($SpreadsheetId -notmatch '^[A-Za-z0-9_-]{20,}$') { throw 'SpreadsheetId: ID таблицы заявок из её ссылки (между /d/ и /edit)' }
$marker = '::ZAKUPKI-PS1::'
$installer = Join-Path $PSScriptRoot '..\installer\local-install.ps1'
$body = [IO.File]::ReadAllText($installer, [Text.Encoding]::UTF8).TrimEnd("`r", "`n") -split "`r?`n"
if ($body -contains $marker) { throw "В local-install.ps1 встречается служебная метка $marker — переименуйте её в Make-Installer.ps1" }
$client = Get-Content -LiteralPath $GoogleClientPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $client.installed.client_id) { throw 'Нужен OAuth client JSON типа Desktop app (ключ "installed")' }
$encoded = [Convert]::ToBase64String([IO.File]::ReadAllBytes($GoogleClientPath))
$sheet = $SheetName

# Командная часть — после chcp 65001 cmd читает .bat как UTF-8, поэтому -SheetName
# передаём литералом в ДВОЙНЫХ кавычках: cmd не разворачивает одинарные кавычки,
# а powershell -File не вычисляет выражения в аргументах.
# Извлечённый ps1 пишем с BOM, иначе PowerShell 5.1 читает его как ANSI и ломает кириллицу.
$head = @'
@echo off
chcp 65001 >nul
title Zakupki MCP setup
set "WORKDIR=%TEMP%\zakupki-mcp-install"
if not exist "%WORKDIR%" mkdir "%WORKDIR%"
powershell.exe -NoLogo -NoProfile -Command "$c = [IO.File]::ReadAllLines('%~f0', [Text.Encoding]::UTF8); $i = 0; while ($i -lt $c.Count -and $c[$i] -ne '{{MARKER}}') { $i++ }; if ($i -ge $c.Count) { exit 1 }; [IO.File]::WriteAllLines($env:WORKDIR + '\local-install.ps1', [string[]]$c[($i+1)..($c.Count-1)], (New-Object System.Text.UTF8Encoding($true)))"
if errorlevel 1 (echo Install script extraction failed.& pause& exit /b 1)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\local-install.ps1" -Repo "{{REPO}}" -Branch "{{BRANCH}}" -GoogleClient "{{CLIENT}}" -SpreadsheetId "{{SPREADSHEET}}" -SheetName "{{SHEET}}"
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
{{MARKER}}
'@
$head = $head.Replace('{{MARKER}}', $marker).Replace('{{REPO}}', $Repo).Replace('{{BRANCH}}', $Branch).Replace('{{CLIENT}}', $encoded).Replace('{{SPREADSHEET}}', $SpreadsheetId).Replace('{{SHEET}}', $sheet)
$bat = (($head -split "`n" | ForEach-Object { $_.TrimEnd("`r") }) -join "`r`n") + "`r`n" + ($body -join "`r`n") + "`r`n"
[IO.File]::WriteAllText($Output, $bat, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Готово: $Output" -ForegroundColor Green
Write-Host 'Раздайте файл сотрудникам внутри компании (не выкладывайте в интернет): в нём ключ Google-клиента.'
Write-Host 'Если антивирус всё равно блокирует — добавьте файл в исключения через администратора'
Write-Host 'Kaspersky Security Center или отправьте его как ложное срабатывание на opentip.kaspersky.com.'
