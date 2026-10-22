# Установщик MCP «Закупки» на ПК сотрудника: OpenCode + MCP с GitHub + Google + LetAI.
# Запускается файлом «Установить-MCP.bat» (его собирает администратор: tools\Make-Installer.ps1)
# или вручную: & ([scriptblock]::Create((irm <raw-ссылка на этот файл>))) -Repo owner/repo -GoogleClient <base64>
# Повторный запуск обновляет MCP и OpenCode до последних версий; данные и вход в Google сохраняются.
param(
    [Parameter(Mandatory = $true)][string]$Repo,
    [string]$Branch = 'main',
    [string]$GoogleClient = '',
    [Parameter(Mandatory = $true)][string]$SpreadsheetId,
    [string]$SheetName = 'незакр 2025/26',
    [string]$LetAIBaseUrl = 'http://10.19.27.103:4000/v1',
    [string]$LetAIModel = 'LetAI-ONE'
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$root = Join-Path $env:LOCALAPPDATA 'zakupki-mcp'
$state = Join-Path $env:USERPROFILE '.config\zakupki-mcp'
$opencodeDir = Join-Path $env:LOCALAPPDATA 'Programs\opencode'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$serverName = 'zakupki'

function Write-Step([string]$Text) { Write-Host "==> $Text" -ForegroundColor Cyan }
function Write-Secret([string]$Path, [string]$Value) { [IO.File]::WriteAllText($Path, $Value, $utf8) }

function Get-OpenCodeAsset {
    if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { return 'opencode-windows-arm64.zip' }
    Add-Type -Namespace ZakupkiMcp -Name Cpu -MemberDefinition '[DllImport("kernel32.dll")] public static extern bool IsProcessorFeaturePresent(uint feature);'
    # 40 = PF_AVX2_INSTRUCTIONS_AVAILABLE; без AVX2 нужна baseline-сборка.
    if ([ZakupkiMcp.Cpu]::IsProcessorFeaturePresent(40)) { return 'opencode-windows-x64.zip' }
    return 'opencode-windows-x64-baseline.zip'
}

function Get-LatestOpenCode([string]$Asset) {
    # Официальный адрес последней версии; номер берём из перенаправления без скачивания.
    $url = "https://opencode.ai/update/api/latest/cli/opencode/$Asset"
    $request = [Net.HttpWebRequest]::Create($url)
    $request.AllowAutoRedirect = $false
    $response = $request.GetResponse()
    try { $location = $response.Headers['Location'] } finally { $response.Close() }
    if (-not $location) { return @{ Url = $url; Version = $null } }
    $target = New-Object Uri ((New-Object Uri $url), $location)
    $version = if ($target.AbsolutePath -match '/(\d+\.\d+\.\d+)/') { $Matches[1] } else { $null }
    return @{ Url = $target.AbsoluteUri; Version = $version }
}

function Get-OpenCodeVersion([string]$Path) {
    $text = "$(& $Path --version 2>$null | Select-Object -First 1)"
    if ($text -match '(\d+\.\d+\.\d+)') { return $Matches[1] }
    return $null
}

function Find-OpenCode {
    $command = Get-Command opencode -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }
    foreach ($candidate in @((Join-Path $opencodeDir 'opencode.exe'), (Join-Path $env:USERPROFILE '.opencode\bin\opencode.exe'))) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

function Expand-Download([string]$Url, [string]$Name) {
    # Скачать zip во временную папку и распаковать; вернуть путь распаковки.
    $temp = Join-Path ([IO.Path]::GetTempPath()) ('zakupki-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $temp | Out-Null
    $zip = Join-Path $temp $Name
    Write-Host "    Скачиваю $Url"
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $zip
    Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $temp 'x') -Force
    return (Join-Path $temp 'x')
}

function Add-UserPath([string]$Dir) {
    $parts = @([Environment]::GetEnvironmentVariable('Path', 'User') -split ';' | Where-Object { $_ })
    if ($parts -notcontains $Dir) { [Environment]::SetEnvironmentVariable('Path', (($parts + $Dir) -join ';'), 'User') }
    if (($env:Path -split ';') -notcontains $Dir) { $env:Path = "$env:Path;$Dir" }
}

function Install-OpenCode($Latest, [string]$Asset) {
    $unpacked = Expand-Download $Latest.Url $Asset
    $exe = Get-ChildItem -LiteralPath $unpacked -Recurse -Filter 'opencode.exe' | Select-Object -First 1
    if (-not $exe) { throw "В архиве $Asset нет opencode.exe" }
    New-Item -ItemType Directory -Force -Path $opencodeDir | Out-Null
    Copy-Item -LiteralPath $exe.FullName -Destination (Join-Path $opencodeDir 'opencode.exe') -Force
    Add-UserPath $opencodeDir
    return (Join-Path $opencodeDir 'opencode.exe')
}

function Remove-JsonComments([string]$Text) {
    # JSONC -> JSON: комментарии // и /* */ и висячие запятые; строки не трогаем.
    $Text = $Text.TrimStart([char]0xFEFF)
    $out = New-Object System.Text.StringBuilder
    $inString = $false; $escaped = $false; $i = 0
    while ($i -lt $Text.Length) {
        $c = $Text[$i]
        $next = if ($i + 1 -lt $Text.Length) { $Text[$i + 1] } else { [char]0 }
        if ($inString) {
            [void]$out.Append($c)
            if ($escaped) { $escaped = $false } elseif ($c -eq [char]'\') { $escaped = $true } elseif ($c -eq [char]'"') { $inString = $false }
            $i++
            continue
        }
        if ($c -eq [char]'"') { $inString = $true; [void]$out.Append($c); $i++; continue }
        if ($c -eq [char]'/' -and $next -eq [char]'/') {
            while ($i -lt $Text.Length -and $Text[$i] -ne [char]"`n") { $i++ }
            continue
        }
        if ($c -eq [char]'/' -and $next -eq [char]'*') {
            $end = $Text.IndexOf('*/', $i + 2)
            if ($end -lt 0) { throw 'Незакрытый комментарий /* в конфиге OpenCode' }
            [void]$out.Append(' ')
            $i = $end + 2
            continue
        }
        [void]$out.Append($c)
        $i++
    }
    $clean = $out.ToString()
    $result = New-Object System.Text.StringBuilder
    $inString = $false; $escaped = $false
    for ($i = 0; $i -lt $clean.Length; $i++) {
        $c = $clean[$i]
        if ($inString) {
            [void]$result.Append($c)
            if ($escaped) { $escaped = $false } elseif ($c -eq [char]'\') { $escaped = $true } elseif ($c -eq [char]'"') { $inString = $false }
            continue
        }
        if ($c -eq [char]'"') { $inString = $true }
        if ($c -eq [char]',') {
            $j = $i + 1
            while ($j -lt $clean.Length -and [char]::IsWhiteSpace($clean[$j])) { $j++ }
            if ($j -lt $clean.Length -and ($clean[$j] -eq [char]'}' -or $clean[$j] -eq [char]']')) { continue }
        }
        [void]$result.Append($c)
    }
    return $result.ToString()
}

function Get-Child($Parent, [string]$Name) {
    $property = $Parent.PSObject.Properties[$Name]
    if (-not $property -or -not ($property.Value -is [System.Management.Automation.PSCustomObject])) {
        $Parent | Add-Member -NotePropertyName $Name -NotePropertyValue ([pscustomobject]@{}) -Force
    }
    return $Parent.PSObject.Properties[$Name].Value
}

function Remove-Child($Parent, [string]$Name) {
    if ($Parent -and $Parent.PSObject.Properties[$Name]) { $Parent.PSObject.Properties.Remove($Name) }
}

function Update-OpenCodeConfig([string]$Schema, $Server, $Provider, [string]$Model) {
    # Вписывает MCP и провайдера в ~/.config/opencode/opencode.json(c) в формате версии:
    # v1 - mcp.<имя>, provider.<имя>; v2 - mcp.servers.<имя>, providers.<имя>. Остальное сохраняется.
    $configDir = Join-Path $env:USERPROFILE '.config\opencode'
    $configFile = $null
    foreach ($name in @('opencode.jsonc', 'opencode.json')) {
        $candidate = Join-Path $configDir $name
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $configFile = $candidate; break }
    }
    $config = $null
    if ($configFile) {
        $text = Remove-JsonComments ([IO.File]::ReadAllText($configFile, [Text.Encoding]::UTF8))
        if ($text.Trim()) { $config = $text | ConvertFrom-Json }
    } else {
        New-Item -ItemType Directory -Force -Path $configDir | Out-Null
        $configFile = Join-Path $configDir 'opencode.json'
    }
    if ($null -eq $config) { $config = [pscustomobject]@{ '$schema' = 'https://opencode.ai/config.json' } }
    if (-not ($config -is [System.Management.Automation.PSCustomObject])) { throw "Корень $configFile должен быть объектом JSON" }
    $before = $config | ConvertTo-Json -Depth 64

    $mcp = Get-Child $config 'mcp'
    # Прежний портал добавлял удалённый google-sheets (сервер без интернета): убираем его.
    $legacy = Join-Path $env:USERPROFILE '.config\mcp-portal\managed.json'
    $old = @($serverName)
    if (Test-Path -LiteralPath $legacy) { foreach ($n in ([IO.File]::ReadAllText($legacy) | ConvertFrom-Json)) { $old += $n } }
    foreach ($n in $old) {
        Remove-Child $mcp $n
        if ($mcp.PSObject.Properties['servers'] -and $mcp.servers -is [System.Management.Automation.PSCustomObject]) { Remove-Child $mcp.servers $n }
    }
    $target = if ($Schema -eq 'v2') { Get-Child $mcp 'servers' } else { $mcp }
    $target | Add-Member -NotePropertyName $serverName -NotePropertyValue $Server -Force

    if ($Provider) {
        $section = if ($Schema -eq 'v2') { 'providers' } else { 'provider' }
        $providers = Get-Child $config $section
        $providers | Add-Member -NotePropertyName 'letai-one' -NotePropertyValue $Provider -Force
        if ($Schema -eq 'v2' -and $config.PSObject.Properties['provider'] -and $config.provider -is [System.Management.Automation.PSCustomObject]) {
            Remove-Child $config.provider 'letai-one'
            if (-not @($config.provider.PSObject.Properties).Count) { Remove-Child $config 'provider' }
        }
        if (-not $config.PSObject.Properties['model']) { $config | Add-Member -NotePropertyName 'model' -NotePropertyValue $Model }
    }

    $after = $config | ConvertTo-Json -Depth 64
    if ($after -ne $before) {
        if (Test-Path -LiteralPath $configFile) {
            $backups = Join-Path $configDir 'backups'
            New-Item -ItemType Directory -Force -Path $backups | Out-Null
            $copy = Join-Path $backups ('opencode-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + [IO.Path]::GetExtension($configFile))
            Copy-Item -LiteralPath $configFile -Destination $copy
            Write-Host "    Резервная копия конфига: $copy"
        }
        $temporary = "$configFile.zakupki.tmp"
        [IO.File]::WriteAllText($temporary, $after + "`n", $utf8)
        Move-Item -LiteralPath $temporary -Destination $configFile -Force
    }
    Write-Host "    MCP «$serverName» подключён в $configFile (OpenCode $Schema)"
}

try {
    Write-Host ''
    Write-Host 'Установка MCP «Закупки» для OpenCode' -ForegroundColor Green
    Write-Host ''
    New-Item -ItemType Directory -Force -Path $root, $state | Out-Null
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    & icacls.exe $state /inheritance:r /grant:r "${identity}:(OI)(CI)F" | Out-Null

    Write-Step 'OpenCode'
    $asset = Get-OpenCodeAsset
    $latest = $null
    try { $latest = Get-LatestOpenCode $asset; Write-Host "    Последняя версия: $($latest.Version)" }
    catch { Write-Host "    Не удалось узнать последнюю версию ($($_.Exception.Message))." -ForegroundColor Yellow }
    $opencode = Find-OpenCode
    $ours = Join-Path $opencodeDir 'opencode.exe'
    if (-not $opencode) {
        if (-not $latest) { throw 'OpenCode не установлен, а opencode.ai недоступен. Проверьте интернет.' }
        Write-Host '    Не найден, устанавливаю.'
        $opencode = Install-OpenCode $latest $asset
    } else {
        $installed = Get-OpenCodeVersion $opencode
        Write-Host "    Установлен: $opencode ($installed)"
        $outdated = $latest -and $latest.Version -and $installed -and ([version]$installed -lt [version]$latest.Version)
        if ($outdated -and [IO.Path]::GetFullPath($opencode) -eq [IO.Path]::GetFullPath($ours)) {
            Write-Host "    Обновляю до $($latest.Version)."
            try { $opencode = Install-OpenCode $latest $asset }
            catch { Write-Host "    Обновить не удалось ($($_.Exception.Message)): закройте OpenCode и запустите установщик снова." -ForegroundColor Yellow }
        } elseif ($outdated) {
            Write-Host "    Доступна $($latest.Version): эту копию обновите тем же способом, каким ставили." -ForegroundColor Yellow
        }
    }
    $version = Get-OpenCodeVersion $opencode
    if (-not $version) { throw "Не удалось определить версию OpenCode ($opencode --version)" }
    $schema = if (([version]$version).Major -ge 2) { 'v2' } else { 'v1' }

    Write-Step 'uv и Python'
    # Python, уже установленный на ПК (или ранее скачанный uv), без скачивания: корпоративная
    # сеть не пускает к архивам Python на GitHub. Скачивание - только если Python 3.12 нет совсем.
    $env:UV_PYTHON_PREFERENCE = 'managed'
    $uv = (Get-Command uv -ErrorAction SilentlyContinue | Select-Object -First 1).Source
    if (-not $uv) {
        $uvDir = Join-Path $root 'uv'
        $uv = Join-Path $uvDir 'uv.exe'
        if (-not (Test-Path -LiteralPath $uv)) {
            $unpacked = Expand-Download 'https://github.com/astral-sh/uv/releases/latest/download/uv-x86_64-pc-windows-msvc.zip' 'uv.zip'
            New-Item -ItemType Directory -Force -Path $uvDir | Out-Null
            Get-ChildItem -LiteralPath $unpacked -Recurse -Filter 'uv*.exe' | Copy-Item -Destination $uvDir -Force
        }
    }
    Write-Host "    uv: $uv"
    # Корпоративная сеть подменяет HTTPS-сертификаты своим: uv должен доверять хранилищу Windows.
    # В новых uv настройка называется UV_SYSTEM_CERTS, в старых - UV_NATIVE_TLS.
    if (((& $uv --help) -join "`n") -match 'system-certs') { $env:UV_SYSTEM_CERTS = '1' } else { $env:UV_NATIVE_TLS = '1' }

    Write-Step "MCP с GitHub ($Repo, ветка $Branch)"
    $unpacked = Expand-Download "https://codeload.github.com/$Repo/zip/refs/heads/$Branch" 'mcp.zip'
    $source = Get-ChildItem -LiteralPath $unpacked -Directory | Select-Object -First 1
    if (-not $source -or -not (Test-Path -LiteralPath (Join-Path $source.FullName 'server.py'))) { throw "В репозитории $Repo нет server.py" }
    $app = Join-Path $root 'app'
    if (Test-Path -LiteralPath $app) { Remove-Item -LiteralPath $app -Recurse -Force }
    Move-Item -LiteralPath $source.FullName -Destination $app
    $venv = Join-Path $root 'venv'
    if (Test-Path -LiteralPath $venv) {
        # Пересоздаём на своём Python; если MCP сейчас запущен OpenCode, файлы заняты - оставляем как есть.
        try { Remove-Item -LiteralPath $venv -Recurse -Force -ErrorAction Stop }
        catch { Write-Host '    Python-окружение занято (OpenCode открыт?) - обновляю на месте.' -ForegroundColor Yellow }
    }
    Write-Host '    Готовлю Python 3.12...'
    & $uv venv --python 3.12 --allow-existing $venv | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw 'Не удалось подготовить Python 3.12: на компьютере его нет, а скачать не дала сеть. Установите Python 3.12 с python.org (или через ИТ) и запустите установщик снова.'
    }
    $python = Join-Path $venv 'Scripts\python.exe'
    Write-Host '    Устанавливаю библиотеки MCP...'
    & $uv pip install --python $python -r (Join-Path $app 'requirements.lock') | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw 'Не удалось установить зависимости MCP. Если выше «UnknownIssuer» или «certificate» - сеть не пускает к pypi.org; если «Connect» - нет доступа в интернет.'
    }

    Write-Step 'Google'
    $credentials = Join-Path $state 'credentials.json'
    $token = Join-Path $state 'token.json'
    $nextToInstaller = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'google-client.json' } else { '' }
    if ($GoogleClient -and (Test-Path -LiteralPath $GoogleClient -PathType Leaf)) {
        Copy-Item -LiteralPath $GoogleClient -Destination $credentials -Force
    } elseif ($GoogleClient) {
        Write-Secret $credentials ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($GoogleClient)))
    } elseif ($nextToInstaller -and (Test-Path -LiteralPath $nextToInstaller)) {
        Copy-Item -LiteralPath $nextToInstaller -Destination $credentials -Force
    }
    if (-not (Test-Path -LiteralPath $credentials)) { throw 'Нет ключа Google-клиента: получите у администратора файл «Установить-MCP.bat».' }
    $env:GOOGLE_AUTH_MODE = 'local'; $env:TOKEN_PATH = $token; $env:LOCAL_STATE_DIR = $state
    $env:REQUESTS_SPREADSHEET_ID = $SpreadsheetId; $env:WORK_SHEET_NAME = $SheetName
    $env:PYTHONUTF8 = '1'; $env:PYTHONIOENCODING = 'utf-8'
    if (-not (Test-Path -LiteralPath $token)) {
        Write-Host '    Откроется браузер: войдите в свой Google-аккаунт (с доступом к таблице заявок) и разрешите доступ.'
        Write-Host '    Если Google напишет «приложение не проверено» - «Дополнительно» -> «Перейти».'
        & $python (Join-Path $app 'authorize_google.py') --local --credentials $credentials --token $token
        if ($LASTEXITCODE -ne 0) { throw 'Вход в Google не завершён.' }
    } else {
        Write-Host '    Вход в Google уже выполнен.'
    }

    Write-Step 'Ключ LetAI (корпоративная модель)'
    $keyPath = Join-Path $state 'letai.key'
    $prompt = if (Test-Path -LiteralPath $keyPath) { 'Вставьте API-ключ LetAI (ввод скрыт; Enter - оставить текущий)' } else { 'Вставьте API-ключ LetAI (ввод скрыт; Enter - пропустить)' }
    $secure = Read-Host $prompt -AsSecureString
    if ($secure -and $secure.Length -gt 0) {
        $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
        Write-Secret $keyPath $plain.Trim()
        $plain = $null
        Write-Host '    Ключ сохранён.'
    }
    $provider = $null
    if (Test-Path -LiteralPath $keyPath) {
        $apiKey = '{file:~/.config/zakupki-mcp/letai.key}'
        $limit = [pscustomobject]@{ context = 262144; output = 65536 }
        $models = [pscustomobject]@{}
        if ($schema -eq 'v2') {
            $models | Add-Member -NotePropertyName $LetAIModel -NotePropertyValue ([pscustomobject]@{ name = $LetAIModel; limit = $limit })
            $provider = [pscustomobject]@{ name = 'LetAI-ONE'; package = '@opencode/ai/providers/openai-compatible'
                settings = [pscustomobject]@{ baseURL = $LetAIBaseUrl; apiKey = $apiKey }; models = $models }
        } else {
            $models | Add-Member -NotePropertyName $LetAIModel -NotePropertyValue ([pscustomobject]@{ name = $LetAIModel; limit = $limit
                modalities = [pscustomobject]@{ input = @('text', 'image'); output = @('text') } })
            $provider = [pscustomobject]@{ npm = '@ai-sdk/openai-compatible'; name = 'LetAI-ONE'
                options = [pscustomobject]@{ baseURL = $LetAIBaseUrl; apiKey = $apiKey; timeout = 600000 }; models = $models }
        }
    } else {
        Write-Host '    Без ключа: модель настроите в OpenCode сами.'
    }

    Write-Step 'Подключаю MCP к OpenCode'
    $environment = [pscustomobject]@{
        GOOGLE_AUTH_MODE = 'local'; TOKEN_PATH = $token; LOCAL_STATE_DIR = $state
        REQUESTS_SPREADSHEET_ID = $SpreadsheetId; WORK_SHEET_NAME = $SheetName; MCP_DRIVE_FOLDER = 'Отчёты MCP'
        PYTHONUTF8 = '1'; PYTHONIOENCODING = 'utf-8'
    }
    $command = @($python, (Join-Path $app 'server.py'), '--stdio')
    if ($schema -eq 'v2') {
        $server = [pscustomobject]@{ type = 'local'; command = $command; environment = $environment
            timeout = [pscustomobject]@{ startup = 60000; catalog = 60000; execution = 600000 } }
    } else {
        $server = [pscustomobject]@{ type = 'local'; command = $command; environment = $environment; enabled = $true; timeout = 60000 }
    }
    Update-OpenCodeConfig $schema $server $provider "letai-one/$LetAIModel"
    # Прежняя схема (портал на сервере) больше не нужна: убираем её автообновление.
    try { Unregister-ScheduledTask -TaskName 'MCP Portal sync' -Confirm:$false -ErrorAction Stop } catch { }

    Write-Step 'Проверка'
    & $python (Join-Path $app 'server.py') --check
    $checked = $LASTEXITCODE -eq 0

    Write-Host ''
    if ($checked) { Write-Host 'Готово.' -ForegroundColor Green } else { Write-Host 'Установлено, но проверка нашла проблему (см. выше).' -ForegroundColor Yellow }
    Write-Host 'Полностью закройте OpenCode и откройте снова. Затем напишите: «кто я», «мои позиции в работе».'
    Write-Host 'Обновить MCP позже - запустите этот же установщик ещё раз.'
    exit 0
} catch {
    Write-Host ''
    Write-Host ('ОШИБКА: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}
