<#
.SYNOPSIS
    Сторож очереди печати: если принтер получил от роутера другой IP, находит его
    по MAC-адресу и переводит очередь на новый IP-порт.

.DESCRIPTION
    Нужен только если на роутере НЕ закреплён постоянный IP принтера.
    Ставится задачей планировщика (от SYSTEM, при старте и каждые 15 минут, без окон).
    Имя очереди, MAC и имя хоста берёт из printer.psd1 рядом со скриптом. При установке
    они записываются в задачу, дальше файл настроек сторожу не нужен.
    На печать ничего не отправляется: проверяются веб-интерфейс принтера и ping.
    Лог: C:\ProgramData\PrinterWiFiFix\watchdog.log

.EXAMPLE
    .\Printer-Watchdog.ps1 -Install      # поставить (запросит UAC)
.EXAMPLE
    .\Printer-Watchdog.ps1 -Uninstall    # убрать
.EXAMPLE
    .\Printer-Watchdog.ps1               # разовая проверка с выводом на экран

.NOTES
    Параметры командной строки важнее значений из файла настроек.
    Код возврата: 0 - готово, 1 - ошибка, 99 - скрипт перезапущен с правами администратора в новом окне.
#>
[CmdletBinding()]
param(
    [switch]$Install,
    [switch]$Uninstall,
    # Файл настроек. По умолчанию printer.psd1 рядом со скриптом.
    [string]$Config,
    [string]$QueueName,
    [string]$PrinterMac,
    [string]$PrinterHostName,
    [int]$PortNumber,
    [int]$IntervalMinutes = 15
)

$ErrorActionPreference = 'Stop'
$Bound = $PSBoundParameters
$InstallDir = Join-Path $env:ProgramData 'PrinterWiFiFix'
$InstalledScript = Join-Path $InstallDir 'Printer-Watchdog.ps1'
$LogFile = Join-Path $InstallDir 'watchdog.log'
$TaskPrefix = 'Printer Wi-Fi Watchdog'
# Порты для проверки связи. 9100 не используем: пустое подключение принтер может принять за задание печати.
$ProbePorts = 80, 443, 631
$ExitHandedOff = 99

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Log([string]$Text) {
    $line = '{0:yyyy-MM-dd HH:mm:ss}  [{1}] {2}' -f (Get-Date), $QueueName, $Text
    Write-Host $line
    try {
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt 512KB) {
            Move-Item $LogFile "$LogFile.old" -Force
        }
        Add-Content -Path $LogFile -Value $line -Encoding UTF8
    } catch { }
}

function ConvertTo-Mac([string]$Value) {
    $hex = ($Value -replace '[^0-9A-Fa-f]', '').ToUpper()
    if ($hex.Length -ne 12) { throw "Некорректный MAC-адрес: '$Value'. Нужен вид AA-BB-CC-DD-EE-FF." }
    ($hex -split '(..)' | Where-Object { $_ }) -join '-'
}

# Аргумент для powershell.exe: в кавычках, с экранированием по правилам командной строки Windows.
function ConvertTo-NativeArgument([string]$Value) {
    $v = $Value -replace '(\\*)"', '$1$1\"'
    $v = $v -replace '(\\+)$', '$1$1'
    '"' + $v + '"'
}

# Читает printer.psd1 как данные: UTF-8 с BOM или без, код из файла не выполняется.
function Read-ConfigFile([string]$Path) {
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
    if ($errors.Count) {
        throw "Файл настроек '$Path' повреждён (строка $($errors[0].Extent.StartLineNumber)): $($errors[0].Message)"
    }
    $hash = $ast.Find({ $args[0] -is [System.Management.Automation.Language.HashtableAst] }, $false)
    if (-not $hash) { throw "В файле настроек '$Path' нет таблицы @{ ... }." }
    try { $hash.SafeGetValue() } catch { throw "В файле настроек '$Path' допустимы только простые значения: $($_.Exception.Message)" }
}

function Measure-TcpConnect([string]$Address, [int]$Port, [int]$TimeoutMs = 1500) {
    $client = New-Object Net.Sockets.TcpClient
    try {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $ok = $client.BeginConnect($Address, $Port, $null, $null).AsyncWaitHandle.WaitOne($TimeoutMs)
        if ($ok -and $client.Connected) { return $sw.ElapsedMilliseconds }
        return $null
    } catch {
        return $null
    } finally {
        $client.Close()
    }
}

function Get-MacByIP([string]$Address) {
    $n = Get-NetNeighbor -IPAddress $Address -ErrorAction SilentlyContinue |
        Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' } |
        Select-Object -First 1
    if ($n) { $n.LinkLayerAddress.ToUpper() }
}

# Устройство по адресу отвечает (веб-интерфейс, IPP или ping) и его MAC - наш.
function Test-IsOurPrinter([string]$Address) {
    $alive = $false
    foreach ($port in $ProbePorts) {
        if ($null -ne (Measure-TcpConnect $Address $port)) { $alive = $true; break }
    }
    if (-not $alive) {
        try { $alive = (New-Object Net.NetworkInformation.Ping).Send($Address, 1500).Status -eq 'Success' } catch { }
    }
    $alive -and (Get-MacByIP $Address) -eq $PrinterMac
}

function Get-IPsByMac {
    Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress.ToUpper() -eq $PrinterMac } |
        Select-Object -ExpandProperty IPAddress
}

function Find-PrinterIP([string]$HintIP) {
    # 1. ARP-кэш и mDNS
    $candidates = @(Get-IPsByMac)
    if ($PrinterHostName) {
        $candidates += @(Resolve-DnsName "$PrinterHostName.local" -Type A -ErrorAction SilentlyContinue |
            Where-Object { $_.Type -eq 'A' } | Select-Object -ExpandProperty IPAddress)
    }
    foreach ($ip in ($candidates | Select-Object -Unique)) {
        if (Test-IsOurPrinter $ip) { return $ip }
    }

    # 2. Пинг всей подсети /24 (наполняет ARP-кэш) и снова поиск по MAC
    $prefixes = @()
    if ($HintIP -match '^(\d+\.\d+\.\d+)\.\d+$') { $prefixes += $Matches[1] }
    $prefixes += @(Get-NetIPConfiguration -ErrorAction SilentlyContinue |
        Where-Object { $_.IPv4DefaultGateway -and $_.IPv4DefaultGateway.NextHop -ne '0.0.0.0' } |
        ForEach-Object { $_.IPv4Address } |
        Where-Object { $_.PrefixLength -eq 24 } |
        ForEach-Object { $_.IPAddress -replace '\.\d+$', '' })

    foreach ($prefix in ($prefixes | Select-Object -Unique)) {
        $tasks = foreach ($i in 1..254) {
            (New-Object Net.NetworkInformation.Ping).SendPingAsync("$prefix.$i", 700)
        }
        try { [Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]$tasks) } catch { }
        foreach ($ip in @(Get-IPsByMac)) {
            if (Test-IsOurPrinter $ip) { return $ip }
        }
    }
    return $null
}

# Перезапуск с правами администратора: передаём все заданные параметры и путь к файлу настроек.
function Start-Elevated {
    Write-Host 'Нужны права администратора - запрашиваю через UAC...' -ForegroundColor Yellow
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', (ConvertTo-NativeArgument $PSCommandPath))
    foreach ($name in $Bound.Keys) {
        if ($name -in 'Config', 'Verbose', 'Debug') { continue }
        if ($Bound[$name] -is [switch]) { if ($Bound[$name]) { $argList += "-$name" } }
        else { $argList += "-$name"; $argList += ConvertTo-NativeArgument ([string]$Bound[$name]) }
    }
    $argList += '-Config'
    $argList += ConvertTo-NativeArgument $ConfigPath
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList
    } catch {
        throw 'Запрос прав администратора отклонён. Ничего не изменено.'
    }
}

$ConfigPath = if ($Config) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config) }
              else { Join-Path $PSScriptRoot 'printer.psd1' }
$exitCode = 0

try {
    # --- Настройки: параметры важнее файла ------------------------------------
    $cfg = @{}
    if (Test-Path -LiteralPath $ConfigPath) { $cfg = Read-ConfigFile $ConfigPath }
    foreach ($name in 'QueueName', 'PrinterMac', 'PrinterHostName', 'PortNumber') {
        if (-not $Bound.ContainsKey($name) -and $cfg.ContainsKey($name)) { Set-Variable -Name $name -Value $cfg[$name] }
    }
    if (-not $PortNumber) { $PortNumber = 9100 }
    if (-not $QueueName) { throw "Не задано имя очереди (QueueName): нет файла настроек $ConfigPath. Сначала запустите Fix-PrinterWiFi.cmd" }
    if (-not $PrinterMac) { throw 'Не задан MAC-адрес принтера (PrinterMac): сторож ищет принтер именно по нему.' }
    $PrinterMac = ConvertTo-Mac $PrinterMac
    $TaskName = "$TaskPrefix - " + ($QueueName -replace '[\\/:*?"<>|]', '_')

    # --- Установка / удаление -------------------------------------------------
    if ($Install) {
        if (-not (Test-IsAdmin)) { Start-Elevated; exit $ExitHandedOff }
        New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        if ($PSCommandPath -ne $InstalledScript) { Copy-Item -Path $PSCommandPath -Destination $InstalledScript -Force }

        $argument = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$InstalledScript`"" +
                    " -QueueName $(ConvertTo-NativeArgument $QueueName) -PrinterMac $PrinterMac" +
                    " -PrinterHostName $(ConvertTo-NativeArgument $PrinterHostName) -PortNumber $PortNumber"
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument
        $triggers = @(
            (New-ScheduledTaskTrigger -AtStartup),
            (New-ScheduledTaskTrigger -Once -At (Get-Date).Date -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes))
        )
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
            -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Principal $principal -Settings $settings `
            -Description "Следит за IP принтера ($PrinterMac) и перенастраивает порт очереди '$QueueName'." -Force | Out-Null
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Задача '$TaskName' установлена: при старте Windows и каждые $IntervalMinutes мин." -ForegroundColor Green
        Write-Host "Лог: $LogFile"
        exit 0
    }

    if ($Uninstall) {
        if (-not (Test-IsAdmin)) { Start-Elevated; exit $ExitHandedOff }
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        # Копию скрипта убираем, только если не осталось сторожей других принтеров
        if (-not (Get-ScheduledTask -TaskName "$TaskPrefix - *" -ErrorAction SilentlyContinue)) {
            Remove-Item $InstalledScript -Force -ErrorAction SilentlyContinue
        }
        Write-Host "Задача '$TaskName' удалена (лог оставлен: $LogFile)." -ForegroundColor Green
        exit 0
    }

    # --- Проверка -------------------------------------------------------------
    $printer = Get-Printer -Name $QueueName -ErrorAction SilentlyContinue
    if (-not $printer) { Write-Log "Очередь '$QueueName' не найдена - сначала запустите Fix-PrinterWiFi.cmd"; exit 0 }

    $port = Get-PrinterPort -Name $printer.PortName -ErrorAction SilentlyContinue
    $currentIP = if ($port) { $port.PrinterHostAddress }

    if ($currentIP -and (Test-IsOurPrinter $currentIP)) {
        Write-Host "OK: принтер на $currentIP"
        exit 0
    }

    $newIP = Find-PrinterIP $currentIP
    if (-not $newIP) {
        Write-Log "Принтер не найден в сети (выключен или вне Wi-Fi). Порт не менялся: $($printer.PortName)"
        exit 0
    }
    if ($newIP -eq $currentIP) { exit 0 }

    if (-not (Test-IsAdmin)) {
        Write-Log "Принтер переехал $currentIP -> $newIP, но для смены порта нужны права администратора (запустите от администратора или -Install)"
        exit 0
    }

    $newPort = if ($PortNumber -eq 9100) { "IP_$newIP" } else { "IP_${newIP}_$PortNumber" }
    if (-not (Get-PrinterPort -Name $newPort -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $newPort -PrinterHostAddress $newIP -PortNumber $PortNumber
    }
    Set-Printer -Name $QueueName -PortName $newPort
    Write-Log "Принтер переехал $currentIP -> ${newIP}: очередь переведена на порт $newPort"

    $oldPort = $printer.PortName
    if ($oldPort -like 'IP_*' -and -not (Get-Printer | Where-Object PortName -eq $oldPort)) {
        try { Remove-PrinterPort -Name $oldPort } catch { }
    }
} catch {
    if ($QueueName) { Write-Log "ОШИБКА: $($_.Exception.Message)" }
    else { Write-Host "ОШИБКА: $($_.Exception.Message)" -ForegroundColor Red }
    $exitCode = 1
}
exit $exitCode
