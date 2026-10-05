<#
.SYNOPSIS
    Переводит печать на сетевой принтер на прямой IP-порт и удаляет его дублирующиеся
    сетевые очереди (WSD/IPP/копии). Подходит для любого сетевого принтера,
    для которого установлен драйвер производителя.

.DESCRIPTION
    Почему печать по Wi-Fi бывает медленной: порт очереди указывает на сетевое имя принтера,
    которое Windows каждый раз ищет через LLMNR/mDNS (секунды, а при потере multicast-пакетов -
    ошибка и повтор задания спулером через минуты), либо очередь работает через WSD
    со стандартным драйвером Microsoft.

    Какой принтер исправлять, скрипт берёт из файла настроек printer.psd1 рядом с собой.
    Файл создаёт мастер (-Setup). Если файла нет и принтер не задан параметрами, мастер запустится сам.

    Скрипт на текущем компьютере:
      1. Проверяет, что по IP отвечает именно этот принтер (по MAC-адресу, если он задан).
      2. Находит сетевые очереди этого принтера: по адресу порта, по адресу WSD-устройства,
         по имени хоста и по шаблону из настроек. Показывает план и спрашивает подтверждение.
      3. Делает резервную копию всех принтеров (PrintBrm).
      4. Оставляет одну очередь с драйвером производителя на порту "IP_<адрес>" (RAW, без SNMP).
      5. Удаляет остальные сетевые копии; если одна из них была принтером по умолчанию,
         по умолчанию становится новая очередь.
    Не трогает: USB-очереди, подключения к общим принтерам других ПК (\\сервер\принтер),
    очереди других устройств, очереди факса (удалить: -RemoveFax) и сетевые очереди,
    которые не удалось уверенно отнести к этому принтеру. На самом принтере ничего не меняет.
    Если через WSD/IPP-устройство в Windows работает сетевой сканер, его WSD-очереди
    оставляются: при их удалении Windows удаляет устройство вместе со сканером
    (удалить всё равно: -RemoveNetworkScanner).

.EXAMPLE
    .\Fix-PrinterWiFi.ps1 -Setup
    Найти сетевые принтеры на этом ПК, выбрать нужный и записать printer.psd1.

.EXAMPLE
    .\Fix-PrinterWiFi.ps1 -WhatIf
    Показать, что будет сделано на этом ПК, ничего не меняя.

.EXAMPLE
    .\Fix-PrinterWiFi.ps1
    Выполнить (сам запросит права администратора через UAC).

.NOTES
    Параметры командной строки важнее значений из файла настроек.
    Код возврата: 0 - готово, 1 - ошибка, 99 - скрипт перезапущен с правами администратора в новом окне.
    Откат: PrintBrm.exe -R -F "<файл .printerExport из C:\ProgramData\PrinterWiFiFix>"
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    # Мастер: найти сетевые принтеры на этом ПК и записать файл настроек. Сам ничего не исправляет.
    [switch]$Setup,
    # Файл настроек. По умолчанию printer.psd1 рядом со скриптом.
    [string]$Config,
    [string]$PrinterIP,
    # MAC-адрес принтера. Если пуст, проверка "по адресу отвечает именно принтер" пропускается.
    [string]$PrinterMac,
    [string]$PrinterHostName,
    # Точное имя драйвера производителя, как в "Управлении печатью".
    [string]$DriverName,
    # Как будет называться итоговая очередь (одинаково на всех ПК).
    [string]$QueueName,
    # Регулярное выражение: дополнительный признак очередей этого принтера (имя очереди или драйвера).
    [string]$MatchPattern,
    [int]$PortNumber,
    # Удалять и сетевые очереди факса этого принтера.
    [switch]$RemoveFax,
    # Удалять WSD-очереди, даже если Windows при этом удалит сетевой сканер.
    [switch]$RemoveNetworkScanner,
    [switch]$TestPage,
    # Не задавать вопросов: принимать ответы по умолчанию.
    [switch]$Yes,
    [string]$WorkDir = (Join-Path $env:ProgramData 'PrinterWiFiFix')
)

$ErrorActionPreference = 'Stop'
# Загружаем модули заранее, иначе в режиме -WhatIf они засоряют вывод строками про свои алиасы
& { $WhatIfPreference = $false; Import-Module CimCmdlets, NetTCPIP, PrintManagement, PnpDevice -ErrorAction SilentlyContinue }

$Bound = $PSBoundParameters
# Порты для проверки связи. 9100 не используем: пустое подключение принтер может принять за задание печати.
$ProbePorts = 80, 443, 631
$ExitHandedOff = 99
$exitCode = 0
$dnsCache = @{}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Step([string]$Text) { Write-Host "`n=== $Text ===" -ForegroundColor Cyan }
function Write-Ok([string]$Text)   { Write-Host "  [OK] $Text" -ForegroundColor Green }
function Write-Info([string]$Text) { Write-Host "       $Text" }
function Write-Warn([string]$Text) { Write-Host "  [!]  $Text" -ForegroundColor Yellow }

# Вопрос пользователю. С ключом -Yes возвращает ответ по умолчанию.
function Read-Answer([string]$Prompt, [string]$Default) {
    $hasDefault = $PSBoundParameters.ContainsKey('Default')
    if ($Yes) {
        if ($hasDefault) { return $Default }
        throw "Нужен выбор пользователя ('$Prompt'), а задан ключ -Yes. Укажите значение параметром."
    }
    # В неинтерактивном сеансе Read-Host бросает исключение либо возвращает $null (ввод закрыт)
    $raw = $null
    try { $raw = Read-Host $Prompt } catch { }
    if ($null -eq $raw) {
        throw "Нужен ответ пользователя ('$Prompt'), а сеанс неинтерактивный. Задайте значение параметром или добавьте -Yes."
    }
    $answer = "$raw".Trim()
    if ($answer -eq '' -and $hasDefault) { return $Default }
    $answer
}

function Confirm-Answer([string]$Prompt) {
    (Read-Answer "$Prompt [Д/н]" 'д') -match '^(д|да|y|yes)$'
}

function Test-IPv4([string]$Value) {
    $ip = $null
    ($Value -match '^\d{1,3}(\.\d{1,3}){3}$') -and [Net.IPAddress]::TryParse($Value, [ref]$ip)
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

# Время TCP-подключения в мс или $null.
function Measure-TcpConnect([string]$Address, [int]$Port, [int]$TimeoutMs = 3000) {
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

# Время отклика принтера в мс или $null: веб-интерфейс, IPP, в крайнем случае ping.
function Measure-PrinterResponse([string]$Address) {
    foreach ($port in $ProbePorts) {
        $ms = Measure-TcpConnect $Address $port
        if ($null -ne $ms) { return $ms }
    }
    try {
        $reply = (New-Object Net.NetworkInformation.Ping).Send($Address, 2000)
        if ($reply.Status -eq 'Success') { return $reply.RoundtripTime }
    } catch { }
    return $null
}

function Get-MacByIP([string]$Address) {
    $n = Get-NetNeighbor -IPAddress $Address -ErrorAction SilentlyContinue |
        Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' } |
        Select-Object -First 1
    if ($n) { $n.LinkLayerAddress.ToUpper() }
}

# IPv4-адреса имени хоста (с ограничением по времени: мёртвые имена ищутся долго).
function Resolve-HostIPs([string]$HostName) {
    if (-not $HostName) { return @() }
    if (Test-IPv4 $HostName) { return @($HostName) }
    $key = $HostName.ToLower()
    if (-not $dnsCache.ContainsKey($key)) {
        $ips = @()
        try {
            $task = [Net.Dns]::GetHostAddressesAsync($HostName)
            if ($task.Wait(4000)) {
                $ips = @($task.Result | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | ForEach-Object { $_.ToString() })
            }
        } catch { }
        $dnsCache[$key] = $ips
    }
    @($dnsCache[$key])
}

function Get-PnpContainerId([string]$InstanceId) {
    try {
        [string](Get-PnpDeviceProperty -InstanceId $InstanceId -KeyName 'DEVPKEY_Device_ContainerId' -ErrorAction Stop).Data
    } catch { '' }
}

# IPv4-адреса сетевого устройства Windows из его свойств PnP-X (IpAddress, RemoteAddress, XAddrs).
function Get-PnpDeviceIPs([string]$InstanceId) {
    $ips = @()
    try {
        foreach ($prop in Get-PnpDeviceProperty -InstanceId $InstanceId -ErrorAction Stop) {
            if ($prop.KeyName -notmatch 'PNPX_(IpAddress|RemoteAddress|XAddrs)|Device_LocationInfo|^\{656A3BB3-ECC0-43FD-8477-4AE0404A96CD\}\s+(12297|4102|4099)$') { continue }
            foreach ($value in @($prop.Data)) {
                foreach ($m in [regex]::Matches([string]$value, '(?<![\d.])\d{1,3}(\.\d{1,3}){3}(?![\d.])')) {
                    if (Test-IPv4 $m.Value) { $ips += $m.Value }
                }
            }
        }
    } catch { }
    @($ips | Select-Object -Unique)
}

# Сетевые устройства Windows (WSD, eSCL, IPP): IP-адреса по контейнеру устройства и список сетевых сканеров.
function Get-NetworkDeviceInfo {
    $info = @{ IPsByContainer = @{}; Scanners = @() }
    $devices = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object { $_.InstanceId -match '^SWD\\(DAFWSDPROVIDER|ESCL|IPP)' })
    $scanners = @()
    foreach ($d in $devices) {
        $ips = @(Get-PnpDeviceIPs $d.InstanceId)
        $cid = Get-PnpContainerId $d.InstanceId
        if ($cid -and $ips) {
            $known = @()
            if ($info.IPsByContainer.ContainsKey($cid)) { $known = $info.IPsByContainer[$cid] }
            $info.IPsByContainer[$cid] = @($known + $ips | Select-Object -Unique)
        }
        if ($d.Class -eq 'Image' -and $d.InstanceId -match '^SWD\\(DAFWSDPROVIDER|ESCL)\\') {
            $scanners += [pscustomobject]@{ Name = [string]$d.FriendlyName; ContainerId = $cid; IPs = $ips }
        }
    }
    # У сканера адреса может не быть, а у соседнего устройства того же контейнера - быть
    foreach ($s in $scanners) {
        if (-not $s.IPs -and $s.ContainerId -and $info.IPsByContainer.ContainsKey($s.ContainerId)) {
            $s.IPs = $info.IPsByContainer[$s.ContainerId]
        }
    }
    $info.Scanners = $scanners
    $info
}

# Сетевые локальные очереди этого ПК с адресом устройства, если его удалось определить.
# USB/LPT/COM, виртуальные принтеры и подключения к общим принтерам (\\ПК\...) сюда не попадают.
# $KnownHosts - имена, которые искать в сети не нужно (имя хоста самого принтера).
function Get-NetworkQueues($DeviceInfo, [string[]]$KnownHosts = @()) {
    $ports = @{}
    Get-PrinterPort | ForEach-Object { $ports[$_.Name] = $_ }
    $queueDevices = $null
    foreach ($p in Get-Printer) {
        if ($p.Type -ne 'Local') { continue }
        if ($p.PortName -match '^(USB|LPT|COM|DOT4|TS)\d' -or $p.PortName -match ':$') { continue }
        $port = $ports[$p.PortName]
        if ($port -and $port.PortMonitor -eq 'Local Monitor') { continue }

        $hostAddress = ''
        if ($port -and $port.PrinterHostAddress) { $hostAddress = [string]$port.PrinterHostAddress }
        elseif ($port -and $port.HostName) { $hostAddress = [string]$port.HostName }

        $ips = @()
        if ($hostAddress) {
            $kind = 'TCP/IP'
            if ($KnownHosts -notcontains $hostAddress) { $ips = @(Resolve-HostIPs $hostAddress) }
        } else {
            $kind = if ($p.PortName -like 'WSD-*') { 'WSD' } else { 'сеть' }
            if ($p.PortName -match '^[a-z]+://([^/:]+)') {
                $hostAddress = $Matches[1]
                if ($KnownHosts -notcontains $hostAddress) { $ips = @(Resolve-HostIPs $hostAddress) }
            } else {
                # Очередь -> её устройство PnP -> контейнер -> адрес сетевого устройства того же контейнера
                if ($null -eq $queueDevices) { $queueDevices = @(Get-PnpDevice -Class PrintQueue -ErrorAction SilentlyContinue) }
                $dev = $queueDevices | Where-Object { $_.FriendlyName -eq $p.Name } | Select-Object -First 1
                if ($dev) {
                    $cid = Get-PnpContainerId $dev.InstanceId
                    if ($cid -and $DeviceInfo.IPsByContainer.ContainsKey($cid)) { $ips = @($DeviceInfo.IPsByContainer[$cid]) }
                }
            }
        }

        [pscustomobject]@{
            Name       = $p.Name
            DriverName = $p.DriverName
            PortName   = $p.PortName
            Kind       = $kind
            Host       = $hostAddress
            IPs        = $ips
            IsFax      = ($p.Name -match 'fax|факс' -or $p.DriverName -match 'fax|факс')
        }
    }
}

# Драйверы производителей (не классовые драйверы Microsoft).
function Get-NativeDrivers {
    Get-PrinterDriver | Where-Object { $_.Manufacturer -ne 'Microsoft' -and $_.Name -notmatch '^(Microsoft|Universal Print|Generic|Remote Desktop)' } |
        Select-Object -ExpandProperty Name -Unique | Sort-Object
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

function Save-ConfigFile([string]$Path, $Values) {
    $quote = { param($s) "'" + ([string]$s).Replace("'", "''") + "'" }
    $lines = @(
        '# Настройки фиксера печати. Создано мастером: Fix-PrinterWiFi.cmd -Setup'
        '# Файл можно править в Блокноте. Он описывает принтер вашей сети - не публикуйте его.'
        '@{'
        '    # IP-адрес принтера. Закрепите его за принтером на роутере, иначе адрес может смениться.'
        "    PrinterIP       = $(& $quote $Values.PrinterIP)"
        ''
        '    # MAC-адрес принтера: по нему проверяется, что по IP отвечает именно он.'
        '    # Пусто - проверка пропускается (и сторож Printer-Watchdog работать не сможет).'
        "    PrinterMac      = $(& $quote $Values.PrinterMac)"
        ''
        '    # Сетевое имя принтера (необязательно): по нему находятся старые порты и WSD-очереди.'
        "    PrinterHostName = $(& $quote $Values.PrinterHostName)"
        ''
        '    # Точное имя драйвера производителя.'
        "    DriverName      = $(& $quote $Values.DriverName)"
        ''
        '    # Имя итоговой очереди - одинаковое на всех компьютерах.'
        "    QueueName       = $(& $quote $Values.QueueName)"
        ''
        '    # Необязательно: регулярное выражение для очередей этого принтера, которые скрипт'
        "    # не смог опознать сам (по имени очереди или драйвера). Пример: 'Pantum1|M660'"
        "    MatchPattern    = $(& $quote $Values.MatchPattern)"
        ''
        '    # Порт RAW-печати.'
        "    PortNumber      = $([int]$Values.PortNumber)"
        '}'
    )
    if ($PSCmdlet.ShouldProcess($Path, 'Записать файл настроек')) {
        [IO.File]::WriteAllLines($Path, [string[]]$lines, (New-Object Text.UTF8Encoding($true)))
        Write-Ok "Настройки сохранены: $Path"
    }
}

# Мастер: выбрать принтер и собрать его настройки. Возвращает таблицу значений.
function Invoke-Setup {
    Write-Step 'Настройка: какой принтер исправляем'
    $s = @{
        PrinterIP = $PrinterIP; PrinterMac = $PrinterMac; PrinterHostName = $PrinterHostName
        DriverName = $DriverName; QueueName = $QueueName; MatchPattern = $MatchPattern; PortNumber = $PortNumber
    }
    if (-not $s.PortNumber) { $s.PortNumber = 9100 }

    Write-Info 'Ищу сетевые принтеры на этом компьютере (до минуты, если есть порты на имя хоста)...'
    $queues = @(Get-NetworkQueues (Get-NetworkDeviceInfo))

    if (-not $s.PrinterIP) {
        $groups = @($queues | Where-Object { $_.IPs } | Group-Object { $_.IPs[0] } | Sort-Object Name)
        $unknown = @($queues | Where-Object { -not $_.IPs })
        for ($i = 0; $i -lt $groups.Count; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $groups[$i].Name) -ForegroundColor White
            $groups[$i].Group | ForEach-Object { Write-Info "    '$($_.Name)' ($($_.DriverName), $($_.Kind))" }
        }
        if ($unknown) {
            Write-Info 'Сетевые очереди, адрес которых определить не удалось:'
            $unknown | ForEach-Object { Write-Info "    '$($_.Name)' ($($_.DriverName), $($_.Kind), порт $($_.PortName))" }
        }
        if (-not $groups) { Write-Warn 'Сетевых очередей с известным адресом на этом компьютере нет.' }

        if ($groups.Count -eq 1 -and $Yes) {
            $s.PrinterIP = $groups[0].Name
        } else {
            $hint = if ($groups) { "Номер принтера (1-$($groups.Count)) или его IP-адрес" } else { 'IP-адрес принтера (виден на роутере или в распечатке настроек сети)' }
            $answer = Read-Answer $hint
            if ($answer -match '^\d+$' -and [int]$answer -ge 1 -and [int]$answer -le $groups.Count) {
                $s.PrinterIP = $groups[[int]$answer - 1].Name
            } else {
                $s.PrinterIP = $answer
            }
        }
    }
    if (-not (Test-IPv4 $s.PrinterIP)) { throw "Некорректный IP-адрес: '$($s.PrinterIP)'." }
    $mine = @($queues | Where-Object { $_.IPs -contains $s.PrinterIP })

    # Связь и MAC
    $ms = Measure-PrinterResponse $s.PrinterIP
    if ($null -eq $ms) { throw "Принтер не отвечает по $($s.PrinterIP). Включите его, проверьте адрес и запустите настройку снова." }
    if (-not $s.PrinterMac) { $s.PrinterMac = [string](Get-MacByIP $s.PrinterIP) }
    if ($s.PrinterMac) {
        $s.PrinterMac = ConvertTo-Mac $s.PrinterMac
        Write-Ok "Принтер $($s.PrinterIP) отвечает за $ms мс, MAC $($s.PrinterMac)"
    } else {
        Write-Warn "Принтер $($s.PrinterIP) отвечает, но MAC определить не удалось (другая подсеть или VPN). Проверка по MAC будет пропускаться."
    }

    # Имя хоста: из портов очередей, из брошенных портов на имя, иначе обратный поиск
    if (-not $s.PrinterHostName) {
        $named = $mine | Where-Object { $_.Host -and -not (Test-IPv4 $_.Host) } | Select-Object -First 1
        if ($named) {
            $s.PrinterHostName = $named.Host
        } else {
            $s.PrinterHostName = Get-PrinterPort |
                Where-Object { $_.PrinterHostAddress -and -not (Test-IPv4 $_.PrinterHostAddress) } |
                Select-Object -ExpandProperty PrinterHostAddress -Unique |
                Where-Object { (Resolve-HostIPs $_) -contains $s.PrinterIP } |
                Select-Object -First 1
        }
        if (-not $s.PrinterHostName) {
            try {
                $task = [Net.Dns]::GetHostEntryAsync($s.PrinterIP)
                if ($task.Wait(3000) -and -not (Test-IPv4 $task.Result.HostName)) { $s.PrinterHostName = $task.Result.HostName }
            } catch { }
        }
        $s.PrinterHostName = ([string]$s.PrinterHostName -split '\.')[0]
    }
    if ($s.PrinterHostName) { Write-Info "Сетевое имя: $($s.PrinterHostName)" }

    # Драйвер производителя
    $native = @(Get-NativeDrivers)
    if (-not $s.DriverName) {
        $used = @($mine | Where-Object { -not $_.IsFax -and $native -contains $_.DriverName } | Select-Object -ExpandProperty DriverName -Unique)
        if ($used.Count -eq 1) {
            $s.DriverName = $used[0]
        } else {
            $choices = if ($used.Count -gt 1) { $used } else { @($native | Where-Object { $_ -notmatch 'fax|факс' }) }
            if (-not $choices) {
                throw 'На этом ПК не установлен драйвер производителя принтера. Установите его с сайта производителя и запустите настройку снова.'
            }
            Write-Info 'Какой драйвер у этого принтера:'
            for ($i = 0; $i -lt $choices.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $choices[$i]) }
            $answer = Read-Answer "Номер драйвера (1-$($choices.Count))"
            if ($answer -notmatch '^\d+$' -or [int]$answer -lt 1 -or [int]$answer -gt $choices.Count) { throw "Нет драйвера с номером '$answer'." }
            $s.DriverName = $choices[[int]$answer - 1]
        }
    }
    if (-not (Get-PrinterDriver -Name $s.DriverName -ErrorAction SilentlyContinue)) {
        throw "Драйвер '$($s.DriverName)' на этом ПК не установлен."
    }
    Write-Info "Драйвер: $($s.DriverName)"

    # Имя итоговой очереди
    if (-not $s.QueueName) {
        $ready = $mine | Where-Object { $_.DriverName -eq $s.DriverName -and $_.Host -eq $s.PrinterIP } | Select-Object -First 1
        $suggested = if ($ready) { $ready.Name } else { "$($s.DriverName) (Wi-Fi)" }
        $s.QueueName = Read-Answer "Имя итоговой очереди [Enter - '$suggested']" $suggested
    }
    Write-Info "Очередь: $($s.QueueName)"
    $s
}

$ConfigPath = if ($Config) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Config) }
              else { Join-Path $PSScriptRoot 'printer.psd1' }
$transcribing = $false

try {
    # --- Настройки: параметры важнее файла ------------------------------------
    $cfg = @{}
    if (-not $Setup -and (Test-Path -LiteralPath $ConfigPath)) { $cfg = Read-ConfigFile $ConfigPath }
    elseif ($Config -and -not $Setup) { throw "Файл настроек не найден: $ConfigPath" }
    foreach ($name in 'PrinterIP', 'PrinterMac', 'PrinterHostName', 'DriverName', 'QueueName', 'MatchPattern', 'PortNumber') {
        # -WhatIf:$false - иначе пробный прогон не применит и чтение настроек
        if (-not $Bound.ContainsKey($name) -and $cfg.ContainsKey($name)) { Set-Variable -Name $name -Value $cfg[$name] -WhatIf:$false }
    }

    # --- Мастер ---------------------------------------------------------------
    if ($Setup -or -not $PrinterIP) {
        if (-not $Setup) {
            if (Test-Path -LiteralPath $ConfigPath) { Write-Warn "В файле настроек не задан PrinterIP ($ConfigPath) - запускаю настройку." }
            else { Write-Warn "Файл настроек не найден ($ConfigPath) - запускаю настройку." }
        }
        $s = Invoke-Setup
        Save-ConfigFile $ConfigPath $s
        foreach ($name in $s.Keys) { Set-Variable -Name $name -Value $s[$name] -WhatIf:$false }
        if ($Setup) {
            Write-Host "`nНастройка завершена. Скопируйте папку со скриптами и файлом printer.psd1 на каждый компьютер" -ForegroundColor Green
            Write-Host 'и запустите Fix-PrinterWiFi.cmd. Посмотреть план заранее: Fix-PrinterWiFi.cmd -WhatIf' -ForegroundColor Green
            exit 0
        }
        if (-not $WhatIfPreference -and -not (Confirm-Answer 'Исправить печать на этом компьютере сейчас?')) {
            Write-Host 'Настройки сохранены, на компьютере ничего не изменено.'
            exit 0
        }
    }

    # --- Проверка настроек ----------------------------------------------------
    if (-not $PortNumber) { $PortNumber = 9100 }
    if (-not (Test-IPv4 $PrinterIP)) { throw "Некорректный IP-адрес принтера: '$PrinterIP'. Запустите настройку: Fix-PrinterWiFi.cmd -Setup" }
    if (-not $DriverName) { throw 'Не задано имя драйвера (DriverName). Запустите настройку: Fix-PrinterWiFi.cmd -Setup' }
    if (-not $QueueName)  { throw 'Не задано имя итоговой очереди (QueueName). Запустите настройку: Fix-PrinterWiFi.cmd -Setup' }
    if ($PrinterMac) { $PrinterMac = ConvertTo-Mac $PrinterMac }
    if ($MatchPattern) {
        try { [void][regex]::new($MatchPattern) } catch { throw "MatchPattern - некорректное регулярное выражение: $($_.Exception.InnerException.Message)" }
    }

    # --- Права администратора -------------------------------------------------
    if (-not $WhatIfPreference -and -not (Test-IsAdmin)) {
        Write-Host 'Нужны права администратора - запрашиваю через UAC...' -ForegroundColor Yellow
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', (ConvertTo-NativeArgument $PSCommandPath))
        foreach ($name in $Bound.Keys) {
            if ($name -in 'Setup', 'Config', 'WhatIf', 'Confirm', 'Verbose', 'Debug') { continue }
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
        exit $ExitHandedOff
    }

    $portName = if ($PortNumber -eq 9100) { "IP_$PrinterIP" } else { "IP_${PrinterIP}_$PortNumber" }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if (-not $WhatIfPreference) {
        New-Item -ItemType Directory -Path $WorkDir -Force -WhatIf:$false | Out-Null
        Start-Transcript -Path (Join-Path $WorkDir "Fix-PrinterWiFi_${env:COMPUTERNAME}_$stamp.log") -WhatIf:$false | Out-Null
        $transcribing = $true
    }

    Write-Step "Компьютер $env:COMPUTERNAME"
    Write-Info "Принтер:  $PrinterIP, порт $PortNumber"
    Write-Info "Драйвер:  $DriverName"
    Write-Info "Очередь:  $QueueName"

    # --- 1. Принтер доступен по IP и это он ---------------------------------
    Write-Step "1. Проверка принтера по адресу $PrinterIP"
    $ms = Measure-PrinterResponse $PrinterIP
    if ($null -eq $ms) {
        throw "Принтер не отвечает по $PrinterIP. Проверьте, что он включён и в сети, и какой у него IP (роутер или веб-интерфейс принтера). Ничего не изменено."
    }
    if ($PrinterMac) {
        $mac = Get-MacByIP $PrinterIP
        if (-not $mac) {
            throw "Не удалось определить MAC устройства по адресу $PrinterIP (компьютер и принтер в разных подсетях или связь идёт через VPN?). Ничего не изменено."
        }
        if ($mac -ne $PrinterMac) {
            throw "По адресу $PrinterIP отвечает устройство с MAC '$mac', а не принтер ($PrinterMac). IP сменился? Ничего не изменено."
        }
        Write-Ok "Отвечает за $ms мс, MAC $mac совпадает"
    } else {
        Write-Ok "Отвечает за $ms мс"
        Write-Warn 'MAC принтера не задан - не проверяю, что по этому адресу отвечает именно он.'
    }

    # --- 2. Что оставить и что удалить --------------------------------------
    Write-Step '2. Анализ очередей'
    if (-not (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue)) {
        throw "На этом ПК не установлен драйвер '$DriverName'. Установите драйвер принтера с сайта производителя и запустите скрипт снова. Ничего не изменено."
    }

    $printerHosts = @($PrinterIP)
    if ($PrinterHostName) { $printerHosts += $PrinterHostName, "$PrinterHostName.local" }
    $deviceInfo = Get-NetworkDeviceInfo
    $queues = @(Get-NetworkQueues $deviceInfo $printerHosts)

    # Чья очередь: 'Ours' - этого принтера, 'Other' - другого устройства, 'Unknown' - непонятно (не трогаем)
    foreach ($q in $queues) {
        $relation = 'Unknown'
        $why = 'адрес устройства определить не удалось'
        if ($q.Name -eq $QueueName) { $relation = 'Ours'; $why = 'итоговая очередь' }
        elseif ($q.IPs -contains $PrinterIP) { $relation = 'Ours'; $why = "адрес порта $($q.Kind)" }
        elseif ($q.Host -and $printerHosts -contains $q.Host) { $relation = 'Ours'; $why = 'порт на имя хоста принтера' }
        elseif ($q.IPs) { $relation = 'Other'; $why = "другое устройство: $($q.IPs -join ', ')" }
        elseif ($PrinterHostName -and $q.Name.IndexOf($PrinterHostName, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $relation = 'Ours'; $why = 'имя хоста принтера в названии' }
        elseif ($MatchPattern -and ($q.Name -match $MatchPattern -or $q.DriverName -match $MatchPattern)) { $relation = 'Ours'; $why = 'шаблон MatchPattern' }
        $q | Add-Member -NotePropertyName Relation -NotePropertyValue $relation
        $q | Add-Member -NotePropertyName Why -NotePropertyValue $why
    }
    $ours = @($queues | Where-Object { $_.Relation -eq 'Ours' })
    $unknown = @($queues | Where-Object { $_.Relation -eq 'Unknown' })

    $clash = Get-Printer -Name $QueueName -ErrorAction SilentlyContinue
    if ($clash -and ($ours.Name -notcontains $QueueName)) {
        throw "Очередь с именем '$QueueName' уже есть и не является сетевой очередью этого принтера (порт $($clash.PortName)). Задайте другое имя (QueueName). Ничего не изменено."
    }

    $keep = $ours | Where-Object { $_.Name -eq $QueueName } | Select-Object -First 1
    if (-not $keep) {
        # Предпочитаем очередь с драйвером производителя, уже смотрящую на принтер по TCP/IP
        $keep = $ours |
            Where-Object { $_.DriverName -eq $DriverName -and -not $_.IsFax } |
            Sort-Object { if ($_.Kind -eq 'TCP/IP') { 0 } else { 1 } }, Name |
            Select-Object -First 1
    }

    $fax = @($ours | Where-Object { $_.IsFax -and -not $RemoveFax -and (-not $keep -or $_.Name -ne $keep.Name) })
    $toRemove = @($ours | Where-Object { (-not $keep -or $_.Name -ne $keep.Name) -and ($fax.Name -notcontains $_.Name) })

    # Очереди на WSD-портах - часть сетевого устройства Windows. При удалении такой очереди
    # Windows удаляет устройство целиком, вместе с сетевым сканером (WSD/eSCL).
    # Сканер считаем своим, пока не доказано, что он от другого устройства.
    $netScanners = @($deviceInfo.Scanners | Where-Object { -not $_.IPs -or $_.IPs -contains $PrinterIP })
    $protected = @()
    if ($netScanners -and -not $RemoveNetworkScanner) {
        $protected = @($toRemove | Where-Object { $_.PortName -like 'WSD-*' })
        $toRemove = @($toRemove | Where-Object { $_.PortName -notlike 'WSD-*' })
    }

    $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
    if ($existingPort -and $existingPort.PrinterHostAddress -ne $PrinterIP) {
        throw "Порт '$portName' уже существует, но указывает на '$($existingPort.PrinterHostAddress)'. Удалите его в 'Управлении печатью' и запустите скрипт снова. Ничего не изменено."
    }

    # Старые TCP/IP-порты этого принтера, которые после изменений никому не будут нужны
    $portsInUse = @(Get-Printer | Where-Object { $toRemove.Name -notcontains $_.Name -and (-not $keep -or $_.Name -ne $keep.Name) } |
        Select-Object -ExpandProperty PortName) + $portName
    $stalePorts = @(Get-PrinterPort | Where-Object {
        $_.PrinterHostAddress -and $printerHosts -contains $_.PrinterHostAddress -and $portsInUse -notcontains $_.Name
    })

    if ($keep) {
        Write-Info "Оставить:  '$($keep.Name)' ($($keep.DriverName), $($keep.PortName)) -> '$QueueName' на $portName"
        if ($keep.DriverName -ne $DriverName) { Write-Info "           драйвер будет заменён на '$DriverName'" }
    } else {
        Write-Info "Оставить:  создать '$QueueName' (драйвер '$DriverName', порт $portName)"
    }
    if ($toRemove) { $toRemove | ForEach-Object { Write-Info "Удалить:   '$($_.Name)' ($($_.DriverName), $($_.PortName)) - $($_.Why)" } }
    else           { Write-Info 'Удалить:   нечего' }
    foreach ($p in $stalePorts) { Write-Info "Удалить:   неиспользуемый порт $($p.Name) ($($p.PrinterHostAddress))" }
    foreach ($q in $fax) { Write-Info "Не трогаю: '$($q.Name)' - очередь факса (удалить: -RemoveFax)" }
    foreach ($q in $protected) {
        Write-Warn "Оставлена '$($q.Name)': вместе с ней Windows удалит сетевой сканер ($($netScanners.Name -join ', ')). Печатать на неё не нужно. Удалить всё равно: -RemoveNetworkScanner"
    }
    foreach ($q in $unknown) {
        Write-Warn "Не трогаю '$($q.Name)' ($($q.DriverName), $($q.PortName)): $($q.Why). Если это тот же принтер, добавьте её имя в MatchPattern в printer.psd1."
    }

    $defaultPrinter = (Get-CimInstance Win32_Printer -Filter 'Default=TRUE' -ErrorAction SilentlyContinue).Name
    $makeDefault = $defaultPrinter -and ($ours.Name -contains $defaultPrinter) -and $defaultPrinter -ne $QueueName

    $nothingToDo = $keep -and $keep.Name -eq $QueueName -and $keep.PortName -eq $portName -and $keep.DriverName -eq $DriverName -and
                   -not $toRemove -and -not $stalePorts -and -not $makeDefault

    if ($nothingToDo) {
        Write-Ok 'Всё уже настроено, менять нечего.'
    } else {
        if (-not $WhatIfPreference -and -not (Confirm-Answer "`nВыполнить этот план?")) {
            Write-Host 'Отменено. Ничего не изменено.' -ForegroundColor Yellow
            exit 0
        }

        # --- 3. Резервная копия -----------------------------------------------
        Write-Step '3. Резервная копия принтеров'
        $backupFile = Join-Path $WorkDir "printers-before_${env:COMPUTERNAME}_$stamp.printerExport"
        if ($PSCmdlet.ShouldProcess($backupFile, 'PrintBrm: резервная копия всех принтеров, портов и драйверов')) {
            $printBrm = Join-Path $env:WINDIR 'System32\spool\tools\PrintBrm.exe'
            & $printBrm -B -F $backupFile | Out-Host
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path $backupFile)) {
                throw "PrintBrm не смог сделать резервную копию (код $LASTEXITCODE). Ничего не изменено."
            }
            Write-Ok "Бэкап: $backupFile"
            Write-Ok "Откат:  `"$printBrm`" -R -F `"$backupFile`""
        }

        # --- 4. Порт на IP ----------------------------------------------------
        Write-Step "4. Порт $portName (RAW $PortNumber, без SNMP)"
        if ($existingPort) {
            Write-Ok 'Порт уже существует'
        } elseif ($PSCmdlet.ShouldProcess($portName, "Создать порт TCP/IP -> ${PrinterIP}:$PortNumber")) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $PrinterIP -PortNumber $PortNumber
            Write-Ok 'Порт создан'
        }

        # --- 5. Основная очередь ----------------------------------------------
        Write-Step "5. Очередь '$QueueName'"
        if (-not $keep) {
            if ($PSCmdlet.ShouldProcess($QueueName, "Создать очередь (драйвер '$DriverName', порт $portName)")) {
                Add-Printer -Name $QueueName -DriverName $DriverName -PortName $portName
                Write-Ok 'Очередь создана'
            }
        } else {
            if ($keep.PortName -ne $portName -and
                $PSCmdlet.ShouldProcess($keep.Name, "Перевести с порта '$($keep.PortName)' на '$portName'")) {
                Set-Printer -Name $keep.Name -PortName $portName
                Write-Ok "Порт очереди: $portName"
            }
            if ($keep.DriverName -ne $DriverName -and
                $PSCmdlet.ShouldProcess($keep.Name, "Заменить драйвер '$($keep.DriverName)' на '$DriverName'")) {
                Set-Printer -Name $keep.Name -DriverName $DriverName
                Write-Ok "Драйвер очереди: $DriverName"
            }
            if ($keep.Name -ne $QueueName -and
                $PSCmdlet.ShouldProcess($keep.Name, "Переименовать в '$QueueName'")) {
                Rename-Printer -Name $keep.Name -NewName $QueueName
                Write-Ok "Переименована в '$QueueName'"
            }
            if ($keep.Name -eq $QueueName -and $keep.PortName -eq $portName -and $keep.DriverName -eq $DriverName) { Write-Ok 'Уже настроена' }
        }

        # --- 6. Сетевые дубли -------------------------------------------------
        Write-Step '6. Удаление сетевых дублей'
        foreach ($q in $toRemove) {
            # WSD/IPP-очереди одного устройства Windows удаляет вместе: вторая может уже исчезнуть
            if (-not $WhatIfPreference -and -not (Get-Printer -Name $q.Name -ErrorAction SilentlyContinue)) {
                Write-Ok "'$($q.Name)' - удалена Windows вместе со связанной очередью"
                continue
            }
            if ($PSCmdlet.ShouldProcess($q.Name, 'Отменить задания и удалить очередь')) {
                try {
                    Get-PrintJob -PrinterName $q.Name -ErrorAction SilentlyContinue | Remove-PrintJob -ErrorAction SilentlyContinue
                    Remove-Printer -Name $q.Name -Confirm:$false
                    Write-Ok "Удалена '$($q.Name)'"
                } catch {
                    Write-Warn "Не удалось удалить '$($q.Name)': $($_.Exception.Message)"
                }
            }
        }
        if (-not $toRemove) { Write-Info 'Дублей нет' }

        foreach ($p in $stalePorts) {
            if ($PSCmdlet.ShouldProcess($p.Name, "Удалить неиспользуемый порт ($($p.PrinterHostAddress))")) {
                try {
                    Remove-PrinterPort -Name $p.Name
                    Write-Ok "Удалён порт $($p.Name)"
                } catch {
                    Write-Warn "Порт $($p.Name) пока занят спулером - удалится при повторном запуске скрипта после перезагрузки"
                }
            }
        }

        if ($makeDefault -and $PSCmdlet.ShouldProcess($QueueName, 'Сделать принтером по умолчанию (был сетевой дубль этого принтера)')) {
            $wp = Get-CimInstance Win32_Printer -Filter "Name='$($QueueName.Replace("'", "''"))'"
            Invoke-CimMethod -InputObject $wp -MethodName SetDefaultPrinter | Out-Null
            Write-Ok "'$QueueName' - принтер по умолчанию"
        }
    }

    # --- 7. Итог --------------------------------------------------------------
    Write-Step '7. Итог'
    Get-Printer | Where-Object { $_.Name -eq $QueueName -or $ours.Name -contains $_.Name } |
        Format-Table Name, DriverName, PortName -AutoSize | Out-String -Width 200 | Write-Host
    Write-Ok "Подключение к принтеру по IP: $(Measure-PrinterResponse $PrinterIP) мс (без поиска имени в сети)"

    if ($TestPage -and -not $WhatIfPreference) {
        $wp = Get-CimInstance Win32_Printer -Filter "Name='$($QueueName.Replace("'", "''"))'"
        Invoke-CimMethod -InputObject $wp -MethodName PrintTestPage | Out-Null
        Write-Ok 'Отправлена пробная страница'
    }

    if ($WhatIfPreference) {
        Write-Host "`nЭто был пробный прогон (-WhatIf): ничего не изменено." -ForegroundColor Yellow
    } else {
        Write-Host "`nГотово. Печатайте на '$QueueName'." -ForegroundColor Green
        Write-Host 'USB-очереди и другие принтеры не изменялись.'
        if ($protected) { Write-Host 'Сетевой сканер сохранён: связанные с ним WSD-очереди оставлены (печатать на них не нужно).' }
    }
} catch {
    Write-Host "`nОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
} finally {
    if ($transcribing) { Stop-Transcript | Out-Null }
}
exit $exitCode
