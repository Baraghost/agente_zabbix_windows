# =========================================
# Script: instalar_agente_zabbix.ps1
# Descripcion: Aplicacion para instalar Agente Zabbix con menu
# Fuente: MSI oficial de Zabbix
# Requiere: PowerShell como Administrador
# Autor: Baraghost
# =========================================

param(
    [Parameter(Mandatory = $false)]
    [string]$ServerIp,

    [Parameter(Mandatory = $false)]
    [string]$Hostname = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [int]$AgentTimeout = 10
)

$ErrorActionPreference = 'Stop'

# ================================
# CONFIG
# ================================
$ZabbixUrl        = "https://cdn.zabbix.com/zabbix/binaries/stable/7.4/latest/zabbix_agent2-7.4-latest-windows-amd64-openssl.msi"
$ZabbixMsiPath    = Join-Path $env:TEMP "zabbix_agent2.msi"
$InstallDir       = "C:\Program Files\Zabbix Agent 2"
$ConfigPath       = Join-Path $InstallDir "zabbix_agent2.conf"
$LegacyDir        = "C:\zabbix"
$FirewallRuleName = "Zabbix Agent 2"
$TranscriptLog    = "C:\Windows\Temp\zabbix_agent2_install.log"
$MsiLog           = "C:\Windows\Temp\zabbix_agent2_msi.log"
$AgentLogFile     = "C:\zabbix\zabbix_agent2.log"

$RegistryPathsToClean = @(
    "HKLM:\SYSTEM\CurrentControlSet\Services\Zabbix Agent 2",
    "HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\Application\Zabbix Agent 2"
)

$FoldersToClean = @(
    "C:\zabbix",
    "C:\Program Files\Zabbix Agent 2"
)

# ================================
# FUNCIONES
# ================================
function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'INFO'  { 'Cyan' }
        'OK'    { 'Green' }
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
    }

    $prefix = switch ($Level) {
        'INFO'  { '[..]' }
        'OK'    { '[OK]' }
        'WARN'  { '[!]' }
        'ERROR' { '[X]' }
    }

    Write-Host "$prefix $Message" -ForegroundColor $color
}

function Test-IsAdministrator {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Set-OrAddConfigValue {
    param(
        [string]$Path,
        [string]$Key,
        [string]$Value
    )

    $content = Get-Content -Path $Path -Raw -ErrorAction Stop

    if ($content -match "(?m)^\s*#?\s*$([regex]::Escape($Key))\s*=.*$") {
        $content = [regex]::Replace(
            $content,
            "(?m)^\s*#?\s*$([regex]::Escape($Key))\s*=.*$",
            "$Key=$Value"
        )
    } else {
        $content += "`r`n$Key=$Value`r`n"
    }

    Set-Content -Path $Path -Value $content -Encoding ASCII -Force
}

function Get-ConfigValue {
    param(
        [string]$Path,
        [string]$Key
    )

    $match = Select-String -Path $Path -Pattern "^\s*$([regex]::Escape($Key))\s*=" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($match) {
        return (($match.Line -split "=", 2)[1]).Trim()
    }
    return $null
}

function Get-ZabbixServices {
    Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match 'zabbix' -or $_.DisplayName -match 'zabbix'
    }
}

function Stop-And-Delete-ZabbixServices {
    $services = Get-ZabbixServices

    if (-not $services) {
        Write-Status "No se detectaron servicios Zabbix previos" "OK"
        return
    }

    foreach ($svc in $services) {
        Write-Status "Deteniendo servicio '$($svc.Name)' / '$($svc.DisplayName)'" "WARN"
        try {
            Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Status "No se pudo detener '$($svc.Name)': $($_.Exception.Message)" "WARN"
        }

        Start-Sleep -Seconds 2

        Write-Status "Eliminando servicio '$($svc.Name)'" "WARN"
        & sc.exe delete $svc.Name | Out-Null
    }

    Start-Sleep -Seconds 3
}

function Remove-ZabbixProducts {
    $uninstallRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $products = Get-ItemProperty $uninstallRoots -ErrorAction SilentlyContinue | Where-Object {
        $_.DisplayName -match '^Zabbix Agent( 2)?'
    }

    if (-not $products) {
        Write-Status "No se detectaron productos MSI previos de Zabbix" "OK"
        return
    }

    foreach ($product in $products) {
        Write-Status "Desinstalando producto previo: $($product.DisplayName)" "WARN"

        if ($product.PSChildName -match '^\{[0-9A-Fa-f\-]+\}$') {
            $guid = $product.PSChildName
            $args = @(
                "/x", $guid,
                "/qn",
                "/norestart",
                "/l*v", "`"$MsiLog`""
            )

            $proc = Start-Process -FilePath "msiexec.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow
            if ($proc.ExitCode -ne 0) {
                Write-Status "La desinstalación devolvió código $($proc.ExitCode). Continúo." "WARN"
            }
        }
        elseif ($product.UninstallString) {
            Write-Status "No se encontró ProductCode, usando UninstallString" "WARN"
            cmd.exe /c $product.UninstallString | Out-Null
        }
    }

    Start-Sleep -Seconds 5
}

function Remove-RegistryKeyIfExists {
    param([string]$Path)

    if (Test-Path $Path) {
        Write-Status "Eliminando clave de registro: $Path" "WARN"
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Status "Clave eliminada: $Path" "OK"
    } else {
        Write-Status "No existe la clave: $Path" "OK"
    }
}

function Remove-FolderIfExists {
    param([string]$Path)

    if (Test-Path $Path) {
        Write-Status "Eliminando carpeta: $Path" "WARN"
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        Write-Status "Carpeta eliminada: $Path" "OK"
    } else {
        Write-Status "No existe la carpeta: $Path" "OK"
    }
}

function Ensure-Folder {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Find-InstalledZabbixService {
    $services = Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | Where-Object {
        $_.PathName -match 'zabbix_agent2.exe' -or $_.Name -match 'zabbix' -or $_.DisplayName -match 'zabbix'
    }

    return $services | Select-Object -First 1
}

function Show-Menu {
    Clear-Host
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host "     ZABBIX AGENT 2 - INSTALADOR" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host "1) Instalación completa (clean install)"
    Write-Host "2) Instalación sin borrar C:\zabbix (mantener logs)"
    Write-Host "3) Salir"
    Write-Host ""
}

function Get-UserOption {
    do {
        $option = Read-Host "Seleccione una opción"
    } while ($option -notin @("1","2","3"))
    return $option
}

# ================================
# INICIO
# ================================
if (-not (Test-IsAdministrator)) {
    Write-Host "[X] Este script debe ejecutarse como Administrador." -ForegroundColor Red
    exit 1
}

$PreserveLogs = $false

Show-Menu
$choice = Get-UserOption

switch ($choice) {
    "1" {
        Write-Host "Modo: instalación completa" -ForegroundColor Yellow
        $PreserveLogs = $false
    }
    "2" {
        Write-Host "Modo: instalación sin borrar C:\zabbix" -ForegroundColor Yellow
        $PreserveLogs = $true
    }
    "3" {
        Write-Host "Saliendo..." -ForegroundColor Yellow
        exit 0
    }
}

if (-not $ServerIp) {
    $ServerIp = Read-Host "Ingrese la IP o DNS del servidor Zabbix"
}

Start-Transcript -Path $TranscriptLog -Append | Out-Null

try {
    Write-Host ""
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host " Instalador Zabbix Agent 2" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host "Servidor Zabbix : $ServerIp" -ForegroundColor Gray
    Write-Host "Hostname        : $Hostname" -ForegroundColor Gray
    Write-Host "MSI URL         : $ZabbixUrl" -ForegroundColor Gray
    Write-Host "Config esperada : $ConfigPath" -ForegroundColor Gray
    Write-Host "Transcript      : $TranscriptLog" -ForegroundColor Gray
    Write-Host "MSI log         : $MsiLog" -ForegroundColor Gray
    Write-Host "PreserveLogs    : $PreserveLogs" -ForegroundColor Gray
    Write-Host ""

    Write-Status "Paso 1/10 - Validando conectividad básica" "INFO"
    $pingOk = Test-Connection -ComputerName $ServerIp -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($pingOk) {
        Write-Status "Conectividad ICMP OK con $ServerIp" "OK"
    } else {
        Write-Status "No respondió al ping. Continúo igualmente." "WARN"
    }

    Write-Status "Paso 2/10 - Deteniendo y eliminando servicios Zabbix previos" "INFO"
    Stop-And-Delete-ZabbixServices
    Write-Status "Limpieza de servicios finalizada" "OK"

    Write-Status "Paso 3/10 - Desinstalando productos MSI previos de Zabbix" "INFO"
    Remove-ZabbixProducts
    Write-Status "Desinstalación previa finalizada" "OK"

    Write-Status "Paso 4/10 - Eliminando residuos de registro" "INFO"
    foreach ($regPath in $RegistryPathsToClean) {
        Remove-RegistryKeyIfExists -Path $regPath
    }
    Write-Status "Limpieza de registro finalizada" "OK"

    Write-Status "Paso 5/10 - Eliminando carpetas previas" "INFO"
    foreach ($folder in $FoldersToClean) {
        if ($PreserveLogs -and $folder -eq "C:\zabbix") {
            Write-Status "Se preserva la carpeta $folder (modo mantener logs)" "INFO"
            continue
        }

        Remove-FolderIfExists -Path $folder
    }
    Write-Status "Limpieza de carpetas finalizada" "OK"

    Write-Status "Esperando unos segundos para estabilizar el sistema..." "INFO"
    Start-Sleep -Seconds 3

    Write-Status "Paso 6/10 - Descargando instalador" "INFO"
    $maxRetries = 3
    $downloaded = $false

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            Write-Status "Intento de descarga $attempt de $maxRetries" "INFO"
            Invoke-WebRequest -Uri $ZabbixUrl -OutFile $ZabbixMsiPath
            if ((Test-Path $ZabbixMsiPath) -and ((Get-Item $ZabbixMsiPath).Length -gt 1MB)) {
                $downloaded = $true
                break
            }
        } catch {
            Write-Status "Fallo descarga: $($_.Exception.Message)" "WARN"
            Start-Sleep -Seconds 3
        }
    }

    if (-not $downloaded) {
        throw "No se pudo descargar correctamente el instalador."
    }

    Write-Status "Instalador descargado en $ZabbixMsiPath" "OK"

    Write-Status "Paso 7/10 - Instalando Zabbix Agent 2" "INFO"
    if (Test-Path $MsiLog) {
        Remove-Item $MsiLog -Force -ErrorAction SilentlyContinue
    }

    Start-Sleep -Seconds 2

    $msiArgs = @(
        "/i", "`"$ZabbixMsiPath`"",
        "/qn",
        "/norestart",
        "/l*v", "`"$MsiLog`"",
        "SERVER=$ServerIp",
        "SERVERACTIVE=$ServerIp",
        "HOSTNAME=$Hostname",
        "ENABLEPATH=1",
        "STARTUPTYPE=delayed"
    )

    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) {
        throw "La instalación MSI falló con código de salida $($process.ExitCode). Revisar: $MsiLog"
    }

    Write-Status "Instalación MSI completada correctamente" "OK"

    Write-Status "Paso 8/10 - Ajustando configuración final" "INFO"
    if (-not (Test-Path $ConfigPath)) {
        throw "No se encontró el archivo esperado de configuración: $ConfigPath"
    }

    Ensure-Folder -Path $LegacyDir

    Set-OrAddConfigValue -Path $ConfigPath -Key "Server"       -Value $ServerIp
    Set-OrAddConfigValue -Path $ConfigPath -Key "ServerActive" -Value $ServerIp
    Set-OrAddConfigValue -Path $ConfigPath -Key "Hostname"     -Value $Hostname
    Set-OrAddConfigValue -Path $ConfigPath -Key "Timeout"      -Value $AgentTimeout
    Set-OrAddConfigValue -Path $ConfigPath -Key "LogType"      -Value "file"
    Set-OrAddConfigValue -Path $ConfigPath -Key "LogFile"      -Value $AgentLogFile

    Write-Status "Configuración aplicada correctamente" "OK"

    Write-Status "Paso 9/10 - Configurando firewall" "INFO"
    $existingRule = Get-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue
    if ($existingRule) {
        Write-Status "La regla de firewall ya existe. Se elimina y recrea." "WARN"
        Remove-NetFirewallRule -DisplayName $FirewallRuleName -ErrorAction SilentlyContinue
    }

    New-NetFirewallRule `
        -DisplayName $FirewallRuleName `
        -Direction Inbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort 10050 `
        -RemoteAddress $ServerIp | Out-Null

    Write-Status "Regla de firewall creada para 10050/TCP desde $ServerIp" "OK"

    Write-Status "Paso 10/10 - Detectando, iniciando y verificando servicio" "INFO"
    $svc = Find-InstalledZabbixService
    if (-not $svc) {
        throw "La instalación terminó pero no se detectó ningún servicio Zabbix. Revisar: $MsiLog"
    }

    Write-Status "Servicio detectado: Name='$($svc.Name)' DisplayName='$($svc.DisplayName)'" "OK"

    $service = Get-Service -Name $svc.Name -ErrorAction Stop
    if ($service.Status -ne 'Running') {
        Start-Service -Name $svc.Name
        Start-Sleep -Seconds 4
        $service = Get-Service -Name $svc.Name -ErrorAction Stop
    }

    if ($service.Status -ne 'Running') {
        throw "El servicio '$($svc.Name)' existe pero no quedó en estado Running."
    }

    Write-Status "Servicio en ejecución" "OK"

    $listenOk = $false
    try {
        $listen = Get-NetTCPConnection -LocalPort 10050 -State Listen -ErrorAction SilentlyContinue
        if ($listen) { $listenOk = $true }
    } catch {
        $listenOk = $false
    }

    if ($listenOk) {
        Write-Status "El agente está escuchando en 10050" "OK"
    } else {
        Write-Status "No se detectó escucha en 10050. Revisar config o logs." "WARN"
    }

    $activeTest = Test-NetConnection -ComputerName $ServerIp -Port 10051 -WarningAction SilentlyContinue
    if ($activeTest.TcpTestSucceeded) {
        Write-Status "Conectividad OK hacia $ServerIp`:10051" "OK"
    } else {
        Write-Status "No se pudo validar conectividad hacia $ServerIp`:10051" "WARN"
    }

    $cfgServer       = Get-ConfigValue -Path $ConfigPath -Key "Server"
    $cfgServerActive = Get-ConfigValue -Path $ConfigPath -Key "ServerActive"
    $cfgHostname     = Get-ConfigValue -Path $ConfigPath -Key "Hostname"
    $cfgTimeout      = Get-ConfigValue -Path $ConfigPath -Key "Timeout"
    $cfgLogFile      = Get-ConfigValue -Path $ConfigPath -Key "LogFile"

    Write-Host ""
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host " Resultado final" -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor DarkCyan
    Write-Host "Servicio      : $($service.Name) [$($service.Status)]" -ForegroundColor Green
    Write-Host "Server        : $cfgServer" -ForegroundColor Green
    Write-Host "ServerActive  : $cfgServerActive" -ForegroundColor Green
    Write-Host "Hostname      : $cfgHostname" -ForegroundColor Green
    Write-Host "Timeout       : $cfgTimeout" -ForegroundColor Green
    Write-Host "LogFile       : $cfgLogFile" -ForegroundColor Green
    Write-Host "Config        : $ConfigPath" -ForegroundColor Green
    Write-Host "Transcript    : $TranscriptLog" -ForegroundColor Green
    Write-Host "MSI log       : $MsiLog" -ForegroundColor Green
    Write-Host ""

    Write-Status "Instalación finalizada correctamente" "OK"
    exit 0
}
catch {
    Write-Status $_.Exception.Message "ERROR"
    Write-Status "La instalación finalizó con errores. Revisar transcript: $TranscriptLog" "ERROR"
    Write-Status "Revisar también MSI log: $MsiLog" "ERROR"
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}
