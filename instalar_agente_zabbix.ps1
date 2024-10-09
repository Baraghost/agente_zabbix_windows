param (
    [Parameter(Mandatory = $true)]
    [string]$ServerIp
)

# Ruta de instalación y archivo ZIP del agente Zabbix
$zabbixFolder = "C:\zabbix"
$zabbixZipUrl = "https://cdn.zabbix.com/zabbix/binaries/stable/6.4/6.4.19/zabbix_agent-6.4.19-windows-amd64.zip"
$zabbixZipPath = "$env:TEMP\zabbix_agent.zip"
$zabbixConf = "$zabbixFolder\conf\zabbix_agentd.conf"
$logFilePath = "$zabbixFolder\zabbix_agentd.log"

# Crear carpeta Zabbix si no existe
if (-not (Test-Path -Path $zabbixFolder)) {
    New-Item -Path $zabbixFolder -ItemType Directory -Force
    Write-Host "Carpeta $zabbixFolder creada."
} else {
    Write-Host "La carpeta $zabbixFolder ya existe."
}

# Descargar el archivo ZIP del agente
Write-Host "Descargando Zabbix Agent desde $zabbixZipUrl..."
try {
    Invoke-WebRequest -Uri $zabbixZipUrl -OutFile $zabbixZipPath -ErrorAction Stop
    Write-Host "Descarga completada: $zabbixZipPath"
} catch {
    Write-Error "Error al descargar el agente Zabbix: $_"
    exit 1
}

# Extraer el contenido del ZIP a la carpeta de Zabbix
Write-Host "Extrayendo Zabbix Agent en $zabbixFolder..."
try {
    Expand-Archive -Path $zabbixZipPath -DestinationPath $zabbixFolder -Force
    Write-Host "Extracción completada."
} catch {
    Write-Error "Error al extraer el agente Zabbix: $_"
    exit 1
}

# Verificar si el archivo de configuración existe
if (-not (Test-Path -Path $zabbixConf)) {
    Write-Error "El archivo de configuración no se encontró: $zabbixConf"
    exit 1
}

# Editar el archivo de configuración para agregar el Server y LogFile
Write-Host "Editando el archivo de configuración..."
try {
    # Reemplazar la línea de LogFile
    (Get-Content -Path $zabbixConf) -replace "^LogFile=.*", "LogFile=$logFilePath" | Set-Content -Path $zabbixConf
    
    # Reemplazar la línea de Server
    (Get-Content -Path $zabbixConf) -replace "^Server=.*", "Server=$ServerIp" | Set-Content -Path $zabbixConf
    
    Write-Host "Archivo de configuración actualizado."
} catch {
    Write-Error "Error al editar el archivo de configuración: $_"
    exit 1
}

# Instalar el agente Zabbix
Write-Host "Instalando Zabbix Agent..."
try {
    Start-Process -FilePath "$zabbixFolder\bin\zabbix_agentd.exe" -ArgumentList "--config `"$zabbixConf`" --install" -NoNewWindow -Wait
    Write-Host "Zabbix Agent instalado correctamente."
} catch {
    Write-Error "Error al instalar el agente Zabbix: $_"
    exit 1
}

# Iniciar el agente Zabbix
Write-Host "Iniciando Zabbix Agent..."
try {
    Start-Process -FilePath "$zabbixFolder\bin\zabbix_agentd.exe" -ArgumentList "--config `"$zabbixConf`" --start" -NoNewWindow -Wait
    Write-Host "Zabbix Agent iniciado correctamente."
} catch {
    Write-Error "Error al iniciar el agente Zabbix: $_"
    exit 1
}

# Configurar las reglas del firewall
Write-Host "Configurando reglas de firewall..."
$firewallRuleName = "Zabbix Agent"
$existingRule = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue

if ($existingRule) {
    Write-Host "La regla de firewall '$firewallRuleName' ya existe."
} else {
    try {
        New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort 10050 -RemoteAddress $ServerIp -Enabled True
        Write-Host "Regla de firewall creada para el puerto 10050/TCP desde $ServerIp."
    } catch {
        Write-Error "Error al crear la regla de firewall: $_"
        exit 1
    }
}

Write-Host "Instalación y configuración completa."
