# Ruta de instalación y archivo ZIP del agente Zabbix
$zabbixFolder = "C:\zabbix"
$zabbixZipUrl = "https://cdn.zabbix.com/zabbix/binaries/stable/6.4/6.4.19/zabbix_agent-6.4.19-windows-amd64.zip"
$zabbixZipPath = "$env:TEMP\zabbix_agent.zip"
$zabbixConf = "$zabbixFolder\conf\zabbix_agentd.conf"
$logFilePath = "$zabbixFolder\zabbix_agentd.log"

# Crear carpeta Zabbix si no existe
if (-not (Test-Path -Path $zabbixFolder)) {
    New-Item -Path $zabbixFolder -ItemType Directory -Force
    Write-Host "Carpeta $zabbixFolder creada." -ForegroundColor Green
}
else {
    Write-Host "La carpeta $zabbixFolder ya existe." -ForegroundColor Cyan
}

# Descargar el archivo ZIP del agente
Write-Host "Descargando Zabbix Agent desde $zabbixZipUrl..." -ForegroundColor Yellow
try {
    Invoke-WebRequest -Uri $zabbixZipUrl -OutFile $zabbixZipPath -ErrorAction Stop
    Write-Host "Descarga completada: $zabbixZipPath" -ForegroundColor Green
}
catch {
    Write-Error "Error al descargar el agente Zabbix: $_" 
    
}

# Extraer el contenido del ZIP a la carpeta de Zabbix
Write-Host "Extrayendo Zabbix Agent en $zabbixFolder..." -ForegroundColor Yellow
try {
    Expand-Archive -Path $zabbixZipPath -DestinationPath $zabbixFolder -Force
    Write-Host "Extracción completada." -ForegroundColor Green
}
catch {
    Write-Error "Error al extraer el agente Zabbix: $_" 
    
}

# Verificar si el archivo de configuración existe
if (-not (Test-Path -Path $zabbixConf)) {
    Write-Error "El archivo de configuración no se encontró: $zabbixConf" 
    
}

# Solicitar la IP del Servidor Zabbix
$ServerIp = Read-Host "Ingrese la IP del servidor Zabbix"

# Editar el archivo de configuración para agregar el Server y LogFile
Write-Host "Editando el archivo de configuración..." -ForegroundColor Yellow
try {
    # Reemplazar la línea de LogFile
    (Get-Content -Path $zabbixConf) -replace "^LogFile=.*", "LogFile=$logFilePath" | Set-Content -Path $zabbixConf
    
    # Reemplazar la línea de Server
    (Get-Content -Path $zabbixConf) -replace "^Server=.*", "Server=$ServerIp" | Set-Content -Path $zabbixConf
    
    Write-Host "Archivo de configuración actualizado." -ForegroundColor Green
}
catch {
    Write-Error "Error al editar el archivo de configuración: $_" 
    
}

# Instalar el agente Zabbix
Write-Host "Instalando Zabbix Agent..." -ForegroundColor Yellow
try {
    Start-Process -FilePath "$zabbixFolder\bin\zabbix_agentd.exe" -ArgumentList "--config `"$zabbixConf`" --install" -NoNewWindow -Wait
    Write-Host "Zabbix Agent instalado correctamente." -ForegroundColor Green
}
catch {
    Write-Error "Error al instalar el agente Zabbix: $_" 
    
}

# Iniciar el agente Zabbix
Write-Host "Iniciando Zabbix Agent..." -ForegroundColor Yellow
try {
    Start-Process -FilePath "$zabbixFolder\bin\zabbix_agentd.exe" -ArgumentList "--config `"$zabbixConf`" --start" -NoNewWindow -Wait
    Write-Host "Zabbix Agent iniciado correctamente." -ForegroundColor Green
}
catch {
    Write-Error "Error al iniciar el agente Zabbix: $_" 
    
}

# Configurar las reglas del firewall
Write-Host "Configurando reglas de firewall..." -ForegroundColor Yellow
$firewallRuleName = "Zabbix Agent"
$existingRule = Get-NetFirewallRule -DisplayName $firewallRuleName -ErrorAction SilentlyContinue

if ($existingRule) {
    Write-Host "La regla de firewall '$firewallRuleName' ya existe." -ForegroundColor Cyan
}
else {
    try {
        New-NetFirewallRule -DisplayName $firewallRuleName -Direction Inbound -Action Allow -Protocol TCP -LocalPort 10050 -RemoteAddress $ServerIp -Enabled True
        Write-Host "Regla de firewall creada para el puerto 10050/TCP desde $ServerIp." -ForegroundColor Green
    }
    catch {
        Write-Error "Error al crear la regla de firewall: $_" 
        
    }
}

# Crear archivo ZabbixAgent.bat para reiniciar el agente
Write-Host "Creando archivo ZabbixAgent.bat para reiniciar el agente..." -ForegroundColor Yellow
$batContent = @"
@echo off
C:\zabbix\bin\zabbix_agentd.exe --config C:\zabbix\conf\zabbix_agentd.conf --stop
timeout /t 5 /nobreak
C:\zabbix\bin\zabbix_agentd.exe --config C:\zabbix\conf\zabbix_agentd.conf --start
"@

$batFilePath = "$zabbixFolder\ZabbixAgent.bat"

try {
    Set-Content -Path $batFilePath -Value $batContent -Force
    Write-Host "Archivo ZabbixAgent.bat creado en $batFilePath." -ForegroundColor Green
}
catch {
    Write-Error "Error al crear el archivo ZabbixAgent.bat: $_" 
    
}

# Crear una tarea programada para reiniciar el agente
Write-Host "Creando tarea programada para reiniciar el agente Zabbix..." -ForegroundColor Yellow
$taskName = "Zabbix Service Restart"

$action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batFilePath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)

try {
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $taskName -Description "Reinicia el Agente Zabbix" -User "SYSTEM" -RunLevel Highest -Force
    Write-Host "Tarea programada '$taskName' creada exitosamente." -ForegroundColor Green
} catch {
    Write-Error "Error al crear la tarea programada: $_" 
}


Write-Host "Instalación y configuración completa." -ForegroundColor Green
