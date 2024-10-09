## Descripción del Script

El script `instalar_agente_zabbix.ps1` automatiza la instalación y configuración del agente de Zabbix en servidores Windows. Acepta parámetros de línea de comandos para mayor flexibilidad y facilidad de uso, permitiendo a los usuarios especificar la dirección IP del servidor Zabbix y acceder a una ayuda integrada.

#### Comando de Instalación en 1 linea:

```powershell
Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Baraghost/agente_zabbix_windows/refs/heads/main/instalar_agente_zabbix.ps1" -UseBasicParsing).Content -ServerIp "192.168.1.100"
```

**Importante:** Asegúrate de reemplazar la ip por la de tu Server de Zabbix.
