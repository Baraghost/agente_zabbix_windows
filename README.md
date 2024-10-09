## Descripción del Script

El script `instalar_agente_zabbix.ps1` automatiza la instalación y configuración del agente de Zabbix en servidores Windows. Para mayor flexibilidad y facilidad de uso, permite a los usuarios especificar la dirección IP del servidor Zabbix.

#### Comando de Instalación en 1 linea:

```powershell
Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/Baraghost/agente_zabbix_windows/main/instalar_agente_zabbix.ps1" -UseBasicParsing).Content
```

**Importante:** Asegúrate de colocar la IP del servidor de Zabbix al momento que te lo solicite el script.
