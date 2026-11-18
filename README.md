# Ejemplos de automatizacion con PowerShell

Coleccion de scripts para administrar y consultar recursos de Microsoft Fabric, Power BI, Azure, Microsoft 365 y SharePoint. Los ejemplos estan pensados como punto de partida y pueden requerir ajustes segun el tenant, suscripcion y permisos disponibles.

## Requisitos

- PowerShell 7 recomendado.
- Modulos segun el script: `MicrosoftPowerBIMgmt`, `Az`, `PnP.PowerShell`, `Microsoft.Graph`, `Az.DataFactory` y `DatabricksPS`.
- Sesion autenticada en cada servicio antes de ejecutar el script.
- Permisos suficientes sobre las areas de Fabric/Power BI, Azure, Microsoft 365, SharePoint o Azure DevOps que se consulten.

```powershell
Install-Module MicrosoftPowerBIMgmt, Az, PnP.PowerShell, Microsoft.Graph -Scope CurrentUser
```

Instale tambien los modulos especificos que indique el script antes de ejecutarlo.

## Uso

Revise las variables y rutas configuradas dentro de cada archivo antes de ejecutar. Algunos scripts usan exportacion a CSV o PDF y necesitan que la ruta de destino exista.

```powershell
Set-Location .\scripts
.\Nombre-del-script.ps1
```

> [!WARNING]
> Algunos scripts contienen identificadores de tenant, suscripcion, usuarios o valores de prueba. Sustituya esas referencias por variables de entorno o por un administrador de secretos. No confirme credenciales, secretos, tokens ni contrasenas en el repositorio. Consulte un administrador antes de ejecutar operaciones que puedan crear, modificar o eliminar recursos.

## Catalogo de scripts

| Script | Funcion |
|---|---|
| [Case 3 LastRefresh.ps1](./scripts/Case%203%20LastRefresh.ps1) | Exporta la fecha y el estado de la ultima actualizacion de los conjuntos de datos. |
| [Case 3 RefreshSchedule.ps1](./scripts/Case%203%20RefreshSchedule.ps1) | Consulta y exporta los horarios de actualizacion configurados. |
| [Case 3 Refreshes.ps1](./scripts/Case%203%20Refreshes.ps1) | Obtiene el historial de actualizaciones de Power BI. |
| [Case 4 Invoke Workspaces.ps1](./scripts/Case%204%20Invoke%20Workspaces.ps1) | Lista workspaces, datasets y sus objetos mediante la API de Power BI. |
| [Case 4.1 Invoke Workspaces Reports Users.ps1](./scripts/Case%204.1%20Invoke%20Workspaces%20Reports%20Users.ps1) | Relaciona workspaces, informes y usuarios con acceso. |
| [Case 4.2 Invoke Reports Subscriptions.ps1](./scripts/Case%204.2%20Invoke%20Reports%20Subscriptions.ps1) | Lista suscripciones asociadas a informes de Power BI. |
| [Case 5 DataSource.ps1](./scripts/Case%205%20DataSource.ps1) | Inventaria fuentes de datos de los modelos de Power BI. |
| [Case 5 DataSource V2.ps1](./scripts/Case%205%20DataSource%20V2.ps1) | Inventario ampliado de fuentes de datos de Power BI. |
| [Case 6 DataGateWay.ps1](./scripts/Case%206%20DataGateWay.ps1) | Lista gateways de datos y exporta su informacion. |
| [Case 7 ADUser.ps1](./scripts/Case%207%20ADUser.ps1) | Consulta usuarios de Microsoft Entra ID y exporta sus datos. |
| [Case 8 Exportando (Ejemplo).ps1](./scripts/Case%208%20Exportando%20(Ejemplo).ps1) | Ejemplo de exportacion de objetos de Power BI a archivos. |
| [Case 9 Send Email (Ejemplo).ps1](./scripts/Case%209%20Send%20Email%20(Ejemplo).ps1) | Ejemplo basico para enviar correo electronico con PowerShell. |
| [Case 10 Export and Send v1.ps1](./scripts/Case%2010%20Export%20and%20Send%20v1.ps1) | Exporta un informe de Power BI y lo envia por correo, primera version. |
| [Case 10 Export and Send v2.ps1](./scripts/Case%2010%20Export%20and%20Send%20v2.ps1) | Exporta un informe de Power BI y lo envia por correo, segunda version. |
| [Case 10 Export and Send v3.ps1](./scripts/Case%2010%20Export%20and%20Send%20v3.ps1) | Exporta un informe de Power BI y lo envia por correo, tercera version. |
| [Case 10 Export and Send v4.ps1](./scripts/Case%2010%20Export%20and%20Send%20v4.ps1) | Exporta un informe de Power BI y lo envia por correo, cuarta version. |
| [Case 11 Pnp password and login.ps1](./scripts/Case%2011%20Pnp%20password%20and%20login.ps1) | Gestiona credenciales almacenadas y autenticacion con PnP. |
| [Case 12 Pbix.ps1](./scripts/Case%2012%20Pbix.ps1) | Automatiza operaciones de descarga o distribucion de informes PBIX. |
| [Case 13 Fabric Lakehouses.ps1](./scripts/Case%2013%20Fabric%20Lakehouses.ps1) | Consulta lakehouses de Microsoft Fabric. |
| [Case 14 Fabric Connections.ps1](./scripts/Case%2014%20Fabric%20Connections.ps1) | Inventaria conexiones y configuraciones de Fabric. |
| [Case 15 Update Pipeline.ps1](./scripts/Case%2015%20Update%20Pipeline.ps1) | Actualiza la definicion de una pipeline de Azure Data Factory. |
| [Case 16 Update All Pipelines.ps1](./scripts/Case%2016%20Update%20All%20Pipelines.ps1) | Aplica cambios masivos a pipelines de Azure Data Factory. |
| [Case 17 List All Pipelines.ps1](./scripts/Case%2017%20List%20All%20Pipelines.ps1) | Lista las pipelines de Azure Data Factory. |
| [Case 18 List DevOps.ps1](./scripts/Case%2018%20List%20DevOps.ps1) | Consulta repositorios y pipelines de Azure DevOps. |
| [Case 19 List Pipeline Trigger.ps1](./scripts/Case%2019%20List%20Pipeline%20Trigger.ps1) | Lista triggers configurados en pipelines de Azure Data Factory. |
| [Case 20 List Pipeline Activities V2.ps1](./scripts/Case%2020%20List%20Pipeline%20Activities%20V2.ps1) | Obtiene las actividades de ejecuciones de pipelines. |
| [Case 21 List Error Activities.ps1](./scripts/Case%2021%20List%20Error%20Activities.ps1) | Busca actividades fallidas y exporta sus errores. |
| [Case 22 AssignToCapacity.ps1](./scripts/Case%2022%20AssignToCapacity.ps1) | Asigna workspaces o recursos a una capacidad de Fabric. |
| [Case 23 Get Azure Metric Alerts.ps1](./scripts/Case%2023%20Get%20Azure%20Metric%20Alerts.ps1) | Lista alertas de metricas configuradas en Azure. |
| [Case 24 Create Azure Metric Alerts - Pipeline.ps1](./scripts/Case%2024%20Create%20Azure%20Metric%20Alerts%20-%20Pipeline.ps1) | Crea alertas de metricas para una pipeline de Azure Data Factory. |
| [Case 25 Create Azure Metric Alerts - Activity.ps1](./scripts/Case%2025%20Create%20Azure%20Metric%20Alerts%20-%20Activity.ps1) | Crea alertas de metricas para actividades de Data Factory. |
| [Case 26 Create Azure Action Group.ps1](./scripts/Case%2026%20Create%20Azure%20Action%20Group.ps1) | Crea un grupo de acciones para alertas de Azure Monitor. |
| [Case 27 List Azure Resources.ps1](./scripts/Case%2027%20List%20Azure%20Resources.ps1) | Inventaria recursos disponibles en Azure. |
| [Case 28 Service Principal.ps1](./scripts/Case%2028%20Service%20Principal.ps1) | Demuestra autenticacion y uso de un service principal. |
| [Case 29 List Vaults And Secrets.ps1](./scripts/Case%2029%20List%20Vaults%20And%20Secrets.ps1) | Lista Key Vaults y nombres de secretos sin exportar sus valores. |
| [Case 30 Run ADF Pipeline.ps1](./scripts/Case%2030%20Run%20ADF%20Pipeline.ps1) | Ejecuta una pipeline de Azure Data Factory mediante API REST. |
| [Case 31 List Databricks Clusters.ps1](./scripts/Case%2031%20List%20Databricks%20Clusters.ps1) | Lista clusters y detalles disponibles en Azure Databricks. |
| [Case 32 Azure Role Assigments.ps1](./scripts/Case%2032%20Azure%20Role%20Assigments.ps1) | Consulta asignaciones de roles RBAC en Azure. |
| [Case 33 Azure Add Storage Account Network Rule.ps1](./scripts/Case%2033%20Azure%20Add%20Storage%20Account%20Network%20Rule.ps1) | Agrega una regla de red a una cuenta de Azure Storage. |
| [Case 34 Azure Group and Members AD.ps1](./scripts/Case%2034%20Azure%20Group%20and%20Members%20AD.ps1) | Exporta grupos de Microsoft Entra ID y sus miembros. |
| [Case 35 Blob Storage Detailed Settings.ps1](./scripts/Case%2035%20Blob%20Storage%20Detailed%20Settings.ps1) | Obtiene una configuracion detallada de cuentas y Blob Storage. |
| [Case 36 Invoke Reports Subscriptions.ps1](./scripts/Case%2036%20Invoke%20Reports%20Subscriptions.ps1) | Consulta suscripciones de informes y sus destinatarios. |
| [Case 37 Grant Access PowerBI To Sharepoint.ps1](./scripts/Case%2037%20Grant%20Access%20PowerBI%20To%20Sharepoint.ps1) | Otorga permisos o integra acceso entre Power BI y SharePoint mediante Microsoft Graph. |

## Contribuciones

Los cambios deben conservar rutas relativas, evitar datos reales y documentar los modulos y permisos necesarios para cada script.
