---
name: extraer-power-bi
description: Cómo extraer las tablas de un archivo Power BI (.pbix) a CSV para analizarlas en Excel - instalar Power BI Desktop, conectarse a su motor local y exportar. Usar cuando parte de la data de un proyecto viene en un .pbix o cuando un análisis parece "bloqueado" por un archivo de Power BI.
---

# Extraer datos de un .pbix

## Por qué importa
Un `.pbix` guarda los datos comprimidos dentro de un modelo (modo importación). No se lee como Excel, pero **los datos sí están en el archivo**. Tratarlo como bloqueo y pedirle un export al cliente suele ser innecesario.

## Paso a paso (Windows)
1. **Instala Power BI Desktop** (gratuito), con autorización del usuario: `winget install --id Microsoft.PowerBI -e`. Pide permisos de administrador.
2. **Ejecuta `scripts/pbix_export.ps1 -Pbix <ruta> -Out <carpeta> -ListOnly`**: abre el archivo, espera a que levante el motor local (`msmdsrv`), se conecta con la librería AdomdClient que trae Power BI y lista las tablas con su número de filas.
3. Revisa la lista: si las tablas tienen filas, los datos están adentro. Si al abrir pide credenciales o las tablas están vacías, el modelo es DirectQuery y ahí sí hace falta el cliente.
4. **Ejecuta sin `-ListOnly`** para exportar todas las tablas a CSV (separador `;`, UTF-8). Unas 200 mil filas toman menos de un minuto.
5. Cierra Power BI Desktop sin guardar.

## Detalles técnicos
- El puerto del motor está en `%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\…\Data\msmdsrv.port.txt` (UTF-16).
- Tras conectarte, fija la base del modelo: `GetSchemaDataSet('DBSCHEMA_CATALOGS')` y `ChangeDatabase`. El modelo puede tardar en cargar; reintenta.
- Lista tablas con `SELECT [Name] FROM $SYSTEM.TMSCHEMA_TABLES` y exporta con `EVALUATE 'Tabla'`.
- Omite las tablas automáticas `LocalDateTable_*` y `DateTableTemplate_*`.

## Después de exportar
- Perfila cada columna (valores distintos y más frecuentes) antes de analizar.
- Busca columnas de auditoría (usuario que crea, modifica o **elimina**) y excluye los registros eliminados.
- Una tabla con nombre de un año ("(2025)") puede traer varios años: revisa la columna de año.
- Cuadra los totales contra otra fuente (reportes del cliente, láminas anteriores) antes de usarlos.
