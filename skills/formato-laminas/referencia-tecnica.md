# Referencia técnica: PowerPoint y Excel por COM desde PowerShell 5.1

Problemas reales encontrados y su solución.

## Configuración regional (es-CL) y COM
- **Tamaños de fuente decimales fallan** ("La conversión especificada no es válida") cuando la cultura del hilo es es-CL. Asigna con `InvokeMember` y `CultureInfo.InvariantCulture` (función `SetP` en `pptlib.ps1`), no con `$obj.Size = 11.5`. Cambiar `CurrentCulture` no sirve: PowerShell 5.1 la restablece en cada instrucción.
- **Formatos numéricos de etiquetas de gráfico** se interpretan según la cultura de la llamada: `#,##0` puede salir como `26539,0`. Solución robusta: escribe el texto de cada etiqueta ya formateado (`FmtNum` en `pptlib.ps1`).
- **Excel `NumberFormat`** también se interpreta localmente en algunos casos; usa `NumberFormatLocal` con códigos locales (`#.##0`, `0,0%`).
- Alturas de fila y anchos de columna decimales: igual que las fuentes, usa `SetP`.

## Gráficos
- `ChartData.Activate()` abre el Excel embebido de forma asíncrona: reintenta hasta que `Workbook` responda y **verifica** que la data quedó escrita (nombre de la serie en la celda B1) antes de seguir.
- La hoja embebida en Office en español se llama `Hoja1`, no `Sheet1`: usa `$ws.Name` en `SetSourceData`.
- Categorías como "2019" se convierten en número y el gráfico las toma como serie: formatea la columna A y la fila 1 como texto (`@`) antes de escribir.
- Gráficos de una serie muestran el nombre como título automático: deja el nombre en blanco (`' '`) y elimina el título al final.
- El refresco del ChartData puede pisar formatos: aplica etiquetas en una segunda pasada, después de una pausa.
- Cascadas: barras apiladas con una serie "base" invisible (`color = -1`).
- Etiquetas en barras horizontales con negativos: `TickLabelPosition = Low`.

## Estabilidad
- PowerPoint rechaza llamadas mientras recalcula (errores intermitentes E_FAIL u "Object required"): envuelve las operaciones en reintentos (`Rt`) y reintenta la lámina o el gráfico completo si falla.
- Corre cada etapa del deck en un proceso de PowerShell separado y guarda un checkpoint `.pptx` al terminar cada una.
- Cierra las instancias de Excel de automatización que queden huérfanas (`CommandLine` contiene `/automation` o `-Embedding`). Nunca cierres un Excel que el usuario abrió.
- Abrir PDFs con Word COM puede colgarse por un diálogo invisible de conversión: lánzalo en segundo plano con timeout.

## Scripts y codificación
- Guarda los `.ps1` en **UTF-8 con BOM** si tienen tildes; sin BOM, PowerShell 5.1 los lee como ANSI y las comparaciones de texto ("Mecánica/Eléctrica") fallan en silencio.
- PowerShell no distingue mayúsculas: `$L` y `$l` son la misma variable (un ciclo `foreach($l …)` pisa un rango `$L`).
- `?` es válido en nombres de variable: `"/sheets/$SID?include=…"` se rompe; usa `"/sheets/${SID}?include=…"`.
- Una función que devuelve un `DataReader` debe devolverlo con coma (`return ,$rd`) para que PowerShell no lo enumere.
