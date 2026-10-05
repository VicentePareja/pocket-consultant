---
name: formato-laminas
description: Estándar para crear o revisar láminas de consultoría en PowerPoint (títulos de acción, estructura, notas al pie con fuente y respaldo, gráficos, densidad) y cómo construirlas y verificarlas con PowerShell + PowerPoint COM. Usar al armar un deck, editar láminas o revisar la forma de una presentación.
---

# Formato de láminas

## La lámina se entiende leyendo solo el título
- **Título de acción, no descriptivo:** una oración con el hallazgo y su número. "Los cambios de formato son 47% de las horas perdidas" y no "Análisis de pérdidas".
- **Máximo dos líneas** (unos 130 caracteres a 22 pt en 16:9). Si no cabe, se reescribe; nunca se achica la letra ni se deja en tres líneas.
- Un número clave en el título, no cinco.
- Evita cortar cifras entre líneas ("$691–" / "815 MM"): reformula ("entre $691 y $815 MM").

## Estructura de cada lámina
1. **Pestaña de sección** a la izquierda: el capítulo al que pertenece.
2. **Título de acción.**
3. **Subtítulos de bloque** (izquierda y derecha) con su **unidad** en cursiva debajo ("Millones de pesos, enero–junio").
4. **Un gráfico o tabla que prueba el título** y, al lado, una caja de "Implicancia" o "Qué nos dice" con 2 a 4 bullets.
5. **Nota al pie en tres partes:** `Notas:` (definiciones, periodo, exclusiones) · `Fuente:` (archivo original del cliente) · `Respaldo:` (Excel > hoja > celda).
6. **"PRELIMINAR"** en la esquina mientras las cifras no estén validadas con el cliente.

## Densidad
- **Deck de respaldo** (técnico, para el equipo o como anexo): puede tener tablas completas, conciliaciones y limitaciones.
- **Deck ejecutivo** (comité): una idea y un gráfico por lámina, lenguaje de negocio, máximo 3 bullets. El detalle vive en el respaldo.
- Si una caja queda con mucho espacio en blanco, se ajusta su alto; si queda llena, se recorta el texto, no la letra.

## Gráficos
- Nativos de PowerPoint (editables), no imágenes.
- **Cascada (waterfall)** para explicar cómo se va de un total a otro (margen, horas, ahorro por palanca).
- **Barras horizontales** para rankings con nombres largos; ordenadas de mayor a menor.
- **100% apiladas** para mezcla (composición del costo, SKUs vs venta).
- **Líneas** solo para series de tiempo.
- Etiquetas de datos con formato local (es-CL: `1.234` y `12,5%`), sin ejes innecesarios ni títulos automáticos.
- Colores con significado: rojo para pérdida o caída, verde para ahorro o alza, azul corporativo para el dato principal.

## Verificación antes de entregar (obligatoria)
- Exporta cada lámina a PNG (`scripts/export_png.ps1`) y **mírala**: títulos en dos líneas, cajas que no se pisan, etiquetas legibles, números con formato correcto.
- Revisa que el total de la lámina de resumen calce con la suma de las palancas.
- Busca en todo el deck cualquier cifra que cambió (por ejemplo, un porcentaje recalculado) y actualízala en todas sus apariciones: resumen ejecutivo, síntesis, anexos, guion.
- **Cero placeholders vacíos.** Un placeholder sin texto muestra "Haga clic para agregar título" o "Inserte título" al abrir el archivo, y el PNG exportado **no lo muestra**, así que mirar las imágenes no basta. Corre `scripts/qa_pptx.ps1 -Path <carpeta>` antes de entregar (con `-Fix` borra los vacíos) y termina cada script de construcción con `Remove-EmptyPlaceholders $pres`.
- La versión que se envía se guarda en `04 Entregables` como `AAAAMMDD - Nombre vN.pptx` y no se sobrescribe; los checkpoints de construcción quedan en `03 Trabajo`.
- En un deck para el cliente no quedan marcas internas: franjas "A validar", "PRELIMINAR" de trabajo, referencias a hojas del Excel de respaldo.

## Construcción con PowerShell + PowerPoint COM
Ver `referencia-tecnica.md` para los detalles. Lo esencial:
- Trabaja siempre sobre una **copia** de la plantilla corporativa y construye por etapas con **puntos de guardado**: si una etapa falla, se retoma desde la última copia buena en vez de reconstruir todo.
- Usa `scripts/pptlib.ps1` (títulos, tablas, gráficos, callouts, pestañas de sección).
- Abre la presentación **con ventana** (algunas operaciones de gráficos fallan sin ella).

## Lecciones
- Rehacer un deck completo por un error puntual hace perder tiempo: parchea la lámina afectada o retoma desde el checkpoint.
- Un título de tres líneas pisa el subtítulo: se detecta solo mirando el PNG.
- El número de la lámina ejecutiva es un redondeo del respaldo; dilo explícitamente para que nadie lo lea como error.
- En PowerShell, un texto entre comillas dobles con montos (`"Rango de $320 MM"`) pierde el monto porque `$320` se lee como variable. Usa comillas simples o escribe `` `$320 ``.
- Duplicar una lámina y borrarle subtítulos puede dejar placeholders vacíos que no se ven en el PNG: por eso el escaneo es obligatorio.
