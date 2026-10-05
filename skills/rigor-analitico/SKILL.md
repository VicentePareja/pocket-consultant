---
name: rigor-analitico
description: Estándar de rigor para construir y defender números de consultoría - reconstruir desde la data cruda, Excel de respaldo trazable, conciliación con material previo, supuestos explícitos y verificación de bloqueos. Usar al calcular cualquier cifra que irá a una lámina, al revisar el trabajo de un analista o cuando un número "no calza".
---

# Rigor analítico

## Principio
Cada número presentado sale de la **data cruda del cliente** y se puede reproducir desde un **Excel de respaldo** hasta la celda. Las láminas anteriores (talleres, propuestas) se usan para **conciliar**, nunca como fuente.

## El Excel de respaldo
- Vive en `03 Trabajo` del proyecto y lee la data cruda desde `01 Insumos`, que no se edita (skill `organizar-proyecto`).
- Hoja `README`: propósito, lista de hojas, fuente exacta de cada una (archivo > hoja > columnas > filtros), supuestos, y tabla **"número presentado → hoja!celda"**.
- Hojas `DATA_*`: extracto de la data usada (o agregado documentado si es muy grande).
- Hojas `A1_…`, `A2_…`: un análisis por hoja, con **fórmulas** (SUMIFS, COUNTIFS) sobre las `DATA_*`, no valores pegados.
- **Supuestos en celdas de color** (azul o celeste) para que se vean y se puedan cambiar.
- Revisa que no haya errores de fórmula antes de entregar.

## Conciliar con el material previo
- Para cada cifra del material anterior: ¿la reproduzco? Si sí, dilo ("reproducido"). Si no, explica por qué difiere: otra base, otro periodo, un error.
- Los errores del material previo se documentan en una tabla de conciliación (número, versión anterior, versión reconstruida, por qué). Cómo se comunican al cliente lo decide el senior.
- Lo que no se puede reproducir se marca **"no reproducido"** en la nota al pie de la lámina.

## Trampas que ya encontramos
- **Denominador del indicador:** un OEE sobre horas trabajadas y otro sobre horas calendario pueden diferir 3 veces. Explicita siempre la base.
- **Unidades mezcladas:** litros vs unidades, miles vs millones, con o sin IVA. Un should cost que compara precio con IVA contra costo sin IVA infla la brecha.
- **Registros eliminados o anulados** en bases transaccionales (columna de "usuario que elimina", notas de crédito): exclúyelos y dilo.
- **Cambios de perímetro:** una reorganización que mueve personas entre áreas crea "alzas" falsas; compara por persona o con perímetro constante.
- **Periodos no comparables:** acumulado a septiembre vs año completo. Anualiza con criterio explícito (× 12/9) o compara el mismo periodo.
- **Codificación de texto:** una comparación con tildes que falla en silencio deja fuera registros. Cuadra los totales contra otra fuente.
- **No sumar peras con manzanas:** separa ahorro recurrente ($/año), caja por única vez, riesgos y oportunidades con inversión. Evita el doble conteo entre frentes.
- **Redistribución entre unidades:** antes de decir que una unidad "crece" (una línea, una planta), revisa el total. A veces solo absorbió lo que otra dejó de hacer.
- **Efecto mezcla:** un promedio por evento puede subir solo porque cambió la mezcla de tipos. Compara el mismo tipo de evento entre años.
- **Valores estándar en registros manuales:** si muchas duraciones se repiten exactas (13 o 30 minutos), parte son tiempos estándar y no mediciones. Dilo en la nota y propone medir antes de fijar metas.
- **Cifras parecidas con distinta base:** "la planta" y "sus líneas" no son lo mismo. Define el perímetro y usa uno solo en la lámina.

## Cuantificar
- Todo hallazgo importante termina en horas, unidades o $/año, con **rango** (mínimo–máximo) y **factibilidad**.
- Un techo teórico se presenta como techo, con la captura prudente al lado.
- Un escenario se verifica **unidad por unidad**, no solo en el total. Si una unidad no cumple la condición, agrega la palanca que falta (por ejemplo, mover volumen a otra línea que ya ha hecho esos productos) y muéstralo.

## ¿Es realmente un bloqueo?
Antes de marcar algo como "bloqueado" o escalarlo al senior:
1. ¿El dato existe en lo que ya tenemos? (por ejemplo, dentro de un .pbix).
2. ¿Lo que falta es una herramienta que puedo conseguir? (Power BI Desktop es gratuito; ver skill `extraer-power-bi`).
3. Solo si el dato no está o requiere acceso del cliente, es un bloqueo real, y se pide con el dato exacto requerido.

Marcar como bloqueado algo que el analista puede resolver solo se ve débil frente al senior.

## Verificación final
- Los totales cuadran entre láminas (resumen = suma de palancas).
- Un número que cambia se actualiza en **todas** sus apariciones (deck de respaldo, deck ejecutivo, guion, plan de trabajo).
- Relee cada cifra del resumen ejecutivo contra su celda de respaldo.
