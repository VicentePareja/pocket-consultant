---
name: plan-de-trabajo-smartsheet
description: Cómo crear y mantener un plan de trabajo en Smartsheet con tareas bien definidas, estados honestos y respaldo trazable, y cómo editarlo por API sin contactar al equipo. Usar al armar o actualizar un plan de trabajo (PDT), al preparar el estado para un check-in o al leer o escribir una planilla de Smartsheet.
---

# Plan de trabajo en Smartsheet

## REGLA: Claude nunca contacta al equipo
**Prohibido** etiquetar personas (@menciones), publicar comentarios dirigidos a alguien, responder hilos del equipo, asignar responsables de forma que disparen notificaciones o enviar cualquier mensaje por Smartsheet, correo o chat.

En su lugar, Claude:
1. Redacta el mensaje en el chat.
2. Dice **a quién** va y **dónde** conviene publicarlo (fila, hilo, canal).
3. Le pide al usuario que lo envíe él, manualmente.

Así el usuario controla qué recibe el equipo y cuándo. Esto vale aunque haya acceso por API.

## Una tarea bien definida
| Campo | Estándar |
|---|---|
| Nombre | Verbo + objeto + alcance ("Pareto de fuentes de pérdida por línea") |
| Comentarios / descripción | Qué se espera, fuente de datos, referencia (lámina, archivo) |
| Estado | Solo valores del desplegable, exactamente como están escritos ("En Curso", no "En curso") |
| Fechas | Inicio y fin realistas; nada con fechas absurdas por errores de formato |
| Responsable | Quien realmente mueve la tarea; si depende del senior, el senior también |
| Actualización | Una línea con fecha: qué se hizo, el número clave y dónde está el respaldo ("30-sep: … Respaldo: Excel > hoja") |

## Armar un plan nuevo
- **Fases que terminan en un check-in con una decisión** ("decidir alcance del MVP", "elegir plataforma"), no en una reunión de avance genérica.
- **Línea base temprano:** si el proyecto promete ahorro o mejora, la primera fase mide el punto de partida (horas, costo) para poder demostrarlo al cierre.
- Lo pospuesto va en una fase **Backlog** sin fechas, para que no se pierda ni ensucie el Gantt.
- Supuestos (quién trabaja, fechas comprometidas, feriados) y preguntas al supervisor se entregan junto con el plan, con una recomendación por pregunta.

## Estados honestos
- **Completo** solo si todo lo que pide la tarea está hecho. Si falta una parte, es "En Curso" con la nota de qué falta.
- **Bloqueado** solo si el bloqueo es real (ver skill `rigor-analitico`, "¿Es realmente un bloqueo?").
- Nunca marcar "Completo" algo que no se hizo para que el plan se vea bien: el senior lo detecta y se pierde confianza.
- Mantén el plan dentro del alcance del frente propio; lo que es de otro frente o de la síntesis se reasigna.

## Adjuntos
Cada entregable se adjunta en su fila. Cuando el archivo cambia, se sube la nueva versión a la misma fila para que nadie revise una versión antigua.

## Trabajar por API
Usa `scripts/smartsheet.ps1`:
- La clave `SMARTSHEET_API_KEY` se lee desde la variable de entorno o desde `~\.pocket-consultant\.env` (skill `organizar-proyecto`). **Nunca** se imprime ni se copia a un proyecto o al repositorio.
- **Respalda antes de escribir:** guarda el JSON de la planilla antes de cualquier cambio.
- Busca la planilla por nombre en `GET /sheets`; los links `app.smartsheet.com/sheets/<token>` no son el id numérico.
- Lee `columns` (tipos y opciones del desplegable) antes de escribir estados.
- En la API, "Row N" de Smartsheet es el `rowNumber`; en un export a Excel es la fila N+1 (por el encabezado).
- Para columnas de contacto múltiple: tipo `MULTI_CONTACT_LIST`, valor con `objectValue` de tipo `MULTI_CONTACT`.
- Los comentarios con tildes y símbolos se publican de forma confiable con `curl.exe` y JSON compacto en UTF-8 sin BOM (Invoke-RestMethod dio error 400).
- Verifica leyendo la planilla de vuelta después de escribir. Las celdas de fecha no traen `displayValue`: lee `value` (`AAAA-MM-DD`).
- Al crear una planilla, `column.validation` no se acepta (error 1032); restringir el desplegable se activa después en la columna.
- Guarda el id y el link de la planilla en la ficha del proyecto (`README.md`).

### Crear un plan nuevo
`New-PlanSheet '<nombre>' '<ruta>.json'` crea la planilla con las columnas estándar (Nombre Tarea, Comentarios, Estado, Fecha de inicio, Fecha de Fin, Responsable) y cada tarea como fila hija de su fase. Responsable va como texto para no notificar a nadie. El JSON:
```json
[
  { "fase": "FASE 0 · Entender el proceso actual",
    "tareas": [
      { "nombre": "Mapear el proceso actual", "comentarios": "Qué se espera, fuente y output",
        "estado": "Pendiente", "inicio": "2026-10-06", "fin": "2026-10-08", "responsable": "Vicente P." }
    ] }
]
```
Estados válidos: Completo, Pendiente, Cancelado, En Curso, Bloqueado. Un backlog es una fase cuyas tareas no tienen fechas.

## Sin API
Si no hay acceso, pide al usuario exportar la planilla a Excel, actualízala en una **copia** en `03 Trabajo` del proyecto (no sobre su original) y entrega el paso a paso para pegarla: bloques de columnas, adjuntos por fila y revisión final por estado.
