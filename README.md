# pocket-consultant

Plugin de Claude Code con los aprendizajes de trabajo de consultoría: cómo armar láminas, cómo sostener cada número, cómo contar la historia a un comité, cómo llevar un plan de trabajo en Smartsheet y cómo preparar un check-in con el senior.

La idea es mejorarlo progresivamente: cada proyecto deja lecciones que se agregan como skills, referencias o scripts.

## Instalación

```
/plugin marketplace add VicentePareja/pocket-consultant
/plugin install pocket-consultant@pocket-consultant
```

## Skills

| Skill | Cuándo se usa |
|---|---|
| `formato-laminas` | Al crear o revisar láminas (PowerPoint): títulos de acción, estructura, notas al pie, gráficos, verificación visual |
| `rigor-analitico` | Al construir cualquier número que irá a una lámina: reconstrucción desde la data cruda, Excel de respaldo, conciliación |
| `storytelling-ejecutivo` | Al preparar una presentación para alta dirección: respuesta primero, palancas, decisiones, guion |
| `plan-de-trabajo-smartsheet` | Al crear o actualizar un plan de trabajo en Smartsheet: tareas bien definidas, estados honestos, uso de la API |
| `preparar-checkin` | Antes de una revisión con el senior: qué mostrar, qué validar, qué preguntar |
| `extraer-power-bi` | Cuando la data viene en un archivo .pbix |
| `organizar-proyecto` | Al iniciar un proyecto, guardar archivos, usar una clave de API o cerrar un proyecto |

## Reglas de oro (valen para todas las skills)

1. **Nunca contactar al equipo en nombre del usuario.** Prohibido etiquetar (@menciones), comentar dirigido a personas, asignar responsables que disparen notificaciones o enviar mensajes por Smartsheet, correo o chat. Claude redacta el mensaje, dice a quién va y el usuario lo envía manualmente.
2. **Cada número sale de la data cruda y queda trazable** a una hoja y celda de un Excel de respaldo.
3. **Un "bloqueo" se verifica antes de escalarlo.** Si el dato existe pero falta una herramienta, primero se intenta resolver.
4. **La respuesta va primero.** En láminas, correos y check-ins.
5. **No se publica ni se sube nada sin mirarlo.** Las láminas se exportan a imagen y se revisan antes de entregar.
6. **Las claves viven en un solo lugar,** `~\.pocket-consultant\.env`, nunca en un proyecto, en este repo ni en una carpeta sincronizada.

## Configuración local

Crea `~\.pocket-consultant\.env` (fuera de OneDrive y de cualquier proyecto):

```
SMARTSHEET_API_KEY=...
PROYECTOS_DIR=C:\ruta\a\proyectos
```

Cualquier variable también puede venir del entorno, que tiene prioridad. Los scripts la leen con `scripts/config.ps1`. Para usar otra carpeta, define `POCKET_CONSULTANT_HOME`.

## Espacio de trabajo

```
Trabajo/MB/
├── CLAUDE.md               convenciones del espacio de trabajo
├── pocket-consultant/      este repo: skills, scripts y plantillas
├── proyectos/              "AAAA-MM Nombre": README.md, 01 Insumos, 02 Notas, 03 Trabajo, 04 Entregables
└── _sandbox/               pruebas sin proyecto
```

Un proyecto nuevo se crea con `scripts/nuevo-proyecto.ps1` (ver skill `organizar-proyecto`).

## Scripts

| Script | Uso |
|---|---|
| `scripts/pptlib.ps1` | Librería PowerShell + PowerPoint COM para construir láminas con plantilla corporativa (títulos, tablas, gráficos nativos, callouts) |
| `scripts/export_png.ps1` | Exporta láminas a PNG para revisión visual |
| `scripts/config.ps1` | Lee claves y configuración desde el entorno o `~\.pocket-consultant\.env` (`Get-PcSetting`) |
| `scripts/smartsheet.ps1` | Cliente mínimo de la API de Smartsheet; `New-PlanSheet` crea un plan de trabajo desde un JSON |
| `scripts/nuevo-proyecto.ps1` | Crea la carpeta estándar de un proyecto con su ficha `README.md` |
| `scripts/pbix_export.ps1` | Abre un .pbix en Power BI Desktop y exporta todas sus tablas a CSV |

Los scripts son para Windows (PowerShell 5.1 con Office instalado). Guárdalos en UTF-8 con BOM: sin BOM, PowerShell 5.1 lee mal las tildes y los símbolos.

## Conectores

Hoy Smartsheet se usa por API con `scripts/smartsheet.ps1`. `New-PlanSheet` crea planes de trabajo desde un JSON. Próximo paso: conector MCP para Smartsheet.

## Cómo contribuir

Después de cada proyecto, agrega lo aprendido a la skill que corresponda (sección "Lecciones") o crea una nueva. Nunca subas datos de clientes, claves ni archivos `.env`.
