---
name: organizar-proyecto
description: Dónde vive cada cosa en el espacio de trabajo - carpeta estándar de cada proyecto, nombres de archivo, dónde se guardan las claves y qué va en el repositorio de skills. Usar al iniciar un proyecto, al guardar un archivo nuevo (notas, análisis, entregables), al necesitar una clave de API o al cerrar un proyecto.
---

# Organizar un proyecto

## Espacio de trabajo
```
Trabajo/MB/
├── CLAUDE.md               convenciones (este resumen)
├── pocket-consultant/      repo de skills y scripts. Nunca datos de clientes.
├── proyectos/              un proyecto por carpeta: "AAAA-MM Nombre"
└── _sandbox/               pruebas y borradores sin proyecto
~\.pocket-consultant\.env   claves y configuración local (fuera de todo lo anterior)
```

## Un proyecto nuevo
Créalo con el script, no a mano:
```
& <repo>\scripts\nuevo-proyecto.ps1 -Nombre 'Diagnóstico Compras' -Cliente 'Cliente X' -Supervisor 'Nombre del senior'
```
Deja `README.md` y cuatro carpetas:

| Carpeta | Qué va |
|---|---|
| `01 Insumos` | Lo que entrega el cliente o el equipo: data cruda, propuestas, documentos. No se edita; se trabaja sobre copias en `03 Trabajo`. |
| `02 Notas` | Notas de reuniones: `AAAA-MM-DD Tema.md`. |
| `03 Trabajo` | Análisis, Excel de respaldo, borradores y scripts del proyecto. |
| `04 Entregables` | Lo que se envía, con fecha y versión: `AAAAMMDD - Nombre vN.ext`. Nunca se sobrescribe una versión enviada. |

El `README.md` es la ficha del proyecto: cliente, supervisor, objetivo, alcance, link e id de la planilla de Smartsheet y la tabla de decisiones. Se actualiza cuando cambia algo de eso, no a diario.

## Claves y configuración
- **Una sola ubicación:** `~\.pocket-consultant\.env` (o la variable de entorno del mismo nombre). Los scripts la leen con `scripts/config.ps1` (`Get-PcSetting`).
- **Nunca** dentro de un proyecto, del repo o de una carpeta sincronizada (OneDrive, SharePoint, Drive). Si aparece una, se mueve a esa ubicación y se borra la copia.
- Nunca se imprime una clave: para verificar, muestra solo el largo o prueba una llamada a la API.
- Variables actuales: `SMARTSHEET_API_KEY`, `PROYECTOS_DIR`.

## Qué va en el repositorio y qué no
- **Repo:** skills, scripts, plantillas y lecciones generalizadas, sin nombres de clientes ni cifras.
- **Proyecto:** todo lo demás (data, notas, análisis, entregables). El `.gitignore` del repo bloquea `.env`, `.xlsx`, `.pptx`, `.pbix` y `.csv` como red de seguridad, no como regla principal.

## Al cerrar un proyecto
1. Cambia el estado de la ficha a "Cerrado" y deja el link a la versión final de cada entregable.
2. Lleva lo aprendido a la skill que corresponda del repo, sin datos del cliente.
3. Si el proyecto usó una clave propia, revócala.
