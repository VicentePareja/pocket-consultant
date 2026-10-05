# Crea la carpeta estándar de un proyecto en PROYECTOS_DIR (ver config.ps1):
#   <PROYECTOS_DIR>\AAAA-MM Nombre\  README.md, 01 Insumos, 02 Notas, 03 Trabajo, 04 Entregables
# Uso:  & <repo>\scripts\nuevo-proyecto.ps1 -Nombre 'Diagnóstico Compras' -Cliente 'Cliente X' -Supervisor 'Nombre del senior'
param(
  [Parameter(Mandatory = $true)][string]$Nombre,
  [string]$Cliente = '',
  [string]$Supervisor = '',
  [datetime]$Inicio = (Get-Date)
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'config.ps1')

$raiz = Get-PcSetting 'PROYECTOS_DIR' -Required
$dir = Join-Path $raiz ('{0:yyyy-MM} {1}' -f $Inicio, $Nombre)
if(Test-Path $dir){ throw "Ya existe $dir" }

foreach($d in '01 Insumos','02 Notas','03 Trabajo','04 Entregables'){ New-Item -ItemType Directory -Force (Join-Path $dir $d) | Out-Null }

$plantilla = Join-Path $PSScriptRoot '..\plantillas\proyecto\README.md'
$readme = [IO.File]::ReadAllText($plantilla, [Text.Encoding]::UTF8)
$readme = $readme.Replace('{{NOMBRE}}', $Nombre).Replace('{{CLIENTE}}', $Cliente).Replace('{{SUPERVISOR}}', $Supervisor).Replace('{{INICIO}}', $Inicio.ToString('yyyy-MM-dd'))
[IO.File]::WriteAllText((Join-Path $dir 'README.md'), $readme, (New-Object Text.UTF8Encoding($false)))
$dir
