# Cliente mínimo de la API de Smartsheet (PowerShell 5.1).
# Uso:  . <repo>\scripts\smartsheet.ps1 ; $sheets = SS GET '/sheets?includeAll=true'
# La clave es SMARTSHEET_API_KEY: variable de entorno o ~\.pocket-consultant\.env (ver config.ps1).
# Nunca imprimas la clave ni la copies a una carpeta de proyecto o a un repositorio.
# REGLA: este cliente no debe usarse para contactar al equipo (comentarios dirigidos, @menciones, asignaciones que notifiquen).
#        Los mensajes los redacta Claude en el chat y los envía el usuario manualmente.
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'config.ps1')
$script:SSKEY = Get-PcSetting 'SMARTSHEET_API_KEY' -Required
$script:SSH = @{ Authorization = "Bearer $script:SSKEY" }
$script:SSBASE = 'https://api.smartsheet.com/2.0'
$script:SS_ESTADOS = @('Completo','Pendiente','Cancelado','En Curso','Bloqueado')

# Llamada genérica. Ojo: usa "${var}?param" en las rutas, porque "?" es válido en nombres de variable.
# -InputObject mantiene los arreglos de un elemento como arreglo. Ante un error muestra el mensaje de Smartsheet.
function SS([string]$method, [string]$path, $body = $null){
  $p = @{ Uri = "$script:SSBASE$path"; Method = $method; Headers = $script:SSH; ContentType = 'application/json; charset=utf-8' }
  if($body -ne $null){ $p.Body = [Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $body -Depth 10)) }
  try { return Invoke-RestMethod @p }
  catch {
    $msg = $_.Exception.Message
    if($_.Exception.Response){ try { $msg = (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() } catch {} }
    throw "Smartsheet $method $path -> $msg"
  }
}

# Busca una planilla por nombre exacto y devuelve su id numérico.
function Find-Sheet([string]$name){
  $r = SS GET '/sheets?includeAll=true'
  return ($r.data | Where-Object { $_.name -eq $name } | Select-Object -First 1).id
}

# Guarda un respaldo JSON de la planilla antes de modificarla.
function Backup-Sheet($sheetId, [string]$outFile){
  $s = SS GET "/sheets/${sheetId}?include=objectValue,discussions,attachments"
  ($s | ConvertTo-Json -Depth 20) | Out-File -Encoding utf8 $outFile
  return $s
}

# Actualiza celdas. $rows = @(@{ id = <rowId>; cells = @(@{ columnId = <colId>; value = '...' }) })
function Update-Rows($sheetId, $rows){ return SS PUT "/sheets/$sheetId/rows" $rows }

# Crea un plan de trabajo desde un JSON con fases y tareas (formato en la skill plan-de-trabajo-smartsheet).
# Las tareas quedan como filas hijas de su fase. Responsable va como texto, no como contacto, para no notificar a nadie.
function New-PlanSheet([string]$name, [string]$jsonPath){
  if(Find-Sheet $name){ throw "Ya existe una planilla llamada '$name'" }
  $def = @{
    name = $name
    columns = @(
      @{ title = 'Nombre Tarea';    type = 'TEXT_NUMBER'; primary = $true; width = 380 },
      @{ title = 'Comentarios';     type = 'TEXT_NUMBER'; width = 520 },
      @{ title = 'Estado';          type = 'PICKLIST'; options = $script:SS_ESTADOS },
      @{ title = 'Fecha de inicio'; type = 'DATE' },
      @{ title = 'Fecha de Fin';    type = 'DATE' },
      @{ title = 'Responsable';     type = 'TEXT_NUMBER'; width = 160 }
    )
  }
  $sheet = (SS POST '/sheets' $def).result
  $col = @{}; $sheet.columns | ForEach-Object { $col[$_.title] = $_.id }

  foreach($fase in (Get-Content -Raw -Encoding UTF8 $jsonPath | ConvertFrom-Json)){
    $parent = SS POST "/sheets/$($sheet.id)/rows" @(@{ toBottom = $true; cells = @(@{ columnId = $col['Nombre Tarea']; value = $fase.fase }) })
    $hijas = @(foreach($t in $fase.tareas){
      if($t.estado -and $script:SS_ESTADOS -notcontains $t.estado){ throw "Estado inválido '$($t.estado)' en '$($t.nombre)'" }
      $cells = @(@{ columnId = $col['Nombre Tarea']; value = $t.nombre })
      if($t.comentarios){ $cells += @{ columnId = $col['Comentarios'];     value = $t.comentarios } }
      if($t.estado){      $cells += @{ columnId = $col['Estado'];          value = $t.estado; strict = $true } }
      if($t.inicio){      $cells += @{ columnId = $col['Fecha de inicio']; value = $t.inicio } }
      if($t.fin){         $cells += @{ columnId = $col['Fecha de Fin'];    value = $t.fin } }
      if($t.responsable){ $cells += @{ columnId = $col['Responsable'];     value = $t.responsable } }
      @{ parentId = $parent.result[0].id; toBottom = $true; cells = $cells }
    })
    if($hijas.Count){ $null = SS POST "/sheets/$($sheet.id)/rows" $hijas }
  }
  return $sheet
}
