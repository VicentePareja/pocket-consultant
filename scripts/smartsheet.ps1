# Cliente mínimo de la API de Smartsheet (PowerShell 5.1).
# Uso:  . .\smartsheet.ps1 ; $sheets = SS GET '/sheets?includeAll=true'
# La clave se toma de $env:SMARTSHEET_API_KEY o de un archivo .env (SMARTSHEET_API_KEY=...) en el directorio actual.
# Nunca imprimas la clave ni subas el .env a un repositorio.
# REGLA: este cliente no debe usarse para contactar al equipo (comentarios dirigidos, @menciones, asignaciones que notifiquen).
#        Los mensajes los redacta Claude en el chat y los envía el usuario manualmente.
$ErrorActionPreference = 'Stop'
$script:SSKEY = $env:SMARTSHEET_API_KEY
if(-not $script:SSKEY -and (Test-Path '.env')){
  $line = Get-Content '.env' | Where-Object { $_ -match '^\s*(SMARTSHEET_API_KEY|SHH_API_KEY)\s*=' } | Select-Object -First 1
  if($line){ $script:SSKEY = ($line -replace '^[^=]*=','').Trim().Trim('"').Trim("'") }
}
if(-not $script:SSKEY){ throw 'No encuentro SMARTSHEET_API_KEY (variable de entorno o .env)' }
$script:SSH = @{ Authorization = "Bearer $script:SSKEY" }
$script:SSBASE = 'https://api.smartsheet.com/2.0'

# Llamada genérica. Ojo: usa "${var}?param" en las rutas, porque "?" es válido en nombres de variable.
function SS([string]$method, [string]$path, $body = $null){
  $p = @{ Uri = "$script:SSBASE$path"; Method = $method; Headers = $script:SSH; ContentType = 'application/json; charset=utf-8' }
  if($body -ne $null){ $p.Body = [Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Depth 10)) }
  return Invoke-RestMethod @p
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
