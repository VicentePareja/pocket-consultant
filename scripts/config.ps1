# Configuración y secretos locales de pocket-consultant (PowerShell 5.1).
# Fuente única: la variable de entorno y, si no existe, el archivo ~\.pocket-consultant\.env,
# que vive fuera de los proyectos, del repositorio y de carpetas sincronizadas (OneDrive, SharePoint).
# Uso:  . "$PSScriptRoot\config.ps1" ; $k = Get-PcSetting 'SMARTSHEET_API_KEY' -Required
# Nunca imprimas un secreto ni lo copies a una carpeta de proyecto.
$script:PC_HOME = if ($env:POCKET_CONSULTANT_HOME) { $env:POCKET_CONSULTANT_HOME } else { Join-Path $HOME '.pocket-consultant' }
$script:PC_ENV = Join-Path $script:PC_HOME '.env'

function Get-PcSetting([string]$name, [switch]$Required){
  $v = [Environment]::GetEnvironmentVariable($name)
  if(-not $v -and (Test-Path $script:PC_ENV)){
    $pattern = '^\s*' + [regex]::Escape($name) + '\s*='
    $line = Get-Content -Encoding UTF8 $script:PC_ENV | Where-Object { $_ -match $pattern } | Select-Object -First 1
    if($line){ $v = ($line -replace '^[^=]*=', '').Trim().Trim('"').Trim("'") }
  }
  if(-not $v -and $Required){ throw "Falta $name. Agrégalo a $script:PC_ENV (ver README, 'Configuración local')." }
  return $v
}
