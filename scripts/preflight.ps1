# Chequeo previo de identidades: confirma que git, gh, Vercel y Supabase apuntan a la cuenta
# y al proyecto correctos ANTES de trabajar. Solo lee; no cambia nada.
# Uso:  & <repo>\scripts\preflight.ps1 -Config '<proyecto>\identidades.json'
# Formato de identidades.json en la skill organizar-proyecto.
param([Parameter(Mandatory = $true)][string]$Config)
$ErrorActionPreference = 'Stop'

$c = Get-Content -Raw -Encoding UTF8 $Config | ConvertFrom-Json
$repo = [Environment]::ExpandEnvironmentVariables($c.repo)
$fallas = 0

function Check([string]$que, [string]$esperado, [string]$real){
  if($real -eq $esperado){ Write-Host ("  OK     {0}: {1}" -f $que, $real) -ForegroundColor Green }
  else { Write-Host ("  FALLA  {0}: esperado '{1}', obtenido '{2}'" -f $que, $esperado, $real) -ForegroundColor Red; $script:fallas++ }
}
# Algunas CLI (vercel) escriben su salida normal en stderr: se captura todo, sin que PowerShell 5.1 la trate como error.
function Run([scriptblock]$cmd){ $ErrorActionPreference = 'Continue'; try { (& $cmd 2>&1 | ForEach-Object { "$_" } | Out-String).Trim() } catch { '' } }

Write-Host "Git"
Check 'remoto origin' $c.git_remote (Run { git -C $repo remote get-url origin })
Check 'correo de commits' $c.git_email (Run { git -C $repo config user.email })

Write-Host "GitHub CLI"
Check 'cuenta activa de gh' $c.gh_user (Run { gh api user --jq .login })

if($c.vercel_team){
  Write-Host "Vercel"
  $vc = $c.vercel_config -replace '^~', $HOME
  $equipos = Run { vercel teams ls --global-config $vc }
  $tieneEquipo = if($equipos -match [regex]::Escape($c.vercel_team)){ $c.vercel_team } else { '(sin acceso)' }
  Check 'equipo visible con la sesión de trabajo' $c.vercel_team $tieneEquipo
  $link = Join-Path $repo ($c.vercel_link_dir + '\.vercel\project.json')
  if(Test-Path $link){
    $p = Get-Content -Raw $link | ConvertFrom-Json
    Check 'proyecto vinculado (.vercel)' $c.vercel_project $p.projectName
    # El nombre puede repetirse entre equipos; el id no.
    if($c.vercel_project_id){ Check 'id del proyecto vinculado' $c.vercel_project_id $p.projectId }
  } else { Write-Host "  --     proyecto aún no vinculado ($link)" -ForegroundColor Yellow }
}

if($c.supabase_ref){
  Write-Host "Supabase (MCP de Claude Code)"
  $cfg = Get-Content -Raw (Join-Path $HOME '.claude.json') | ConvertFrom-Json
  $proj = $cfg.projects.PSObject.Properties | Where-Object { $_.Name -replace '\\','/' -eq ($repo -replace '\\','/') } | Select-Object -First 1
  $urls = @($proj.Value.mcpServers.PSObject.Properties | ForEach-Object { $_.Value.url } | Where-Object { $_ -match 'supabase' })
  $ok = $urls.Count -gt 0 -and @($urls | Where-Object { $_ -notmatch "project_ref=$($c.supabase_ref)" -or ($c.supabase_read_only -and $_ -notmatch 'read_only=true') }).Count -eq 0
  Check 'MCP fijo al proyecto (y solo lectura si aplica)' 'si' $(if($ok){'si'}else{"no: $($urls -join ', ')"})
  # Un MCP de Supabase global sin project_ref daría acceso a todos los proyectos de la cuenta.
  $globales = @($cfg.mcpServers.PSObject.Properties | ForEach-Object { $_.Value.url } | Where-Object { $_ -match 'supabase' -and $_ -notmatch 'project_ref=' })
  Check 'sin MCP de Supabase global sin fijar' 'ninguno' $(if($globales.Count){ $globales -join ', ' } else { 'ninguno' })
}

if($fallas){ Write-Host "`n$fallas chequeo(s) fallaron. No trabajes hasta corregirlos." -ForegroundColor Red; exit 1 }
Write-Host "`nTodo apunta a la cuenta y al proyecto correctos." -ForegroundColor Green
