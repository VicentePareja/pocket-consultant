# Abre un .pbix en Power BI Desktop, se conecta a su motor local (msmdsrv) y exporta cada tabla a CSV.
# Requiere Power BI Desktop instalado (winget install --id Microsoft.PowerBI -e).
# Uso: .\pbix_export.ps1 -Pbix archivo.pbix -Out carpeta [-ListOnly]
param([Parameter(Mandatory=$true)][string]$Pbix, [string]$Out = 'pbix_export', [switch]$ListOnly)
$ErrorActionPreference = 'Stop'
$Pbix = (Resolve-Path $Pbix).Path
New-Item -ItemType Directory -Force $Out | Out-Null

$pbiExe = @("$env:ProgramFiles\Microsoft Power BI Desktop\bin\PBIDesktop.exe", "${env:ProgramFiles(x86)}\Microsoft Power BI Desktop\bin\PBIDesktop.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if(-not $pbiExe){ throw 'No encuentro PBIDesktop.exe: instala Power BI Desktop' }
$bin = Split-Path $pbiExe
$wsRoot = "$env:LOCALAPPDATA\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces"

if(-not (Get-Process msmdsrv -ErrorAction SilentlyContinue)){ Start-Process -FilePath $pbiExe -ArgumentList "`"$Pbix`"" }
$port = $null
for($i = 0; $i -lt 120 -and -not $port; $i++){
  Start-Sleep -Seconds 5
  $pf = Get-ChildItem $wsRoot -Recurse -Filter 'msmdsrv.port.txt' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if($pf){ $port = ([IO.File]::ReadAllText($pf.FullName, [Text.Encoding]::Unicode)).Trim() }
}
if(-not $port){ throw 'El motor de Power BI no levantó' }
"Puerto del motor: $port"

Add-Type -Path (Get-ChildItem $bin -Filter 'Microsoft.PowerBI.AdomdClient.dll' | Select-Object -First 1).FullName
$conn = New-Object Microsoft.AnalysisServices.AdomdClient.AdomdConnection "Data Source=localhost:$port"
$cat = $null
for($i = 0; $i -lt 60 -and -not $cat; $i++){
  try { if($conn.State -ne 'Open'){ $conn.Open() }; $ds = $conn.GetSchemaDataSet('DBSCHEMA_CATALOGS', $null); if($ds.Tables[0].Rows.Count){ $cat = $ds.Tables[0].Rows[0][0] } } catch {}
  if(-not $cat){ Start-Sleep -Seconds 5 }
}
if(-not $cat){ throw 'El modelo no terminó de cargar' }
$conn.ChangeDatabase($cat)
function Q([string]$dax){ $cmd = $conn.CreateCommand(); $cmd.CommandText = $dax; $cmd.CommandTimeout = 600; $rd = $cmd.ExecuteReader(); return ,$rd }

$tables = @()
$rd = Q 'SELECT [Name] FROM $SYSTEM.TMSCHEMA_TABLES'
while($rd.Read()){ $tables += [string]$rd.GetValue(0) }
$rd.Close()
$tables = $tables | Where-Object { $_ -notlike 'LocalDateTable_*' -and $_ -notlike 'DateTableTemplate_*' }
foreach($t in $tables){
  $c = Q "EVALUATE ROW(""n"", COUNTROWS('$($t.Replace("'","''"))'))"; $null = $c.Read(); "TABLA: $t | filas=$($c.GetValue(0))"; $c.Close()
}
if($ListOnly){ $conn.Close(); return }

foreach($t in $tables){
  $safe = ($t -replace '[^\w\-]','_')
  $rd = Q "EVALUATE '$($t.Replace("'","''"))'"
  $sw = New-Object IO.StreamWriter((Join-Path (Resolve-Path $Out).Path "$safe.csv"), $false, (New-Object Text.UTF8Encoding($true)))
  $hdr = @(); for($c = 0; $c -lt $rd.FieldCount; $c++){ $hdr += ('"' + ($rd.GetName($c) -replace '"','""') + '"') }
  $sw.WriteLine(($hdr -join ';'))
  $n = 0
  while($rd.Read()){
    $vals = @()
    for($c = 0; $c -lt $rd.FieldCount; $c++){
      $v = $rd.GetValue($c)
      if($v -is [datetime]){ $v = $v.ToString('yyyy-MM-dd HH:mm:ss') } elseif($v -is [double] -or $v -is [decimal]){ $v = $v.ToString([Globalization.CultureInfo]::InvariantCulture) }
      $vals += ('"' + ([string]$v -replace '"','""') + '"')
    }
    $sw.WriteLine(($vals -join ';')); $n++
  }
  $sw.Close(); $rd.Close()
  "exportada ${t}: $n filas"
}
$conn.Close()
