# Exporta láminas de un .pptx a PNG para revisión visual.
# Uso: .\export_png.ps1 -File deck.pptx -Out png [-From 1] [-To 999]
param([Parameter(Mandatory=$true)][string]$File, [string]$Out = 'png', [int]$From = 1, [int]$To = 999)
$File = (Resolve-Path $File).Path
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$pp = New-Object -ComObject PowerPoint.Application
try {
  $p = $pp.Presentations.Open($File, $true, $false, $false)
  $n = $p.Slides.Count
  for($i = $From; $i -le [Math]::Min($To, $n); $i++){ $p.Slides.Item($i).Export("$Out\s$i.png", 'PNG', 1280, 720) }
  "láminas: $n"
  $p.Close()
} finally { $pp.Quit() }
