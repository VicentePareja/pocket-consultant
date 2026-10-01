# Revisa todos los .pptx de una carpeta: placeholders vacíos (muestran "Haga clic para agregar..." al editar)
# y textos de plantilla olvidados. Con -Fix borra los placeholders vacíos y guarda.
# Uso: powershell -File qa_pptx.ps1 -Path "C:\carpeta\entregables" [-Fix]
param([Parameter(Mandatory = $true)][string]$Path, [switch]$Fix)
$ErrorActionPreference = 'Stop'
$files = Get-ChildItem $Path -Recurse -Filter '*.pptx' | Where-Object { $_.Name -notlike '~$*' }
$pat = 'Inserte|Insertar|Haga clic|Hacer clic|Click to|Agregar t|Añadir t|Lorem|Título de|Subtítulo|XXX|TBD|\?\?\?'
$pp = New-Object -ComObject PowerPoint.Application
try {
  foreach($f in $files){
    $rel = $f.FullName.Substring((Resolve-Path $Path).Path.Length).TrimStart('\')
    try { $pr = $pp.Presentations.Open($f.FullName, (-not $Fix), $false, $false) } catch { "!! no se pudo abrir (¿está abierto?): $rel"; continue }
    $issues = 0; $fixed = 0
    foreach($s in $pr.Slides){
      $del = @()
      foreach($sh in $s.Shapes){
        $pt = $null; try { $pt = $sh.PlaceholderFormat.Type } catch {}
        $tx = $null; if($sh.HasTextFrame){ $tx = $sh.TextFrame.TextRange.Text }
        # 13 = número de lámina, 15 = pie, 16 = fecha: se llenan solos
        if($pt -ne $null -and $pt -notin 13,15,16){
          $empty = $false
          if($sh.HasTextFrame){ $empty = ($tx.Trim() -eq '') } elseif(-not ($sh.HasChart -or $sh.HasTable)){ try { $empty = ($sh.PlaceholderFormat.ContainedType -eq 1) } catch {} }
          if($empty){ $issues++; "  $rel | lámina $($s.SlideIndex) | placeholder vacío tipo $pt ($($sh.Name))"; $del += $sh }
        }
        if($tx -and $tx -match $pat){ $issues++; "  $rel | lámina $($s.SlideIndex) | texto sospechoso: $($tx.Substring(0,[Math]::Min(80,$tx.Length)) -replace "`r",' / ')" }
        if($sh.HasTable){ foreach($r in 1..$sh.Table.Rows.Count){ foreach($c in 1..$sh.Table.Columns.Count){ $ct = $sh.Table.Cell($r,$c).Shape.TextFrame.TextRange.Text; if($ct -match $pat){ $issues++; "  $rel | lámina $($s.SlideIndex) | tabla: $ct" } } } }
      }
      if($Fix){ foreach($sh in $del){ $sh.Delete(); $fixed++ } }
    }
    "$rel : $($pr.Slides.Count) láminas, $issues observaciones" + $(if($Fix){", $fixed placeholders borrados"}else{''})
    if($Fix -and $fixed -gt 0){ $pr.Save() }
    $pr.Close()
  }
} finally { $pp.Quit() }
