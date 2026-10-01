# pptlib.ps1 - Librería para construir láminas de consultoría con PowerPoint COM (PowerShell 5.1).
# Supone una plantilla corporativa cuyas láminas 1 a 5 son: portada, agenda, contenido 1 columna, contenido 2 columnas, resumen.
# Funciones: New-ContentSlide, New-SectionSlide, Add-Box, Add-Bullets, Add-Table, Add-Chart, Add-Callout, Add-Kpi, Add-Tab, FmtNum, SetP, Rt.
# Ver skills/formato-laminas/referencia-tecnica.md para los problemas que resuelve (cultura es-CL, ChartData asíncrono, reintentos).
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
# Libreria de construccion de laminas estilo corporativo (PowerPoint COM)
# Colores en formato COM (BGR): usar RGB() helper
# Reintenta llamadas COM intermitentes (PowerPoint rechaza llamadas mientras recalcula texto)
function Rt([scriptblock]$sb){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  for($i=0; $i -lt 40; $i++){
    try { return (& $sb) } catch { if($i -eq 39){ throw }; Start-Sleep -Milliseconds 300 }
  }
}
function SetP($o, [string]$n, $v){ [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, @($v), [Globalization.CultureInfo]::InvariantCulture) }
# Convierte formato numerico ingles a local (es-CL): intercambia , y . fuera de comillas
function LocalFmt([string]$f){
  $sb = New-Object Text.StringBuilder; $q = $false
  foreach($ch in $f.ToCharArray()){
    if($ch -eq '"'){ $q = -not $q; [void]$sb.Append($ch); continue }
    if(-not $q -and $ch -eq ','){ [void]$sb.Append('.') } elseif(-not $q -and $ch -eq '.'){ [void]$sb.Append(',') } else { [void]$sb.Append($ch) }
  }
  return $sb.ToString()
}
# Formatea numero al estilo es-CL segun un codigo tipo Excel simple: [+|-]#,##0[.0..]["%"]
function FmtNum($v, [string]$fmt){
  $ci = [Globalization.CultureInfo]'es-CL'
  $pre = ''; if($fmt.StartsWith('+')){ $pre = '+' } elseif($fmt.StartsWith('-')){ $pre = '-' }
  $suf = ''; if($fmt -match '"%"' -or $fmt.EndsWith('%')){ $suf = '%' }
  $core = ($fmt -replace '"[^"]*"','') -replace '[+\-%]',''
  $dec = 0; $dot = $core.IndexOf('.'); if($dot -ge 0){ $dec = ($core.Substring($dot+1) -replace '[^0#]','').Length }
  $thou = $core.Contains(',')
  $n = [double]$v
  if($pre -eq '-' -and $n -gt 0){ $n = $n } elseif($pre -eq '+' -and $n -lt 0){ $pre = '' }
  if($thou){ $s = $n.ToString("N$dec", $ci) } else { $s = $n.ToString("F$dec", $ci) }
  if($pre -eq '-' -and $s.StartsWith('-')){ $pre = '' }
  return "$pre$s$suf"
}
function RGBc([int]$r,[int]$g,[int]$b){ return $r + 256*$g + 65536*$b }
$script:NAVY  = RGBc 27 54 93      # 1B365D
$script:NAVY2 = RGBc 139 160 181   # gris azulado para series secundarias
$script:TAN   = RGBc 201 180 132   # dorado/tan
$script:RED   = RGBc 176 76 76     # rojo perdida
$script:GREEN = RGBc 84 130 53
$script:GREY  = RGBc 127 127 127
$script:DGREY = RGBc 64 64 64
$script:LGREY = RGBc 231 234 238   # fondo cajas
$script:WHITE = RGBc 255 255 255
$script:FONT  = 'Tahoma'
$script:LEGPOS = -4160
$script:AXMIN = $null
$script:OVERLAP = $false
$script:AX2MAX = $null
$script:LOWTICKS = $false
$script:AXMAX = $null

# Plantilla: 1 portada, 2 agenda, 3 contenido 1 col (grafico), 4 contenido 2 col, 5 resumen tablas
function Copy-TemplateSlide($pres, [int]$templateIndex){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $r = $pres.Slides.Item($templateIndex).Duplicate()
  $s = $r.Item(1)
  $s.MoveTo($pres.Slides.Count)
  return $s
}

function Clear-NonPlaceholders($s){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  for($i=$s.Shapes.Count; $i -ge 1; $i--){
    $sh = $s.Shapes.Item($i)
    if($sh.Type -ne 14){ $sh.Delete() }
  }
}

function Get-Ph($s, [int]$type, [int]$nth = 1){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $k = 0
  foreach($sh in $s.Shapes){ try { if($sh.PlaceholderFormat.Type -eq $type){ $k++; if($k -eq $nth){ return $sh } } } catch {} }
  return $null
}

# Crea lamina de contenido. $twoCol: usa plantilla 2 columnas
function New-ContentSlideCore($pres, [string]$title, [string]$section, [string]$sub1, [string]$unit1, [string]$sub2 = '', [string]$unit2 = '', [string]$note = '', [switch]$twoCol, [switch]$preliminar){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $tpl = 3; if($twoCol){ $tpl = 4 }
  $s = Copy-TemplateSlide $pres $tpl
  Clear-NonPlaceholders $s
  $title_ph = Get-Ph $s 1
  $title_ph.TextFrame.TextRange.Text = $title
  try { $title_ph.TextFrame2.AutoSize = 2 } catch {}
  # placeholders de texto (type 2): orden en la plantilla
  $bodies = @(); foreach($sh in $s.Shapes){ try { if($sh.PlaceholderFormat.Type -eq 2){ $bodies += $sh } } catch {} }
  foreach($b in $bodies){
    if($b.Width -lt 30){ $b.TextFrame.TextRange.Text = $section; $b.TextFrame.Orientation = 2; $b.TextFrame.TextRange.Font.Color.RGB = $script:WHITE; $b.TextFrame.TextRange.Font.Bold = -1; $b.TextFrame.TextRange.Font.Size = 10; $b.TextFrame.TextRange.ParagraphFormat.Alignment = 2; $b.TextFrame.VerticalAnchor = 3; continue }
    $isUnit = ($b.TextFrame.TextRange.Font.Size -le 12.5)
    if($twoCol){
      $isLeft = ($b.Left -lt 300)
      if($isLeft -and -not $isUnit){ $b.TextFrame.TextRange.Text = $sub1 }
      elseif($isLeft -and $isUnit){ $b.TextFrame.TextRange.Text = $unit1 }
      elseif(-not $isLeft -and -not $isUnit){ $b.TextFrame.TextRange.Text = $sub2 }
      else { $b.TextFrame.TextRange.Text = $unit2 }
    } else {
      if(-not $isUnit){ $b.TextFrame.TextRange.Text = $sub1 } else { $b.TextFrame.TextRange.Text = $unit1 }
    }
    if($b.TextFrame.TextRange.Text -eq ''){ $b.Delete() }
    else { $b.TextFrame.WordWrap = 0; $b.TextFrame.AutoSize = 1 }
  }
  $n = Get-Ph $s 7
  if($n){ $n.TextFrame.TextRange.Text = $note; $n.TextFrame.TextRange.Font.Size = 8; $n.Width = 620 }
  if(-not $twoCol){ Add-Tab $s $section }
  if($preliminar){ Add-Box $s 59 9 130 18 'PRELIMINAR' 12 $true $script:NAVY $script:WHITE $script:NAVY 'center' | Out-Null }
  return $s
}

function Add-Box($s, $l, $t, $w, $h, [string]$text, $size = 11, $bold = $false, $fontColor = $script:DGREY, $fill = $null, $line = $null, $align = 'left'){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $b = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $tf = $b.TextFrame
  $tf.WordWrap = -1; $tf.AutoSize = 0
  $tf.MarginLeft = 5; $tf.MarginRight = 5; $tf.MarginTop = 3; $tf.MarginBottom = 3
  $tr = $tf.TextRange; Rt { $tr.Font.Name = $script:FONT }; Rt { SetP $tr.Font 'Size' ([single]$size) }; Rt { $tr.Font.Bold = [int]$bold }; Rt { $tr.Font.Color.RGB = $fontColor }; $tr.Text = $text; $tr = $tf.TextRange; $tr.ParagraphFormat.Bullet.Visible = 0; $tf.Ruler.Levels.Item(1).FirstMargin = 0; $tf.Ruler.Levels.Item(1).LeftMargin = 0
  switch($align){ 'center' { $tr.ParagraphFormat.Alignment = 2 } 'right' { $tr.ParagraphFormat.Alignment = 3 } default { $tr.ParagraphFormat.Alignment = 1 } }
  if($fill -ne $null){ $b.Fill.Visible = -1; $b.Fill.Solid(); $b.Fill.ForeColor.RGB = $fill } else { $b.Fill.Visible = 0 }
  if($line -ne $null){ $b.Line.Visible = -1; $b.Line.ForeColor.RGB = $line; $b.Line.Weight = 1 } else { $b.Line.Visible = 0 }
  $b.Height = $h
  return $b
}

# Bullets: $items = array de strings; linea que empieza con '>' = sub-bullet
function Add-Bullets($s, $l, $t, $w, $h, $items, $size = 11, $fill = $null){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $b = Add-Box $s $l $t $w $h (($items | ForEach-Object { $_ -replace '^>\s*','' }) -join "`r") $size $false $script:DGREY $fill
  $tr = $b.TextFrame.TextRange
  for($i=1; $i -le $items.Count; $i++){
    $p = $tr.Paragraphs($i)
    $p.ParagraphFormat.Bullet.Visible = -1
    $p.ParagraphFormat.Bullet.Character = 8226
    $p.ParagraphFormat.SpaceAfter = 4
    if($items[$i-1] -like '>*'){ Rt { $p.IndentLevel = 2 }; Rt { SetP $p.Font 'Size' ([single]($size - 1)) } }
  }
  return $b
}

function Set-BoldPrefix($shape, [string]$prefix){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $tr = $shape.TextFrame.TextRange
  $f = $tr.Find($prefix)
  if($f){ $f.Font.Bold = -1 }
}

# Tabla estilo corporativo. $rows = array de arrays (primera fila = encabezado). $colW = anchos
function Add-Table($s, $l, $t, $colW, $rows, $size = 10, $rowH = 20, $boldLast = $false){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $nr = $rows.Count; $nc = $rows[0].Count
  $w = ($colW | Measure-Object -Sum).Sum
  $shp = $s.Shapes.AddTable($nr, $nc, $l, $t, $w, $nr * $rowH)
  $tb = $shp.Table
  for($c=1; $c -le $nc; $c++){ Rt { SetP $tb.Columns.Item($c) 'Width' ([single]$colW[$c-1]) } }
  for($r=1; $r -le $nr; $r++){
    for($c=1; $c -le $nc; $c++){ Rt {
      $cell = $tb.Cell($r,$c)
      $tr = $cell.Shape.TextFrame.TextRange
      $tr.Text = [string]$rows[$r-1][$c-1]
      Rt { $tr.Font.Name = $script:FONT }; Rt { SetP $tr.Font 'Size' ([single]$size) }
      $cell.Shape.TextFrame.MarginTop = 2; $cell.Shape.TextFrame.MarginBottom = 2; $cell.Shape.TextFrame.MarginLeft = 4; $cell.Shape.TextFrame.MarginRight = 4
      $cell.Shape.TextFrame.VerticalAnchor = 3
      if($c -gt 1){ $tr.ParagraphFormat.Alignment = 2 } else { $tr.ParagraphFormat.Alignment = 1 }
      if($r -eq 1){
        $cell.Shape.Fill.ForeColor.RGB = $script:NAVY; $tr.Font.Color.RGB = $script:WHITE; $tr.Font.Bold = -1; $tr.ParagraphFormat.Alignment = 2
      } elseif($boldLast -and $r -eq $nr){
        $cell.Shape.Fill.ForeColor.RGB = $script:NAVY; $tr.Font.Color.RGB = $script:WHITE; $tr.Font.Bold = -1
      } else {
        if($r % 2 -eq 0){ $cell.Shape.Fill.ForeColor.RGB = $script:WHITE } else { $cell.Shape.Fill.ForeColor.RGB = $script:LGREY }
        $tr.Font.Color.RGB = $script:DGREY; $tr.Font.Bold = 0
      }
    } }
    Rt { SetP $tb.Rows.Item($r) 'Height' ([single]$rowH) }
  }
  return $shp
}

# Grafico nativo. $type: 51 col clustered, 52 col stacked, 53 col stacked100, 57 bar clustered, 58 bar stacked, 4 line, 65 line markers, 5 pie, -4120 doughnut
# $cats: array categorias; $series: array de @{name=..; values=@(...); color=..}
function Add-ChartCore($s, [int]$type, $l, $t, $w, $h, $cats, $series, [string]$numFmt = '#,##0', [switch]$legend, [switch]$noLabels, $gap = 60, [switch]$labelsInside){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  try { $win = $s.Parent.Windows.Item(1); $win.Activate(); $win.View.GotoSlide($s.SlideIndex) } catch {}
  $shp = Rt { $s.Shapes.AddChart2(-1, $type, $l, $t, $w, $h) }
  $ch = $shp.Chart
  $cd = $ch.ChartData
  $nr = $cats.Count + 1; $nc = $series.Count + 1
  $arr = New-Object 'object[,]' $nr, $nc
  $arr[0,0] = ' '
  for($i=0; $i -lt $cats.Count; $i++){ $arr[($i+1),0] = [string]$cats[$i] }
  for($j=0; $j -lt $series.Count; $j++){
    if($series.Count -eq 1){ $arr[0,($j+1)] = ' ' } else { $arr[0,($j+1)] = [string]$series[$j].name }
    for($i=0; $i -lt $cats.Count; $i++){ $v = $series[$j].values[$i]; if($v -eq $null){ $arr[($i+1),($j+1)] = $null } else { $arr[($i+1),($j+1)] = [double]$v } }
  }
  $colL = [char](64 + $nc)
  $ok = $false
  for($attempt=0; $attempt -lt 6 -and -not $ok; $attempt++){
    $wb = $null
    for($try=0; $try -lt 40 -and $wb -eq $null; $try++){
      try { $cd.Activate() } catch {}
      Start-Sleep -Milliseconds 400
      try { $wb = $cd.Workbook; $null = $wb.Worksheets.Count } catch { $wb = $null }
    }
    if($wb -eq $null){ continue }
    try { $wb.Application.WindowState = -4140 } catch {}
    try {
      $ws = $wb.Worksheets.Item(1)
      $ws.Cells.Clear()
      $rng = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($nr,$nc))
      try { $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($nr,1)).NumberFormat = '@'; $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1,$nc)).NumberFormat = '@' } catch {}
      $rng.Value2 = $arr
      for($j=0; $j -lt $series.Count; $j++){ $fmt = $numFmt; if($series[$j].lblfmt){ $fmt = $series[$j].lblfmt }; try { $ws.Range($ws.Cells.Item(2, $j + 2), $ws.Cells.Item($nr, $j + 2)).NumberFormat = $fmt } catch {} }
      Start-Sleep -Milliseconds 200
      $ch.SetSourceData("='" + $ws.Name + "'!`$A`$1:`$$colL`$$nr")
      try { $ch.PlotBy = 2 } catch {}
      $chk = [string]$ws.Cells.Item(1,2).Value2
      $ok = ($ch.SeriesCollection().Count -eq $series.Count) -and ($chk -eq [string]$arr[0,1])
    } catch { $ok = $false }
    try { $wb.Close() } catch {}
  }
  if(-not $ok){ throw "Datos del grafico no quedaron escritos" }
  Start-Sleep -Milliseconds 1500
  try { $ch.Refresh() } catch {}
  try { $ch.SetElement(0) } catch {}
  try { $ch.HasTitle = $false } catch {}
  $ch.HasLegend = [bool]$legend
  if($legend){ $ch.Legend.Position = $script:LEGPOS; $ch.Legend.Font.Name = $script:FONT; $ch.Legend.Font.Size = 10 }
  $ch.ChartArea.Format.TextFrame2.TextRange.Font.Name = $script:FONT
  $ch.ChartArea.Format.Line.Visible = 0
  $ch.ChartArea.Fill.Visible = 0
  $ch.PlotArea.Fill.Visible = 0
  for($j=1; $j -le $series.Count; $j++){
    $ser = $ch.SeriesCollection($j)
    $sd = $series[$j-1]
    if($sd.ctype){ $ser.ChartType = $sd.ctype }
    if($sd.axis2){ $ser.AxisGroup = 2 }
    $col = $sd.color
    if($col -eq -1){ $ser.Format.Fill.Visible = 0; $ser.Format.Line.Visible = 0; $ser.HasDataLabels = $false; continue }
    $isLine = ($type -eq 4 -or $type -eq 65 -or $sd.ctype -eq 4 -or $sd.ctype -eq 65)
    if($col -ne $null){
      if($isLine){ $ser.Format.Line.ForeColor.RGB = $col; SetP $ser.Format.Line 'Weight' ([single]2.25); try { $ser.MarkerBackgroundColor = $col; $ser.MarkerForegroundColor = $col } catch {} }
      else { $ser.Format.Fill.ForeColor.RGB = $col }
    }
    if(-not $noLabels -and -not $sd.nolabel){
      $ser.HasDataLabels = $true
      $dl = $ser.DataLabels()
      $f = $numFmt; if($sd.lblfmt){ $f = $sd.lblfmt }; SetP $dl 'NumberFormat' $f
      $dl.Font.Name = $script:FONT; $dl.Font.Size = 10

      if($isLine){ try { $dl.Position = 0 } catch {}; $dl.Font.Color = $col }
      elseif($sd.lblout){ try { $dl.Position = 2 } catch {}; $dl.Font.Color = $script:DGREY }
      elseif($labelsInside -or $type -eq 52 -or $type -eq 53 -or $type -eq 58 -or $type -eq 59){ try { $dl.Position = -4108 } catch {}; $dl.Font.Color = $(if($sd.lblcolor){$sd.lblcolor}else{$script:WHITE}) }
      else { $dl.Font.Color = $script:DGREY }
    }
  }
  if($type -ne 5 -and $type -ne -4120){
    try { $ch.ChartGroups(1).GapWidth = $gap } catch {}
    try { if($type -eq 52 -or $type -eq 53 -or $type -eq 58 -or $script:OVERLAP){ $ch.ChartGroups(1).Overlap = 100 } } catch {}
    $ax = $ch.Axes(2)
    $ax.HasMajorGridlines = $false
    if($script:AXMIN -ne $null){ $ax.MinimumScale = $script:AXMIN }
    if($script:AXMAX -ne $null){ $ax.MaximumScale = $script:AXMAX }
    $ax.Delete()
    try { $ax2 = $ch.Axes(2, 2); $ax2.HasMajorGridlines = $false; $ax2.MinimumScale = 0; if($script:AX2MAX -ne $null){ $ax2.MaximumScale = $script:AX2MAX }; $ax2.TickLabelPosition = -4142; $ax2.Format.Line.Visible = 0; $ax2.MajorTickMark = -4142 } catch {}
    $cx = $ch.Axes(1)
    $cx.TickLabels.Font.Size = 10; $cx.TickLabels.Font.Color = $script:DGREY
    $cx.Format.Line.ForeColor.RGB = $script:GREY
    $cx.MajorTickMark = -4142
    if($type -eq 57 -or $type -eq 58 -or $script:LOWTICKS){ try { $cx.TickLabelPosition = -4134 } catch {} }
  }
  # segunda pasada: el refresco asincrono del ChartData puede pisar formatos y titulo
  Start-Sleep -Milliseconds 600
  for($j=1; $j -le $series.Count; $j++){
    $sd = $series[$j-1]
    if($sd.color -eq -1 -or $noLabels -or $sd.nolabel){ continue }
    $f = $numFmt; if($sd.lblfmt){ $f = $sd.lblfmt }
    $serx = $ch.SeriesCollection($j)
    for($i=1; $i -le $cats.Count; $i++){
      $v = $sd.values[$i-1]
      try {
        $pt = $serx.Points($i)
        if($v -eq $null){ $pt.HasDataLabel = $false }
        else { $pt.HasDataLabel = $true; $pt.DataLabel.Text = (FmtNum $v $f) }
      } catch {}
    }
  }
  try { SetP $ch 'HasTitle' $false } catch {}
  try { if($ch.HasTitle){ $ch.ChartTitle.Delete() } } catch {}
  return $shp
}

function Add-Line($s, $x1, $y1, $x2, $y2, $color = $script:NAVY, $weight = 1, [switch]$dash){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $ln = $s.Shapes.AddLine($x1, $y1, $x2, $y2)
  $ln.Line.ForeColor.RGB = $color; $ln.Line.Weight = $weight
  if($dash){ $ln.Line.DashStyle = 4 }
  return $ln
}

# Numero grande + etiqueta (KPI tile)
function Add-Kpi($s, $l, $t, $w, [string]$value, [string]$label, $color = $script:NAVY){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $v = Add-Box $s $l $t $w 34 $value 24 $true $color $null $null 'center'
  $lb = Add-Box $s $l ($t + 34) $w 34 $label 10 $false $script:DGREY $null $null 'center'
  return $v
}





# Separador de seccion basado en agenda (plantilla 2). $items = lista, $k = item activo (1-based)
function New-SectionSlide($pres, $items, [int]$k){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $s = Copy-TemplateSlide $pres 2
  foreach($sh in $s.Shapes){
    if($sh.Name -eq 'Content Placeholder 4'){ $sh.TextFrame.TextRange.Text = ($items -join "`r") }
  }
  foreach($sh in $s.Shapes){
    if($sh.Name -eq 'Content Placeholder 3'){ SetP $sh 'Top' ([single](96 + ($k - 1) * 35.3)) }
  }
  return $s
}

# Caja con titulo en barra navy y cuerpo gris (callout)
function Add-Callout($s, $l, $t, $w, $h, [string]$head, $items, $size = 10){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $hb = Add-Box $s $l $t $w 20 $head 11 $true $script:WHITE $script:NAVY $null 'left'
  $b = Add-Bullets $s $l ($t + 20) $w ($h - 20) $items $size $script:LGREY
  return $b
}

# Resalta texto (negrita) de todas las ocurrencias
function Set-Bold($shape, [string]$txt){
  [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]'en-US'
  $tr = $shape.TextFrame.TextRange
  $f = $tr.Find($txt)
  while($f -ne $null){ $f.Font.Bold = -1; $start = $f.Start + $f.Length; $f = $tr.Find($txt, $start) ; if($f -and $f.Start -lt $start){ break } }
}











# Envoltorio con reintento: si PowerPoint rechaza una llamada, borra el grafico parcial y reintenta
function Add-Chart($s, [int]$type, $l, $t, $w, $h, $cats, $series, [string]$numFmt = '#,##0', [switch]$legend, [switch]$noLabels, $gap = 60, [switch]$labelsInside){
  for($k = 0; $k -lt 4; $k++){
    $n0 = $s.Shapes.Count
    try { return (Add-ChartCore @PSBoundParameters) }
    catch {
      if($k -eq 3){ throw }
      Start-Sleep -Seconds 3
      try { while($s.Shapes.Count -gt $n0){ $s.Shapes.Item($s.Shapes.Count).Delete() } } catch {}
      Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | Where-Object { $_.CommandLine -match 'automation|Embedding' } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force } catch {} }
      Start-Sleep -Seconds 2
    }
  }
}
function New-ContentSlide($pres, [string]$title, [string]$section, [string]$sub1, [string]$unit1, [string]$sub2 = '', [string]$unit2 = '', [string]$note = '', [switch]$twoCol, [switch]$preliminar){
  for($k = 0; $k -lt 4; $k++){
    $n0 = $pres.Slides.Count
    try { return (New-ContentSlideCore @PSBoundParameters) }
    catch {
      if($k -eq 3){ throw }
      Start-Sleep -Seconds 3
      try { while($pres.Slides.Count -gt $n0){ $pres.Slides.Item($pres.Slides.Count).Delete() } } catch {}
    }
  }
}
# Pestana lateral de seccion para laminas de una columna (la plantilla 1 columna no la trae)
function Add-Tab($s, [string]$text){
  foreach($sh in $s.Shapes){ if($sh.Name -eq 'TabSeccion'){ return } }
  $r = $s.Shapes.AddShape(1, 0, 0, 20, 540); $r.Name = 'TabSeccion'; $r.Fill.ForeColor.RGB = $script:NAVY; $r.Line.Visible = 0
  $tf = $r.TextFrame; $tf.Orientation = 2; $tf.MarginLeft = 0; $tf.MarginRight = 0
  $tr = $tf.TextRange; $tr.Text = $text; $tr.Font.Name = $script:FONT; $tr.Font.Size = 10; $tr.Font.Bold = -1; $tr.Font.Color.RGB = $script:WHITE; $tr.ParagraphFormat.Alignment = 2
}

# Borra placeholders vacíos de todas las láminas (en edición muestran "Haga clic para agregar..."). Llamar antes de guardar.
function Remove-EmptyPlaceholders($pres){
  $n = 0
  foreach($s in $pres.Slides){
    $del = @()
    foreach($sh in $s.Shapes){
      $pt = $null; try { $pt = $sh.PlaceholderFormat.Type } catch {}
      if($pt -eq $null -or $pt -in 13,15,16){ continue }
      if($sh.HasTextFrame){ if($sh.TextFrame.TextRange.Text.Trim() -eq ''){ $del += $sh } }
      elseif(-not ($sh.HasChart -or $sh.HasTable)){ try { if($sh.PlaceholderFormat.ContainedType -eq 1){ $del += $sh } } catch {} }
    }
    foreach($sh in $del){ $sh.Delete(); $n++ }
  }
  return $n
}