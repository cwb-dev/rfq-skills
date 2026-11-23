<#
Generate a Chinese Runsun purchase contract as .xls + PDF.

powershell -ExecutionPolicy Bypass -File make_pc.ps1 -Spec <pc_spec.json> [-Force]

See pc_spec.example.json for the complete field structure. Quantities and prices must use
matching units, for example KPCS and CNY/KPCS or PCS and CNY/PCS.
#>
param([Parameter(Mandatory=$true)][string]$Spec, [switch]$Force)
$ErrorActionPreference = 'Stop'
$skill = Split-Path -Parent $MyInvocation.MyCommand.Path
$s = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json

function V($o, [string]$name) {
  if ($null -ne $o -and $o.PSObject.Properties.Name -contains $name -and $null -ne $o.$name) { return [string]$o.$name }
  return ''
}
function Width([string]$text) {
  $w = 0
  foreach ($ch in $text.ToCharArray()) { $w += $(if ([int]$ch -gt 127) { 2 } else { 1 }) }
  return $w
}
function Est([string]$text, [int]$chars) {
  $lines = 0
  foreach ($part in ($text -split "`n")) { $lines += [math]::Max(1, [math]::Ceiling((Width $part) / $chars)) }
  return $lines
}
function SectionWords([int]$n) {
  $digits = @('零','壹','贰','叁','肆','伍','陆','柒','捌','玖')
  $units = @('','拾','佰','仟')
  $result = ''; $pendingZero = $false
  for ($pos = 3; $pos -ge 0; $pos--) {
    $d = [int]([math]::Floor($n / [math]::Pow(10, $pos))) % 10
    if ($d -eq 0) { if ($result) { $pendingZero = $true }; continue }
    if ($pendingZero) { $result += '零'; $pendingZero = $false }
    $result += $digits[$d] + $units[$pos]
  }
  return $result
}
function IntegerWords([decimal]$n) {
  if ($n -eq 0) { return '零' }
  $big = @('','万','亿','兆')
  $groups = @()
  while ($n -gt 0) {
    $groups += [int]($n % 10000)
    $n = [decimal]::Floor($n / 10000)
  }
  if ($groups.Count -gt $big.Count) { throw 'amount is too large for RMB words' }
  $result = ''; $zeroGroup = $false
  for ($i = $groups.Count - 1; $i -ge 0; $i--) {
    $g = $groups[$i]
    if ($g -eq 0) { if ($result) { $zeroGroup = $true }; continue }
    if ($result -and ($zeroGroup -or $g -lt 1000)) { $result += '零' }
    $result += (SectionWords $g) + $big[$i]
    $zeroGroup = $false
  }
  return $result
}
function RmbWords([decimal]$amount) {
  if ($amount -lt 0) { throw 'negative contract amount' }
  $cents = [long][decimal]::Round($amount * 100, 0)
  $integer = [decimal]::Floor($cents / 100)
  $remain = [int]($cents % 100)
  $jiao = [int][math]::Floor($remain / 10); $fen = $remain % 10
  $digits = @('零','壹','贰','叁','肆','伍','陆','柒','捌','玖')
  $result = (IntegerWords $integer) + '元'
  if ($remain -eq 0) { return $result + '整' }
  if ($jiao -gt 0) { $result += $digits[$jiao] + '角' }
  if ($fen -gt 0) { if ($jiao -eq 0) { $result += '零' }; $result += $digits[$fen] + '分' }
  return $result
}

$items = @($s.items)
if ($items.Count -lt 1) { throw 'no items' }
$terms = @($s.terms | Where-Object { $_ })
$labels = @($s.labels | Where-Object { $_ })
if ($labels.Count -gt 6) { throw 'the label panel supports at most 6 rows' }
if ($s.qty_unit -ne 'KPCS' -and $s.qty_unit -ne 'PCS') { throw 'qty_unit must be KPCS or PCS' }
if ($s.price_unit -ne "CNY/$($s.qty_unit)") { throw 'price_unit must match qty_unit, e.g. CNY/KPCS' }
$total = [decimal]0; $weightTotal = [decimal]0
foreach ($it in $items) {
  $qty = [decimal]$it.qty; $price = [decimal]$it.unit_price
  if ($qty -le 0 -or $price -lt 0) { throw 'item quantity and price must be positive' }
  $total += [decimal]::Round($qty * $price, 2)
  if ($null -ne $it.weight_total) { $weightTotal += [decimal]$it.weight_total }
  if ($null -ne $it.unit_weight -and $null -ne $it.weight_total) {
    $expectedWeight = [decimal]::Round($qty * [decimal]$it.unit_weight, 2)
    if ([math]::Abs($expectedWeight - [decimal]$it.weight_total) -gt 0.02) {
      throw "weight mismatch for $($it.customer_drawing): calculated $expectedWeight vs stated $($it.weight_total)"
    }
  }
}
if ($null -ne $s.expected_total -and [math]::Abs($total - [decimal]$s.expected_total) -gt 0.005) {
  throw "total mismatch: calculated $total vs expected_total $($s.expected_total)"
}
$outDir = [string]$s.out_dir
New-Item -ItemType Directory -Force $outDir | Out-Null
$base = $(if (V $s 'basename') { [string]$s.basename } else { [string]$s.contract_no + '（' + [string]$s.supplier_short + '）' })
if ($base -match '[\\/:*?"<>|]') { throw 'basename contains invalid filename characters' }
$xlsPath = Join-Path $outDir ($base + '.xls')
$pdfPath = Join-Path $outDir ($base + '.pdf')
if (-not $Force -and ((Test-Path $xlsPath) -or (Test-Path $pdfPath))) { throw "$xlsPath or .pdf exists; use -Force only when replacement is intended" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AskToUpdateLinks = $false; $xl.AutomationSecurity = 3
try {
  $wb = $xl.Workbooks.Open((Join-Path $skill 'templates\PC-RUNSUN.xlsx'), 0, $true)
  $ws = $wb.Worksheets.Item(1)
  function Put([string]$addr, $value) {
    if ($value -is [psobject]) { $value = $value.psobject.BaseObject }
    $cell = $ws.Range($addr).MergeArea.Cells.Item(1)
    if ($value -is [string] -or $null -eq $value) { $cell.Value2 = [string]$value }
    else { $cell.Value2 = [double]$value }
  }

  Put 'M7' $s.contract_no
  Put 'M8' ([string]$s.place + '  |  ' + [string]$s.date)
  Put 'A11' ([string]$s.buyer.name + "`n" + [string]$s.buyer.address)
  Put 'I11' ([string]$s.supplier.name + "`n" + [string]$s.supplier.address)
  $qtyText = $(if ($s.qty_unit -eq 'KPCS') { '千个' } else { '个' })
  $weightUnit = $(if ($s.qty_unit -eq 'KPCS') { 'KG/千个' } else { 'KG/个' })
  Put 'I15' ("数量`n" + $qtyText)
  Put 'J15' ("千重`n" + $weightUnit)
  Put 'M15' ("价格`n元/" + $qtyText + "`n含税运价")
  Put 'O15' ("合计`n元")

  $n = $items.Count
  if ($n -gt 1) {
    $ws.Range("17:$(16+$n)").EntireRow.Insert() | Out-Null
    $ws.Rows.Item(16).Copy($ws.Range("17:$(16+$n)")) | Out-Null
  }
  $d = $n - 1
  $qtyFmt = $(if (V $s 'qty_format') { [string]$s.qty_format } elseif ($s.qty_unit -eq 'KPCS') { '#,##0.000' } else { '#,##0' })
  for ($i = 0; $i -lt $n; $i++) {
    $r = 16 + $i; $it = $items[$i]
    Put "A$r" ([string]$it.name)
    Put "D$r" ([string]$it.product_drawing)
    Put "E$r" ([string]$it.customer_drawing)
    Put "F$r" ([string]$it.description_en)
    Put "I$r" ([double]$it.qty)
    $ws.Range("I$r").NumberFormat = $qtyFmt
    if ($null -ne $it.unit_weight) { Put "J$r" ([double]$it.unit_weight); $ws.Range("J$r").NumberFormat = '#,##0.00' }
    if ($null -ne $it.weight_total) { Put "K$r" ([double]$it.weight_total); $ws.Range("K$r").NumberFormat = '#,##0.00' }
    Put "L$r" ([string]$s.packing)
    Put "M$r" ([double]$it.unit_price)
    $ws.Range("M$r").NumberFormat = '#,##0.00'
    $ws.Range("O$r").Formula = "=ROUND(I$r*M$r,2)"
    $ws.Range("O$r").NumberFormat = '#,##0.00'
    $lines = [math]::Max((Est ([string]$it.name) 28), (Est ([string]$it.description_en) 20))
    $lines = [math]::Max($lines, (Est ([string]$it.product_drawing) 16))
    $lines = [math]::Max($lines, (Est ([string]$it.customer_drawing) 16))
    $ws.Rows.Item($r).RowHeight = [math]::Max(64, $lines * 14 + 10)
  }
  $tr = 17 + $d; $ar = 18 + $d; $body = 20 + $d; $tf = 21 + $d; $sh = 37 + $d
  $lastItem = 16 + $n - 1
  Put "A$tr" '合计：'
  $ws.Range("I$tr").Formula = "=SUM(I16:I$lastItem)"
  $ws.Range("I$tr").NumberFormat = $qtyFmt
  Put "K$tr" ([double]$weightTotal); $ws.Range("K$tr").NumberFormat = '#,##0.00'
  $ws.Range("O$tr").Formula = "=SUM(O16:O$lastItem)"
  $ws.Range("O$tr").NumberFormat = '#,##0.00'
  Put "A$ar" ('合计人民币金额（大写）：' + (RmbWords $total) + '。')

  Put "A$tf" ([string]$s.label_note)
  $labelHeights = @{}
  for ($i = 0; $i -lt 6; $i++) {
    $r = $tf + 1 + [int][math]::Floor($i / 2)
    if ($i % 2 -eq 0) { $c1 = 'A'; $c2 = 'C'; $c3 = 'E' } else { $c1 = 'I'; $c2 = 'K'; $c3 = 'M' }
    if ($i -lt $labels.Count) {
      Put "$c1$r" $labels[$i][0]; Put "$c2$r" $labels[$i][1]; Put "$c3$r" $labels[$i][2]
      $h = [double][math]::Max(30, (Est ([string]$labels[$i][2]) 24) * 12 + 7)
      if ($labelHeights.ContainsKey($r)) { $h = [math]::Max($labelHeights[$r], $h) }
      $labelHeights[$r] = $h
    } else { Put "$c1$r" ''; Put "$c2$r" ''; Put "$c3$r" '' }
  }
  foreach ($r in $labelHeights.Keys) { $ws.Rows.Item($r).RowHeight = [double]$labelHeights[$r] }
  if ($terms.Count -gt 9) {
    $extra = $terms.Count - 9
    $ws.Range("${sh}:$($sh+$extra-1)").EntireRow.Insert() | Out-Null
    $sh += $extra
  }
  for ($i = 0; $i -lt [math]::Max(9, $terms.Count); $i++) {
    $r = $tf + 6 + $i
    if ($i -lt $terms.Count) {
      $text = "$($i+1). " + [string]$terms[$i]
      Put "A$r" $text
      $ws.Rows.Item($r).RowHeight = [double][math]::Max(18, (Est $text 100) * 14 + 5)
    } else { Put "A$r" '' }
  }

  Put "A$($sh+1)" ('单位名称（章）：' + [string]$s.buyer.name)
  Put "I$($sh+1)" ('单位名称（章）：' + [string]$s.supplier.name)
  Put "A$($sh+2)" ('单位地址：' + [string]$s.buyer.address)
  Put "I$($sh+2)" ('单位地址：' + [string]$s.supplier.address)
  Put "A$($sh+3)" ('法定代表人：' + (V $s.buyer 'legal_representative') + '    委托代理人：' + (V $s.buyer 'agent'))
  Put "I$($sh+3)" ('法定代表人：' + (V $s.supplier 'legal_representative') + '    委托代理人：' + (V $s.supplier 'agent'))
  Put "A$($sh+4)" ('电话：' + (V $s.buyer 'phone'))
  Put "I$($sh+4)" ('电话：' + (V $s.supplier 'phone'))
  Put "A$($sh+5)" ('开户银行：' + (V $s.buyer 'bank'))
  Put "I$($sh+5)" ('开户银行：' + (V $s.supplier 'bank'))
  Put "A$($sh+6)" ('账号：' + (V $s.buyer 'account'))
  Put "I$($sh+6)" ('账号：' + (V $s.supplier 'account'))
  Put "A$($sh+7)" ('税号：' + (V $s.buyer 'tax_id'))
  $supplierLast = ''
  if (V $s.supplier 'tax_id') { $supplierLast = '税号：' + (V $s.supplier 'tax_id') }
  if (V $s.supplier 'bank_no') { if ($supplierLast) { $supplierLast += '    ' }; $supplierLast += '行号：' + (V $s.supplier 'bank_no') }
  Put "I$($sh+7)" $supplierLast
  $ws.PageSetup.PrintArea = "A1:P$($sh+10)"

  $xl.Calculate()
  $calc = [decimal]$ws.Range("O$tr").Value2
  if ([math]::Abs($calc - $total) -gt 0.005) { throw "sheet total mismatch: $calc vs $total" }
  $wb.SaveAs($xlsPath, 56)
  $wb.ExportAsFixedFormat(0, $pdfPath)
  $wb.Close($false)

  "OK  total=$('{0:N2}' -f $total) CNY  items=$n  amount_words=$(RmbWords $total)"
  "    $xlsPath"
  "    $pdfPath"
  $drawings = @($s.drawings | Where-Object { $_ })
  if ($drawings.Count -gt 0) {
    $mergedPath = Join-Path $outDir ($base + '-含图纸.pdf')
    $mergeArgs = @('-I', (Join-Path $skill 'merge_drawings.py'), $pdfPath, $mergedPath) + $drawings
    & python @mergeArgs
    if ($LASTEXITCODE -ne 0) { throw 'drawing merge failed' }
    "    $mergedPath"
  }
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
