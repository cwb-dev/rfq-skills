<#
Fill a packing-list template from a JSON spec and save <invoice_no>.xls + <invoice_no>.pdf.

  powershell -ExecutionPolicy Bypass -File make_pl.ps1 -Spec <spec.json> [-Force]

Spec (UTF-8 JSON):
{
  "seller": "runsun" | "cwb",
  "invoice_no": "26RSIT-RU-CMT-0010",
  "date": "2026/11/20",
  "buyer": ["CENTER OF HARDWARE TECHNOLOGIES LLC", "188693, Russia, ..."],
  "refs": [["PI NO.:", "RSIT-CMT-26-016"]],
  "qty_unit": "KPCS",
  "unit_weight_unit": "KGS/ KPCS",
  "qty_format": "#,##0.000",
  "items": [{
    "pallet_no": "1", "product": "DIN 934", "description": "ZINC\nM10",
    "qty": 10, "unit_weight": 10.6, "qty_per_ctn": 1, "ctns": 10,
    "gross": 110, "net": 106, "measurement": "60 X 40 X 40", "cbm": 0.96
  }],
  "expected_totals": {"qty": 10, "ctns": 10, "gross": 110, "net": 106, "cbm": 0.96},
  "shipping_mark": "CMT\n26RSIT-RU-CMT-0010\nPALLET NO. 1",
  "packing_summary": "10 CTNS",
  "notes": ["COUNTRY OF ORIGIN: CHINA."],
  "out_dir": "F:/Claude/<test-or-shipment-folder>"
}

Quantity, qty_per_ctn and unit_weight use the units named above. gross/net are line totals;
cbm is the total volume for that line, not the volume of one carton. If the source gives
only a shipment-level volume, leave item cbm empty and set "total_cbm": 0.4128.
#>
param([Parameter(Mandatory=$true)][string]$Spec, [switch]$Force)
$ErrorActionPreference = 'Stop'
$skill = Split-Path -Parent $MyInvocation.MyCommand.Path
$s = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json

$maps = @{
  runsun = @{ file='PL-RUNSUN.xls' }
  cwb    = @{ file='PL-CWB.xls' }
}
$m = $maps[$s.seller]
if (-not $m) { throw "seller must be 'runsun' or 'cwb'" }

$items = @($s.items)
$n = $items.Count
if ($n -lt 1) { throw 'no items' }
$qtyUnit = $(if ($s.qty_unit) { [string]$s.qty_unit } else { 'KPCS' })
$weightUnit = $(if ($s.unit_weight_unit) { [string]$s.unit_weight_unit } else { "KGS/ $qtyUnit" })
$totalQty = 0.0; $totalCtns = 0.0; $totalGross = 0.0; $totalNet = 0.0; $totalCbm = 0.0
$cbmIncomplete = $false
for ($i=0; $i -lt $n; $i++) {
  $it = $items[$i]
  foreach ($field in @('qty','ctns','gross','net')) {
    if ($null -eq $it.$field) { throw "items[$i].$field is required" }
  }
  if ([double]$it.gross -lt [double]$it.net) { throw "items[$i]: gross weight is below net weight" }
  $totalQty += [double]$it.qty
  $totalCtns += [double]$it.ctns
  $totalGross += [double]$it.gross
  $totalNet += [double]$it.net
  if ($null -eq $it.cbm) { $cbmIncomplete = $true } else { $totalCbm += [double]$it.cbm }
}
if ($null -ne $s.total_cbm) { $totalCbm = [double]$s.total_cbm }
elseif ($cbmIncomplete) { throw 'item cbm values are missing and total_cbm was not supplied' }
if ($s.expected_totals) {
  $e = $s.expected_totals
  $checks = @(@('qty',$totalQty),@('ctns',$totalCtns),@('gross',$totalGross),@('net',$totalNet),@('cbm',$totalCbm))
  foreach ($check in $checks) {
    $name = $check[0]; $actual = [double]$check[1]
    if ($null -ne $e.$name -and [math]::Abs($actual - [double]$e.$name) -gt 0.005) {
      throw "$name total $actual does not match expected_totals.$name $($e.$name)"
    }
  }
}

function Fmt([double]$v) { $v.ToString('#,##0.###', [Globalization.CultureInfo]::InvariantCulture) }
function FmtCbm([double]$v) { $v.ToString('#,##0.####', [Globalization.CultureInfo]::InvariantCulture) }
function Est([string]$t, [int]$chars) {
  $l = 0
  foreach ($p in ($t -split "`n")) { $l += [math]::Max(1, [math]::Ceiling($p.Length / $chars)) }
  $l
}

$outDir = $s.out_dir
New-Item -ItemType Directory -Force $outDir | Out-Null
$xlsPath = Join-Path $outDir ($s.invoice_no + '.xls')
$pdfPath = Join-Path $outDir ($s.invoice_no + '.pdf')
if (-not $Force -and ((Test-Path $xlsPath) -or (Test-Path $pdfPath))) { throw "$xlsPath or .pdf exists; use -Force to overwrite" }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.AutomationSecurity = 3
try {
  $wb = $xl.Workbooks.Open((Join-Path $skill "templates\$($m.file)"), 0, $true)
  $ws = $wb.Worksheets.Item(1)
  function Put([string]$addr, $v) {
    if ($v -is [psobject]) { $v = $v.psobject.BaseObject }
    $cell = $ws.Range($addr).MergeArea.Cells.Item(1)
    try { if ($v -is [string] -or $null -eq $v) { $cell.Value2 = [string]$v } else { $cell.Value2 = [double]$v } }
    catch { throw "Put $addr <- [$($v.GetType().Name)] '$v': $($_.Exception.Message)" }
  }

  Put 'K7' $s.invoice_no
  $noRange = $ws.Range('K7')
  $noRange.ShrinkToFit = $true
  if (([string]$s.invoice_no).Length -gt 22) { $noRange.Font.Size = 7 }
  $ws.Range('K8').NumberFormat = '@'
  Put 'K8' $s.date
  Put 'A11' (($s.buyer) -join "`n")
  for ($i=0; $i -lt 3; $i++) {
    $r = 11 + $i
    Put "H$r" ''
    Put "K$r" ''
  }
  $refs = @($s.refs | Where-Object { $_ })
  if ($refs.Count -gt 3) { throw 'at most 3 refs are supported' }
  for ($i=0; $i -lt $refs.Count; $i++) {
    $r = 11 + $i
    Put "H$r" $refs[$i][0]
    $ws.Range("K$r").NumberFormat = '@'
    Put "K$r" $refs[$i][1]
    $refRange = $ws.Range("K$r")
    $refRange.ShrinkToFit = $true
    if (([string]$refs[$i][1]).Length -gt 16) { $refRange.Font.Size = 8 }
  }
  Put 'E16' "Q'TY $qtyUnit"
  Put 'F16' $weightUnit
  Put 'G16' "$qtyUnit/ CTN"

  $itemRow = 17
  $itemCapacity = 20
  $templateTotalRow = 37
  $d = [math]::Max(0, $n - $itemCapacity)
  if ($d -gt 0) {
    $ws.Range("${templateTotalRow}:$($templateTotalRow+$d-1)").EntireRow.Insert() | Out-Null
    $ws.Rows.Item($templateTotalRow-1).Copy($ws.Range("${templateTotalRow}:$($templateTotalRow+$d-1)")) | Out-Null
  }
  $tr = $templateTotalRow + $d
  $lastTemplateItem = $templateTotalRow - 1 + $d
  $ws.Range("A${itemRow}:L$lastTemplateItem").ClearContents()
  $allInt = -not ($items | Where-Object { [double]$_.qty -ne [math]::Floor([double]$_.qty) })
  $qtyFmt = $(if ($s.qty_format) { $s.qty_format } elseif ($allInt) { '#,##0_ ' } else { '#,##0.000_ ' })
  for ($i=0; $i -lt $n; $i++) {
    $r = $itemRow + $i
    $it = $items[$i]
    Put "A$r" ([string]$it.pallet_no)
    Put "B$r" ([string]$it.product)
    Put "C$r" ([string]$it.description)
    Put "E$r" ([double]$it.qty)
    $ws.Range("E$r").NumberFormat = $qtyFmt
    Put "F$r" $(if ($null -ne $it.unit_weight) { [double]$it.unit_weight } else { '' })
    Put "G$r" $(if ($null -ne $it.qty_per_ctn) { [double]$it.qty_per_ctn } else { '' })
    Put "H$r" ([double]$it.ctns)
    $ws.Range("I$r").NumberFormat = '@'
    Put "I$r" ("$(Fmt ([double]$it.gross)) / $(Fmt ([double]$it.net))")
    $ws.Range("K$r").NumberFormat = '@'
    Put "K$r" ([string]$it.measurement)
    $lines = [math]::Max((Est ([string]$it.product) 17), (Est ([string]$it.description) 22))
    $ws.Rows.Item($r).RowHeight = [math]::Max(18, $lines * 14 + 5)
  }
  $ws.Range("E$tr").Formula = "=SUM(E${itemRow}:E$lastTemplateItem)"
  $ws.Range("E$tr").NumberFormat = $qtyFmt
  $ws.Range("H$tr").Formula = "=SUM(H${itemRow}:H$lastTemplateItem)"
  $ws.Range("I$tr").NumberFormat = '@'
  Put "I$tr" ("$(Fmt $totalGross) / $(Fmt $totalNet)")
  Put "K$tr" $totalCbm
  $ws.Range("K$tr").NumberFormat = '#,##0.0000'

  $summaryRow = 40 + $d
  if ($s.shipping_mark) {
    $ws.Range("A${summaryRow}:F$($summaryRow+4)").Merge()
    $c = $ws.Range("A$summaryRow")
    $c.Value2 = (($s.shipping_mark) -join "`n")
    $c.Font.Name = 'Arial'
    $c.Font.Size = 10
    $c.WrapText = $true
    $c.HorizontalAlignment = -4131
    $c.VerticalAlignment = -4160
  }
  Put "K$summaryRow" ("$(Fmt $totalQty) $qtyUnit")
  Put "K$($summaryRow+1)" $(if ($s.packing_summary) { [string]$s.packing_summary } else { "$(Fmt $totalCtns) CTNS" })
  Put "K$($summaryRow+2)" ("$(Fmt $totalGross) KGS")
  Put "K$($summaryRow+3)" ("$(Fmt $totalNet) KGS")
  Put "K$($summaryRow+4)" ("$(FmtCbm $totalCbm) CBM")

  $nf = 47 + $d
  $footer = 51 + $d
  $notes = @($s.notes | Where-Object { $_ })
  if ($notes.Count -gt ($footer - $nf)) {
    $k = $notes.Count - ($footer - $nf)
    $ws.Range("${footer}:$($footer+$k-1)").EntireRow.Insert() | Out-Null
    $footer += $k
  }
  $ws.Range("A${nf}:L$($footer-1)").UnMerge()
  $ws.Range("A${nf}:L$($footer-1)").ClearContents()
  for ($i=0; $i -lt $notes.Count; $i++) {
    $r = $nf + $i
    $ws.Range("A${r}:L$r").Merge()
    $c = $ws.Range("A$r")
    $c.Value2 = "$($i+1). $($notes[$i])"
    $c.Font.Name = 'Arial'
    $c.Font.Size = 10
    $c.HorizontalAlignment = -4131
    $c.VerticalAlignment = -4160
    $c.WrapText = $true
    $ws.Rows.Item($r).RowHeight = (Est $c.Value2 95) * 14 + 6
  }

  $xl.Calculate()
  $sheetQty = [double]$ws.Range("E$tr").Value2
  $sheetCtns = [double]$ws.Range("H$tr").Value2
  if ([math]::Abs($sheetQty - $totalQty) -gt 0.005) { throw "quantity mismatch: sheet $sheetQty vs spec $totalQty" }
  if ([math]::Abs($sheetCtns - $totalCtns) -gt 0.005) { throw "carton mismatch: sheet $sheetCtns vs spec $totalCtns" }
  $wb.SaveAs($xlsPath, 56)
  $wb.ExportAsFixedFormat(0, $pdfPath)
  $wb.Close($false)
  "OK  qty=$(Fmt $totalQty) $qtyUnit  ctns=$(Fmt $totalCtns)  gross=$(Fmt $totalGross) KG  net=$(Fmt $totalNet) KG  cbm=$(FmtCbm $totalCbm)  items=$n"
  "    $xlsPath"
  "    $pdfPath"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
