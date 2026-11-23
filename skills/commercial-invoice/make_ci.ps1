<#
Fill a commercial-invoice template from a JSON spec and save <no>.xls + <no>.pdf.

  powershell -ExecutionPolicy Bypass -File make_ci.ps1 -Spec <spec.json> [-Force]

Spec (UTF-8 JSON):
{
  "seller": "runsun" | "cwb",
  "no": "26RSIT-XX-ABC-0001",
  "date": "2026/11/20",
  "buyer": ["EXAMPLE BUYER LLC", "1 EXAMPLE STREET, EXAMPLE CITY, COUNTRY"],
  "refs": [["PI NO.:", "RSIT-ABC-26-001"], ["CONTRACT NO.:", "EXAMPLE-001"]],
  "incoterm": "FOB SHANGHAI",
  "currency": "CNY",
  "qty_unit": "MPCS", "price_unit": "CNY/MPCS",
  "weight_unit": "KG/MPCS",
  "qty_format": "#,##0", "price_format": "#,##0.00",
  "items": [{"product": "DIN 934", "description": "ZINC\nM10", "qty": 10, "price": 100, "weight": 10.6}],
  "expected_total": 1000,
  "notes": ["PAYMENT: ...", "COUNTRY OF ORIGIN: CHINA."],
  "buyer_sign_name": "EXAMPLE BUYER LLC",
  "say_total": true,
  "out_dir": "./output"
}
#>
param([Parameter(Mandatory=$true)][string]$Spec, [switch]$Force)
$ErrorActionPreference = 'Stop'
$skill = Split-Path -Parent $MyInvocation.MyCommand.Path
$s = Get-Content -Raw -Encoding UTF8 $Spec | ConvertFrom-Json

$maps = @{
  runsun = @{ file='CI-RUNSUN.xls'; noCell='K7'; dateCell='K8'; buyerCell='A11'; refLabel='H'; refVal='K'; refRow=11
              incoterm='J16'; currency='J17'; qtyUnit='F17'; weightHdr='G15'; weightUnit='G17'; priceUnit='H17'
              itemRow=18; itemCapacity=8; totalRow=26; cNo='A'; cProduct='B'; cDesc='C'; cQty='F'; cWeight='G'; cPrice='H'; cAmount='J'
              productChars=17; descChars=34; notesFirst=38; signRow=44; buyerSign='A'; lastCol='L' }
  cwb    = @{ file='CI-CWB.xls'; noCell='K7'; dateCell='K8'; buyerCell='A11'; refLabel='H'; refVal='K'; refRow=11
              incoterm='J16'; currency='J17'; qtyUnit='F17'; weightHdr=$null; weightUnit=$null; priceUnit='H17'
              itemRow=18; itemCapacity=8; totalRow=26; cNo='A'; cProduct='B'; cDesc='C'; cQty='F'; cWeight=$null; cPrice='H'; cAmount='J'
              productChars=17; descChars=34; notesFirst=39; signRow=46; buyerSign='A'; lastCol='L' }
}
$m = $maps[$s.seller]
if (-not $m) { throw "seller must be 'runsun' or 'cwb'" }
if ($s.seller -eq 'runsun' -and $s.currency -ne 'CNY') { throw 'runsun commercial invoices use CNY' }
if ($s.seller -eq 'cwb' -and $s.currency -ne 'USD') { throw 'cwb commercial invoices use USD' }

function Words([long]$n) {
  $ones = 'ZERO ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ELEVEN TWELVE THIRTEEN FOURTEEN FIFTEEN SIXTEEN SEVENTEEN EIGHTEEN NINETEEN'.Split(' ')
  $tens = '_ _ TWENTY THIRTY FORTY FIFTY SIXTY SEVENTY EIGHTY NINETY'.Split(' ')
  if ($n -lt 20) { return $ones[$n] }
  if ($n -lt 100) { return $tens[[math]::Floor($n/10)] + $(if ($n%10) { '-' + $ones[$n%10] } else { '' }) }
  if ($n -lt 1000) { return $ones[[math]::Floor($n/100)] + ' HUNDRED' + $(if ($n%100) { ' AND ' + (Words ($n%100)) } else { '' }) }
  foreach ($u in @(@(1000000000,'BILLION'),@(1000000,'MILLION'),@(1000,'THOUSAND'))) {
    if ($n -ge $u[0]) { $r = $n % $u[0]; return (Words ([math]::Floor($n/$u[0]))) + ' ' + $u[1] + $(if ($r) { $(if ($r -lt 100) {' AND '} else {' '}) + (Words $r) } else { '' }) }
  }
}
$curName = @{ CNY='CHINESE YUAN'; USD='US DOLLARS' }
$centName = @{ CNY='FEN'; USD='CENTS' }

$items = @($s.items)
$n = $items.Count
if ($n -lt 1) { throw 'no items' }
$total = 0.0
foreach ($it in $items) { $total += [math]::Round([double]$it.qty * [double]$it.price, 2) }
if ($null -ne $s.expected_total -and [math]::Abs($total - [double]$s.expected_total) -gt 0.005) {
  throw "item total $total does not match expected_total $($s.expected_total)"
}
$outDir = $s.out_dir
New-Item -ItemType Directory -Force $outDir | Out-Null
$xlsPath = Join-Path $outDir ($s.no + '.xls')
$pdfPath = Join-Path $outDir ($s.no + '.pdf')
if (-not $Force -and ((Test-Path $xlsPath) -or (Test-Path $pdfPath))) { throw "$xlsPath or .pdf exists; use -Force to overwrite" }

function Est([string]$t, [int]$chars) {
  $l = 0
  foreach ($p in ($t -split "`n")) { $l += [math]::Max(1, [math]::Ceiling($p.Length / $chars)) }
  $l
}

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

  Put $m.noCell $s.no
  $noRange = $ws.Range($m.noCell)
  $noRange.ShrinkToFit = $true
  if (([string]$s.no).Length -gt 22) { $noRange.Font.Size = 7 }
  $ws.Range($m.dateCell).NumberFormat = '@'
  Put $m.dateCell $s.date
  Put $m.buyerCell (($s.buyer) -join "`n")
  for ($i=0; $i -lt 3; $i++) {
    $r = $m.refRow + $i
    Put "$($m.refLabel)$r" ''
    Put "$($m.refVal)$r" ''
  }
  $refs = @($s.refs | Where-Object { $_ })
  if ($refs.Count -gt 3) { throw 'at most 3 refs are supported' }
  for ($i=0; $i -lt $refs.Count; $i++) {
    $r = $m.refRow + $i
    Put "$($m.refLabel)$r" $refs[$i][0]
    $ws.Range("$($m.refVal)$r").NumberFormat = '@'
    Put "$($m.refVal)$r" $refs[$i][1]
    $refRange = $ws.Range("$($m.refVal)$r")
    $refRange.ShrinkToFit = $true
    if (([string]$refs[$i][1]).Length -gt 16) { $refRange.Font.Size = 8 }
  }
  Put $m.incoterm $s.incoterm
  Put $m.currency $s.currency
  Put $m.qtyUnit $s.qty_unit
  Put $m.priceUnit $s.price_unit
  if ($m.weightHdr) {
    if ($s.weight_unit) { Put $m.weightHdr 'WEIGHT'; Put $m.weightUnit $s.weight_unit }
    else { Put $m.weightHdr ''; Put $m.weightUnit '' }
  }

  $d = [math]::Max(0, $n - $m.itemCapacity)
  if ($d -gt 0) {
    $insertAt = $m.totalRow
    $ws.Range("${insertAt}:$($insertAt+$d-1)").EntireRow.Insert() | Out-Null
    $ws.Rows.Item($m.totalRow-1).Copy($ws.Range("${insertAt}:$($insertAt+$d-1)")) | Out-Null
  }
  $ir = $m.itemRow
  $tr = $m.totalRow + $d
  $lastTemplateItem = $m.totalRow - 1 + $d
  $ws.Range("A${ir}:$($m.lastCol)$lastTemplateItem").ClearContents()
  $allInt = -not ($items | Where-Object { [double]$_.qty -ne [math]::Floor([double]$_.qty) })
  $qtyFmt = $(if ($s.qty_format) { $s.qty_format } elseif ($allInt) { '#,##0_ ' } else { '#,##0.000_ ' })
  for ($i=0; $i -lt $n; $i++) {
    $r = $ir + $i
    $it = $items[$i]
    Put "$($m.cNo)$r" ($i+1)
    Put "$($m.cProduct)$r" ([string]$it.product)
    Put "$($m.cDesc)$r" ([string]$it.description)
    Put "$($m.cQty)$r" ([double]$it.qty)
    $ws.Range("$($m.cQty)$r").NumberFormat = $qtyFmt
    if ($m.cWeight) { Put "$($m.cWeight)$r" $(if ($null -ne $it.weight) { [double]$it.weight } else { '' }) }
    Put "$($m.cPrice)$r" ([double]$it.price)
    if ($s.price_format) { $ws.Range("$($m.cPrice)$r").NumberFormat = $s.price_format }
    $ws.Range("$($m.cAmount)$r").Formula = "=ROUND($($m.cPrice)$r*$($m.cQty)$r,2)"
    $lines = [math]::Max((Est ([string]$it.product) $m.productChars), (Est ([string]$it.description) $m.descChars))
    $ws.Rows.Item($r).RowHeight = [math]::Max(24, $lines * 14 + 8)
  }
  $ws.Range("$($m.cQty)$tr").Formula = "=SUM($($m.cQty)${ir}:$($m.cQty)$lastTemplateItem)"
  $ws.Range("$($m.cQty)$tr").NumberFormat = $qtyFmt
  $ws.Range("$($m.cAmount)$tr").Formula = "=SUM($($m.cAmount)${ir}:$($m.cAmount)$lastTemplateItem)"

  if ($s.say_total) {
    $cur = [string]$s.currency
    $int = [long][math]::Floor($total)
    $cents = [int][math]::Round(($total - $int) * 100)
    $w = "SAY TOTAL: $($curName[$cur]) $(Words $int)" + $(if ($cents) { " AND $(Words $cents) $($centName[$cur])" } else { '' }) + ' ONLY.'
    $ws.Range("B$($tr+1):$($m.lastCol)$($tr+1)").Merge()
    $c = $ws.Range("B$($tr+1)")
    $c.Value2 = $w
    $c.Font.Name = 'Arial'
    $c.Font.Size = 9
    $c.Font.Italic = $true
    $c.Font.Color = 0x555555
    $c.HorizontalAlignment = -4131
  }

  $nf = $m.notesFirst + $d
  $sr = $m.signRow + $d
  $notes = @($s.notes | Where-Object { $_ })
  if ($notes.Count -gt ($sr - $nf)) {
    $k = $notes.Count - ($sr - $nf)
    $ws.Range("${sr}:$($sr+$k-1)").EntireRow.Insert() | Out-Null
    $sr += $k
  }
  $ws.Range("A${nf}:$($m.lastCol)$($sr-1)").UnMerge()
  $ws.Range("A${nf}:$($m.lastCol)$($sr-1)").ClearContents()
  for ($i=0; $i -lt $notes.Count; $i++) {
    $r = $nf + $i
    $ws.Range("A${r}:$($m.lastCol)$r").Merge()
    $c = $ws.Range("A$r")
    $c.Value2 = "$($i+1). $($notes[$i])"
    $c.Font.Name = 'Arial'
    $c.Font.Size = 10
    $c.HorizontalAlignment = -4131
    $c.VerticalAlignment = -4160
    $c.WrapText = $true
    $ws.Rows.Item($r).RowHeight = (Est $c.Value2 95) * 14 + 6
  }
  if ($s.buyer_sign_name) {
    $c = $ws.Range("$($m.buyerSign)$sr")
    $c.Value2 = $s.buyer_sign_name
    $c.Font.Name = 'Arial'
    $c.Font.Size = 10
    $c.Font.Bold = $true
    $c.HorizontalAlignment = -4131
  }

  $xl.Calculate()
  $calc = [double]$ws.Range("$($m.cAmount)$tr").Value2
  if ([math]::Abs($calc - $total) -gt 0.005) { throw "total mismatch: sheet $calc vs spec $total" }
  $wb.SaveAs($xlsPath, 56)
  $wb.ExportAsFixedFormat(0, $pdfPath)
  $wb.Close($false)
  "OK  total=$('{0:N2}' -f $total) $($s.currency)  items=$n"
  "    $xlsPath"
  "    $pdfPath"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
