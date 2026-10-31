$ErrorActionPreference='Stop'
$app=New-Object -ComObject Excel.Application
$app.Visible=$false
$app.DisplayAlerts=$false
try {
 $path=Join-Path $PWD 'outputs/tracker/Financial Tracker.xlsx'
 $book=$app.Workbooks.Open($path,0,$false,5,"","",$true,1,"",$false,$false,0,$false,$true,1)
 $o=$book.Worksheets.Item('Overview');$i=$book.Worksheets.Item('Income');$e=$book.Worksheets.Item('Outgoings')
 $o.Columns.Item('D:XFD').Hidden=$true
 $o.Range('A19:A20').WrapText=$true;$o.Range('A19:A20').RowHeight=34
 $i.Range('B3:E3').Merge();$i.Range('B4:E4').Merge()
 $e.Range('B3:F3').Merge();$e.Range('B4:F4').Merge();$e.Range('H3:L3').Merge()
 $o.Range('B4').Validation.Delete();$o.Range('B4').Validation.Add(4,1,7,'=$A$31','31/12/9999')
 $o.Range('A31').Validation.Delete();$o.Range('A31').Validation.Add(7,1,1,'=AND(ISNUMBER(A31),DAY(A31)=1,A31>0)')
 foreach($s in @($i,$e)) {
  $s.Range('D7:D10000').Validation.ShowError=$true
  $s.Range('C7:C10000').Validation.ShowError=$true
 }
 $e.Range('E7:E10000').Validation.ShowError=$true
 function DateNum($y,$m,$d){([datetime]::new($y,$m,$d)).ToOADate()}
 function Check($cell,$expected){$app.CalculateFullRebuild();$actual=$o.Range($cell).Value2;if([math]::Abs($actual-$expected)-gt 0.001){throw "$cell actual $actual expected $expected"}}
 $it=$i.ListObjects.Item('IncomeTable');$ot=$e.ListObjects.Item('OutgoingsTable')
 $it.Resize($i.Range('A6:E7'));$ot.Resize($e.Range('A6:F7'))
 $i.Range('A7').Value2=DateNum 2026 10 1;$i.Range('C7').Value2='Salary';$i.Range('D7').Value2=1000.0
 $null=$it.ListRows.Add();$i.Range('A8').Value2=DateNum 2026 10 31;$i.Range('C8').Value2='Freelance';$i.Range('D8').Value2=200.0
 $null=$it.ListRows.Add();$i.Range('A9').Value2=DateNum 2026 10 8;$i.Range('C9').Value2='Transfer between own accounts';$i.Range('D9').Value2=900.0
 $null=$it.ListRows.Add();$i.Range('A10').Value2=DateNum 2026 11 1;$i.Range('C10').Value2='Salary';$i.Range('D10').Value2=300.0
 $null=$it.ListRows.Add();$i.Range('A11').Value2=DateNum 2027 1 1;$i.Range('C11').Value2='Interest';$i.Range('D11').Value2=25.0
 $e.Range('A7').Value2=DateNum 2026 10 31;$e.Range('C7').Value2='Debt Repayments';$e.Range('D7').Value2=100.0
 $null=$ot.ListRows.Add();$e.Range('A8').Value2=DateNum 2026 10 8;$e.Range('C8').Value2='Transfer between own accounts';$e.Range('D8').Value2=900.0
 $null=$ot.ListRows.Add();$e.Range('A9').Value2=DateNum 2026 11 1;$e.Range('C9').Value2='Groceries';$e.Range('D9').Value2=50.0
 $e.Range('I7').Value2=800.0;$o.Range('A33').Value2=500.0
 Check 'B35' 1200;Check 'C35' 100;Check 'A13' 1100;Check 'A17' 1600;Check 'A23' 800
 $i.Range('C8').Value2='Gifts';Check 'B11' 0;Check 'B15' 200
 $o.Range('B4').Value2=DateNum 2026 11 1;Check 'B35' 300;Check 'C35' 50;Check 'A15' 1600;Check 'A17' 1850
 $o.Range('B4').Value2=DateNum 2027 1 1;Check 'B35' 25;Check 'A15' 1850;Check 'A17' 1875
 $i.Range('A7:E11').ClearContents();$e.Range('A7:F9').ClearContents();$e.Range('H7:L7').ClearContents();$it.Resize($i.Range('A6:E7'));$ot.Resize($e.Range('A6:F7'))
 $o.Range('A33').Value2=0.0;$o.Range('B4').Value2=DateNum 2026 10 1
 Check 'B35' 0;Check 'C35' 0;Check 'A17' 0;Check 'A23' 0
 foreach($s in @($o,$i,$e)){
  try{$errors=$s.UsedRange.SpecialCells(-4123,16);if($errors.Count-gt 0){throw "Formula errors in $($s.Name)"}}catch{if($_.Exception.Message -like 'Formula errors*'){throw}}
  $s.Activate();$app.ActiveWindow.Zoom=85;$s.Range('A1').Select()
 }
 $o.Activate();$o.Range('B4').Select();$app.Calculation=-4105
 $book.SaveAs((Join-Path $PWD "outputs/tracker/Financial Tracker Ready.xlsx"),51)
 Write-Output 'Excel native tests passed: row expansion, category edits, transfers, debt separation, month/year changes, carry-forward and blank zero summaries.'
 $book.Close($true)
} finally {$app.Quit()}


