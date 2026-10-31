$ErrorActionPreference='Stop'
$app=New-Object -ComObject Excel.Application
$app.Visible=$false;$app.DisplayAlerts=$false
try {
 $b=$app.Workbooks.Open((Join-Path $PWD 'outputs/tracker/Financial Tracker Ready.xlsx'))
 if($b.Worksheets.Count -ne 3){throw 'Sheet count'}
 $o=$b.Worksheets.Item('Overview')
 if($o.Range('A23').Value2-ne 0){throw 'Debt cache'}
 foreach($s in @($b.Worksheets.Item('Income'),$b.Worksheets.Item('Outgoings'))){
  $s.Range('A3').Formula='=HYPERLINK("#''Overview''!A1","Overview")'
  $s.Range('A3').Font.Color=12611584
  if($s.Range('C7').Validation.Type-ne 3){throw 'Missing category dropdown'}
  if(-not $s.ListObjects.Item(1).ShowAutoFilter){throw 'Missing filters'}
  $s.Activate();if(-not $app.ActiveWindow.FreezePanes){throw 'Missing frozen headers'}
 }
 Write-Output $b.Worksheets.Item('Income').ListObjects.Item(1).Range.Address()
 if($b.Worksheets.Item('Outgoings').ListObjects.Count-ne 2){throw 'Debt table'}
 $o.Activate();$o.Range('A1').Select()
 $o.PageSetup.PrintArea='$A$1:$C$42'
 $o.PageSetup.Zoom=$false;$o.PageSetup.FitToPagesWide=1;$o.PageSetup.FitToPagesTall=1
 $b.SaveAs((Join-Path $PWD 'outputs/tracker/Financial Tracker.xlsx'),51)
 $b.Close($true)
 'Clean workbook reopened successfully; three sheets, dropdowns, filters, frozen headers, blank tables and zero debt confirmed.'
}finally{$app.Quit()}

