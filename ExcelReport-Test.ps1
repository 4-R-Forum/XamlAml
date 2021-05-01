function Global:ExcelReport-Test {
    Param(
         [parameter(Mandatory=$true)]
         [Object]
         $innov
        ,[parameter(Mandatory=$true)]
         [String]
         $report_uri
    )
# Sample code from https://github.com/dfinke/ImportExcel/blob/master/Examples/PivotTable/TableAndPivotTable.ps1

# Delete existing sheet
Remove-Item $report_uri -ErrorAction Ignore

# create Excel Workbook
$excel = Export-Excel -PassThru -Path  $xlReportFile
$pivot = $excel.Workbook.Worksheets["InBasketStatus"]
$xl = $excel.Workbook.Worksheets["Sheet1"]
$xl.Cells[1,1].Value = "LoginName"
$xl.Cells[1,2].Value = "KeyedName"
$xl.Cells[1,3].Value = "Email"
# get data from MilHub
$aml1 = @"
<AML>
	<Item type='USer' action='get' page='1' select='login_name,keyed_name,email' />
</AML>
"@
$res = $innov.applyAML($aml1)
if ($res.isError()) { Exit }
# populate the sheet
for ($i=0; $i -lt $res.getItemCOunt();$i++){
    $this_item = $res.getItemByIndex($i)
    $xl.Cells[($i + 2),1].Value = $this_item.getProperty("login_name")
    $xl.Cells[($i + 2),2].Value = $this_item.getProperty("keyed_name")
    $xl.Cells[($i + 2),3].Value = $this_item.getProperty("email","")
}

#Save and open in excel
#
Close-ExcelPackage $excel  $report_uri -Show 
 }
 