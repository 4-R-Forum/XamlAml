function Global:ExcelReport-InBasket {
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
Add-Worksheet -ExcelWorkbook $excel.Workbook -WorksheetName InBasketStatus
$pivot = $excel.Workbook.Worksheets["InBasketStatus"]
$xl = $excel.Workbook.Worksheets["Sheet1"]
$xl.Cells[1,1].Value = "Assigned"
$xl.Cells[1,2].Value = "Task"
$xl.Cells[1,3].Value = "Status"
# get data from MilHub
$aml1 = @"
<AML>
	<Item type='InBasket Task' action='get' page='1' select='assigned_to,container,status'>
		<OR>
			<status condition='eq'>Active</status>
			<status condition='eq'>Closed</status>
		</OR>
		<container>
			<Item action='get'>
				<OR>
					<keyed_name condition='like'>*AR*</keyed_name>
					<keyed_name condition='like'>*CCP*</keyed_name>
				</OR>
			</Item>
		</container>
	</Item>
</AML>
"@
$res = $innov.applyAML($aml1)
if ($res.isError()) { Exit }
# populate the sheet
for ($i=0; $i -lt $res.getItemCOunt();$i++){
    $this_task = $res.getItemByIndex($i)
    $xl.Cells[($i + 2),1].Value = $this_task.getPropertyAttribute("assigned_to","keyed_name")
    $xl.Cells[($i + 2),2].Value = $this_task.getPropertyAttribute("container","keyed_name")
    $xl.Cells[($i + 2),3].Value = $this_task.getProperty("status")
}

#Add a pivot table, 
$source_range = $xl.Dimension.Address
Add-PivotTable -PivotTableName InBasketStatus -Activate -Address $pivot.Cells["A1"] -PassThru -PivotColumns Status -PivotData Task -PivotRows Assigned -PivotTotals Rows -SourceRange $source_range -SourceWorksheet $xl
#Save and open in excel
Select-Worksheet -ExcelPackage $excel -WorksheetName InBasketStatus
Close-ExcelPackage $excel  $report_uri -Show 
 }
 