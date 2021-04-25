function Global:ExcelReport {
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $sd
        ,[parameter(Mandatory=$true)]
         [Object]
         $innov
        ,[parameter(Mandatory=$true)]
         [String]
         $report_script
        ,[parameter(Mandatory=$true)]
         [String]
         $report_uri
    )
    Set-Location $sd

    & $report_script -sd $sd -innov $innov -report_uri $report_uri
 }
 