function Global:Load-Sql {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $sd
      , [parameter(Mandatory = $true)]
        [Object]
        $innov
    )
    Set-Location $sd
    $applyAML = $false
    $output = "C:\Users\JonHodge\OneDrive - Anautics, Inc\Workspace\2022Projects\PStruct\Staging\aml.xml"


    $data_rows = Read-SqlTableData -ServerInstance "localhost\SQL2019" -DatabaseName "B52PDB" -SchemaName "dbo" -TableName "TinyPart"
    $data_columns = $data_rows.Columns
    if (! $applyAML) { Add-content $output -Value "<AML>" }

    Foreach ($row in $data_rows){
	    $this_item = $innov.newItem("Part", "add") 
	    $this_item.setID( (New-Guid).ToString().replace("-", "").ToUpper())
        $this_item.setProperty("is_ci",$row[0])
	    $this_item.setProperty("j_cage_code",$row[1])
        $this_item.setProperty("j_part_number",$row[2])
        $this_item.setProperty("mh_smr_code",$row[3])
        $this_item.setProperty("name",$row[4])
        $this_item.setProperty("smr_s",$row[7])
        $this_item.setProperty("smr_r",$row[8])
        $this_item.setProperty("oem_rev",$row[9])
        $this_item.setProperty("supply_plan",$row[10])        
        $this_item.setProperty("classification",$row[11])
        $this_item.setProperty("fsc",$row[12])
        $this_item.setProperty("niin",$row[13])

        if ($applyAML) {
             $res = $this_item.apply()
                if ($res.isError()) { $this_item.dom.OuterXml | Out-Host ; $res.dom.OuterXml | Out-Host ; Exit }
         }
         else {
	        Add-content $output -Value $this_item.node.OuterXml
        }
     }
     if (! $applyAML) { Add-content $output -Value "</AML>" }
}