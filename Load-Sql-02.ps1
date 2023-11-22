function Global:Load-Sql-02 {
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
    
    $data_rows = Read-SqlTableData -ServerInstance "localhost\SQL2019" -DatabaseName "B52PDB" -SchemaName "dbo" -TableName "TinyCAD2"
       if (! $applyAML) { Add-content $output -Value "<AML>" }
     Foreach ($row in $data_rows){
	    $this_item = $innov.newItem("CAD", "add") 
	    $this_item.setID( (New-Guid).ToString().replace("-", "").ToUpper())
        Foreach ($col in $row.Table.Columns){ # Squirrely syntax here
            $this_col_name = $col.ColumnName
            $this_item.setProperty($this_col_name,$row[$this_col_name])
        }
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