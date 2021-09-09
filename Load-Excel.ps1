function Global:Load-Excel {
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $sd
        , [parameter(Mandatory = $true)]
        [String]
        $ExcelFile        
        , [parameter(Mandatory = $true)]
        [Boolean]
        $applyAML
        , [parameter(Mandatory = $true)]
        [String]
        $output
        , [parameter(Mandatory = $true)]
        [Object]
        $innov
        , [parameter(Mandatory = $true)]
        [String]
        $ignore_pfx
    )
    Set-Location $sd
    .\Get-PropItems.ps1 

    $xl = Open-ExcelPackage -Path $ExcelFile
    $regex1 = "([a-z0-9_]*)\(([a-z0-9_]*)\)" 
    $sht_ct = $xl.Workbook.Worksheets.Count
    $itemtype_property_value = @{}

    if (! $applyAML) { Add-content $output -Value "<AML>" }
    For ($s = 1; $s -le $sht_ct; $s++) {
        
        $this_sht = $xl.Workbook.Worksheets[$s]
        if (-not (($this_sht.Name).StartsWith($ignore_pfx))) {
            
            $row_ct = $this_sht.Dimension.Rows
            $col_ct = $this_sht.Dimension.Columns
           
            $propTypes = Get-PropItems -this_sht $this_sht -innov $innov -ignore_pfx $ignore_pfx
           
            for ($r = 2 ; $r -le $row_ct; $r++) {
  
                $this_item = $innov.newItem($this_sht.Name, "add") 
                $this_item.setID( (New-Guid).ToString().replace("-", "").ToUpper())
                for ($c = 1; $c -le $col_ct; $c++) {
                    
                    $this_col_name = $this_sht.Cells[1, $c].Value
                    if (-not ($this_col_name.StartsWith($ignore_pfx))) {
                        
                        $prop = [regex]::match($this_col_name, $regex1).Groups[1].Value
                        $rel_prop = [regex]::match($this_col_name, $regex1).Groups[2].Value
                        if (-not ($this_col_name -eq "physical_file")) {
                            
                            $this_cell_value = $this_sht.Cells[$r, $c].Value
                            if (-not ([string]::IsNullOrEmpty($this_cell_value))) {
                                
                                if ($this_col_name -match $regex1) {
                                    
                                    $itemtype = $propTypes.$prop
                                    
                                    $this_prop_item = $innov.newItem( $propTypes.$prop) # Create a new Item for data_source
                                    if (-not ($itemtype -eq "File")) {
                                        $key = "${itemtype}_${rel_prop}_${this_cell_value}"
                                        if (-not $itemtype_property_value.ContainsKey($key)) {
                                            $this_prop_item = $innov.newItem( $propTypes.$prop, "get")
                                            $this_prop_item.setProperty($rel_prop, $this_cell_value)
                                            $this_prop_item.setAttribute("select", "id")
                                            $this_prop_item.setPropertyAttribute($rel_prop, "condition", "eq")
                                            $this_prop_item_res = $this_prop_item.apply()
                                            if ($this_prop_item_res.isError()) {
                                                $this_prop_item_res.dom.OuterXml | Out-Host ; Exit
                                            }
                                            $item_count = $this_prop_item_res.getItemCount()
                                            if($item_count -gt 1){
                                                "Query returned more than 1 item for itemtype: ${itemtype} using property ${rel_prop} with value ${this_cell_value} and attribute 'condition eq' on ${rel_prop}" | Out-Host ; Exit
                                            }
                                            if($item_count -eq 0){
                                                "Query returned 0 items for itemtype: ${itemtype} using property ${rel_prop} with value ${this_cell_value} and attribute 'condition eq' on ${rel_prop}" | Out-Host ; Exit
                                            }
                                            $itemtype_property_value[$key] = $this_prop_item_res.getItemByIndex(0).getProperty("id")
                                        }
                                        $this_item.setProperty($prop , $itemtype_property_value[$key]) 
                                    }
                                    else {
                                        <# use IOM.setFileProperty(...), special handling for type File      #
                                         # it creates a new File Item, with new guid created automatically,  #
                                         # and Located relationship with id of vault for user.               #
                                         # physical file will be loaded, which cannot be done with AML alone #
                                         # any File replaced may be orphaned                                 #>
                                        $filepath = $this_sht.Cells[$r, $Global:pf_col].Value
                                        $this_item.setFileProperty($prop, $filepath)
                                    }
                                }
                                else {
                                    $this_item.setProperty($this_col_name , $this_Cell_value)
                                } 
                            }
                        }
                    }
                }
                if ($applyAML) {
                    
                    $res = $this_item.apply()
                    if ($res.isError()) { $this_item.dom.OuterXml | Out-Host ; $res.dom.OuterXml | Out-Host ; Exit }
                }
                else {
                    
                    Add-content $output -Value $this_item.node.OuterXml
                }
            }
        }
    }
    if (! $applyAML) { Add-content $output -Value "</AML>" } 
    Close-ExcelPackage -ExcelPackage $xl -NoSave
}
 