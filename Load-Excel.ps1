function Global:Load-Excel {
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $sd
        ,[parameter(Mandatory=$true)]
         [String]
         $ExcelFile        
        ,[parameter(Mandatory=$true)]
         [Boolean]
         $applyAML
        ,[parameter(Mandatory=$true)]
         [String]
         $output
        ,[parameter(Mandatory=$true)]
         [Object]
         $innov
        ,[parameter(Mandatory=$true)]
         [String]
         $ignore_pfx
    )
    Set-Location $sd
    .\Get-PropItems.ps1 # returns hash table of properties of type Item for the currrent sheet

    $xl = Open-ExcelPackage -Path $ExcelFile
    $regex1 = "([a-z0-9_]*)\(([a-z0-9_]*)\)" # used to parse property names
    $sht_ct= $xl.Workbook.Worksheets.Count
    if (! $applyAML) {Add-content $output -Value "<AML>" } # open aml
    For ($s=1; $s -le $sht_ct; $s++) # for each sheet
    {
        $this_sht = $xl.Workbook.Worksheets[$s]
        if (-not (($this_sht.Name).StartsWith($ignore_pfx))) # ignore sheets named to ignore
        {
            $row_ct = $this_sht.Dimension.Rows
            $col_ct = $this_sht.Dimension.Columns
            # ///TODO error if Dimension is wrong
            #Get-PropItems gets column names and types for properties of type Item in this sheet
            $propTypes= Get-PropItems -this_sht $this_sht -innov $innov

            # step 5.4 iterate across sheets, and their rows to load data
            for ($r = 2 ; $r -le $row_ct; $r++) # for each row
            {   
                $this_item = $innov.newItem($this_sht.Name,"add") 
                $this_item.setID( (New-Guid).ToString().replace("-","").ToUpper())
                for ($c = 1; $c -le $col_ct; $c++) # for each column
                {
                    $this_col_name = $this_sht.Cells[1, $c].Value
                    if (-not ($this_col_name.StartsWith($ignore_pfx))) # ignore columns starting with _
                    {
                        $prop= [regex]::match($this_col_name, $regex1).Groups[1].Value
                        $rel_prop = [regex]::match($this_col_name, $regex1).Groups[2].Value
                        if (-not ($this_col_name -eq "physical_file")) # ignore column "physical_file", reserved for loading Files
                        {
                            $this_cell_value = $this_sht.Cells[$r,$c].Value
                            if ($this_col_name -match $regex1)
                            {           
                               $this_prop_item = $innov.newItem( $propTypes.$prop) # Create a new Item for data_source
                               if (-not ($propTypes.$prop -eq "File")) 
                                {   # all data_sources except File
                                    $this_prop_item.setProperty($rel_prop, $this_cell_value)
                                    $res = $this_item.setPropertyItem($prop , $this_prop_item) 
                                    $res.setAction("get")
                                }
                                else
                                {   <# use IOM.setFileProperty(...), special handling for type File      #
                                     # it creates a new File Item, with new guid created automatically,  #
                                     # and Located relationship with id of vault for user.               #
                                     # physical file will be loaded, which cannot be done with AML alone #
                                     # any File replaced may be orphaned                                 #>
                                    $filepath =   $this_sht.Cells[$r,$Global:pf_col].Value
                                    $this_item.setFileProperty("primary_file",$filepath)
                                }
                            }
                            elseif ($this_cell_value -ne $null) { $this_item.setProperty($this_col_name , $this_Cell_value) }
                         }
                    }
                }
                if ($applyAML) # apply item
                {
                    $res = $this_item.apply()
                    if ($res.isError()) { $this_item.dom.OuterXml | Out-Host ; $res.dom.OuterXml | Out-Host ; Exit}
                }
                else # or write AML to file
                {
                    Add-content $output -Value $this_item.node.OuterXml
                }
            }
        }
    }
    if (! $applyAML) {Add-content $output -Value "</AML>" } # close aml
    Close-ExcelPackage -ExcelPackage $xl -NoSave # close the Excel File, to release it. File will not open with other apps if not closed here.   
 }
 