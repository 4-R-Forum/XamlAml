function Global:Get-PropItems{
    Param(
         [parameter(Mandatory=$true)]
         [Object]
         $this_sht
        , [parameter(Mandatory=$true)]
         [Object]
         $innov
    )        
    # put column names like property_name(Type) in comma separated_list
    # to be used in AML property condition attribute
    $prop_names = "" # comma delimited list of properties of type Item
    $propTypes = @{} # hashtable to be returned
    for ($c = 1; $c -le $col_ct; $c++) # for each column
    {
        $this_col_name = $this_sht.Cells[1, $c].Value
        if ($this_col_name -eq "physical_file") { $Global:pf_col =  $c} # save physical_file column for this sheet
        if ($this_col_name -match $regex1) # add to list of properties of type item
        {           
            $prop= [regex]::match($this_col_name, $regex1).Groups[1].Value # ///TODO function
            if ($prop_names -ne "") {$prop_names += ","}
            $prop_names += ("'" + $prop + "'")            
        }
    }
    if ($prop_names -ne "") # get properties of type Item and find data_sources, if any 
    {
        $this_type = $innov.newItem("ItemType","get")
        $this_type.setProperty("name", $this_sht.Name)
        $this_type.setAttribute("select", "is_relationship")
        $this_type_rel = $innov.newItem("Property","get")
        $this_type_rel.setProperty("name",$prop_names)
        $this_type_rel.setPropertyAttribute("name", "condition","in")
        $this_type_rel.setAttribute("select","name,data_source")
        $this_type.addRelationship($this_type_rel)
        # get all types for this sheet from the server
        $res_type = $this_type.apply()
        if ($res_type.isError()) { $res_type.dom.OuterXml | Out-Host ; Exit} # report error and exit script
        $prop_items = $res_type.getItemsByXPath("//Item[@type='Property']");
        # populate hashtable key=prop-name value=itemtype-name       
        for ($p = 0; $p -lt $prop_items.getItemCount(); $p++)
        {
            $this_prop = $prop_items.getItemByIndex($p)
            $propTypes.Add( $this_prop.getProperty("name"), $this_prop.getPropertyAttribute("data_source","keyed_name"))
        }
    }
    return $propTypes # empty if none
}