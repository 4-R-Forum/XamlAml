function Audit-Package
{
    $pa =  $tbPackage.Text
    if ($pd -eq "")
    { [System.Windows.MessageBox]::Show("Please enter Packge Name(s) as comma separated string")}
    else
    {
      # remove the stored procedure, and temp tables if they exist
      $s_qry =
@"
      if exists (
        SELECT id FROM sysobjects
        WHERE  sysobjects.name = 'SelectPackageElementForItemType'
        and sysobjects.type='P'
      )
      DROP PROCEDURE SelectPackageElementForItemType
"@
      Invoke-Sqlcmd -ServerInstance $s_inst -Database $s_db -Username $s_user -Password $s_pw -Query $s_qry     
      
      # create a stored procedure for one time use by this script
        $qry_fname=resolve-path -path "./SelectPackagElementForItemType.sql"
        $s_qry = [IO.File]::ReadAllText($qry_fname)
        Invoke-Sqlcmd -ServerInstance $s_inst -Database $s_db -Username $s_user -Password $s_pw -Query $s_qry

        # get the audit results
        $qry_fname=resolve-path -path "./SelectPackageAudit.sql"
        $s_qry = [IO.File]::ReadAllText($qry_fname)
        $s_qry = [string]::Format($s_qry,$pa)
        $pa_result = Invoke-Sqlcmd -ServerInstance $s_inst -Database $s_db -Username $s_user -Password $s_pw -Query $s_qry

        # display result
        $pa_result | Out-Gridview

    }
}

function refresh-grid
{
  $compare_datetime = Get-CompareDate
  $s_qry = [string]::Format($s_qry_template, $compare_datetime  ,$time_zone)
  $Global:changes = Invoke-Sqlcmd -ServerInstance $s_inst -Database $s_db -Username $s_user -Password $s_pw -Query $s_qry
  $dataGrid1.ItemsSource = $Global:changes
  $lRepo.Content = "Grid shows changes in local Database since Compare Date " + $compare_datetime
  $lRepo.Content += ". Last commit in current branch " +$gitBranch+" on "+$last_commit
  $dataGrid1.ItemsSource = $Global:changes
}

function Export-Items
{
    #create a reference to libs.dll for import/export functionality
    Add-Type -path  ($libs_folder.Path+"Libs.dll")
    <#* **************************************************
    * The following lines use Libs.dll, copied to Server/bin
    * and reference added to method_config.xml.
    * Build number of libs.dll probably needs to match IOM.dll
    * which it does v11SP9.
    * Research by inspecting source code in Visual Studio, aka "the documentation"
    * from libs source code: public CItemHelper(string Url, string Password, string DbName, string UserName, string Folder)
    * It is necessary to log in as well for ItemTypes and Relationship types to be exported!
    * This function adapted from C:\Repos\SelfDocumentingAras\ConfigurationManager\Import\Method\cm_export_selected_2.xml
    *************************************************** #>
    Get-ChildItem $export_folder -Force -Directory -Recurse | Remove-Item -Force

    $cih = New-Object Aras.Tools.SolutionUpgrade.CItemHelper($url,$pw,$db,$user,$export_folder);
    $cih.Login();
    $cei = New-Object Aras.Tools.SolutionUpgrade.CExportItems($cih);

    $selectrows = $dataGrid1.SelectedItems
    foreach ($item in $selectrows) 
    { 
        $this_pd=$item["Package"]
        $this_pg=$item["Group"]
        $this_name=$item["keyed_name"]
        $this_pe_id=$item["ElementId"]
        if ($this_pg -eq "RelationshipType (ItemType)") {$this_pg="RelationshipType"}
        $cih.Folder=$export_folder.Path+"`\"+$this_pd+"`\Import\"
        $h = @{}
        $ei = New-Object Aras.Tools.SolutionUpgrade.ExportItem($this_name,$this_pe_id,$this_pg);
        $cei.Export($ei,$this_pd,"1",$h)
    } 
    refresh-grid  
}
