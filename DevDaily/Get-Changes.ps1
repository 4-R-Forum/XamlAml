
# get list of changes since last commit and populate DataGrid

# get params from CI config
$machine_specific_includes = "C:\_machine_specific_includes\"+$project_prefix+"-"+$gitBranch+".Settings.include"
[xml] $settings = get-content $machine_specific_includes

$s_inst =$settings.SelectSingleNode("/project/property[@name='MSSQL.Server']/@value").Value
$s_db = $settings.SelectSingleNode("/project/property[@name='MSSQL.Database.Name']/@value").Value
$s_user =$settings.SelectSingleNode("/project/property[@name='MSSQL.Innovator.User']/@value").Value
$s_pw =$settings.SelectSingleNode("/project/property[@name='MSSQL.Innovator.Password']/@value").Value
$qry_fname=resolve-path -path "./ConfiguationReportDateTime.sql"
$s_qry_template = [IO.File]::ReadAllText($qry_fname)
$time_zone = "Eastern Standard Time"
function Get-CompareDate
{
  $compare_datetime= $tbCompareDateTime.Text
  if ([string]::IsNullOrEmpty($compare_datetime))   {$compare_datetime = $last_commit }
  else { $compare_datetime = get-date -date $compare_datetime -Format s }
  return $compare_datetime
}
$compare_datetime = Get-CompareDate
$s_qry = [string]::Format($s_qry_template, $compare_datetime  ,$time_zone)
$Global:changes = Invoke-Sqlcmd -ServerInstance $s_inst -Database $s_db -Username $s_user -Password $s_pw -Query $s_qry 



$repo_folder= (resolve-path "../")
$lRepo.Content += (" "+$repo_folder.Path +" in branch : "+$gitBranch+" on "+$last_commit)
$dataGrid1.ItemsSource = $Global:changes
