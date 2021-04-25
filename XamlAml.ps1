function Get-ScriptDirectory{
    # like https://blogs.msdn.microsoft.com/powershell/2007/06/19/get-scriptdirectory-to-the-rescue/
    # this script wants to find files in the same folder
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    try {
        Split-Path $Invocation.MyCommand.Path -ea 0
    }
    catch {
        Write-Warning 'You need to call this function from within a saved script.'
    }
}
# set location. include scripts, and set variables
$sd = Get-ScriptDirectory
Set-Location $sd
.\Create-XamlAmlForm.ps1 # function that implements the GUI

# create parameters to pass to XamlForm
$xamlFile =   $sd + '\XamlAml.xaml'
$configFile = $sd + '\config.xml'
[xml]$Global:config =  Get-Content $configFile
# getDBList needs to use IOM before Xaml form can select IOM matching target url, use a default for getDBList
$dbList_iom = $config.selectSingleNode("config/SDKIOM").'#text'
if ([string]::IsNullOrEmpty($dbList_iom )) { $dbList_iom= $config.selectSingleNode("config/defaultIOM").'#text' }

# show the Xaml GUI
$Form = Create-XamlAmlForm -sd $sd -xaml $xamlFile -dbList_iom $dbList_iom -configFile $configFile
$Form.ShowDialog() | Out-Null