# Step 0 Set up with utility scripts
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
# change directory to this folder
$sd = Get-ScriptDirectory
Set-Location $sd

.\DevDaily\CreateXaml-DevDaily.ps1 # function that implements the GUI

# create parameters to pass to XamlForm
$xamlFile =   $sd + '\DevDaily\DevDaily.xaml'
$iom =        $sd + '\12SP3\iom.dll' # iom must match server service pack
$configFile = $sd + '\DevDaily\config-devdaily.xml'

# show the Xaml GUI
$Form = CreateXaml-DevDaily -sd $sd -xaml $xamlFile -configFile $configFile -repo "C:\Repos\B-52-V12"
$Form.ShowDialog() | Out-Null
