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
.\Connect-IOM.ps1 # returns an Aras.Innovator object with authenticated connection to the server
.\Load-Sql-02.ps1  # loops through Sql Table to write or apply AML
# set variables for connection
$SDKIOM = "C:\InnovatorTemp\12SP18\Aras Innovator 12.0 SP18 IOM SDK\.NET\IOM.dll"
$url = "http://localhost/Innovator12SP18"
$dbase = "Concept22"
$user = "admin"
$pw = "innovator"
# create parameters to pass to Load-Sql
$innov = Connect-IOM -iom $SDKIOM -url $url -dbase $dbase -user  $user -pw $pw


# Call Load-Sql
Load-Sql-02 -sd $sd -innov $innov