function Global:Create-XamlInBasket{
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $sd
        ,[parameter(Mandatory=$true)]
         [String]
         $xaml
        ,[parameter(Mandatory=$true)]
         [String]
         $configFile
        ,[parameter(Mandatory=$true)]
         [String]
         $iom
    )
    #==============================================================================================
    # example from:https://docs.microsoft.com/en-us/archive/blogs/platformspfe/integrating-xaml-into-powershell/
    # XAML file created in Visual Studio WPF Application, and saved in script folder.
    # See blog for namespaces used by Visual Studio that need to be removed!
    #==============================================================================================
    # also https://stackoverflow.com/questions/27791783/powershell-unable-to-find-type-system-windows-forms-keyeventhandler
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    Set-Location $sd
    .\Connect-IOM.ps1 # returns an Aras.Innovator object with authenticated connection to the server
    .\Get-DbList.ps1  # gets list of Dbs for url to populate db dropdown
    .\ExcelReport-InBasket.ps1  # loops through Excel File to write or apply AML

    # load XAML and create variables for Named elements
    [xml]$xaml = [IO.File]::ReadAllText($xamlFile)
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
    catch{Write-Host "Unable to load Windows.Markup.XamlReader. Invalid XAML."; Exit}
    # Store Form Objects In PowerShell, any named elements in the XAML are created as variables like $name_value
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object{Set-Variable -Scope global -Name  ($_.Name) -Value $Form.FindName($_.Name)}
    # variables set here
    $Global:ignore_pfx ="_"
    $Global:applyAML = $True
    $Global:FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
    [xml]$Global:config =  Get-Content $configFile 

    $bSDKIOM.Add_Click({Get-SDKIOM})
    function Global:Get-SDKIOM {
        $null = $FileBrowser.ShowDialog()
        $tbSDKIOM.Text = $FileBrowser.FileName
        $Form.UpdateLayout()
    }

    
    $bDbase.Add_Click({Set-Dbase})
    function Global:Set-Dbase {
        $db_list = Get-DbList -iom $iom -url $tbUrl.Text
        foreach ($db in $db_list) {
            $null = $cbDbase.Items.Add($db)
        }
        $Form.UpdateLayout()
    } 

    $bRun.Add_Click({Do-Load})
    function Global:Do-load  {
        $Global:config.selectSingleNode("config/url").InnerText =  $tbUrl.Text
        $Global:config.selectSingleNode("config/dbase").InnerText =  $cbDbase.Text
        $Global:config.selectSingleNode("config/user").InnerText =  $tbUser.Text
        $Global:config.Save($configFile)
        $lStatus.Content = "Status: Loading ..."
        $Form.UpdateLayout()
        $innov = Connect-IOM -iom $iom -url $tbUrl.Text -dbase $cbDbase.Text -user  $tbUser.Text -pw $pwbPw.Password
        ExcelReport-InBasket -innov $innov -sd $sd 
        $lStatus.Content = "Status: Finished"
    }

    $bExit.Add_Click({Exit-Form})
    function Global:Exit-Form  {
        $Global:config.selectSingleNode("config/url").InnerText =  $tbUrl.Text
        $Global:config.selectSingleNode("config/dbase").InnerText =  $cbDbase.Text
        $Global:config.selectSingleNode("config/user").InnerText =  $tbUser.Text
        $Global:config.selectSingleNode("config/SDKIOM").InnerText =  $tbSDKIOM.Text
        $Global:config.Save($configFile)
        $Form.Close()  | Out-Null
    }

    # populate form from config file
 <#   $tbAMLFile.Text=$config.selectSingleNode("config/AMLFile").'#text'
    $tbExcelFile.Text=$config.selectSingleNode("config/ExcelFile").'#text'
    #>
    $tbSDKIOM.Text = $config.selectSingleNode("config/SDKIOM").'#text'
    $tbUrl.Text=$config.selectSingleNode("config/url").'#text'
    Set-DBase
    $idx = [array]::indexof($cbDbase.Items,$config.selectSingleNode("config/dbase").'#text')
    $cbDbase.SelectedIndex=$idx
    $tbUser.Text=$config.selectSingleNode("config/user").'#text'

    return $Form
}