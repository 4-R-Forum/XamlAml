function Global:Create-XamlForm{
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $FormLocation
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

    # Step 0.1 load XAML and create variables for Named elements
    [xml]$xaml = [IO.File]::ReadAllText($FormLocation)
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
    catch{Write-Host "Unable to load Windows.Markup.XamlReader. Invalid XAML."; Exit}
    # Store Form Objects In PowerShell, any named elements in the XAML are created as variables like $name_value
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object{Set-Variable -Scope global -Name  ($_.Name) -Value $Form.FindName($_.Name)}
    # variables set here
    $Global:ignore_pfx ="_"
    $Global:applyAML = $True
    $Global:FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }

    $tbIgnore.Add_TextChanged({Set-Ignore})
    function Global:Set-Ignore {
        $global:ignore_pfx = $tbIgnore.Text
    }

    $cbApplyAML.Add_Click({Toggle-ApplyAML})
    function Global:Toggle-ApplyAML{
        $Global:applyAML = (-not $Global:applyAML)
    }

    $bExcelFile.Add_Click({Get-Excel})
    function Global:Get-Excel {
        $null = $FileBrowser.ShowDialog()
        $tbExcelFile.Text = $FileBrowser.FileName
        $Form.UpdateLayout()
    }

    $bAMLFile.Add_Click({Set-AMLFile})
    function Global:Set-AMLFile {
        $null = $FileBrowser.ShowDialog()
        $tbAMLFile.Text = $FileBrowser.FileName
        $Form.UpdateLayout()
    } 

    $bLoad.Add_Click({Do-Load})
    function Global:Do-load  {
        $Global:config.selectSingleNode("config/AMLFile").InnerText =  $tbAMLFile.Text
        $Global:config.selectSingleNode("config/ExcelFile").InnerText =  $tbExcelFile.Text
        $Global:config.selectSingleNode("config/url").InnerText =  $tbUrl.Text
        $Global:config.selectSingleNode("config/dbase").InnerText =  $tbDbase.Text
        $Global:config.selectSingleNode("config/user").InnerText =  $tbUser.Text
        $Global:config.Save("$Global:sd\config.xml")
        $lStatus.Content = "Status: Loading ..."
        $Form.UpdateLayout()
        $Global:xl =  Open-ExcelPackage -Path $tbExcelFile.Text
        $Global:innov = Connect-IOM -iom $Global:iom -url $tbUrl.Text -dbase $tbDbase.Text -user  $tbUser.Text -pw $pwbPw.Password
        Load-Excel
        $lStatus.Content = "Status: Finished"
    }

    $bExit.Add_Click({Exit-Form})
    function Global:Exit-Form  {
      $Form.Close()  | Out-Null
    }

    $tbAMLFile.Text=$Global:config.selectSingleNode("config/AMLFile").'#text'
    $tbExcelFile.Text=$Global:config.selectSingleNode("config/ExcelFile").'#text'
    $tbUrl.Text=$Global:config.selectSingleNode("config/url").'#text'
    $tbDbase.Text=$Global:config.selectSingleNode("config/dbase").'#text'
    $tbUser.Text=$Global:config.selectSingleNode("config/user").'#text'

    return $Form
}