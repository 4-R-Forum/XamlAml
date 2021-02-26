function Global:Get-DbList{
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $iom
         ,[parameter(Mandatory=$true)]
         [String]
         $url
    )
    Add-Type -path $iom
    # return dbList
    $serverConnection =  [Aras.IOM.IomFactory]::CreateHttpServerConnection($url);
    return $serverConnection.GetDatabases();
}