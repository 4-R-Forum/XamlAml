function Global:Connect-IOM{
    Param(
         [parameter(Mandatory=$true)]
         [String]
         $iom 
        ,[parameter(Mandatory=$true)]
         [String]
         $url
        ,[parameter(Mandatory=$true)]
         [String]
         $dbase
        ,[parameter(Mandatory=$true)]
         [String]
         $user
        ,[parameter(Mandatory=$true)]
         [String]
         $pw
    )
    Add-Type -path $iom
    # create a connection 
    $conn =[Aras.IOM.IomFactory]::createHttpserverconnection($url,$dbase,$user,$pw) 
    $res = $conn.Login()
    if ($res.isError())
    {
        $res.ToString()  
        Exit
    }
    return [Aras.IOM.IomFactory]::createInnovator($conn) 
}