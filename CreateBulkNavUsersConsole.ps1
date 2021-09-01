$OFS = "`r`n"
$createfiledir = "C:\Temp\"
$createfile = "UsersPasswords.txt"
$filepath = $createfiledir + "\" + $createfile

If (!(test-path $createfiledir)) {
    New-Item -ItemType Directory -Force -Path $createfiledir
} 
$fileExist = [System.IO.File]::Exists($filepath)

if ($fileExist -eq $false) {
    Write-Host ""
    Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Write-Host "File does not exist" -ForegroundColor  Magenta
    Write-Host "Creating File: $createfile To $createfiledir" -ForegroundColor  Yellow
    Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    #New-Item -path "C:\Program Files\" -name "myps" -Itemtype "directory"
    New-Item $filepath -type File
}
else {
       
    Write-Host "File  Exist in the  Location" -ForegroundColor Green
         
}
function RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}
Write-Host ""
Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

Write-Host "Example: C:\PathLocation\FileName.extension" -ForegroundColor Green
$TargetFile = (Read-Host User Bulk FilePath)

Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

$bcPackage = Get-Package  | Where-Object { $_.Name -match '(Microsoft Dynamics 365 Business Central Server)' }

Try {
    $fullbcversion = $bcPackage.Version
    $bcversion = $bcPackage.Version.Substring(0, 2)
    $bcvsnstring = "BC Version $bcversion"
    $bcInstVersion = $bcvsnstring
    if ($bcversion -eq "14") {
        $bcmoduleswap = "140"
        #$cbxbcversion.Items.Add("BC Version 15")
        # $cbxbcversion.Items.Add("BC Version 16")
        [System.Windows.MessageBox]::Show("Business Central Server, Build Version: $fullbcversion")
    }
    if ($bcversion -eq "15") {
        $bcmoduleswap = "150"    
        #$cbxbcversion.Items.Add("BC Version 14")
        #$cbxbcversion.Items.Add("BC Version 16")
        [System.Windows.MessageBox]::Show("Business Central Server, Build Version: $fullbcversion")
    }
    if ($bcversion -eq "16") {
        $bcmoduleswap = "160"
        #$cbxbcversion.Items.Add("BC Version 14")
        #$cbxbcversion.Items.Add("BC Version 15") 
        [System.Windows.MessageBox]::Show("Business Central Server, Build Version: $fullbcversion")
    }

    Import-Module "${env:ProgramFiles}\Microsoft Dynamics 365 Business Central\$bcmoduleswap\Service\NavAdminTool.ps1" -ErrorAction Stop
}
catch { 
    [System.Windows.MessageBox]::Show("No Busineess Central Server Installation:  $_", 'Error No Install', 'OK', 'Error')
    Break
}

Get-NAVServerInstance | Format-Table -Property "State", "DisplayName" -AutoSize 
Write-Host "Input Dynamics NAV Service Instance Name" -ForegroundColor Green
$instanceSelected = (Read-Host Instance:)

$navUserNames = (Get-Content -Path $TargetFile)

foreach ($navUserName in $navUserNames) {
    function ScramblePassword([string]$inputString) {     
        $characterArray = $inputString.ToCharArray()   
        $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
        $outputString = -join $scrambledStringArray
        return $outputString 
    }
    $password = RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
    $password += RandomCharacters -length 1 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += RandomCharacters -length 3 -characters '1234567890'
    $password += RandomCharacters -length 2 -characters '!?@#$_-*'
  
   
    $PlainPassword = $password
    $SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
 
    $contentfile = "User Name: $navUserName" + $OFS + " Password: $password" + $OFS
    Add-Content $filepath "$contentfile" -NoNewline

    Get-NAVServerInstance | Format-Table -Property "State", "DisplayName" -AutoSize 
    Write-Host "Input Dynamics NAV Service Instance Name" -ForegroundColor Green
    
    $instance = $instanceSelected

    New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -ChangePasswordAtNextLogOn -Verbose -LicenseType Full
    New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId SUPER -Verbose
}
Write-Host "File has been created with Users & Password" -ForegroundColor Green
Write-Host "File Location $filepath" -ForegroundColor Green