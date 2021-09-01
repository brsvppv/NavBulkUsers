##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlCDjvPxzAtZyn/Rb0cdS/myurmp173trLq46nKAGcxEdl9nmSL5FgW0Wv1y
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlCDjpH5zAFT2X/rQ2VrWOyokJmJhMz83d/AnATrYLtUZUBzqhrdSWa8V/MVUPgQusVRGF0ZLOAC8qbDLOi7SaY209F6auGXlbE7HErM8K+1/Tik877SGw12U37qYKlxFD3u12TPWFG5m4FokWabU9jUuJt0pn2K2HEG30dFUK+Ek1k=
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba

Add-Type -AssemblyName System.Windows.Forms
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     
        Title="Create NavUsers From File" Height="175" Width="300" ResizeMode="NoResize" ShowInTaskbar="True">
    <Grid Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" Margin="0,0,0,0" Height="155" VerticalAlignment="Top" HorizontalAlignment="Left" Width="300">
        <TextBox Name="txtUserFile" HorizontalAlignment="Left" Height="20" Width="230" Margin="26,15,0,0" TextWrapping="Wrap" Text="Select CSV User List" VerticalAlignment="Top"  FontSize="11" IsReadOnly="true"/>
        <Label Name="lblsrvc" Content="Service Instance" HorizontalAlignment="Left" Margin="26,79,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="cbxNavInstance" HorizontalAlignment="Left" Margin="117,79,0,0" VerticalAlignment="Top" Width="158"/>
        <TextBox Name="txtPswdFile" HorizontalAlignment="Left" Height="20" Margin="26,51,0,0" TextWrapping="Wrap" Text="Dir To Export" VerticalAlignment="Top" Width="231" FontSize="11" IsReadOnly="true"/>
        <Button Name="btnUserSourceFile" Content="..." Margin="255,15,0,0" VerticalAlignment="Top" Height="20" HorizontalAlignment="Left" Width="20"/>
        <Button Name="btnSavePswdFile" Content="..." HorizontalAlignment="Left" Margin="255,51,0,0" VerticalAlignment="Top" Width="20" Height="20"/>
        <Button Name="btnPerformAction" Content="Perform Action" Margin="27,115,0,0" VerticalAlignment="Top" Height="23" HorizontalAlignment="Left" Width="248"/>
    </Grid>
</Window>
"@

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
try { $Form = [Windows.Markup.XamlReader]::Load( $reader ) }
catch { Write-Host "Unable to load Windows.Markup.XamlReader"; exit }
# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }
$fullPath
$bcPackage = Get-Package  | Where-Object { $_.Name -match '(Microsoft Dynamics 365 Business Central Server)' }
function RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}

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
        #    $cbxbcversion.Items.Add("BC Version 16")
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

$navServicesList = Get-NAVServerInstance | Where-Object { ($_.State -eq "Running") }
foreach ($service in $navServicesList) {
    $cbxNavInstance.Items.Add($service.ServerInstance.Substring(27)) 
}

Function PickUserFile { 
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $false # Multiple files can be chosen
        Filter      = 'Text Documents (*.txt)|*.txt' # Specified file types
    }
    [void]$FileBrowser.ShowDialog()
    
    If ($FileBrowser.FileNames -like "*\*")
    { $fullPath = $FileBrowser.FileNames }
    else {
        Write-Host
        [System.Windows.MessageBox]::Show( "Cancelled by user. `nYou need to select file.  `n $_", 'No FIle Selected', 'OK', 'Error')
    }
    $filedirectory = Split-Path -Parent $FileBrowser.FileName
    $fname = [System.IO.Path]::GetFileName($fullPath)          
    $txtUserFile.Text = "$filedirectory" + "\" + "$fname"
    $navUserNames = $txtUserFile.Text
    return $FileBrowser.FileNames 
  
}
function FindFolders {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = "C:\"
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Select a directory"

    $loop = $true
    while ($loop) {
        if ($browse.ShowDialog() -eq "OK") {
            $loop = $false
		
            #Insert your script here
		
        }
        else {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($res -eq "No") {
                #Ends script
                return
            }
        }
    }
    $browse.SelectedPath
    $txtPswdFile.Text = $browse.SelectedPath
    $browse.Dispose()
    
}

$btnUserSourceFile.Add_click( {
        PickUserFile    
        #$navUserNames = $txtUserFile.Text
    })

$btnSavePswdFile.Add_click( {
        FindFolders
        #$filepath = $txtPswdFile.Text

    })

#$ConfirmResult = -1
$OFS = "`r`n"

$btnPerformAction.Add_click( {

        $navUserNames = (Get-Content -Path $txtUserFile.Text)
        
        
        $createfile = "UsersPasswords.txt"
    
        $filepath = $txtPswdFile.Text + "\" + $createfile
        
        New-Item $filepath -type File

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
    
            $instance = $cbxNavInstance.SelectedValue
            

            New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -ChangePasswordAtNextLogOn -Verbose -LicenseType Full
            New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId SUPER -Verbose

        }
    })

$Form.ShowDialog() | out-null