##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlGDjqzx7Bp48WbhVG01UuyYtri0+K+56MPPiBneQZUHXRp+lSac
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
##KsfMAp/KUzWI0g==
##OsfOAYaPHGbQvbyVvnQnqxugEiZ7PKU=
##LNzNAIWJGmPcoKHc7Do3uAu+DDhlPovL69Y=
##LNzNAIWJGnvYv7eVvnRU907vVm1rStyVuLuux5L80cva+xDKTIgHKQ==
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnRU907vVm1rStyVuLuux5L80cva9Af6CbgBRV83ozr5Flj9f+AdWLUzvd0UNQ==
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlGDjpLm+idj4EbKS3oubdGUq7+i172Y+vnnryrJdbsQTXp2gBzvAVmuF/cKUJU=
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
[Net.ServicePointManager]::SecurityProtocol = 'Tls12'
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationCore, PresentationFramework
[xml]$XAML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     
        Title="Create NavUsers From File" Height="370" Width="355" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" ShowInTaskbar="True">
        <Grid Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" Height="365" VerticalAlignment="Center" HorizontalAlignment="Center" Width="345">
            <TextBox Name="txtUserFile" HorizontalAlignment="Left" Height="20" Width="311" Margin="5,31,0,0" Text="Select Username File List" VerticalAlignment="Top"  FontSize="11" IsReadOnly="true"/>
            <Label Name="lblsrvc" Content="Service Instance" HorizontalAlignment="Left" Margin="3,105,0,0" VerticalAlignment="Top"/>
            <ComboBox Name="cbxNavInstance" HorizontalAlignment="Left" Margin="128,105,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0"/>
            <Label Name="lblLicType" Content="License Type" HorizontalAlignment="Left" Margin="3,130,0,0" VerticalAlignment="Top"/>
            <ComboBox Name="cbxUsersLicenseTypes" HorizontalAlignment="Left" Margin="128,130,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0">
                <ComboBoxItem Content="Full"/>
                <ComboBoxItem Content="Limited"/>
            </ComboBox>
            <TextBox Name="txtPswdFile" HorizontalAlignment="Left" Height="20" Margin="5,56,0,0" Text="Directory To Export Password File" VerticalAlignment="Top" Width="312" FontSize="11" IsReadOnly="true"/>
            <Button Name="btnUserSourceFile" Content="..." Margin="314,31,0,0" VerticalAlignment="Top" Height="20" HorizontalAlignment="Left" Width="27"/>
            <Button Name="btnSavePswdFile" Content="..." HorizontalAlignment="Left" Margin="314,56,0,0" VerticalAlignment="Top" Width="27" Height="20"/>
            <Button Name="btnPerformAction" Content="Perform Action" Margin="0,323,0,0" VerticalAlignment="Top" Height="23" HorizontalAlignment="Center" Width="335"/>
            <ComboBox Name="cbxAuthType" HorizontalAlignment="Left" Margin="128,80,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0">
                <ComboBoxItem Content="Windows Authentication"/>
                <ComboBoxItem Content="NAV User Password"/>
            </ComboBox>
            <Label Name="lblAuthType" Content="Authentication " HorizontalAlignment="Left" Margin="3,80,0,0" VerticalAlignment="Top"/>
            <ListBox Name="listBoxPermission" Width="160"  Height="150" Margin="5,163,179,42"/>
            <ListBox Name="listBoxPermissionSelected" Width="160"  Height="150" Margin="0,163,5,42" HorizontalAlignment="right"/>
            <Button Name="btnAddPerm" Content=">" Margin="0,168,0,0" VerticalAlignment="Top" Height="75" HorizontalAlignment="Center" Width="16"/>
            <Button Name="btnRemoveSelcted" Content="&lt;" Margin="0,243,0,0" VerticalAlignment="Top" Height="75" HorizontalAlignment="Center" Width="16"/>      
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


function RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}

[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Function ImportNavModule{
$RootDir = "C:\Program Files\"
$navdir = "Microsoft Dynamics NAV"
#$global:globalpath = $null
$globaldir = Join-Path $RootDir -ChildPath $navdir 
$fileversion = "Microsoft.Dynamics.Nav.Service.dll"
$buildversion = $null
$exist = [System.IO.Directory]::Exists($globaldir)
$OFS = "`r`n"
if ($exist -eq $true) {

    $vdir = Get-Childitem $globaldir -ErrorAction Stop   

    $globaldir = Join-Path $globaldir -ChildPath  $vdir.Name | Join-Path -ChildPath "Service"
}
            
if ($exist -eq $false ) {
    $bcdir = "Microsoft Dynamics 365 Business Central"       
    $globaldir = Join-Path $RootDir -ChildPath $bcdir 
    $vdir = Get-ChildItem $globaldir -ErrorAction Stop | Out-Null
    $globaldir = Join-Path "$globaldir" -ChildPath $vdir.Name  | Join-Path -ChildPath "Service"
            
}          
$file = Join-Path $globaldir -ChildPath $fileversion  
try {  
    $buildversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($file).FileVersion  
}
catch {
   
}
$boolBuildExist = [string]::IsNullOrEmpty($buildversion) 
if ($boolBuildExist -eq $false) {
    $Packages = Get-Package | Where-Object { $_.version -eq $buildversion } | Format-Table -AutoSize
    [System.Windows.MessageBox]::Show("Dynamics NAV/BC Detected. Build Version: $buildversion", 'Info', 'OK', 'information')
    
}
else {
    [System.Windows.MessageBox]::Show("No Dynamcis NAV/BC Version Detected", 'Error', 'OK', 'Error')
    break
}
    
$psmodpath = "$globaldir\NavAdminTool.ps1"
#$global:globalpath = $psmodpath
$exist = [System.IO.file]::Exists($psmodpath)
            
if ($exist -eq $false) {
    [System.Windows.MessageBox]::Show("MODULE NOT DETECTED $_", 'FILE NOT FOUND', 'OK', 'Error')
}
else {
    try {
        Import-Module "$psmodpath" -ErrorAction Stop | Out-Null
        #Import-Module "$global:globalpath" | Out-Null                
    }
    catch [System.Management.Automation.ItemNotFoundException] {     
        [System.Windows.MessageBox]::Show("An Error Occured During Import  $_", 'Version: $buildversion', 'OK', 'Infromation')        
        #Import-Module "$global:globalpath" | Out-Null
        Import-Module "$psmodpath" -ErrorAction Stop | Out-Null
        Import-Module "$globaldir\NavAdminTool.ps1" -ErrorAction Stop
        [System.Windows.MessageBox]::Show("No Server Installation:  $_", 'Error No Install', 'OK', 'Error')
    }
}
}



$Form.Add_Loaded( {       
        
        $Permissions = Get-NAVServerPermissionSet -ServerInstance $cbxNavInstance.SelectedItem
        foreach ($Permission in $Permissions) {    
            $listBoxPermission.Items.Add($Permission.PermissionSetID)
        }
        $txtPswdFile.Visibility = 'Hidden'
        $btnSavePswdFile.Visibility = 'Hidden'
        $btnPerformAction.IsEnabled = $false
    })
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
        
        $res = [System.Windows.Forms.MessageBox]::Show("No File Selected, Try Again ?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($res -eq "No") {
            return
        }
    }
    $filedirectory = Split-Path -Parent $FileBrowser.FileName
    $fname = [System.IO.Path]::GetFileName($fullPath)          
    $txtUserFile.Text = "$filedirectory" + "\" + "$fname"
    #$navUserNames = $txtUserFile.Text
    return $FileBrowser.FileNames   
}

function SelectDirectory {
 
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = "C:\"
    $browse.ShowNewFolderButton = $true
    $browse.Description = "Select a directory"

    $loop = $true
    while ($loop) {
        if ($browse.ShowDialog() -eq "OK") {
            $loop = $false
		
            $browse.SelectedPath
            $txtPswdFile.Text = $browse.SelectedPath
            $browse.Dispose()
		
        }
        else {
            $res = [System.Windows.Forms.MessageBox]::Show("No Directroy Selected. Try Again?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::YesNo)
            if ($res -eq "No") {
                #####
                return
            }
        }
    }
}

$cbxNavInstance.Add_SelectionChanged( {
        $listBoxPermission.Items.Clear()
        $listBoxPermissionSelected.Items.Clear()
        $Permissions = Get-NAVServerPermissionSet -ServerInstance $cbxNavInstance.SelectedItem
        foreach ($Permission in $Permissions) {    
            $listBoxPermission.Items.Add($Permission.PermissionSetID)

        }
    })

function EnableButton {
    if ($cbxAuthType.SelectedIndex -eq 0 -And $txtUserFile.Text -eq "Select Username File List") {
        $btnPerformAction.IsEnabled = $false
    }
    elseif ($cbxAuthType.SelectedIndex -eq 1 -And $txtUserFile.Text -eq "Select Username File List") {
        $btnPerformAction.IsEnabled = $false
    }
    elseif ($cbxAuthType.SelectedIndex -eq 0 -and $txtUserFile.Text -ne "Select Username File List") {
        $btnPerformAction.IsEnabled = $true
    }
    elseif ($cbxAuthType.SelectedIndex -eq 1 -and $txtPswdFile.Text -ne "Directory To Export Password File") {
        $btnPerformAction.IsEnabled = $true
    }
}
$txtUserFile.Add_TextChanged( {
        EnableButton
    })

$txtPswdFile.Add_TextChanged( {
        EnableButton
    })
$cbxAuthType.Add_SelectionChanged( {

        if ($cbxAuthType.SelectedIndex -eq 0 ) {
            
            $txtPswdFile.Visibility = 'Hidden'
            $btnSavePswdFile.Visibility = 'Hidden'
            #$btnPerformAction.IsEnabled = $true
            EnableButton              
        }
        if ($cbxAuthType.SelectedIndex -eq 1) {

            #$btnPerformAction.IsEnabled = $false
            $txtPswdFile.Visibility = 'Visible'
            $btnSavePswdFile.Visibility = 'Visible'
            if ($txtUserFile.Text -eq "Select Username File List" -and $txtPswdFile.Text -eq "Directory To Export Password File") {
                $btnPerformAction.IsEnabled = $false
            }
            else {
                $btnPerformAction.IsEnabled = $false
            }
            EnableButton
        }
    })
$btnUserSourceFile.Add_click( {
        PickUserFile    
        #$navUserNames = $txtUserFile.Text
    })
$btnSavePswdFile.Add_click( {
        SelectDirectory
        #$filepath = $txtPswdFile.Text
    })
$btnAddPerm.Add_Click({
        $PermSelected = $listBoxPermission.SelectedItem
        $listBoxPermission.Items.Remove($PermSelected)
        $listBoxPermissionSelected.Items.Add($PermSelected)
    
    })
$btnRemoveSelcted.Add_Click({
    $listBoxPermission.Items.Add($listBoxPermissionSelected.Selecteditem)
    $listBoxPermissionSelected.Items.Remove($listBoxPermissionSelected.Selecteditem)
})
$btnPerformAction.Add_click( {

        $navUserNames = (Get-Content -Path $txtUserFile.Text)       
        
        $createfile = "UsersPasswords.txt"
    
        $filepath = $txtPswdFile.Text + "\" + $createfile
        
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
            $LicenseType = $cbxUsersLicenseTypes.Text
            $instance = $cbxNavInstance.SelectedValue
        }
        try {

            if ($cbxAuthType.SelectedIndex -eq 0) {

                $logonName = $navUserName + "@" + $env:USERDNSDOMAIN 

                New-NAVServerUser -WindowsAccount $logonName -ServerInstance $instance -LicenseType "$LicenseType"
                foreach ($PermSet in $listBoxPermissionSelected.Items) {
                    New-NavServerUserPermissionSet -WindowsAccount $logonName -ServerInstance $instance -PermissionSetId "$PermSet"
                }
            }
            
            if ($cbxAuthType.SelectedIndex -eq 1) {
                $exist = [System.IO.Directory]::Exists($filepath)
                if ($exist -eq $false) {
                    Write-Host "File does not exist" -ForegroundColor  Magenta
                    New-Item $filepath -Type File
                }

                $contentfile = "User Name: $navUserName" + $OFS + "Password: $password" + $OFS
                Add-Content $filepath "$contentfile" -NoNewline 

                New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -LicenseType "$LicenseType"  -ChangePasswordAtNextLogOn 
                foreach ($PermSet in $listBoxPermissionSelected.Items) {
                    New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId "$PermSet"
                }
                start-Sleep -Seconds 3
                explorer $txtPswdFile.Text
            }

        }
        catch {
            [System.Windows.MessageBox]::Show("An Error Occured during creation" + $OFS + $_ , 'Error', 'OK', 'Error')
            break
        }

        [System.Windows.MessageBox]::Show("Created Successfully", 'Info Massage', 'OK', 'Information')
        Exit

    })
$Form.ShowDialog() | out-null