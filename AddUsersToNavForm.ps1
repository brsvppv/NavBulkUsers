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
     
        Title="Create NavUsers From File" Height="375" Width="355" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" ShowInTaskbar="True">
        <Grid Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" Height="365" VerticalAlignment="Center" HorizontalAlignment="Center" Width="345">
            <TextBox Name="txtUserFile" HorizontalAlignment="Left" Height="20" Width="311" Margin="5,31,0,0" Text="Select Username File List" VerticalAlignment="Top"  FontSize="11" IsReadOnly="true"/>
            <Label Name="lblsrvc" Content="Service Instance" HorizontalAlignment="Left" Margin="3,102,0,0" VerticalAlignment="Top"/>
            <ComboBox Name="cbxNavInstance" HorizontalAlignment="Left" Margin="128,102,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0"/>
            <Label Name="lblLicType" Content="License Type" HorizontalAlignment="Left" Margin="3,124,0,0" VerticalAlignment="Top"/>
            <ComboBox Name="cbxUsersLicenseTypes" HorizontalAlignment="Left" Margin="128,126,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0">
                <ComboBoxItem Content="Full"/>
                <ComboBoxItem Content="Limited"/>
            </ComboBox>
            <Label Name="lblUsrState" Content="User State" HorizontalAlignment="Left" Margin="3,148,0,0" VerticalAlignment="Top"/>
             <ComboBox Name="cbxUserState" HorizontalAlignment="Left" Margin="128,150,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0">
                <ComboBoxItem Content="Enabled"/>
                <ComboBoxItem Content="Disabled"/>
            </ComboBox>
            <TextBox Name="txtPswdFile" HorizontalAlignment="Left" Height="20" Margin="5,56,0,0" Text="Directory To Export Password File" VerticalAlignment="Top" Width="312" FontSize="11" IsReadOnly="true"/>
            <Button Name="btnUserSourceFile" Content="..." Margin="314,31,0,0" VerticalAlignment="Top" Height="20" HorizontalAlignment="Left" Width="27"/>
            <Button Name="btnSavePswdFile" Content="..." HorizontalAlignment="Left" Margin="314,56,0,0" VerticalAlignment="Top" Width="27" Height="20"/>
            <Button Name="btnPerformAction" Content="Perform Action" Margin="0,330,0,0" VerticalAlignment="Top" Height="23" HorizontalAlignment="Center" Width="335"/>
            <ComboBox Name="cbxAuthType" HorizontalAlignment="Left" Margin="128,78,0,0" VerticalAlignment="Top" Width="213" SelectedIndex="0">
                <ComboBoxItem Content="Windows Authentication"/>
                <ComboBoxItem Content="NAV User Password"/>
            </ComboBox>
            <Label Name="lblAuthType" Content="Authentication " HorizontalAlignment="Left" Margin="3,80,0,0" VerticalAlignment="Top"/>
            <ListBox Name="listBoxPermission" Width="160"  Height="150" Margin="5,174,179,42"/>
            <ListBox Name="listBoxPermissionSelected" Width="160"  Height="150" Margin="0,174,5,42" HorizontalAlignment="right"/>
            <Button Name="btnAddPerm" Content=">" Margin="0,174,0,0" VerticalAlignment="Top" Height="75" HorizontalAlignment="Center" Width="16"/>
            <Button Name="btnRemoveSelcted" Content="&lt;" Margin="0,248,0,0" VerticalAlignment="Top" Height="75" HorizontalAlignment="Center" Width="16"/>      
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
$OFS = "`r`n"
function RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Function AutoNavModuleImport {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $RootDir = "C:\Program Files\"
    $navdir = "Microsoft Dynamics NAV"
    #$global:globalpath = $null
    $globaldir = Join-Path $RootDir -ChildPath $navdir 
    $fileversion = "Microsoft.Dynamics.Nav.Service.dll"
    $buildversion = $null
    $gdirexist = [System.IO.Directory]::Exists($globaldir)
    $WriteLine = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"

    #if($variable -isnot [system.array]){do some code expecting the $variable is not an array}
    #if($variable -is [system.array]){do some other stuff with $variable[0] being an array}
    if ($gdirexist -eq $true) {
        $vdir = Get-Childitem $globaldir -ErrorAction Stop  
        
        if ($vdir -is [system.array]) {

            $version = $vdir.Count 
            ForEach-Object {
                write-warning "Total of $version has been found installed"
                Write-Host $WriteLine
                for ($index = 0; $index -lt $version; $index++) {
                    #$globaldir = Join-Path $globaldir -ChildPath  $vdir[$index].Name  | Join-Path -ChildPath "Service"
                    $file = Join-Path $globaldir -ChildPath  $vdir[$index].Name  | Join-Path -ChildPath "Service" | Join-Path -ChildPath $fileversion 
                    $boolVersionFile = [System.IO.file]::Exists($file)
                    if ($boolVersionFile -eq $false) {
                        Write-Host "Unknown Version Directory:" $vdir[$index].Name -ForegroundColor Yellow; Write-Warning  "Module cannot be found"
                        Write-Host $WriteLine
                    }
                    else { 
                        $buildversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($file).FileVersion
                        Write-Host "Available Module in Directory:" -NoNewline; write-host $vdir[$index].Name -ForegroundColor Yellow -NoNewline; write-host " Version:" -NoNewline; write-host $buildversion -ForegroundColor Cyan
                        Write-Host $WriteLine
                    }
                    
                } }
            $vdir = Read-Host "Please Input Available Module in Directory:"
        }
        
        $globaldir = Join-Path $globaldir -ChildPath  $vdir | Join-Path -ChildPath "Service"
    }
    if ($gdirexist -eq $false ) {

        $bcdir = "Microsoft Dynamics 365 Business Central"
    
        $globaldir = Join-Path $RootDir -ChildPath $bcdir 
        
        $vdir = Get-ChildItem $globaldir -ErrorAction Stop
        
        $globaldir = Join-Path "$globaldir" -ChildPath $vdir.Name  | Join-Path -ChildPath "Service"   
    }          

    $file = Join-Path $globaldir -ChildPath $fileversion   
    $buildversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($file).FileVersion
    $boolBuildExist = [string]::IsNullOrEmpty($buildversion) 

    if ($boolBuildExist -eq $false) { 
        $psmodpath = "$globaldir\NavAdminTool.ps1"
        $psmexist = [System.IO.file]::Exists($psmodpath)   
    }
    else {
        write-host "NO VERSION DETECTED" -BackgroundColor Black -ForegroundColor Red
        Write-Host "No Dynamics NAV Packages"
        BREAK
    }
    # Check if the modile exist  after path change.
    if ($psmexist -eq $true) {
        try {
            Import-Module "$psmodpath" -ErrorAction Stop -Verbose 

        }
        catch [System.Management.Automation.ItemNotFoundException] {     
            Write-Host " Error Found $_" -ForegroundColor Green
            Write-Host "Retry: $psmodpath"
    
            Import-Module "$psmodpath" -ErrorAction Stop -Verbose
            [System.Windows.MessageBox]::Show("No Server Installation:  $_", 'Error No Install', 'OK', 'Error')
        }
        $intVersion = [int]::Parse($vdir)
        Clear-Host
        Write-Host "Loading Packages Locations" -ForegroundColor Yellow
        Get-Package | Where-Object { $_.version -eq $buildversion } | Format-Table -AutoSize 
        Write-Host "Version: $buildversion" -ForegroundColor Green

    }
    else {
        Write-Host "Microsft Dynamcis DLL File was not found in $intVersion" -ForegroundColor  Magenta
    }
}

AutoNavModuleImport

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
        #Get Content of Users from the text file
        $navUserNames = (Get-Content -Path $txtUserFile.Text) 
        $createfile = "UsersPasswords.txt"
        $filepath = $txtPswdFile.Text + "\" + $createfile

        $exist = [System.IO.Directory]::Exists($filepath)
        if ($exist -eq $false) {
            Write-Host "File does not exist" -ForegroundColor  Magenta
            Write-HOst "Creating UserPasswords Files"
            New-Item $filepath -Type File
        }

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
            $UserState = $cbxUserState.Text
            $instance = $cbxNavInstance.SelectedValue
        
            try {
                if ($cbxAuthType.SelectedIndex -eq 0) {

                    $logonName = $navUserName + "@" + $env:USERDNSDOMAIN 

                    New-NAVServerUser -WindowsAccount $logonName -ServerInstance $instance -State $UserState -LicenseType "$LicenseType" -ErrorAction Stop
                    foreach ($PermSet in $listBoxPermissionSelected.Items) {
                        New-NavServerUserPermissionSet -WindowsAccount $logonName -ServerInstance $instance -PermissionSetId "$PermSet"
                    }
                }
                if ($cbxAuthType.SelectedIndex -eq 1) {

                    $contentfile = "User Name: $navUserName" + $OFS + "Password: $password" + $OFS
                    Add-Content $filepath "$contentfile" -NoNewline 
                    New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -LicenseType "$LicenseType" -State $UserState -ChangePasswordAtNextLogOn -ErrorAction Stop
                    foreach ($PermSet in $listBoxPermissionSelected.Items) {
                        New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId "$PermSet"
                    }
                    Start-Sleep -Milliseconds 1000                  
                }
                
            }
            catch {
                [System.Windows.MessageBox]::Show("An Error Occured during creation" + $OFS + $_ , 'Error', 'OK', 'Error')
                break
            }
            Write-Host "User has been added: $navUserName"
            }
            [System.Windows.MessageBox]::Show("Operation Completed", 'Info Massage', 'OK', 'Information')
        start-Sleep -Seconds 1
        Exit
        
        explorer $txtPswdFile.Text
    })
$Form.ShowDialog() | out-null
