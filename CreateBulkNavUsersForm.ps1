[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Windows.Forms.Application]::EnableVisualStyles()
[xml]$XAML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
     
        Title="Create NavUsers From File" Height="255" Width="300" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" ShowInTaskbar="True">
    <Grid Background="{DynamicResource {x:Static SystemColors.WindowBrushKey}}" Margin="0,0,0,0" Height="255" VerticalAlignment="Top" HorizontalAlignment="Left" Width="300">
        <TextBox Name="txtUserFile" HorizontalAlignment="Left" Height="20" Width="230" Margin="26,15,0,0" TextWrapping="Wrap" Text="Select CSV User List" VerticalAlignment="Top"  FontSize="11" IsReadOnly="true"/>
        <Label Name="lblsrvc" Content="Service Instance" HorizontalAlignment="Left" Margin="26,79,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="cbxNavInstance" HorizontalAlignment="Left" Margin="117,80,0,0" VerticalAlignment="Top" Width="158" SelectedIndex="0"/>
        <Label Name="lblPermSet" Content="Permission Set" HorizontalAlignment="Left" Margin="26,105,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="cbxPermissionSets" HorizontalAlignment="Left" Margin="117,105,0,0" VerticalAlignment="Top" Width="158" SelectedIndex="0">
            <ComboBoxItem Content="BASIC"/>
            <ComboBoxItem Content="SUPER"/>
        </ComboBox>
        <Label Name="lblLicType" Content="License Type" HorizontalAlignment="Left" Margin="26,130,0,0" VerticalAlignment="Top"/>
          <ComboBox Name="cbxUsersLicenseTypes" HorizontalAlignment="Left" Margin="117,130,0,0" VerticalAlignment="Top" Width="158" SelectedIndex="0">
            <ComboBoxItem Content="Full"/>
            <ComboBoxItem Content="Limited"/>
        </ComboBox>
        <TextBox Name="txtPswdFile" HorizontalAlignment="Left" Height="20" Margin="26,51,0,0" TextWrapping="Wrap" Text="Dir To Export" VerticalAlignment="Top" Width="231" FontSize="11" IsReadOnly="true"/>
        <Button Name="btnUserSourceFile" Content="..." Margin="255,15,0,0" VerticalAlignment="Top" Height="20" HorizontalAlignment="Left" Width="20"/>
        <Button Name="btnSavePswdFile" Content="..." HorizontalAlignment="Left" Margin="255,51,0,0" VerticalAlignment="Top" Width="20" Height="20"/>
        <Button Name="btnPerformAction" Content="Perform Action" Margin="27,190,0,0" VerticalAlignment="Top" Height="23" HorizontalAlignment="Left" Width="248"/>
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
$RootDir = "C:\Program Files\"
$navdir = "Microsoft Dynamics NAV"
$global:globalpath = $null
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
    $vdir = Get-ChildItem $globaldir -ErrorAction Stop
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
    [System.Windows.MessageBox]::Show("Dynamics NAV/BC Detected." + $OFS + "Version: $buildversion", 'Error No Install', 'OK', 'information')
    
}
else {
    [System.Windows.MessageBox]::Show("No Dynamcis NAV/BC Version Detected", 'Error', 'OK', 'Error')
    break
}
    
$psmodpath = "$globaldir\NavAdminTool.ps1"
$global:globalpath = $psmodpath
$exist = [System.IO.file]::Exists($psmodpath)
            
if ($exist -eq $false) {
    [System.Windows.MessageBox]::Show("MODULE NOT DETECTED $_", 'FILE NOT FOUND', 'OK', 'Error')
}
else {
    try {
        Import-Module "$psmodpath" -ErrorAction Stop -Verbose
        Import-Module "$global:globalpath" | Out-Null  -Verbose                
    }
    catch [System.Management.Automation.ItemNotFoundException] {     
        [System.Windows.MessageBox]::Show("An Error Occured During Import  $_", 'Version: $buildversion', 'OK', 'Infromation')        
        Import-Module "$global:globalpath" | Out-Null  -Verbose
        Import-Module "$globaldir\NavAdminTool.ps1" -ErrorAction Stop -Verbose
        [System.Windows.MessageBox]::Show("No Server Installation:  $_", 'Error No Install', 'OK', 'Error')
    }
    $intVersion = [int]::Parse($vdir)
}
$FORM.Add_Loaded( {       

        $Permissions = Get-NAVServerPermissionSet -ServerInstance $cbxNavInstance.Text
        foreach ($Permission in $Permissions) {    
            $cbxPermissionSets.Items.Add($Permission.PermissionSetID)

        }
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
        
        [System.Windows.MessageBox]::Show( "Cancelled by user. `nYou need to select file.  `n $_", 'No FIle Selected', 'OK', 'Error')
    }
    $filedirectory = Split-Path -Parent $FileBrowser.FileName
    $fname = [System.IO.Path]::GetFileName($fullPath)          
    $txtUserFile.Text = "$filedirectory" + "\" + "$fname"
    $navUserNames = $txtUserFile.Text
    return $FileBrowser.FileNames 
}

function FindFolders {
 
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

$cbxNavInstance.Add_SelectionChanged( {
        $cbxPermissionSets.Items.Clear()
        $Permissions = Get-NAVServerPermissionSet -ServerInstance $cbxNavInstance.Text
        foreach ($Permission in $Permissions) {    
            $cbxPermissionSets.Items.Add($Permission.PermissionSetID)
        }
    })

$cbxNavInstance.Add_GotFocus( {
  
    })

$btnUserSourceFile.Add_click( {
        PickUserFile    
        #$navUserNames = $txtUserFile.Text
    })

$btnSavePswdFile.Add_click( {
        FindFolders
        #$filepath = $txtPswdFile.Text
    })

#$ConfirmResult = -1


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
            $PermissionSet = $cbxPermissionSet.Text
            $LicenseType = $cbxUsersTypes.Text
            $contentfile = "User Name: $navUserName" + $OFS + "Password: $password" + $OFS

            Add-Content $filepath "$contentfile" -NoNewline

            Get-NAVServerInstance | Format-Table -Property "State", "DisplayName" -AutoSize 
            
    
            $instance = $cbxNavInstance.SelectedValue
            
            New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -ChangePasswordAtNextLogOn -Verbose -LicenseType $LicenseType
            New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId $PermissionSet -Verbose

        }
    })

$Form.ShowDialog() | out-null