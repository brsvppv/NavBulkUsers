##https://bcloadtest.tvbg:7048/BGBC14LoadTest/WebClient/OData4/
#https://bcloadtest.tvbg:7048/BGBC14LoadTest/WebClient/OData4/
#https://bcloadtest.tvbg/BGBC14LoadTest/WebClient/
#DynamicsNav://bcloadtest.tvbg:7046/BGBC14LoadTest/
Function Add-NavBulkUsers {
    Param(

        $OFS = "`r`n",
        # Parameter help description
        [Parameter(Mandatory)]
        [String]$FileDirectory = "C:\Temp\",
        [Parameter(Mandatory)]
        $NewFileName = "UsersPasswords.txt",
        [Parameter(Mandatory)]
        $PermSet = 'SUPER',
        $filepath = $FileDirectory + "\" + $NewFileName

    )
    If (!(test-path $FileDirectory)) {
        New-Item -ItemType Directory -Force -Path $FileDirectory
    } 
    $fileExist = [System.IO.File]::Exists($filepath)

    if ($fileExist -eq $false) {
        Write-Host ""
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
        Write-Host "File does not exist" -ForegroundColor  Magenta
        Write-Host "Creating File: $NewFileName To $FileDirectory" -ForegroundColor  Yellow
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
    $UserNameFile = (Read-Host User Bulk FilePath)

    Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    Try {
        $ModuleFileName = "NavAdminTool.ps1"
        $VersionFile = "Microsoft.Dynamics.Nav.Service.dll"
        
    
        $RootDirectory = Get-ChildItem $env:ProgramFiles | Where-Object { $_.Name -like "*Microsoft Dynamics*" }
    
        $VersionDirectory = Get-ChildItem $RootDirectory.FullName
    
        if ($VersionDirectory -is [system.array]) {
            Write-host "Multiyple version installed"
    
            Break
        }
        elseif ($VersionDirectory -isnot [system.array]) {
    
            $ModulePath = Join-Path $VersionDirectory.FullName -ChildPath "Service" | Join-Path -ChildPath $ModuleFileName
            $VersionFile = Join-Path $VersionDirectory.FullName -ChildPath "Service" | Join-Path -ChildPath $VersionFile
            $boolModuleFile = [System.IO.file]::Exists($ModulePath)
            $boolVersionFile = [System.IO.file]::Exists($VersionFile)
        
            Switch ($false) {
                $boolVersionFile {
                    Write-Host "Version File Not FOund"
                }
                $boolModuleFile {
                    Write-host "NO MODULE FILE"
                    break    
                }
            }
            switch ($true) {
                $boolVersionFile {
                    $buildversion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($VersionFile).FileVersion
                    Write-Host $buildversion
                }
                $boolModuleFile {
                    Import-Module "$ModulePath" | Write-Output
                }
    
            }
    
        }

    }
    catch {
        Write-Warning "ERROR $_"
    }

    Get-NAVServerInstance | Format-Table -Property "State", "DisplayName" -AutoSize 
    Write-Host "Input Dynamics NAV Service Instance Name" -ForegroundColor Green
    $instanceSelected = (Read-Host Instance:)

    $NavUserNames = (Get-Content -Path $UserNameFile)

    foreach ($navUserName in $NavUserNames) {
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
        $PermissionSet = $PermSet
        $contentfile = "User Name: $navUserName" + $OFS + " Password: $password" + $OFS
        Add-Content $filepath "$contentfile" -NoNewline
    
        $instance = $instanceSelected
        
        Write-Host "Adding user $navUserName to Service Instance: $instance" -ForegroundColor Yellow

        New-NAVServerUser $instance -UserName $navUserName -Password $SecurePassword -ChangePasswordAtNextLogOn -Verbose -LicenseType Full
        New-NAVServerUserPermissionSet $instance -UserName $navUserName -PermissionSetId $PermissionSet -Verbose
        Write-Host "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    }
    Write-Host "File has been created with Users & Password" -ForegroundColor Green
    Write-Host "File Location $filepath" -ForegroundColor Green
}
Add-NavBulkUsers