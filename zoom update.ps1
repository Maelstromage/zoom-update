$debug = $true
$log = $true
$global:fullversion = "5.1.28642" 
$doinstallflag = $true
$outlookpluginflag = $false
$global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}
$sleeptime = 10

function write-debug{
    param($message)
    if ($debug -eq $true){
        Write-Output $message

    }
    if ($log -eq $true){
        Add-Content -Value $message -Path "$PSScriptRoot\log.txt"
      

    }

}
function check-inmeeting{
    $zoommeeting = 0
    if(Get-Process -Name Zoom -ErrorAction SilentlyContinue){
        $zoommeeting = (Get-NetUDPEndpoint -OwningProcess (Get-Process -Name Zoom).Id -ErrorAction SilentlyContinue | measure).count
    }
    Switch ($zoommeeting)
    {
        0 {$false}
        default {$true}
    }
}
function check-outlookplugin{
    $outlookplugin = $false
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom Outlook Plugin"){$outlookplugin = $true}
    }
    $outlookplugin
}
function check-fullinstall{
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom"){return $true}
    }
    return $false
}
function check-fullinstallversion{
    $fullinstallversion = ""
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom"){$fullinstallversion = $zoominstall.version}
    }
    $fullinstallversion
}
function check-outlookpluginversion{
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom Outlook Plugin"){$zoominstall.version}
    }
}
function check-appdatainstallversion{
    $userfolders = (Get-ChildItem -Directory c:\users).name
    $userfolders = $userfolders + "default"
    foreach ($userfolder in $userfolders){
        
        if (test-path "c:\users\$userfolder\appdata\roaming\zoom\bin\zoom.exe"){
                
                $version = (Get-Command "c:\users\$userfolder\appdata\roaming\zoom\bin\zoom.exe").FileVersionInfo.FileVersion
                Write-debug -message "c:\users\$userfolder\appdata\roaming\zoom\bin\zoom.exe $version"
        }
    }
    if (test-path "C:\windows\temp\zoom\bin"){
                
            $version = (Get-Command "c:\windows\temp\zoom\bin\zoom.exe").FileVersionInfo.FileVersion
            Write-debug -message "c:\windows\temp\zoom\bin $version"
    }

}
function check-appdatainstall{
    $userfolders = (Get-ChildItem -Directory c:\users).name
    $userfolders = $userfolders + "default"
    foreach ($userfolder in $userfolders){
        if (test-path "c:\users\$userfolder\appdata\roaming\zoom\bin\zoom.exe"){
            return $true
        }
    }
    if (test-path "C:\windows\temp\zoom\bin"){
        return $true
    }
    return $false
}
function check-fulluptodate{
    $currentversion = check-fullinstallversion
    if(!(check-fullinstall)){return $false}
    $currentversion = $currentversion.split(".")
    $installversion = $global:fullversion.split(".")
    if($currentversion[0] -lt $installversion[0]){return $false}
    if($currentversion[1] -lt $installversion[1]){return $false}
    if($currentversion[2] -lt $installversion[2]){return $false}
    return $true
}


if($log = $true){Add-Content -Value "`r`n $(get-date) `r`n" -Path "$PSScriptRoot\log.txt"}

if(check-inmeeting){
    $doinstallflag = $false
    write-debug -message "Zoom is in a meeting install. Install aborted"
}

if((check-fulluptodate) -and (!(check-appdatainstall))){
    $doinstallflag = $false
    write-debug -message "Zoom is up to date and no appdata versions found. Install aborted"
}



write-debug -message "Zoom in meeting: $(check-inmeeting)"
write-debug -message "Zoom Full version installed: $(check-fullinstall) $(check-fullinstallversion)"
write-debug -message "Zoom appdata installed: $(check-appdatainstall) $(check-appdatainstallversion)"
write-debug -message "Zoom outlook plugin installed: $(check-outlookplugin) $(check-outlookpluginversion)"
write-debug -message "Zoom uptodate: $(check-fulluptodate)"


#$doinstallflag = $false #remove me

if($doinstallflag -eq $true){
    write-debug -message "Running Clean Zoom"
    start -wait $PSScriptRoot\CleanZoom.exe
    write-debug -message "Installing Full version of Zoom"
    msiexec /i $PSScriptRoot\ZoomInstallerFull.msi /quiet /norestart MSIRESTARTMANAGERCONTROL="Disable" ZoomAutoUpdate="False" ZNoDesktopShortCut="true" ZConfig="nogoogle=1;nofacebook=1" /log install.log 
    if(check-outlookplugin){
        write-debug -message "Installing outlook plugin"
        sleep $sleeptime
        msiexec /i $PSScriptRoot\ZoomOutlookPluginSetup.msi /quiet /norestart
    }
}


if($debug -eq $true){
    sleep $sleeptime
    $global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}
}

write-debug -message ""
write-debug -message "Zoom Full version installed: $(check-fullinstall) $(check-fullinstallversion)"
write-debug -message "Zoom appdata installed: $(check-appdatainstall) $(check-appdatainstallversion)"
write-debug -message "Zoom outlook plugin installed: $(check-outlookplugin) $(check-outlookpluginversion)"
write-debug -message "Zoom up to date: $(check-fulluptodate)"
    
