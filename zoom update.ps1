$debug = $true
$global:fullversion = "5.1.28642" 
$doinstallflag = $true
$outlookpluginflag = $false
$global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}


function write-debug{
    param($message)
    if ($debug -eq $true){
        Write-Output $message

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
    start -wait C:\ZoomTest\CleanZoom.exe
    msiexec /i c:\zoomtest\ZoomInstallerFull.msi /quiet /norestart MSIRESTARTMANAGERCONTROL="Disable" ZoomAutoUpdate="False" ZNoDesktopShortCut="true" ZConfig="nogoogle=1;nofacebook=1" /log install.log 
    if($outlookpluginflag -eq $true){
        sleep 5
        msiexec /i C:\zoomtest\ZoomOutlookPluginSetup.msi /quiet /norestart
    }
}


    


<#

Zoom.msi - We want installed
Zoom.exe - We want removed
Zoom Outlook - We want to keep

To get rid of zoom.exe you must run Cleanzoom.exe
Cleanzoom.exe will end a in progress meeting. (BAD)
Cleanzoom.exe will get rid of Zoom Outlook. (BAD)

Cleanzoom if zoom.exe is installed and no one is in a meeting
Cleanzoom if zoom.msi is out of date and no one is in a meeting
Cleanzoom if zoom.msi is not installed and no one is in a meeting

if Cleanzoom has been run then zoom.msi is installed
if outlook plugin was installed it will reinstall it since outlook plugin is removed by cleanzoom


install zoom = true

Zoom is in a meeting install = false
Zoom is up to date and appdata is not installed = false

cleanzoom
install msi
if outlook plugin was installed, install outlook plugin
#>
