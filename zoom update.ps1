###Change the Below Varibles as required###

#writes messages to the console
$debug = $true
#writes messages to the log 
$log = $true 
#location of log
$loglocation = "$PSScriptRoot\log.txt" 
#Location of CleanZoom
$cleanzoomlocation = "$PSScriptRoot\CleanZoom.exe"
#Location of MSI
$MSIlocation = "$PSScriptRoot\ZoomInstallerFull.msi" 
#Location of Outlook Plugin
$Outlookpluginlocation = "$PSScriptRoot\ZoomOutlookPluginSetup.msi"
#This must be set to the version of the install file or it will continually install everytime the script is run.
$global:fullversion = "5.1.28656" 
#Sleeptime how long it will wait to before it runs the next install
$sleeptime = 10

###Varible Initialization###
###DO NOT CHANGE###
#doinstallflag if true will at end of script will run the install. It will be set to false if zoom is running or if already installed(with no appdata installation)
$doinstallflag = $true
$outlookpluginflag = $false

#This function writes to console or log if set to true
function write-debug{
    param($message)
    if ($debug -eq $true){
        Write-Output $message
    }
    if ($log -eq $true){
        Add-Content -Value $message -Path $loglocation
    }
}
#This function checks if the user is in a meeting
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
#This function checks if outlook plugin is installed
function check-outlookplugin{
    $outlookplugin = $false
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom Outlook Plugin"){$outlookplugin = $true}
    }
    $outlookplugin
}
#This function checks for the full MSI install
function check-fullinstall{
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom"){return $true}
    }
    return $false
}
#This function checks the version of the full install
function check-fullinstallversion{
    $fullinstallversion = ""
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom"){$fullinstallversion = $zoominstall.version}
    }
    $fullinstallversion
}
#This function checks the version of the outlook plugin(not used)
function check-outlookpluginversion{
    foreach ($zoominstall in $global:zoominstalls){
        if($zoominstall.name -eq "Zoom Outlook Plugin"){$zoominstall.version}
    }
}
#This function checks the version of the appdata install(not used)
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
#This function checks if the appdata version is installed. It check each directory in users folder and also checks the temp file and the default folder.
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
#This function checks against the $global:fullversion varible to see if the MSI is up to date
#The $global:fullversion varible must be set to the version of the install file this script references($MSILocation)
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

###Main Code###

write-debug -message (get-date)
Write-Debug -message "`r`nChecking installed programs through WMI`r`n"

#searches for zoom MSI installs with WMI
$global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}

#sets $doinstallflag to falls if user is in a meeting.
if(check-inmeeting){
    $doinstallflag = $false
    write-debug -message "Zoom is in a meeting install. Install aborted"
}

#checks if zoom is up date and if there is no appdata installed. If both are true then it will set $doinstallflag to false
if((check-fulluptodate) -and (!(check-appdatainstall))){
    $doinstallflag = $false
    write-debug -message "Zoom is up to date and no appdata versions found. Install aborted"
}

write-debug -message "Zoom in meeting: $(check-inmeeting)"
write-debug -message "Zoom Full version installed: $(check-fullinstall) $(check-fullinstallversion)"
write-debug -message "Zoom appdata installed: $(check-appdatainstall) $(check-appdatainstallversion)"
write-debug -message "Zoom outlook plugin installed: $(check-outlookplugin) $(check-outlookpluginversion)"
write-debug -message "Zoom uptodate: $(check-fulluptodate)"

#if $doinstallflag is still true it will run cleanzoom.exe, MSI install, and reinstall outlook for users who had outlook.
if($doinstallflag -eq $true){
    write-debug -message "Running Clean Zoom"
    start -wait $cleanzoomlocation
    write-debug -message "Installing Full version of Zoom"
    msiexec /i $MSIlocation /quiet /norestart MSIRESTARTMANAGERCONTROL="Disable" ZoomAutoUpdate="False" ZNoDesktopShortCut="true" ZConfig="nogoogle=1;nofacebook=1" /log install.log
    if(check-outlookplugin){
        write-debug -message "Installing outlook plugin"
        sleep $sleeptime
        msiexec /i $Outlookpluginlocation /quiet /norestart
    }
}

#checks version installed applications matching zoom to report if the installation was successful.
if($debug -eq $true -or $log -eq $true){
    sleep $sleeptime
    $global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}
}

write-debug -message ""
write-debug -message (get-date)
write-debug -message ""
write-debug -message "Zoom Full version installed: $(check-fullinstall) $(check-fullinstallversion)"
write-debug -message "Zoom appdata installed: $(check-appdatainstall) $(check-appdatainstallversion)"
write-debug -message "Zoom outlook plugin installed: $(check-outlookplugin) $(check-outlookpluginversion)"
write-debug -message "Zoom up to date: $(check-fulluptodate)"
    
