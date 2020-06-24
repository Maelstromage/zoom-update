function check-inmeeting{
    $zoommeeting = (Get-NetUDPEndpoint -OwningProcess (Get-Process -Name Zoom).Id -ErrorAction SilentlyContinue | measure).count
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
        if($zoominstall.name -eq "Zoom"){$true}
    }
}
function check-fullinstallversion{
    $fullinstallversion = "Not Installed"
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
                    Write-Output "c:\users\$userfolder\appdata\roaming\zoom\bin\zoom.exe $version"
            }
        }
        if (test-path "C:\windows\temp\zoom\bin"){
                
                $version = (Get-Command "c:\windows\temp\zoom\bin\zoom.exe").FileVersionInfo.FileVersion
                Write-Output "c:\windows\temp\zoom\bin $version"
        }

}

$version = "5,1,27830,0612"
$fullversion = "5.1.27830"
$doinstallflag = $false
$outlookpluginflag = $false

$global:zoominstalls = Get-WmiObject -Class Win32_Product | where {$_.name -like "*zoom*"}

if(!(Get-Process -Name Zoom -ErrorAction SilentlyContinue)){
    Write-Output "zoom is not running"
}else{
    Write-Output "zoom is running"
    $checkmeeting = check-inmeeting
    write-output "In Meeting: $checkmeeting"
    if (!($checkmeeting)){
      
        $outlookpluginflag = check-outlookplugin
        if ((check-fullinstallversion) -ne "Not Installed"){
            if ((check-fullinstallversion) -ne $fullversion){
                $doinstallflag = $true
            }

        }
        foreach ($appdataversion in (check-appdatainstallversion)){
            if ($appdataversion.StartsWith("c:\")){
                if (($appdataversion -split(" "))[1] -ne $version){
                    $doinstallflag = $true
                }

            }

        }
        if((check-fullinstallversion) -eq "Not Installed" -and (check-appdatainstallversion).StartsWith("c:\")){
            $doinstallflag = $true
        } 
        if($doinstallflag -eq $true){
            start -wait C:\ZoomTest\CleanZoom.exe
            msiexec /i c:\zoomtest\ZoomInstallerFull.msi /quiet /norestart MSIRESTARTMANAGERCONTROL="Disable" ZoomAutoUpdate="False" ZNoDesktopShortCut="true" ZConfig="nogoogle=1;nofacebook=1" /log install.log 
            if($outlookpluginflag -eq $true){
                sleep 5
                msiexec /i C:\zoomtest\ZoomOutlookPluginSetup.msi /quiet /norestart
            }
        }

    }
}
