<#Author       : Akash Chawla
# Usage        : Teams Optimization
#>

#######################################
#    Teams Optimization               #
#######################################

# Reference: https://learn.microsoft.com/en-us/azure/virtual-desktop/teams-on-avd

[CmdletBinding()]
Param (
    # Används endast om du kör Teams 1.0-vägen (MSI). Lämnas orörd om du vill köra nya Teams 2.0 via bootstrappern.
    [Parameter()]
    [string]$TeamsDownloadLink = "https://go.microsoft.com/fwlink/?linkid=2196106",

    # VC++ 2015–2022 x64 (permalink)
    [Parameter()]
    [string]$VCRedistributableLink = "https://aka.ms/vs/17/release/vc_redist.x64.exe",

    # Remote Desktop WebRTC Redirector Service (MSI, aka.ms-länk)
    [Parameter()]
    [string]$WebRTCInstaller = "https://aka.ms/msrdcwebrtcsvc/msi",

    # Nya Teams (TeamsBootstrapper.exe) – maskinvid installation
    [Parameter()]
    [string]$TeamsBootStrapperUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
)
 
function InstallTeamsOptimizationforAVD($TeamsDownloadLink, $VCRedistributableLink, $WebRTCInstaller, $TeamsBootStrapperUrl) {
    Begin {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $templateFilePathFolder = "C:\AVDImage"
        Write-Host "Starting AVD AIB Customization: Teams Optimization : $((Get-Date).ToUniversalTime()) "

        $guid = [guid]::NewGuid().Guid
        $tempFolder = Join-Path "C:\temp" $guid

        if (!(Test-Path -Path $tempFolder)) {
            New-Item -Path $tempFolder -ItemType Directory | Out-Null
        }
        Write-Host "AVD AIB Customization: Teams Optimization: Created temp folder $tempFolder"
    }

    Process {
        try {
            # Markera miljön som AVD för Teams
            New-Item -Path HKLM:\SOFTWARE\Microsoft -Name "Teams" -Force -ErrorAction Ignore | Out-Null
            $registryPath  = "HKLM:\SOFTWARE\Microsoft\Teams"
            $registryKey   = "IsWVDEnvironment"
            $registryValue = "1"
            Set-RegKey -registryPath $registryPath -registryKey $registryKey -registryValue $registryValue 
            
            # Installera Microsoft Visual C++ Redistributable (x64)
            Write-Host "AVD AIB Customization: Teams Optimization - Installing Microsoft Visual C++ Redistributable"
            $appName     = 'teams'
            New-Item -Path $tempFolder -Name $appName -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
            $LocalPath   = Join-Path $tempFolder $appName
            $VCRedistExe = 'vc_redist.x64.exe'
            $outputPath  = Join-Path $LocalPath $VCRedistExe
            Invoke-WebRequest -Uri $VCRedistributableLink -OutFile $outputPath
            Start-Process -FilePath $outputPath -ArgumentList "/install /quiet /norestart /log vcredist.log" -Wait
            Write-Host "AVD AIB Customization: Teams Optimization - VC++ Redistributable installed"

            # Installera Remote Desktop WebRTC Redirector Service
            $webRTCMSI  = 'webSocketSvc.msi'
            $outputPath = Join-Path $LocalPath $webRTCMSI
            Invoke-WebRequest -Uri $WebRTCInstaller -OutFile $outputPath
            Start-Process -FilePath msiexec.exe -ArgumentList "/I `"$outputPath`" /quiet /norestart /log webSocket.log" -Wait
            Write-Host "AVD AIB Customization: Teams Optimization - WebRTC Redirector installed"

            # Installera Teams
            if (-not [string]::IsNullOrWhiteSpace($TeamsBootStrapperUrl)) {
                Write-Host "AVD AIB Customization: Teams Optimization - Installing NEW Teams (2.0) via bootstrapper"

                # Tillåt trusted sideloading (krävs ibland vid provisionering)
                New-Item -Path "HKLM:\Software\Policies\Microsoft\Windows" -Name "Appx" -Force -ErrorAction Ignore | Out-Null
                New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\Appx" -Name "AllowAllTrustedApps" -Value 1 -Force | Out-Null
                New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\Appx" -Name "AllowDevelopmentWithoutDevLicense" -Value 1 -Force | Out-Null

                # WebView2 – ofta krav för nya Teams
                Write-Host "AVD AIB Customization: Teams Optimization - Installing Edge WebView2"
                $EdgeWebView = Join-Path $LocalPath 'WebView.exe'
                $webviewUrl  = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
                Invoke-WebRequest -Uri $webviewUrl -OutFile $EdgeWebView
                Start-Process -FilePath $EdgeWebView -NoNewWindow -Wait -ArgumentList "/silent /install"
                Write-Host "AVD AIB Customization: Teams Optimization - Edge WebView2 installed"

                # Teams 2.0 bootstrapper
                $teamsBootStrapperPath = Join-Path $LocalPath 'teamsbootstrapper.exe'
                Invoke-WebRequest -Uri $TeamsBootStrapperUrl -OutFile $teamsBootStrapperPath
                Start-Process -FilePath $teamsBootStrapperPath -NoNewWindow -Wait -ArgumentList "-p"
                Write-Host "AVD AIB Customization: Teams Optimization - Teams 2.0 installation complete"

                # Verifiering (provisionerade paket)
                $provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like '*MSTeams*' }
                if ($provisioned) {
                    Write-Host "AVD AIB Customization: Teams Optimization - Provisioned package found for MSTeams."
                } else {
                    Write-Host "AVD AIB Customization: Teams Optimization - MSTeams NOT found in provisioned packages."
                }
            } 
            else {
                Write-Host "AVD AIB Customization: Teams Optimization - Installing classic Teams (MSI)"
                $teamsMsi   = 'teams.msi'
                $outputPath = Join-Path $LocalPath $teamsMsi
                Invoke-WebRequest -Uri $TeamsDownloadLink -OutFile $outputPath
                Start-Process -FilePath msiexec.exe -ArgumentList "/I `"$outputPath`" /quiet /norestart /log teams_msi.log ALLUSER=1 ALLUSERS=1" -Wait
                Write-Host "AVD AIB Customization: Teams Optimization - Classic Teams installation complete"
            }
        }
        catch {
            Write-Host "*** AVD AIB CUSTOMIZER PHASE *** Teams Optimization - Exception occured: [$($_.Exception.Message)]"
        }    
    }
        
    End {
        # Cleanup
        if (Test-Path -Path $templateFilePathFolder -ErrorAction SilentlyContinue) {
            Remove-Item -Path $templateFilePathFolder -Force -Recurse -ErrorAction Continue
        }
        if (Test-Path -Path $tempFolder -ErrorAction SilentlyContinue) {
            Remove-Item -Path $tempFolder -Force -Recurse -ErrorAction Continue
        }

        $stopwatch.Stop()
        $elapsedTime = $stopwatch.Elapsed
        Write-Host "*** AVD AIB CUSTOMIZER PHASE : Teams Optimization - Exit Code: $LASTEXITCODE ***"    
        Write-Host "Ending AVD AIB Customization : Teams Optimization - Time taken: $elapsedTime"
    }
}

function Set-RegKey($registryPath, $registryKey, $registryValue) {
    try {
        Write-Host "*** AVD AIB CUSTOMIZER PHASE *** Teams Optimization - Setting $registryKey to $registryValue ***"
        New-ItemProperty -Path $registryPath -Name $registryKey -Value $registryValue -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "*** AVD AIB CUSTOMIZER PHASE *** Teams Optimization - Cannot add registry key $registryKey : [$($_.Exception.Message)]"
    }
}

InstallTeamsOptimizationforAVD -TeamsDownloadLink $TeamsDownloadLink -VCRedistributableLink $VCRedistributableLink -WebRTCInstaller $WebRTCInstaller -TeamsBootStrapperUrl $TeamsBootStrapperUrl
