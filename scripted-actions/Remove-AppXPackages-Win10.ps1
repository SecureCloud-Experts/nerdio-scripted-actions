<#Author       : Akash Chawla
# Usage        : Remove Appx Packages
#>

#######################################
#   Remove Appx Packages        #######
#######################################


[CmdletBinding()]
  Param (
        [Parameter()]
        [System.String[]]$AppxPackages="Clipchamp.Clipchamp","Microsoft.BingNews","Microsoft.BingWeather","Microsoft.GamingApp","Microsoft.GetHelp","Microsoft.Getstarted","Microsoft.MicrosoftOfficeHub","Microsoft.MicrosoftSolitaireCollection","Microsoft.MicrosoftStickyNotes","Microsoft.MixedReality.Portal","Microsoft.MSPaint","Microsoft.Office.OneNote","Microsoft.People","Microsoft.PowerAutomateDesktop","Microsoft.ScreenSketch","Microsoft.SkypeApp","Microsoft.WindowsAlarms","Microsoft.WindowsCamera","Microsoft.Windows.DevHome","Microsoft.windowscommunicationsapps","Microsoft.WindowsFeedbackHub","Microsoft.WindowsMaps","Microsoft.WindowsSoundRecorder","Microsoft.WindowsTerminal","Microsoft.XboxApp","Microsoft.Xbox.TCUI","Microsoft.XboxGameCallableUI","Microsoft.XboxGameCallableUI","Microsoft.XboxGameOverlay","Microsoft.XboxGamingOverlay","Microsoft.XboxIdentityProvider","Microsoft.XboxSpeechToTextOverlay","Microsoft.YourPhone","Microsoft.ZuneMusic","Microsoft.ZuneVideo"
 )

 function Remove-ProvidedAppxPackages($AppxPackages) {
   
        Begin {
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $templateFilePathFolder = "C:\AVDImage"
            Write-host "Starting AVD AIB Customization: Remove Appx Packages : $((Get-Date).ToUniversalTime()) "
        }

        Process {
            Foreach ($App in $AppxPackages) {
                try {                
                    Write-Host "AVD AIB CUSTOMIZER PHASE : Removing Provisioned Package $($App)"
                    Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like ("*{0}*" -f $App) } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue  | Out-Null
                            
                    Write-Host "AVD AIB CUSTOMIZER PHASE : Attempting to remove [All Users] $App "
                    Get-AppxPackage -AllUsers -Name ("*{0}*" -f $App) | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue 
                            
                    Write-Host "AVD AIB CUSTOMIZER PHASE : Attempting to remove $App"
                    Get-AppxPackage -Name ("*{0}*" -f $App) | Remove-AppxPackage -ErrorAction SilentlyContinue  | Out-Null

                    if($App -eq "Microsoft.MSPaint") {
                        $PaintWindowsName = "Microsoft.Windows.MSPaint"
                        Get-WindowsCapability -Online -Name ("*{0}*" -f $PaintWindowsName) | Remove-WindowsCapability -Online -ErrorAction SilentlyContinue
                    }
                }
                catch {
                    Write-Host "AVD AIB CUSTOMIZER PHASE : Failed to remove Appx Package $App - $($_.Exception.Message)"
                }
            } 
        }
        
        End {

            #Cleanup
            if ((Test-Path -Path $templateFilePathFolder -ErrorAction SilentlyContinue)) {
                Remove-Item -Path $templateFilePathFolder -Force -Recurse -ErrorAction Continue
            }
    
            $stopwatch.Stop()
            $elapsedTime = $stopwatch.Elapsed
            Write-Host "*** AVD AIB CUSTOMIZER PHASE : Remove Appx Packages -  Exit Code: $LASTEXITCODE ***"    
            Write-Host "Ending AVD AIB Customization : Remove Appx Packages - Time taken: $elapsedTime"
        }
 }

 Remove-ProvidedAppxPackages -AppxPackages $AppxPackages
