#description: Installs/Updates Office 365 Apps to newest version and disables Auto-Update. Recommended to run on desktop images.
#execution mode: IndividualWithRestart
#tags: Nerdio, Apps install
<#
.SYNOPSIS
Installs/Updates Office 365 Apps to newest version and disables Auto-Update. Recommended to run on desktop images.

.DESCRIPTION
Installs/Updates Office 365 Apps to newest version and disables Auto-Update. Recommended to run on desktop images.
It downloads automatically downloads the latest version of ODT and uses it to update M365 Apps.

.NOTES
Update the $ODTConfig variable to match the installation to your needs. For more information to ODT XML look at the documentation at https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool and https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options.
This script is our version with german language pack based on the script in the Nerdio repository.

.LINK
Original Script found at Nerdio on GitHub: https://github.com/Get-Nerdio/NMM/blob/main/scripted-actions/windows-scripts/Install%20Microsoft%20365%20Office%20Apps.ps1
#>

# Configure powershell logging
$SaveVerbosePreference = $VerbosePreference
$VerbosePreference = 'continue'
$VMTime = Get-Date
$LogTime = $VMTime.ToUniversalTime()
mkdir "$env:TEMP\NerdioManagerLogs\ScriptedActions" -Force
Start-Transcript -Path "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER.ps1" -Append -IncludeInvocationHeader

Write-Host "################# New Script Run #################"
Write-host "Current time (UTC-0): $LogTime"

# Create directory to store ODT and setup files
$workingDirectory = "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER"
New-Item -ItemType Directory -Path $workingDirectory -Force

# Parse through the MS Download Center page to get the most up-to-date download link
$MSDlSite2 = Invoke-WebRequest "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117" -UseBasicParsing
ForEach ($Href in $MSDlSite2.Links.Href)
{
    if ($Href -match "officedeploymenttool" ){
        $DLink = $href
    }
}

# Download office deployment tool using up-to-date link grabed eariler
Invoke-WebRequest -Uri $DLink -OutFile "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER\odt_sadata.exe" -UseBasicParsing

# unpack the ODT executable to get the setup.exe
Start-Process -filepath "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER\odt_sadata.exe" -ArgumentList "/extract:$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER /quiet" -Wait

# create a base config XML for ODT to use, this one has auto-update disabled
$ODTConfig = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise">
    <Product ID="O365ProPlusRetail">
      <Language ID="de-de" />
      <Language ID="MatchOS" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="FALSE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@ 
$ODTConfig | Out-File "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER\odtconfig.xml"

# execute odt.exe using the newly created odtconfig.xml. This updates/installs office (takes a while)
Start-Process -filepath "$env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER\setup.exe" -ArgumentList "/configure $env:TEMP\NerdioManagerLogs\ScriptedActions\Install-M365-Apps-GER\odtconfig.xml" -Wait

# End Logging
Stop-Transcript
$VerbosePreference=$SaveVerbosePreference
