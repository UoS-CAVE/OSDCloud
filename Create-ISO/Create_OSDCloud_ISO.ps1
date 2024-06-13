# Install deployment tools: https://go.microsoft.com/fwlink/?linkid=2120254
# Install WinPE add-on for Windows ADK: https://go.microsoft.com/fwlink/?linkid=2120253
# Install Module: Install-Module OSD -Force

# English/GB Template
New-OSDCloudTemplate -Language en-GB -SetInputLocale en-GB -Verbose

# Set Workshpace Folder
$WorkingDir = "C:\OSDCloud"
New-Item -ItemType Directory $WorkingDir
Set-OSDCloudWorkspace -WorkspacePath $WorkingDir

$Startnet = @'
start /wait PowerShell -NoL -C Install-Module OSD -Force -Verbose
start /wait PowerShell -NoL -C Start-OSDCloud -OSVersion 'Windows 11' -OSBuild 23H2 -OSEdition Pro -OSLanguage en-GB -OSLicense Retail
'@
Edit-OSDCloudWinPE -Startnet $Startnet -StartOSDCloudGUI -Brand 'University of Surrey' -Wallpaper "C:\temp\Surrey_2023.png" -CloudDriver *

# create OSD ISO
New-OSDCloudISO -WorkspacePath $WorkingDir
Write-Host "ISO Created: $WorkingDir" 

<# 
#Uninstall
Get-Module -ListAvailable -Name OSD
Remove-Module OSD
Uninstall-Module OSD -AllVersions -Force -Verbose
#>