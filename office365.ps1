# Always Run as Administrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs;
exit 
}

# Variables
$OfficeDeploymentToolUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
$OfficeConfigFileContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
$TempDir = "$env:TEMP\OfficeDeploymentTool"
$OfficeConfigFilePath = "$TempDir\Configuration.xml"
$OfficeDeploymentToolPath = "$TempDir\OfficeDeploymentTool.exe"

# Create temp directory
if (-not (Test-Path -Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir | Out-Null
}

# Download Office Deployment Tool
Write-Host "Downloading Office Deployment Tool..."
Invoke-WebRequest -Uri $OfficeDeploymentToolUrl -OutFile $OfficeDeploymentToolPath -UseBasicParsing

# Extract Office Deployment Tool
Write-Host "Extracting Office Deployment Tool..."
Start-Process -FilePath $OfficeDeploymentToolPath -ArgumentList "/quiet /extract:$TempDir" -Wait

# Write configuration file
Write-Host "Creating configuration file..."
$OfficeConfigFileContent | Out-File -FilePath $OfficeConfigFilePath -Encoding UTF8

# Uninstall existing Office versions
Write-Host "Uninstalling existing Office installations..."
$OfficeRemovalTool = "$TempDir\setup.exe"
Start-Process -FilePath $OfficeRemovalTool -ArgumentList "/configure $OfficeConfigFilePath" -Wait

# Install Office 365
Write-Host "Installing Office 365..."
Start-Process -FilePath $OfficeRemovalTool -ArgumentList "/configure $OfficeConfigFilePath" -Wait

# Cleanup
Write-Host "Cleaning up temporary files..."
Remove-Item -Path $TempDir -Recurse -Force

Write-Host "Office 365 installation completed successfully!"
