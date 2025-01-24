# Variables
$OfficeDeploymentToolUrl = "https://github.com/johnOnk048/ODT365/blob/main/officedeploymenttool_18129-20158.exe?raw=true"
$OfficeConfigFileContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
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
