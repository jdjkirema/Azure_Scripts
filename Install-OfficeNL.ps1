# Installs Microsoft 365 Apps for Business (Business Premium) in Dutch (nl-NL)
# Excludes Teams, Groove, and Lync, uses Current Channel
$ErrorActionPreference = "Stop"
$work = "C:\Temp\ODT"
New-Item -ItemType Directory -Path $work -Force | Out-Null

# Download and extract Office Deployment Tool
$odtExe = Join-Path $work "officedeploymenttool.exe"
Invoke-WebRequest -Uri "https://download.microsoft.com/download/2/1/8/2185FA6B-5110-4A7D-BDAD-381BFD4A1F31/officedeploymenttool.exe" -OutFile $odtExe
Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:$work" -Wait

# Create Dutch configuration (Apps for business, exclude Teams, Groove, Lync)
$configPath = Join-Path $work "config-business-nl.xml"
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="nl-nl"/>
      <ExcludeApp ID="Teams"/>
      <ExcludeApp ID="Groove"/>
      <ExcludeApp ID="Lync"/>
    </Product>
  </Add>
  <RemoveMSI All="TRUE"/>
  <Property Name="SharedComputerLicensing" Value="1"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
  <Display Level="None" AcceptEULA="TRUE"/>
</Configuration>
"@ | Set-Content -Path $configPath -Encoding UTF8

# Install Office
$setup = Join-Path $work "setup.exe"
Start-Process -FilePath $setup -ArgumentList "/configure `"$configPath`"" -Wait
