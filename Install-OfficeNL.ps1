# ===== Config =====
$WorkDir   = "C:\ODT"
$ConfigXml = Join-Path $WorkDir "config.xml"
$LogDir    = "C:\ODT\InstallLogs"
$DownloadODT = $true

# ===== Prep =====
New-Item -Path $WorkDir -ItemType Directory -Force | Out-Null
New-Item -Path $LogDir  -ItemType Directory -Force | Out-Null

# ===== Download latest Office Deployment Tool dynamically =====
if ($DownloadODT) {
    Write-Host "Fetching latest Office Deployment Tool download link..."
    $odtPage = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
    $odtUrl  = ($odtPage.Links | Where-Object { $_.href -match "officedeploymenttool.*\.exe" }).href

    if (-not $odtUrl) { throw "Could not find latest ODT download URL." }

    $odtExe = Join-Path $WorkDir "officedeploymenttool.exe"
    Write-Host "Downloading Office Deployment Tool from: $odtUrl"
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtExe -UseBasicParsing

    Write-Host "Extracting ODT..."
    & $odtExe /quiet /extract:$WorkDir | Out-Null
}

$SetupExe = Join-Path $WorkDir "setup.exe"
if (-not (Test-Path $SetupExe)) { throw "setup.exe not found in $WorkDir." }

# ===== Configuration for AVD (nl-NL, Shared Computer Activation) =====
@"
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" AllowCdnFallback="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="nl-nl" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Groove" />
    </Product>
  </Add>

  <Display Level="None" AcceptEULA="TRUE" />

  <Property Name="SharedComputerLicensing" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />

  <RemoveMSI />

  <Updates Enabled="TRUE" Channel="MonthlyEnterprise" />
</Configuration>
"@ | Set-Content -Encoding UTF8 -Path $ConfigXml

# ===== Pre-cache Office payload =====
Write-Host "Downloading Office payload..."
Start-Process -FilePath $SetupExe -ArgumentList "/download `"$ConfigXml`"" -Wait

# ===== Install silently =====
Write-Host "Installing Microsoft 365 Apps silently..."
$inst = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$ConfigXml`"" -Wait -PassThru

# ===== Collect logs =====
$temps = @("$env:TEMP","C:\Windows\Temp") | Where-Object { Test-Path $_ }
foreach ($t in $temps) {
    Get-ChildItem -Path $t -Filter "OfficeClickToRun*.log" -ErrorAction SilentlyContinue |
        Copy-Item -Destination $LogDir -Force -ErrorAction SilentlyContinue
}

if ($inst.ExitCode -eq 0) {
    Write-Host "Microsoft 365 Apps installed successfully for AVD. Logs in $LogDir."
    exit 0
} else {
    Write-Error "Install failed with exit code $($inst.ExitCode). Check logs in $LogDir and %TEMP%."
    exit $inst.ExitCode
}
