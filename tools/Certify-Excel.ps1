<#
.SYNOPSIS
Certify the checked-out commit in Excel with one command and write a
validated evidence bundle (docs\EXCEL_EVIDENCE.md).
.DESCRIPTION
For the commit checked out in a repository, with a clean working tree:
1. creates the bundle folder, by default excel-evidence-SHORTSHA next to the
   checkout, and starts session.log there;
2. records the environment as found: Windows, Excel version and build, Office
   bitness, the macro setting and "Trust access to the VBA project object
   model". It never changes any of them;
3. starts a visible Excel, adds a new workbook and imports every configured
   VBA component of .github\repository-profile.json from the checkout, in
   role order: internal, public, test, example;
4. compiles the project (Debug > Compile VBAProject);
5. runs every test module's entry point, Module.RunModule, the policy's
   harness first, each with its own log opened by TestLog.OpenLog so that
   every printed line is written to harness.log or Module.log with the
   commit's SHA as its first line; then runs every example's
   Module.RunModule and records whether it ran without error;
6. saves the workbook in the bundle, closes it and quits Excel;
7. writes session.json and runs tools\write_excel_evidence.py, which writes
   host.json from the logs, validates the bundle with
   tools\check_excel_evidence.py, checks every suite's counts against its
   source and writes certification.md.
The exit code is 0 only when host.json validates as PASS, every suite
passed with its expected counts and every example ran. The manual
procedure of docs\EXCEL_EVIDENCE.md stays available as a fallback.
.REQUIREMENTS
- Windows desktop Excel, and Python 3.10 or later on the PATH (or -Python).
- "Trust access to the VBA project object model" already on: the import
  needs it, and the script never turns it on.
- A clean checkout of the commit to certify. -RepoRoot may name a generated
  project; the tools of this checkout then certify that project.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Certify-Excel.ps1 -RepoRoot C:\work\my-project -Operator "Owner, release workstation"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BundleDir,
    [Parameter(Mandatory = $false)]
    [string]$Operator = "Repository owner, local Windows workstation",
    [Parameter(Mandatory = $false)]
    [string]$Python = "python",
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$toolsRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = $toolsRoot
}
$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

$compileControlId = 578
$xlOpenXMLWorkbookMacroEnabled = 52
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$roleOrder = @("internal", "public", "test", "example")

function Get-Timestamp {
    return (Get-Date).ToString("yyyy-MM-ddTHH:mm:sszzz")
}

function Write-Session {
    param([string]$Text)
    $line = "{0}  {1}" -f (Get-Date).ToString("HH:mm:ss"), $Text
    Write-Host $line
    [System.IO.File]::AppendAllText($script:sessionLog, $line + "`r`n", $script:utf8NoBom)
}

function Get-RegistryValue {
    param([string]$Path, [string]$Name)
    $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) {
        return $null
    }
    return $item.$Name
}

function Get-SecurityValue {
    # A value set by group policy takes precedence over the user's own setting.
    param([string]$OfficeVersion, [string]$Name)
    foreach ($root in @("HKCU:\Software\Policies\Microsoft\Office", "HKCU:\Software\Microsoft\Office")) {
        $value = Get-RegistryValue -Path "$root\$OfficeVersion\Excel\Security" -Name $Name
        if ($null -ne $value) {
            return $value
        }
    }
    return $null
}

function Get-MacroPolicy {
    param([string]$OfficeVersion)
    switch (Get-SecurityValue -OfficeVersion $OfficeVersion -Name "VBAWarnings") {
        1 { return "Enable VBA macros" }
        2 { return "Disable VBA macros with notification" }
        3 { return "Disable VBA macros except digitally signed macros" }
        4 { return "Disable VBA macros without notification" }
        default { return "Disable VBA macros with notification (default; no setting found)" }
    }
}

function Get-VbaProjectAccess {
    param([string]$OfficeVersion)
    if ((Get-SecurityValue -OfficeVersion $OfficeVersion -Name "AccessVBOM") -eq 1) {
        return "On"
    }
    return "Off"
}

function Invoke-Macro {
    # Application.Run with the workbook-qualified name; arguments are passed through.
    param([object]$Excel, [string]$Qualified, [string]$Procedure, [object[]]$Arguments = @())
    switch ($Arguments.Count) {
        0 { return $Excel.Run($Qualified + $Procedure) }
        1 { return $Excel.Run($Qualified + $Procedure, $Arguments[0]) }
        default { return $Excel.Run($Qualified + $Procedure, $Arguments[0], $Arguments[1]) }
    }
}

# ----------------------------------------------------------------------------
# Candidate, configuration and bundle
# ----------------------------------------------------------------------------
$sha = (& git -C $RepoRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $sha -notmatch '^[0-9a-f]{40}$') {
    throw "Cannot read the checked-out commit of $RepoRoot."
}
$dirty = & git -C $RepoRoot status --porcelain --untracked-files=no
if ($LASTEXITCODE -ne 0 -or -not [string]::IsNullOrWhiteSpace(($dirty | Out-String))) {
    throw "The working tree has changes. Certify a clean checkout: git status must show none."
}
$configPath = Join-Path $RepoRoot ".github\repository-profile.json"
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$policyPath = Join-Path $RepoRoot ".github\excel-evidence-policy.json"
$policy = Get-Content -LiteralPath $policyPath -Raw -Encoding UTF8 | ConvertFrom-Json
$harnessModule = ([string]$policy.entry_point).Split(".")[0]

$components = @()
foreach ($role in $roleOrder) {
    foreach ($property in ($config.vba.components.PSObject.Properties | Sort-Object Name)) {
        if ([string]$property.Value -eq $role) {
            $path = [string]$property.Name
            $components += [pscustomobject]@{
                role   = $role
                path   = $path
                module = [System.IO.Path]::GetFileNameWithoutExtension($path)
                file   = Join-Path $RepoRoot ($path -replace '/', '\')
            }
        }
    }
}
if ($components.Count -eq 0) {
    throw "No VBA components are configured in $configPath."
}
foreach ($component in $components) {
    if (-not (Test-Path -LiteralPath $component.file)) {
        throw "Configured component is missing from the checkout: $($component.path)"
    }
}

if ([string]::IsNullOrWhiteSpace($BundleDir)) {
    $BundleDir = Join-Path (Split-Path -Parent $RepoRoot) ("excel-evidence-" + $sha.Substring(0, 7))
}
if (Test-Path -LiteralPath $BundleDir) {
    if (@(Get-ChildItem -LiteralPath $BundleDir -Force).Count -gt 0) {
        throw "The bundle folder is not empty: $BundleDir. A bundle is never overwritten."
    }
}
else {
    New-Item -ItemType Directory -Path $BundleDir | Out-Null
}
$BundleDir = (Resolve-Path -LiteralPath $BundleDir).Path
if ($BundleDir.StartsWith($RepoRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "The bundle folder must be outside the checkout: $BundleDir"
}
$script:sessionLog = Join-Path $BundleDir "session.log"
$startedAt = Get-Timestamp
Write-Session "Excel certification of candidate $sha"
Write-Session "Checkout $RepoRoot; bundle $BundleDir; started at $startedAt"

# ----------------------------------------------------------------------------
# Environment, as found
# ----------------------------------------------------------------------------
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
$osName = ($osInfo.Caption -replace '^Microsoft\s+', '').Trim()
$osText = "{0} (version {1}, build {2})" -f $osName, $osInfo.Version, $osInfo.BuildNumber
$osArchitecture = switch ($env:PROCESSOR_ARCHITECTURE) {
    "AMD64" { "x64" }
    "ARM64" { "ARM64" }
    "x86" { "x86" }
    default { $env:PROCESSOR_ARCHITECTURE }
}
$clickToRun = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
$platform = Get-RegistryValue -Path $clickToRun -Name "Platform"
$officeBitness = switch ($platform) {
    "x64" { "64-bit" }
    "x86" { "32-bit" }
    default { "unknown" }
}
$versionToReport = Get-RegistryValue -Path $clickToRun -Name "VersionToReport"
Write-Session "Windows: $osText; architecture $osArchitecture"

$pythonVersion = & $Python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    throw "Python was not found as '$Python'. Install Python 3.10 or later, or pass -Python."
}
Write-Session "Python: $pythonVersion"

$session = [ordered]@{
    operator           = $Operator
    candidate_sha      = $sha
    started_at         = $startedAt
    finished_at        = $null
    excel_version      = "unknown"
    excel_build        = "unknown"
    office_bitness     = $officeBitness
    os                 = $osText
    os_architecture    = $osArchitecture
    macro_policy       = "unknown"
    vba_project_access = "unknown"
    references         = @()
    import             = [ordered]@{ status = "NOT_RUN"; detail = "Not run"; components = @() }
    compile            = [ordered]@{ status = "NOT_RUN"; detail = "Not run" }
    suites             = @()
    examples           = @()
    workbook_saved     = $false
    workbook_closed    = $false
}

$workbookPath = Join-Path $BundleDir "Certification.xlsm"
$excel = $null
$workbook = $null
try {
    # ------------------------------------------------------------------------
    # Excel and its security settings, read and never changed
    # ------------------------------------------------------------------------
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $officeVersion = [string]$excel.Version
    $session.excel_version = $officeVersion
    if ($null -ne $versionToReport) {
        $session.excel_build = "Version $versionToReport (build $($excel.Build))"
    }
    else {
        $session.excel_build = "Version $officeVersion (build $($excel.Build))"
    }
    $session.macro_policy = Get-MacroPolicy -OfficeVersion $officeVersion
    $session.vba_project_access = Get-VbaProjectAccess -OfficeVersion $officeVersion
    Write-Session "Excel: $($session.excel_build), $officeBitness"
    Write-Session "Macro setting: $($session.macro_policy); VBA project access: $($session.vba_project_access)"
    if ($session.vba_project_access -ne "On") {
        throw ("'Trust access to the VBA project object model' is off. The import needs it; " +
               "turn it on yourself in Excel > Options > Trust Center if your policy allows. " +
               "This script never changes it.")
    }

    # ------------------------------------------------------------------------
    # Import the exact source into a new workbook
    # ------------------------------------------------------------------------
    $workbook = $excel.Workbooks.Add()
    $qualified = "'" + $workbook.Name + "'!"
    $project = $workbook.VBProject
    $imported = @()
    foreach ($component in $components) {
        $added = $project.VBComponents.Import($component.file)
        $name = [string]$added.Name
        if ($name -ne $component.module) {
            $session.import = [ordered]@{
                status = "FAIL"; detail = "$($component.path) imported as $name"; components = $imported
            }
            throw "The component $($component.path) was imported as $name, not $($component.module)."
        }
        $imported += $component.path
        Write-Session "Imported $($component.path) as $name ($($component.role))"
    }
    foreach ($reference in $project.References) {
        $session.references += [string]$reference.Description
    }
    Write-Session ("References: " + ($session.references -join "; "))
    $session.import = [ordered]@{
        status = "PASS"
        detail = "VBComponents.Import of every configured component from the checkout of $sha into a new workbook"
        components = $imported
    }

    # ------------------------------------------------------------------------
    # Compile
    # ------------------------------------------------------------------------
    # Debug > Compile VBAProject acts on the VBE's active project and is
    # greyed out once that project is compiled. The VBE window is shown, so
    # that the command's state is current, and the state is read afresh.
    $vbe = $excel.VBE
    $vbe.MainWindow.Visible = $true
    try {
        $vbe.ActiveVBProject = $project
    }
    catch {
        Write-Session ("The project could not be made the active one: " + $_.Exception.Message)
    }
    $compile = $vbe.CommandBars.FindControl(1, $compileControlId)
    if ($null -eq $compile) {
        $session.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject was not found in the VBE" }
        throw "Cannot find Debug > Compile VBAProject."
    }
    if ($compile.Enabled) {
        $compile.Execute()
    }
    $compiled = $false
    for ($attempt = 0; $attempt -lt 10 -and -not $compiled; $attempt++) {
        Start-Sleep -Milliseconds 500
        $compiled = -not $vbe.CommandBars.FindControl(1, $compileControlId).Enabled
    }
    $compileDetail = "Debug > Compile VBAProject completed with no error"
    if (-not $compiled) {
        # The command can stay enabled without a compile error, so the
        # operator looks at the VBE instead of the run stopping on a guess.
        Write-Host ""
        Write-Host "Debug > Compile VBAProject still looks enabled. In the VBE window, choose"
        Write-Host "Debug > Compile VBAProject yourself. Type Y if it completes with no error and"
        Write-Host "the command is then greyed out; type N if it reports a compile error."
        $answer = Read-Host "Compiled with no error? (Y/N)"
        if ($answer -notmatch '^\s*y') {
            $session.compile = [ordered]@{ status = "FAIL"; detail = "Debug > Compile VBAProject reported an error, as the operator saw" }
            throw "The project does not compile; see the VBE."
        }
        $compileDetail = "Debug > Compile VBAProject completed with no error, as the operator saw in the VBE"
    }
    $session.compile = [ordered]@{ status = "PASS"; detail = $compileDetail }
    Write-Session ("Compiled with no error" + $(if ($compiled) { "" } else { ", as the operator saw in the VBE" }))

    $ready = [string](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "TestLog.CertificationReady")
    if ($ready -ne "READY") {
        throw "TestLog.CertificationReady answered '$ready', not READY."
    }
    Write-Session "The project's code runs; TestLog answered READY"

    # ------------------------------------------------------------------------
    # Run every suite, the harness first, each into its own log
    # ------------------------------------------------------------------------
    $suites = @($components | Where-Object { $_.role -eq "test" -and $_.module -ne "TestLog" })
    $ordered = @($suites | Where-Object { $_.module -eq $harnessModule }) +
               @($suites | Where-Object { $_.module -ne $harnessModule })
    foreach ($suite in $ordered) {
        $logName = $(if ($suite.module -eq $harnessModule) { "harness.log" } else { $suite.module + ".log" })
        $entry = $suite.module + ".Run" + $suite.module
        $record = [ordered]@{
            module = $suite.module; entry_point = $entry; log = $logName
            opened = ""; status = "NOT_RUN"; error = $null; closed = ""
        }
        $record.opened = [string](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "TestLog.OpenLog" `
            -Arguments @((Join-Path $BundleDir $logName), $sha))
        if ($record.opened -ne "OPEN") {
            Write-Session "$entry not run: $($record.opened)"
            $session.suites += $record
            continue
        }
        Write-Session "Running $entry into $logName"
        try {
            [void](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure $entry)
            $record.status = "RAN"
        }
        catch {
            $record.status = "ERROR"
            $record.error = $_.Exception.Message
            try {
                [void](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "TestLog.Say" `
                    -Arguments @("ERROR=" + $record.error))
            }
            catch {
            }
            Write-Session "$entry raised: $($record.error)"
        }
        $record.closed = [string](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure "TestLog.CloseLog")
        Write-Session "$entry $($record.status); $($record.closed)"
        $session.suites += $record
    }

    # ------------------------------------------------------------------------
    # Run every example; its output stays in the Immediate window
    # ------------------------------------------------------------------------
    foreach ($example in @($components | Where-Object { $_.role -eq "example" })) {
        $entry = $example.module + ".Run" + $example.module
        $record = [ordered]@{ module = $example.module; entry_point = $entry; status = "NOT_RUN"; error = $null }
        try {
            [void](Invoke-Macro -Excel $excel -Qualified $qualified -Procedure $entry)
            $record.status = "RAN"
            Write-Session "$entry ran without error"
        }
        catch {
            $record.status = "ERROR"
            $record.error = $_.Exception.Message
            Write-Session "$entry raised: $($record.error)"
        }
        $session.examples += $record
    }

    # ------------------------------------------------------------------------
    # Keep the workbook for inspection
    # ------------------------------------------------------------------------
    try {
        $workbook.SaveAs($workbookPath, $xlOpenXMLWorkbookMacroEnabled)
        $session.workbook_saved = $true
        Write-Session "Workbook saved as $workbookPath"
    }
    catch {
        Write-Session ("The workbook could not be saved: " + $_.Exception.Message)
    }
}
catch {
    Write-Session ("Stopped: " + $_.Exception.Message)
}
finally {
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
            $session.workbook_closed = $true
            Write-Session "Workbook closed"
        }
        catch {
            Write-Session ("The workbook could not be closed: " + $_.Exception.Message)
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    else {
        $session.workbook_closed = $true
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ----------------------------------------------------------------------------
# Record and validate
# ----------------------------------------------------------------------------
$session.finished_at = Get-Timestamp
Write-Session "Finished at $($session.finished_at)"
[System.IO.File]::WriteAllText((Join-Path $BundleDir "session.json"),
    ($session | ConvertTo-Json -Depth 5), $utf8NoBom)
Write-Host ""
& $Python (Join-Path $toolsRoot "tools\write_excel_evidence.py") --root $RepoRoot --candidate-sha $sha --bundle $BundleDir
$code = $LASTEXITCODE
Write-Host ""
Write-Host "Bundle: $BundleDir"
if ($code -eq 0) {
    Write-Host "RESULT=PASS: the bundle validates for candidate $sha; every suite and example passed."
    exit 0
}
Write-Host "RESULT=FAIL: see certification.md and session.log in the bundle. Keep the bundle as it is."
exit $code
