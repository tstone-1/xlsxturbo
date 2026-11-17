# Opens workbooks in Excel through COM and reports whether Excel accepts each one.
#
# Reading a file back with openpyxl or as XML is not opening it in Excel (see
# dev-docs/testing.md), so a change that writes something Excel has to parse is
# checked here once before it ships. Windows with Excel installed only.
#
# Usage:
#   pwsh -NoProfile -File scripts/excel_open_check.ps1 -Control good.xlsx out\*.xlsx
#
# -Control is a workbook known to open (for example one written with the same
# options minus the case under test). It is opened first, in the same Excel
# instance; if it fails, the probe itself is broken and no verdict is printed.
#
# With DisplayAlerts off, a file Excel would offer to repair fails to open with
# "Unable to get the Open property of the Workbooks class". That error alone
# does not prove a repair prompt: any failed open raises it. So each refused
# file is opened a second time the way a user opens it, in a new visible Excel
# process with alerts on, and the text of the dialog Excel shows is printed
# beside the verdict. That step needs the logged-on desktop session; -NoDialog
# skips it (over ssh, for example), and the verdict is then the COM error only.
#
# Exit codes: 0 every file opened, 1 at least one refused, 2 the control failed
# or no file matched.

# Without this the first file path binds to -DialogTimeoutSeconds by position.
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)][string]$Control,
    [switch]$NoDialog,
    [int]$DialogTimeoutSeconds = 40,
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)][string[]]$Path
)

$files = @(foreach ($p in $Path) { Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue })
if ($files.Count -eq 0) {
    Write-Output "[ERROR] no file matched: $($Path -join ', ')"
    exit 2
}

# One open attempt. Returns $null on success, else the first line of the error.
function Test-Open($xl, [string]$file) {
    try {
        # Arguments: FileName, UpdateLinks = 0, ReadOnly = true.
        $wb = $xl.Workbooks.Open($file, 0, $true)
        $wb.Close($false)
        return $null
    } catch {
        return $_.Exception.Message.Split("`n")[0]
    }
}

# The texts of every dialog window the given Excel process shows right now.
function Get-DialogText([int]$procId) {
    $auto = [System.Windows.Automation.AutomationElement]
    $all = [System.Windows.Automation.Condition]::TrueCondition
    $children = [System.Windows.Automation.TreeScope]::Children
    $descendants = [System.Windows.Automation.TreeScope]::Descendants
    $ofProcess = New-Object System.Windows.Automation.PropertyCondition($auto::ProcessIdProperty, $procId)

    $dialogs = @()
    $loaded = $false
    foreach ($window in $auto::RootElement.FindAll($children, $ofProcess)) {
        $class = $window.Current.ClassName
        if ($class -like 'MsoSplash*') { continue }
        $candidates = @()
        if ($class -eq 'XLMAIN') {
            # The title names the workbook once it is loaded.
            if ($window.Current.Name -match '\.xls') { $loaded = $true }
            # A dialog owned by the main window is one of its children.
            foreach ($child in $window.FindAll($children, $all)) {
                $cc = $child.Current.ClassName
                if ($cc -eq '#32770' -or $cc -like 'NUIDialog*') { $candidates += $child }
            }
        } else {
            $candidates += $window
        }
        foreach ($dialog in $candidates) {
            $texts = @(foreach ($e in $dialog.FindAll($descendants, $all)) {
                    if ($e.Current.Name) { $e.Current.Name }
                }) | Select-Object -Unique
            if ($texts.Count -gt 0) { $dialogs += (($texts -join ' / ') -replace '\s+', ' ') }
        }
    }
    return [pscustomobject]@{ Dialogs = $dialogs; Loaded = $loaded }
}

# Opens one file in a new visible Excel process with alerts on and returns what
# the user would see. Ends only the process it started.
function Read-OpenDialog([string]$excelExe, [string]$file, [int]$timeoutSeconds) {
    # /x forces a new process, so no Excel window the user has open is touched.
    $proc = Start-Process -FilePath $excelExe -ArgumentList '/x', "`"$file`"" -PassThru
    try {
        $deadline = (Get-Date).AddSeconds($timeoutSeconds)
        $settled = 0
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 1500
            $state = Get-DialogText $proc.Id
            if ($state.Dialogs.Count -gt 0) { return "dialog: " + ($state.Dialogs -join ' || ') }
            # Loaded and still no dialog: give a late one four more polls.
            if ($state.Loaded) { $settled++; if ($settled -ge 4) { return "no dialog, the workbook loaded" } }
        }
        return "no dialog read within $timeoutSeconds s"
    } finally {
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$refused = @()
$opened = @()
try {
    Write-Output "Excel $($xl.Version) build $($xl.Build)"
    $excelExe = Join-Path $xl.Path 'EXCEL.EXE'
    $controlFile = (Resolve-Path -LiteralPath $Control).Path
    $err = Test-Open $xl $controlFile
    if ($null -ne $err) {
        Write-Output "[ERROR] the control did not open, so no verdict is valid: $controlFile -- $err"
        exit 2
    }
    Write-Output "[OK] control opens: $controlFile"
    foreach ($f in $files) {
        $err = Test-Open $xl $f.FullName
        if ($null -eq $err) {
            $opened += $f.FullName
        } else {
            $refused += [pscustomobject]@{ File = $f.FullName; Error = $err }
        }
    }
} finally {
    # Quit only the instance this script created; other Excel windows stay open.
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

foreach ($file in $opened) { Write-Output "OPENS    $file" }
if ($refused.Count -gt 0 -and -not $NoDialog) {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
}
foreach ($r in $refused) {
    $seen = if ($NoDialog) { "dialog not read (-NoDialog)" } else { Read-OpenDialog $excelExe $r.File $DialogTimeoutSeconds }
    Write-Output "REFUSED  $($r.File) -- $($r.Error) -- $seen"
}
Write-Output "$($opened.Count) of $($files.Count) opened"
if ($refused.Count -gt 0) { exit 1 }
exit 0
