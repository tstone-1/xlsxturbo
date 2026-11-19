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
# Each -Path item is resolved on its own: an item that exists as written is
# taken literally (so a name with [ or ] is found), anything else is a
# wildcard. An item that matches nothing stops the run before Excel starts.
#
# Each file gets one of three verdicts:
#   OPENS    Excel opened it (a failure to close it afterwards is printed as a
#            note and does not change the verdict).
#   REFUSED  Excel raised an error on open.
#   BUSY     Excel rejected the call (RPC_E_CALL_REJECTED or
#            RPC_E_SERVERCALL_RETRYLATER) on every retry. That is Excel still
#            working, for example refreshing a pivot table on load: it is "could
#            not check", never a refusal. Re-run the file on its own.
#
# With DisplayAlerts off, a file Excel would offer to repair fails to open with
# "Unable to get the Open property of the Workbooks class". That error alone
# does not prove a repair prompt: any failed open raises it. So each refused
# file is opened a second time the way a user opens it, in a new visible Excel
# process with alerts on, and the text of the dialog Excel shows is printed
# beside the verdict. That step needs the logged-on desktop session; -NoDialog
# skips it (over ssh, for example), and the verdict is then the COM error only.
#
# Exit codes:
#   0  every file opened
#   1  at least one file refused, and every other file opened
#   2  could not check: the control is missing or did not open, a -Path item
#      matched nothing, Excel could not be started or failed through COM, or
#      a file stayed BUSY. A run with a BUSY file exits 2 even when another file
#      was refused; the REFUSED lines are still printed.

# Without this the first file path binds to -DialogTimeoutSeconds by position.
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)][string]$Control,
    [switch]$NoDialog,
    [int]$DialogTimeoutSeconds = 40,
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)][string[]]$Path
)

# How often a call Excel rejects as busy is tried, and the pause between tries.
$BusyAttempts = 5
$BusyPauseMilliseconds = 2000

# Resolves every -Path item on its own. Returns the files and the items that
# matched nothing.
function Resolve-Inputs([string[]]$items) {
    $found = @()
    $missing = @()
    foreach ($p in $items) {
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $found += Get-Item -LiteralPath $p
            continue
        }
        $matched = @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue)
        if ($matched.Count -eq 0) { $missing += $p } else { $found += $matched }
    }
    # A file named by two items is opened once.
    $unique = @($found | Sort-Object -Property FullName -Unique)
    return [pscustomobject]@{ Files = $unique; Missing = $missing }
}

# True when an exception, or one it wraps, is Excel rejecting the call because
# it is busy: RPC_E_CALL_REJECTED (0x80010001) or RPC_E_SERVERCALL_RETRYLATER
# (0x8001010A).
function Test-Busy($exception) {
    for ($e = $exception; $null -ne $e; $e = $e.InnerException) {
        $code = '{0:X8}' -f $e.HResult
        if ($code -eq '80010001' -or $code -eq '8001010A') { return $true }
        if ($e.Message -match 'RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER') { return $true }
    }
    return $false
}

function Clear-Com($object) {
    if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object)
    }
}

# Opens one file read-only and closes it again. Retries while Excel is busy.
# Returns State (OPENS, REFUSED or BUSY), Error (first line of the last error)
# and Note (a failure to close an opened workbook).
function Test-Open($xl, [string]$file) {
    $lastError = $null
    for ($attempt = 1; $attempt -le $BusyAttempts; $attempt++) {
        $books = $null
        $wb = $null
        try {
            $books = $xl.Workbooks
            # Arguments: FileName, UpdateLinks = 0, ReadOnly = true.
            $wb = $books.Open($file, 0, $true)
        } catch {
            $lastError = $_.Exception.Message.Split("`n")[0]
            $isBusy = Test-Busy $_.Exception
            Clear-Com $wb
            Clear-Com $books
            if (-not $isBusy) { return [pscustomobject]@{ State = 'REFUSED'; Error = $lastError; Note = $null } }
            if ($attempt -lt $BusyAttempts) { Start-Sleep -Milliseconds $BusyPauseMilliseconds }
            continue
        }
        # The open succeeded; whatever happens on close does not make it a refusal.
        $note = $null
        try {
            $wb.Close($false)
        } catch {
            $note = "opened, but closing it failed: " + $_.Exception.Message.Split("`n")[0]
        } finally {
            Clear-Com $wb
            Clear-Com $books
        }
        return [pscustomobject]@{ State = 'OPENS'; Error = $null; Note = $note }
    }
    return [pscustomobject]@{ State = 'BUSY'; Error = "$lastError (busy on $BusyAttempts attempts)"; Note = $null }
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

# Everything below needs no Excel until New-Object, so a bad argument costs
# nothing and is never mistaken for a verdict.
if (-not (Test-Path -LiteralPath $Control -PathType Leaf)) {
    Write-Output "[ERROR] the control does not exist: $Control"
    exit 2
}
$controlFile = (Get-Item -LiteralPath $Control).FullName

$inputs = Resolve-Inputs $Path
if ($inputs.Missing.Count -gt 0) {
    foreach ($m in $inputs.Missing) { Write-Output "[ERROR] no file matched: $m" }
    exit 2
}
$files = $inputs.Files

$opened = @()
$refused = @()
$busy = @()
$excelExe = $null
$xl = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    Write-Output "Excel $($xl.Version) build $($xl.Build)"
    $excelExe = Join-Path $xl.Path 'EXCEL.EXE'
    $result = Test-Open $xl $controlFile
    if ($result.State -ne 'OPENS') {
        Write-Output "[ERROR] the control did not open ($($result.State)), so no verdict is valid: $controlFile -- $($result.Error)"
        exit 2
    }
    Write-Output "[OK] control opens: $controlFile"
    foreach ($f in $files) {
        $result = Test-Open $xl $f.FullName
        $entry = [pscustomobject]@{ File = $f.FullName; Error = $result.Error; Note = $result.Note }
        switch ($result.State) {
            'OPENS' { $opened += $entry }
            'REFUSED' { $refused += $entry }
            default { $busy += $entry }
        }
    }
} catch {
    Write-Output "[ERROR] Excel could not be started or failed through COM, so no verdict is valid: $($_.Exception.Message.Split("`n")[0])"
    exit 2
} finally {
    # Quit only the instance this script created; other Excel windows stay open.
    if ($null -ne $xl) {
        try { $xl.Quit() } catch { Write-Output "[WARN] Excel did not quit: $($_.Exception.Message.Split("`n")[0])" }
        Clear-Com $xl
    }
}

foreach ($o in $opened) {
    if ($o.Note) { Write-Output "OPENS    $($o.File) -- $($o.Note)" } else { Write-Output "OPENS    $($o.File)" }
}
if ($refused.Count -gt 0 -and -not $NoDialog) {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
}
foreach ($r in $refused) {
    $seen = if ($NoDialog) { "dialog not read (-NoDialog)" } else { Read-OpenDialog $excelExe $r.File $DialogTimeoutSeconds }
    Write-Output "REFUSED  $($r.File) -- $($r.Error) -- $seen"
}
foreach ($b in $busy) { Write-Output "BUSY     $($b.File) -- $($b.Error) -- could not check, re-run it on its own" }
Write-Output "$($opened.Count) of $($files.Count) opened, $($refused.Count) refused, $($busy.Count) busy"
if ($busy.Count -gt 0) { exit 2 }
if ($refused.Count -gt 0) { exit 1 }
exit 0
