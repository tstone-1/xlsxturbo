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
# With DisplayAlerts off, a file Excel would offer to repair fails to open with
# "Unable to get the Open property of the Workbooks class".
#
# Exit codes: 0 every file opened, 1 at least one refused, 2 the control failed
# or no file matched.
param(
    [Parameter(Mandatory = $true)][string]$Control,
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

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$refused = 0
try {
    Write-Output "Excel $($xl.Version) build $($xl.Build)"
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
            Write-Output "OPENS    $($f.FullName)"
        } else {
            Write-Output "REFUSED  $($f.FullName) -- $err"
            $refused++
        }
    }
} finally {
    # Quit only the instance this script created; other Excel windows stay open.
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
Write-Output "$($files.Count - $refused) of $($files.Count) opened"
if ($refused -gt 0) { exit 1 }
exit 0
