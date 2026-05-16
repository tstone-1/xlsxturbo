param(
    [string]$Root = $HOME,
    [string]$OutputCsv = "",
    [switch]$IncludeClaudeAutoMemory,

    # --- Mode: machine inventory ---
    # Emit a "what's actually cloned on this machine" report so agents can
    # verify sibling repos before invoking cross-repo workflows. Adds an
    # ENV/host column to the output.
    [switch]$Inventory,
    [string]$MachineLabel = $env:COMPUTERNAME,

    # --- Mode: public-safe pre-push check ---
    # Scan a single repo's committed text files for sensitive tokens that
    # must not leak into public repositories. Designed to be invoked from
    # a pre-push hook in public repos (xlsxturbo, dblitz). Exits non-zero
    # if any forbidden token is found.
    [switch]$CheckPublicSafe,
    [string]$CheckRepoPath = ""
)

$ErrorActionPreference = "Stop"

# Forbidden tokens for public repos. Curated from the agent-memory policy:
# internal corp domain, employee usernames, internal hostnames, alternate
# GitHub accounts, internal Gerrit base URL. Case-insensitive match.
# `tstone-1` and `48162401+tstone-1@users.noreply.github.com` are NOT forbidden — they're the
# legitimate public-repo identity.
$ForbiddenPatterns = @(
    # Internal corp infrastructure
    'nexperia\.com',
    'eu-rdc01',
    'eurdc01scm001',
    'OneDrive - Nexperia',

    # Corp/laptop usernames
    '\bnx016023\b',

    # Corp GitHub identity vectors (specific — avoids false positives on the
    # legitimate `48162401+tstone-1@users.noreply.github.com` personal email used in public repos).
    'timo\.stein@nexperia\.com',
    '--user\s+timo-stein\b',
    'gh\s+auth\s+switch.*timo-stein',
    'Co-Authored-By:.*timo-stein',
    'github\.com/timo-stein',

    # Alternate GitHub account names (account enumeration leak)
    '\becoproducts\b',
    'eco-products@',
    'github\.com/ecoproducts',

    # Internal project names that are confidential
    'eco-fmd',
    'eco-imds',
    'eco-camds',
    'eco-rt',
    'ecotools',
    'snowticket',
    'snowscreen',
    'sitm-explorer',
    'mdoc-extraction',
    'rt-checker',
    'rt-explorer',
    'chem-contents-server',
    'nexdoc-creator',
    'customer_reply',
    'ticket-creator2'
)
# Allowlist for files that legitimately reference some of the above.
# Paths are repo-relative, forward slashes, case-sensitive on POSIX.
# Add carve-outs sparingly and only for specific files, never wildcards.
$AllowedFiles = @(
    # The audit script itself necessarily contains every forbidden pattern.
    'scripts/audit-agent-memory.ps1'
)
# Directories whose contents we don't scan at all (binaries, vendored,
# generated artifacts, the .git dir itself).
$SkipDirs = @(
    '.git',
    'node_modules',
    'target',
    '.cargo',
    '.rustup',
    'dist',
    'build',
    '.venv',
    'venv',
    '__pycache__',
    '.pytest_cache',
    '.mypy_cache',
    'site-packages',
    'docs/_build',
    '.next',
    '.nuxt',
    'public/build',
    'agent-memory-audit*.csv'
)
# File extensions to scan. Text-only — never open binaries.
$ScanExtensions = @(
    '.md', '.txt', '.rst', '.adoc',
    '.json', '.toml', '.yaml', '.yml', '.ini', '.cfg', '.conf',
    '.ps1', '.psm1', '.sh', '.bash', '.zsh',
    '.py', '.pyi', '.rs', '.go', '.js', '.ts', '.jsx', '.tsx', '.svelte', '.vue',
    '.html', '.css', '.scss',
    '.sql', '.gitignore', '.gitattributes',
    'Dockerfile', 'Makefile'
)

function Get-RelativePathOrEmpty {
    param(
        [string]$Base,
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }

    try {
        return [System.IO.Path]::GetRelativePath($Base, $Path)
    }
    catch {
        return $Path
    }
}

function Test-FileTextContains {
    param(
        [string]$Path,
        [string]$Pattern
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $false
    }

    try {
        return (Select-String -LiteralPath $Path -Pattern $Pattern -SimpleMatch -Quiet)
    }
    catch {
        return $false
    }
}

function Get-ClaudeAutoMemoryPath {
    param([string]$RepoPath)

    # Claude derives project names internally. This audit intentionally avoids
    # relying on an unstable mapping and instead reports likely project memory
    # matches by normalized path suffix where possible.
    $memoryRoot = Join-Path $HOME ".claude\projects"
    if (-not (Test-Path -LiteralPath $memoryRoot -PathType Container)) {
        return ""
    }

    $repoName = Split-Path -Leaf $RepoPath
    $matches = Get-ChildItem -LiteralPath $memoryRoot -Directory -Recurse -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like "*$repoName*" -and
            (Test-Path -LiteralPath (Join-Path $_.FullName "memory\MEMORY.md") -PathType Leaf)
        } |
        Select-Object -First 3 -ExpandProperty FullName

    return ($matches -join ";")
}

function Invoke-PublicSafeCheck {
    param([string]$RepoPath)

    if ([string]::IsNullOrWhiteSpace($RepoPath)) {
        Write-Host "[public-safe-check] ERROR: -CheckRepoPath is required."
        return 2
    }
    $resolved = (Resolve-Path -LiteralPath $RepoPath -ErrorAction SilentlyContinue)
    if (-not $resolved) {
        Write-Host "[public-safe-check] ERROR: repo path not found: $RepoPath"
        return 2
    }
    $repoRoot = $resolved.Path

    Write-Host "[public-safe-check] Scanning $repoRoot for tokens that must not appear in a public repo..."

    # Only scan files that are TRACKED by git. Untracked / gitignored files
    # won't be pushed, so flagging them produces false positives that the
    # user can't fix without committing the local-only content.
    Push-Location $repoRoot
    try {
        $trackedRaw = git ls-files 2>$null
    } catch {
        $trackedRaw = $null
    } finally {
        Pop-Location
    }
    if (-not $trackedRaw) {
        Write-Host "[public-safe-check] WARNING: git ls-files returned nothing. Is this a git repo?"
        return 2
    }
    $tracked = @($trackedRaw -split "`n" | Where-Object { $_ -ne "" })

    $files = foreach ($rel in $tracked) {
        $abs = Join-Path $repoRoot $rel
        if (-not (Test-Path -LiteralPath $abs -PathType Leaf)) { continue }
        $item = Get-Item -LiteralPath $abs
        if ($ScanExtensions -contains $item.Extension -or $ScanExtensions -contains $item.Name) {
            $item
        }
    }

    $hits = @()
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($repoRoot.Length).TrimStart('\','/').Replace('\','/')
        if ($AllowedFiles -contains $rel) { continue }

        try {
            $content = Get-Content -LiteralPath $f.FullName -Raw -ErrorAction Stop
        } catch {
            continue
        }
        if ([string]::IsNullOrEmpty($content)) { continue }

        foreach ($pat in $ForbiddenPatterns) {
            # Avoid the $matches automatic variable name to keep regex state clean.
            $mts = [regex]::Matches($content, $pat, 'IgnoreCase')
            if ($mts.Count -gt 0) {
                # Compute line numbers for nicer reporting (max 3 per file/pattern).
                $lines = ($content -split "`n")
                $matchLines = @()
                for ($i = 0; $i -lt $lines.Length -and $matchLines.Count -lt 3; $i++) {
                    if ([regex]::IsMatch($lines[$i], $pat, 'IgnoreCase')) {
                        $matchLines += ($i + 1)
                    }
                }
                $hits += [pscustomobject]@{
                    File    = $rel
                    Pattern = $pat
                    Hits    = $mts.Count
                    Lines   = ($matchLines -join ',')
                }
            }
        }
    }

    if ($hits.Count -eq 0) {
        Write-Host "[public-safe-check] OK - no forbidden tokens found."
        return 0
    }

    Write-Host ""
    Write-Host "[public-safe-check] FAIL - forbidden tokens found:"
    Write-Host ""
    $tbl = $hits | Sort-Object File, Pattern | Format-Table File, Pattern, Hits, Lines -AutoSize | Out-String
    Write-Host $tbl.Trim()
    Write-Host ""
    Write-Host "These tokens must not appear in a public repository. Either:"
    Write-Host "  1. Remove/rephrase the content to remove the token, or"
    Write-Host "  2. If the match is a documented false positive, add a carve-out to"
    Write-Host "     `$AllowedFiles in audit-agent-memory.ps1 (specific path, never wildcard)."
    Write-Host ""
    Write-Host "To override this check for a single push (use with care):"
    Write-Host "  PowerShell:  `$env:SKIP_AGENT_MEMORY_AUDIT='1'; git push"
    Write-Host "  bash:        SKIP_AGENT_MEMORY_AUDIT=1 git push"
    return 1
}

function Invoke-Inventory {
    param([string]$ResolvedRoot, [string]$Label)

    Write-Host "[inventory] Machine: $Label   Root: $ResolvedRoot"

    $gitDirs = Get-ChildItem -LiteralPath $ResolvedRoot -Directory -Force -Recurse -Filter ".git" -ErrorAction SilentlyContinue |
        Where-Object {
        # Skip non-user repos: deps caches, browser profiles, app-data clones,
        # OS-managed locations, vendored builds. These show up as .git dirs
        # under recursive scans of $HOME and aren't actually user projects.
        $_.FullName -notmatch "\\(\.venv|venv|node_modules|target|\.cargo|\.rustup|AppData|Library|Extensions|browser-profile-test|edge-test|cache|Cache|\.local|\.cache|\.gradle|\.m2|\.npm|\.pnpm-store|dist|build|site-packages)\\"
    }

    $rows = foreach ($gitDir in $gitDirs) {
        $repo = Split-Path -Parent $gitDir.FullName
        $rel  = Get-RelativePathOrEmpty -Base $ResolvedRoot -Path $repo
        $name = Split-Path -Leaf $repo

        $remote = ""
        try {
            Push-Location $repo
            $remote = (git config --get remote.origin.url 2>$null)
        } catch { } finally { Pop-Location }

        [pscustomobject]@{
            Machine     = $Label
            Repo        = $name
            RelativePath = $rel
            AbsolutePath = $repo
            Remote      = $remote
            HasAgents   = (Test-Path -LiteralPath (Join-Path $repo "AGENTS.md") -PathType Leaf)
            HasClaude   = (Test-Path -LiteralPath (Join-Path $repo "CLAUDE.md") -PathType Leaf)
        }
    }

    $sorted = $rows | Sort-Object Repo
    if ($OutputCsv) {
        $outputPath = $OutputCsv
        if (-not [System.IO.Path]::IsPathRooted($outputPath)) {
            $outputPath = Join-Path (Get-Location).Path $outputPath
        }
        $sorted | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
        Write-Host "[inventory] Wrote $($sorted.Count) rows to $outputPath"
    } else {
        $sorted | Format-Table Machine, Repo, RelativePath, HasAgents, HasClaude -AutoSize
    }
}

# --- Dispatch new modes ---
if ($CheckPublicSafe) {
    $rc = Invoke-PublicSafeCheck -RepoPath $CheckRepoPath
    exit $rc
}

$resolvedRoot = (Resolve-Path -LiteralPath $Root).Path

if ($Inventory) {
    Invoke-Inventory -ResolvedRoot $resolvedRoot -Label $MachineLabel
    exit 0
}
$gitDirs = Get-ChildItem -LiteralPath $resolvedRoot -Directory -Force -Recurse -Filter ".git" -ErrorAction SilentlyContinue |
    Where-Object {
        # Skip non-user repos: deps caches, browser profiles, app-data clones,
        # OS-managed locations, vendored builds. These show up as .git dirs
        # under recursive scans of $HOME and aren't actually user projects.
        $_.FullName -notmatch "\\(\.venv|venv|node_modules|target|\.cargo|\.rustup|AppData|Library|Extensions|browser-profile-test|edge-test|cache|Cache|\.local|\.cache|\.gradle|\.m2|\.npm|\.pnpm-store|dist|build|site-packages)\\"
    }

$rows = foreach ($gitDir in $gitDirs) {
    $repo = Split-Path -Parent $gitDir.FullName

    $agents = Join-Path $repo "AGENTS.md"
    $agentsOverride = Join-Path $repo "AGENTS.override.md"
    $rootClaude = Join-Path $repo "CLAUDE.md"
    $dotClaude = Join-Path $repo ".claude\CLAUDE.md"
    $claudeRules = Join-Path $repo ".claude\rules"

    $hasAgents = Test-Path -LiteralPath $agents -PathType Leaf
    $hasAgentsOverride = Test-Path -LiteralPath $agentsOverride -PathType Leaf
    $hasRootClaude = Test-Path -LiteralPath $rootClaude -PathType Leaf
    $hasDotClaude = Test-Path -LiteralPath $dotClaude -PathType Leaf
    $hasClaudeRules = Test-Path -LiteralPath $claudeRules -PathType Container
    $dotClaudeImportsAgents = Test-FileTextContains -Path $dotClaude -Pattern "@../AGENTS.md"
    $rootClaudeImportsAgents = Test-FileTextContains -Path $rootClaude -Pattern "@AGENTS.md"

    $status = if ($hasAgents -and ($dotClaudeImportsAgents -or $rootClaudeImportsAgents -or (-not ($hasRootClaude -or $hasDotClaude)))) {
        "shared-memory-ready"
    }
    elseif ($hasAgents -and ($hasRootClaude -or $hasDotClaude)) {
        "check-for-drift"
    }
    elseif (($hasRootClaude -or $hasDotClaude -or $hasClaudeRules) -and -not $hasAgents) {
        "needs-agents-md"
    }
    elseif ($hasAgents) {
        "codex-only"
    }
    else {
        "no-agent-memory"
    }

    $autoMemory = ""
    if ($IncludeClaudeAutoMemory) {
        $autoMemory = Get-ClaudeAutoMemoryPath -RepoPath $repo
    }

    [pscustomobject]@{
        Status = $status
        Repo = $repo
        HasAgents = $hasAgents
        HasAgentsOverride = $hasAgentsOverride
        HasRootClaude = $hasRootClaude
        HasDotClaude = $hasDotClaude
        HasClaudeRules = $hasClaudeRules
        RootClaudeImportsAgents = $rootClaudeImportsAgents
        DotClaudeImportsAgents = $dotClaudeImportsAgents
        ClaudeAutoMemoryCandidates = $autoMemory
        RelativeRepo = Get-RelativePathOrEmpty -Base $resolvedRoot -Path $repo
    }
}

$sortedRows = $rows | Sort-Object Status, Repo

if ($OutputCsv) {
    $outputPath = $OutputCsv
    if (-not [System.IO.Path]::IsPathRooted($outputPath)) {
        $outputPath = Join-Path (Get-Location).Path $outputPath
    }
    $sortedRows | Export-Csv -LiteralPath $outputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Wrote $($sortedRows.Count) rows to $outputPath"
}
else {
    $sortedRows | Format-Table Status, RelativeRepo, HasAgents, HasRootClaude, HasDotClaude, HasClaudeRules -AutoSize
    Write-Host ""
    $sortedRows | Group-Object Status | Sort-Object Name | Format-Table Name, Count -AutoSize
}
