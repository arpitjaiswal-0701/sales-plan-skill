# ============================================================================
# /sales plan — Adobe Lite DX Business Plan Skill Installer (Windows)
# ============================================================================
#
# Run from PowerShell:
#   .\install.ps1
#
# Or from a local clone:
#   cd sales-plan-skill
#   .\install.ps1

$ErrorActionPreference = "Stop"

Write-Host ""
Write-Host "============================================================" -ForegroundColor Blue
Write-Host "  /sales plan - Adobe Lite DX Business Plan Skill" -ForegroundColor Cyan
Write-Host "  1 skill  *  1 Python script  *  Auto-populated PPT" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Blue
Write-Host ""

# ---------------------------------------------------------------------------
# Locate source directory
# ---------------------------------------------------------------------------
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-Path "$ScriptDir\skills\sales-plan\SKILL.md") {
    $SourceDir = $ScriptDir
    Write-Host "Installing from local directory: $SourceDir" -ForegroundColor Green
} else {
    Write-Host "Clone the repo first, then run install.ps1 from inside it." -ForegroundColor Red
    Write-Host "  git clone https://github.com/YOUR_GITHUB_USERNAME/sales-plan-skill.git"
    exit 1
}

# ---------------------------------------------------------------------------
# Prerequisites
# ---------------------------------------------------------------------------
Write-Host "Checking prerequisites..." -ForegroundColor Blue

# Claude Code
if (Get-Command claude -ErrorAction SilentlyContinue) {
    Write-Host "  OK  Claude Code found" -ForegroundColor Green
} else {
    Write-Host "  !!  Claude Code CLI not found - skills will install but won't run without it" -ForegroundColor Yellow
    Write-Host "      Get it at: https://claude.ai/download" -ForegroundColor Yellow
}

# Python
$PythonCmd = $null
foreach ($cmd in @("py", "python", "python3")) {
    try {
        $ver = & $cmd --version 2>&1
        if ($ver -match "Python 3\.[89]|Python 3\.[1-9][0-9]") {
            $PythonCmd = $cmd
            Write-Host "  OK  Python found: $cmd ($ver)" -ForegroundColor Green
            break
        }
    } catch {}
}
if (-not $PythonCmd) {
    Write-Host "  !!  Python 3.8+ not found" -ForegroundColor Red
    Write-Host "      Install from https://www.python.org/downloads/" -ForegroundColor Red
    exit 1
}

# python-pptx
try {
    & $PythonCmd -c "import pptx" 2>&1 | Out-Null
    Write-Host "  OK  python-pptx installed" -ForegroundColor Green
} catch {
    Write-Host "  ..  Installing python-pptx..." -ForegroundColor Yellow
    & $PythonCmd -m pip install python-pptx lxml -q
    Write-Host "  OK  python-pptx installed" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Prompt: template path
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Configuration" -ForegroundColor Blue
Write-Host ""

Write-Host "  Enter the full path to your Lite DX Business Plan Template .pptx file."
Write-Host "  (Your personal copy synced from OneDrive.)"
Write-Host ""

# Search common OneDrive locations for any "Lite DX Business Plan Template*.pptx"
$DetectedTemplate = $null
$SearchDirs = @(
    "$env:USERPROFILE\OneDrive - Adobe",
    "$env:USERPROFILE\OneDrive",
    "$env:USERPROFILE\Documents"
)
foreach ($dir in $SearchDirs) {
    if (Test-Path $dir) {
        $match = Get-ChildItem -Path $dir -Filter "Lite DX Business Plan Template*.pptx" -Recurse -Depth 2 -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($match) { $DetectedTemplate = $match.FullName; break }
    }
}

if ($DetectedTemplate) {
    Write-Host "  Detected: $DetectedTemplate" -ForegroundColor Cyan
    $UseDetected = Read-Host "  Use this path? [Y/n]"
    if ($UseDetected -match "^[Nn]$") {
        $TemplatePath = Read-Host "  Template path"
    } else {
        $TemplatePath = $DetectedTemplate
    }
} else {
    $TemplatePath = Read-Host "  Template path"
}

if (-not (Test-Path $TemplatePath)) {
    Write-Host "  !!  File not found: $TemplatePath" -ForegroundColor Red
    Write-Host "      Make sure OneDrive is synced and try again." -ForegroundColor Red
    exit 1
}
Write-Host "  OK  Template found" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Prompt: deals root
# ---------------------------------------------------------------------------
$DefaultDeals = "$env:USERPROFILE\Desktop\claude-workspace\deals"
Write-Host ""
Write-Host "  Enter your deals root directory (where account folders live)."
Write-Host "  Default: $DefaultDeals" -ForegroundColor Cyan
$DealsRoot = Read-Host "  [Press Enter to accept default]"
if ([string]::IsNullOrWhiteSpace($DealsRoot)) { $DealsRoot = $DefaultDeals }
New-Item -ItemType Directory -Force -Path $DealsRoot | Out-Null
Write-Host "  OK  Deals root: $DealsRoot" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Install
# ---------------------------------------------------------------------------
$SkillsDir  = "$env:USERPROFILE\.claude\skills"
$SkillDest  = "$SkillsDir\sales-plan"
$ScriptDest = "$SkillDest\sales_plan_ppt.py"

New-Item -ItemType Directory -Force -Path $SkillDest | Out-Null

Write-Host ""
Write-Host "Installing..." -ForegroundColor Blue

# Use forward slashes for paths embedded in the skill/script (cross-platform safe)
$TemplatePathFwd = $TemplatePath.Replace('\', '/')
$DealsRootFwd    = $DealsRoot.Replace('\', '/')
$ScriptDestFwd   = $ScriptDest.Replace('\', '/')

# Patch and copy SKILL.md
$SkillContent = Get-Content "$SourceDir\skills\sales-plan\SKILL.md" -Raw
$SkillContent = $SkillContent.Replace('__DEALS_ROOT__',    $DealsRootFwd)
$SkillContent = $SkillContent.Replace('__PPT_SCRIPT__',    $ScriptDestFwd)
$SkillContent = $SkillContent.Replace('__TEMPLATE_PATH__', $TemplatePathFwd)
$SkillContent = $SkillContent.Replace('__PYTHON_CMD__',    $PythonCmd)
Set-Content -Path "$SkillDest\SKILL.md" -Value $SkillContent -Encoding utf8
Write-Host "  OK  sales-plan skill installed" -ForegroundColor Green

# Patch and copy Python script
$PyContent = Get-Content "$SourceDir\tools\sales_plan_ppt.py" -Raw
$PyContent = $PyContent.Replace('__TEMPLATE_PATH__', $TemplatePathFwd)
Set-Content -Path $ScriptDest -Value $PyContent -Encoding utf8
Write-Host "  OK  sales_plan_ppt.py installed" -ForegroundColor Green

# ---------------------------------------------------------------------------
# Check sub-skills
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "Checking required sub-skills..." -ForegroundColor Blue
$MissingSkills = @()
foreach ($skill in @("sales-prospect","sales-qualify","sales-contacts","sales-competitors","sales-prep","sales-outreach")) {
    if (Test-Path "$SkillsDir\$skill\SKILL.md") {
        Write-Host "  OK  $skill" -ForegroundColor Green
    } else {
        Write-Host "  !!  $skill - NOT FOUND" -ForegroundColor Red
        $MissingSkills += $skill
    }
}
if ($MissingSkills.Count -gt 0) {
    Write-Host ""
    Write-Host "  Missing sub-skills. Install the AI Sales Team first:" -ForegroundColor Yellow
    Write-Host "  curl -fsSL https://raw.githubusercontent.com/zubair-trabzada/ai-sales-team-claude/main/install.sh | bash" -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Skill:    $SkillDest\SKILL.md" -ForegroundColor Cyan
Write-Host "  Script:   $ScriptDest" -ForegroundColor Cyan
Write-Host "  Template: $TemplatePath" -ForegroundColor Cyan
Write-Host "  Deals:    $DealsRoot" -ForegroundColor Cyan
Write-Host ""
Write-Host "Quick start — open Claude Code and run:" -ForegroundColor Blue
Write-Host ""
Write-Host "  /sales plan https://www.example.com" -ForegroundColor Cyan
Write-Host ""
Write-Host "  With deal context:" -ForegroundColor Blue
Write-Host "  /sales plan https://www.example.com --arr=150000 --renewal=Q3 --stage=POC --close-date=2026-09-30" -ForegroundColor Cyan
Write-Host ""
