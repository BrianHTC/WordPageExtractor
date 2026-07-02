<#
.SYNOPSIS
  Extract Word pages into separate or combined .docx files while preserving original visible page numbers. Fixed v2.

.DESCRIPTION
  Output modes used by this customized version:
    S = Combined file: all designated pages/ranges are extracted together into one output file.
    C = Separate files: each comma-separated input range becomes its own output file.

  Example combined input:
    1-3,8
  If user chooses C:
    one .docx containing original pages 1, 2, 3, and 8.
  If user chooses S:
    separate .docx files for 1-3 and 8.

.REQUIREMENTS
  - Windows
  - Microsoft Word desktop app installed

.PAGE RANGE FORMAT
  Examples:
    1-3
    1-3,5,8-10
#>

param(
    [string]$InputPath,
    [string]$Ranges,
    [string]$OutputFolder,
    [ValidateSet('S','C','Separate','Combined')]
    [string]$Mode,
    [switch]$OpenOutputFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Set-ConsoleFontSize {
    param([int]$Size = 18)

    try {
        if (-not ('ConsoleFont.NativeMethods' -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace ConsoleFont {
    [StructLayout(LayoutKind.Sequential)]
    public struct COORD {
        public short X;
        public short Y;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct CONSOLE_FONT_INFO_EX {
        public uint cbSize;
        public uint nFont;
        public COORD dwFontSize;
        public int FontFamily;
        public int FontWeight;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string FaceName;
    }

    public class NativeMethods {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetCurrentConsoleFontEx(
            IntPtr hConsoleOutput,
            bool bMaximumWindow,
            ref CONSOLE_FONT_INFO_EX lpConsoleCurrentFontEx
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetCurrentConsoleFontEx(
            IntPtr hConsoleOutput,
            bool bMaximumWindow,
            ref CONSOLE_FONT_INFO_EX lpConsoleCurrentFontEx
        );
    }
}
"@
        }

        $STD_OUTPUT_HANDLE = -11
        $handle = [ConsoleFont.NativeMethods]::GetStdHandle($STD_OUTPUT_HANDLE)

        $fontInfo = New-Object ConsoleFont.CONSOLE_FONT_INFO_EX
        $fontInfo.cbSize = [Runtime.InteropServices.Marshal]::SizeOf($fontInfo)

        [void][ConsoleFont.NativeMethods]::GetCurrentConsoleFontEx($handle, $false, [ref]$fontInfo)

        $fontInfo.dwFontSize.Y = [short]$Size
        $fontInfo.dwFontSize.X = 0

        [void][ConsoleFont.NativeMethods]::SetCurrentConsoleFontEx($handle, $false, [ref]$fontInfo)
    }
    catch {
        # Ignore font-size errors silently because some console hosts, especially Windows Terminal, override this.
    }
}

Set-ConsoleFontSize 18

function Read-NonEmpty([string]$PromptText) {
    do {
        Write-Host $PromptText -ForegroundColor Green
        $v = $Host.UI.ReadLine()
    } while ([string]::IsNullOrWhiteSpace($v))
    return $v.Trim().Trim('"')
}

function Read-Mode() {
    do {
        Write-Host ''
        Write-Host 'Choose output mode, enter S or C:' -ForegroundColor Green
        Write-Host '  S = Separate files: 1-3,8 creates two files'-ForegroundColor Green
        Write-Host '  C = Combined file: 1-3,8 creates one file containing pages 1-3 and 8'-ForegroundColor Green
        $m = $Host.UI.ReadLine()
        $m = $m.Trim()
    } while ($m -notin @('S','s','C','c'))
    return $m.ToUpperInvariant()
}

function Parse-PageRanges([string]$Text) {
    $items = @()
    foreach ($part in ($Text -split ',')) {
        $p = $part.Trim()
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if ($p -match '^\d+$') {
            $n = [int]$p
            if ($n -lt 1) { throw "Invalid page number: $p" }
            $items += [pscustomobject]@{ Start = $n; End = $n }
        }
        elseif ($p -match '^(\d+)\s*-\s*(\d+)$') {
            $s = [int]$Matches[1]
            $e = [int]$Matches[2]
            if ($s -lt 1 -or $e -lt 1 -or $s -gt $e) { throw "Invalid page range: $p" }
            $items += [pscustomobject]@{ Start = $s; End = $e }
        }
        else {
            throw "Invalid range format: '$p'. Use examples like 1-3,5,8-10"
        }
    }
    if ($items.Count -eq 0) { throw 'No valid page ranges were entered.' }
    return $items
}

function Normalize-PageRanges($RangesList, [int]$TotalPages) {
    $pages = New-Object System.Collections.Generic.List[int]
    foreach ($r in $RangesList) {
        $s = [int]$r.Start
        $e = [int]$r.End
        if ($s -gt $TotalPages) { throw "Range starts after the last page: $s-$e, total pages: $TotalPages" }
        if ($e -gt $TotalPages) {
            Write-Warning "Range $s-$e exceeds total pages. It will be clipped to $TotalPages."
            $e = $TotalPages
        }
        for ($p = $s; $p -le $e; $p++) { [void]$pages.Add($p) }
    }
    $uniqueSorted = @($pages | Sort-Object -Unique)
    return @($uniqueSorted | ForEach-Object { [int]$_ })
}

function Convert-PagesToBlocks([int[]]$Pages) {
    $Pages = @($Pages)
    $blocks = @()
    if ($Pages.Count -eq 0) { return $blocks }

    $sorted = @($Pages | Sort-Object)
    $blockStart = [int]$sorted[0]
    $prev = [int]$sorted[0]
    $outStart = 1
    $len = 1

    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $p = [int]$sorted[$i]
        if ($p -eq ($prev + 1)) {
            $len++
        }
        else {
            $blocks += [pscustomobject]@{ OriginalStart = $blockStart; Length = $len; OutputStart = $outStart }
            $outStart += $len
            $blockStart = $p
            $len = 1
        }
        $prev = $p
    }
    $blocks += [pscustomobject]@{ OriginalStart = $blockStart; Length = $len; OutputStart = $outStart }
    return $blocks
}

function Safe-FileName([string]$Name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) { $Name = $Name.Replace($ch, '_') }
    return $Name
}

function Make-RangeLabel($RangesList, [int]$TotalPages) {
    $parts = @()
    foreach ($r in $RangesList) {
        $s = [int]$r.Start
        $e = [int]$r.End
        if ($s -gt $TotalPages) { throw "Range starts after the last page: $s-$e, total pages: $TotalPages" }
        if ($e -gt $TotalPages) { $e = $TotalPages }
        if ($s -eq $e) { $parts += ('p{0:D3}' -f $s) }
        else { $parts += ('p{0:D3}-p{1:D3}' -f $s, $e) }
    }
    return ($parts -join '_')
}

function Save-WordDocx($Document, [string]$Path) {
    $wdFormatXMLDocument = 12
    try { $Document.SaveAs2($Path, $wdFormatXMLDocument) }
    catch { $Document.SaveAs($Path, $wdFormatXMLDocument) }
}

function Get-PageCount($Document) {
    $Document.Repaginate()
    return [int]$Document.ComputeStatistics(2) # wdStatisticPages = 2
}

function Delete-PageByNumber($WordApp, $Document, [int]$PageNumber) {
    $Document.Activate()
    [void]$WordApp.Selection.GoTo(1, 1, $PageNumber) # wdGoToPage = 1, wdGoToAbsolute = 1
    $pageRange = $WordApp.Selection.Bookmarks.Item('\Page').Range
    [void]$pageRange.Delete()
}

function Keep-Only-PageSet($WordApp, $Document, [int[]]$PagesToKeep, [int]$OriginalTotalPages) {
    $PagesToKeep = @($PagesToKeep)
    $keep = @{}
    foreach ($p in $PagesToKeep) { $keep[[int]$p] = $true }

    for ($p = $OriginalTotalPages; $p -ge 1; $p--) {
        if (-not $keep.ContainsKey($p)) {
            Delete-PageByNumber -WordApp $WordApp -Document $Document -PageNumber $p
        }
    }
}

function Remove-Trailing-BlankPages-IfAny($WordApp, $Document, [int]$ExpectedPages) {
    $maxPasses = 5
    for ($i = 0; $i -lt $maxPasses; $i++) {
        $currentPages = Get-PageCount -Document $Document
        if ($currentPages -le $ExpectedPages) { break }

        $Document.Activate()
        [void]$WordApp.Selection.GoTo(1, 1, $currentPages)
        $lastPageRange = $WordApp.Selection.Bookmarks.Item('\Page').Range
        $txt = [string]$lastPageRange.Text
        $normalized = $txt -replace "[\r\n\f\v\t ]", ""

        if ($normalized.Length -eq 0) { [void]$lastPageRange.Delete() }
        else { break }
    }
}

function Get-SectionAtOutputPage($WordApp, $Document, [int]$OutputPage) {
    $Document.Activate()
    [void]$WordApp.Selection.GoTo(1, 1, $OutputPage)
    $rng = $WordApp.Selection.Range
    return $rng.Sections.Item(1)
}

function Set-SectionPageNumberStart($Section, [int]$StartNumber) {
    foreach ($idx in 1,2,3) {
        try {
            $pn = $Section.Footers.Item($idx).PageNumbers
            $pn.RestartNumberingAtSection = $true
            $pn.StartingNumber = $StartNumber
        } catch { }
    }
    foreach ($idx in 1,2,3) {
        try {
            $pn = $Section.Headers.Item($idx).PageNumbers
            $pn.RestartNumberingAtSection = $true
            $pn.StartingNumber = $StartNumber
        } catch { }
    }
}

function Preserve-OriginalPageNumbers($WordApp, $Document, $Blocks) {
    $Blocks = @($Blocks)
    $wdSectionBreakContinuous = 3

    $Document.Activate()
    $Document.Repaginate()

    for ($i = $Blocks.Count - 1; $i -ge 1; $i--) {
        $outPage = [int]$Blocks[$i].OutputStart
        [void]$WordApp.Selection.GoTo(1, 1, $outPage)
        $WordApp.Selection.InsertBreak($wdSectionBreakContinuous)
    }

    $Document.Repaginate()

    foreach ($b in $Blocks) {
        $outPage = [int]$b.OutputStart
        $origStart = [int]$b.OriginalStart
        $sec = Get-SectionAtOutputPage -WordApp $WordApp -Document $Document -OutputPage $outPage
        Set-SectionPageNumberStart -Section $sec -StartNumber $origStart
    }

    try { [void]$Document.Fields.Update() } catch { }
    foreach ($s in $Document.Sections) {
        foreach ($idx in 1,2,3) {
            try { [void]$s.Headers.Item($idx).Range.Fields.Update() } catch { }
            try { [void]$s.Footers.Item($idx).Range.Fields.Update() } catch { }
        }
    }
    $Document.Repaginate()
}

function Create-ExtractedDoc($WordApp, [string]$InputPath, [string]$OutPath, [int[]]$PagesToKeep, [int]$OriginalTotalPages) {
    $PagesToKeep = @($PagesToKeep)
    $workDoc = $null
    try {
        $blocks = Convert-PagesToBlocks -Pages $PagesToKeep

        $workDoc = $WordApp.Documents.Open($InputPath, $false, $true)
        $workDoc.Repaginate()
        Save-WordDocx -Document $workDoc -Path $OutPath
        $workDoc.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workDoc) | Out-Null
        $workDoc = $null

        $workDoc = $WordApp.Documents.Open($OutPath, $false, $false)
        $workDoc.Repaginate()
        Keep-Only-PageSet -WordApp $WordApp -Document $workDoc -PagesToKeep $PagesToKeep -OriginalTotalPages $OriginalTotalPages
        $workDoc.Repaginate()
        Remove-Trailing-BlankPages-IfAny -WordApp $WordApp -Document $workDoc -ExpectedPages ($PagesToKeep.Count)
        $workDoc.Repaginate()

        Preserve-OriginalPageNumbers -WordApp $WordApp -Document $workDoc -Blocks $blocks

        $finalPages = Get-PageCount -Document $workDoc
        $workDoc.Save()
        Write-Host "Created: $OutPath" -ForegroundColor Cyan
        Write-Host "Final pages detected: $finalPages; expected approximately: $($PagesToKeep.Count)" -ForegroundColor Cyan
        Write-Host "Page numbering starts preserved at: $((@($blocks | ForEach-Object { $_.OriginalStart })) -join ', ')" -ForegroundColor Cyan
    }
    finally {
        if ($null -ne $workDoc) {
            $workDoc.Close($false) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workDoc) | Out-Null
        }
    }
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $InputPath = Read-NonEmpty 'Enter the full path of the Word file (.docx/.doc)'
}
$InputPath = [System.IO.Path]::GetFullPath($InputPath.Trim().Trim('"'))
if (-not (Test-Path -LiteralPath $InputPath)) { throw "File not found: $InputPath" }

if ([string]::IsNullOrWhiteSpace($Ranges)) {
    $Ranges = Read-NonEmpty 'Enter page ranges, e.g. 1-3 or 1-3,5,8-10'
}
$pageRanges = Parse-PageRanges $Ranges

if ([string]::IsNullOrWhiteSpace($Mode)) { $Mode = Read-Mode }
if ($Mode -eq 'Separate') { $Mode = 'C' }
if ($Mode -eq 'Combined') { $Mode = 'S' }
$Mode = $Mode.ToUpperInvariant()

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $defaultOut = Join-Path (Split-Path -Parent $InputPath) (([System.IO.Path]::GetFileNameWithoutExtension($InputPath)) + '_extracted_pages_keep_original_numbers')
    Write-Host "Enter output folder, or press Enter for: $defaultOut" -ForegroundColor Green
    $o = $Host.UI.ReadLine()
    if ([string]::IsNullOrWhiteSpace($o)) { $OutputFolder = $defaultOut } else { $OutputFolder = $o.Trim().Trim('"') }
}
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$word = $null
$sourceDoc = $null
try {
    Write-Host 'Starting Microsoft Word...' -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Write-Host "Opening source for page count: $InputPath" -ForegroundColor Cyan
    $sourceDoc = $word.Documents.Open($InputPath, $false, $true)
    $sourceDoc.Repaginate()
    $totalPages = Get-PageCount -Document $sourceDoc
    Write-Host "Total pages detected by Word: $totalPages" -ForegroundColor Cyan
    $sourceDoc.Close($false) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceDoc) | Out-Null
    $sourceDoc = $null

    $baseName = Safe-FileName ([System.IO.Path]::GetFileNameWithoutExtension($InputPath))

    if ($Mode -eq 'C') {
        $pagesToKeep = Normalize-PageRanges -RangesList $pageRanges -TotalPages $totalPages
        $label = Make-RangeLabel -RangesList $pageRanges -TotalPages $totalPages
        $outPath = Join-Path $OutputFolder ("{0}_{1}_combined_keepOriginalPageNumbers.docx" -f $baseName, $label)
        Write-Host "Creating combined file containing original pages: $($pagesToKeep -join ', ')" -ForegroundColor Cyan
        Create-ExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages
    }
    else {
        foreach ($r in $pageRanges) {
            $singleRange = @($r)
            $pagesToKeep = Normalize-PageRanges -RangesList $singleRange -TotalPages $totalPages
            $label = Make-RangeLabel -RangesList $singleRange -TotalPages $totalPages
            $outPath = Join-Path $OutputFolder ("{0}_{1}_keepOriginalPageNumbers.docx" -f $baseName, $label)
            Write-Host "Creating separate file containing original pages: $($pagesToKeep -join ', ')" -ForegroundColor Cyan
            Create-ExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages
        }
    }

    Write-Host ''
    Write-Host 'Done.' -ForegroundColor Cyan
    Write-Host "Output folder: $OutputFolder" -ForegroundColor Cyan
    if ($OpenOutputFolder) { Start-Process explorer.exe $OutputFolder }
}
catch {
    Write-Host ''
    Write-Host 'Error:' -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    if ($null -ne $sourceDoc) {
        $sourceDoc.Close($false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sourceDoc) | Out-Null
    }
    if ($null -ne $word) {
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
