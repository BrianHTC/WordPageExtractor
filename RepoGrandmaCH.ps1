<#
.SYNOPSIS
  Editable Word page extractor with extra-page hard trim.

.DESCRIPTION
  C = combined output file.
  S = separate output files by comma-separated ranges.
  The output remains editable. The script saves a full copy first, deletes unselected pages, then aggressively trims extra trailing pages.
#>

param(
    [string]$InputPath,
    [string]$Ranges,
    [string]$OutputFolder,
    [ValidateSet("S","C","Separate","Combined")]
    [string]$Mode,
    [switch]$OpenOutputFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Read-NonEmpty([string]$PromptText) {
    do {
        Write-Host $PromptText -ForegroundColor Green
        $v = $Host.UI.ReadLine()
    } while ([string]::IsNullOrWhiteSpace($v))
    return $v.Trim().Trim([char]34)
}

function Read-Mode() {
    do {
        Write-Host ""
        Write-Host "選擇輸出模式，要一個檔案的話按C，要用逗號分開的範圍(1,2,3-4分成3個)分成不同的檔案的話按S" -ForegroundColor Green
        $m = $Host.UI.ReadLine()
        $m = $m.Trim()
    } while ($m -notin @("S","s","C","c"))
    return $m.ToUpperInvariant()
}

function Parse-PageRanges([string]$Text) {
    $items = @()
    foreach ($part in ($Text -split ",")) {
        $p = $part.Trim()
        if ([string]::IsNullOrWhiteSpace($p)) { continue }

        if ($p -match "^\d+$") {
            $n = [int]$p
            if ($n -lt 1) { throw "Invalid page number: $p" }
            $items += [pscustomobject]@{ Start = $n; End = $n }
        }
        elseif ($p -match "^(\d+)\s*-\s*(\d+)$") {
            $s = [int]$Matches[1]
            $e = [int]$Matches[2]
            if ($s -lt 1 -or $e -lt 1 -or $s -gt $e) { throw "Invalid page range: $p" }
            $items += [pscustomobject]@{ Start = $s; End = $e }
        }
        else {
            throw "Invalid range format: $p. Use examples like 1-3,5,8-10"
        }
    }
    if ($items.Count -eq 0) { throw "No valid page ranges were entered." }
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

function Safe-FileName([string]$Name) {
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($ch in $invalid) { $Name = $Name.Replace($ch, "_") }
    return $Name
}

function Make-RangeLabel($RangesList, [int]$TotalPages) {
    $parts = @()
    foreach ($r in $RangesList) {
        $s = [int]$r.Start
        $e = [int]$r.End
        if ($s -gt $TotalPages) { throw "Range starts after the last page: $s-$e, total pages: $TotalPages" }
        if ($e -gt $TotalPages) { $e = $TotalPages }
        if ($s -eq $e) { $parts += ("p{0:D3}" -f $s) }
        else { $parts += ("p{0:D3}-p{1:D3}" -f $s, $e) }
    }
    return ($parts -join "_")
}

function Save-WordDocx($Document, [string]$Path) {
    $wdFormatXMLDocument = 12
    try { $Document.SaveAs2($Path, $wdFormatXMLDocument) }
    catch { $Document.SaveAs($Path, $wdFormatXMLDocument) }
}

function Get-PageCount($Document) {
    $Document.Repaginate()
    return [int]$Document.ComputeStatistics(2)
}

function Safe-CloseDoc($Document) {
    if ($null -ne $Document) {
        try { $Document.Close($false) | Out-Null } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null } catch { }
    }
}

function Delete-PageByNumber($WordApp, $Document, [int]$PageNumber) {
    $Document.Activate()
    [void]$WordApp.Selection.GoTo(1, 1, $PageNumber)
    $pageRange = $WordApp.Selection.Bookmarks.Item("\Page").Range
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

function Trim-EndingEmptyParagraphs($Document) {
    try {
        for ($i = 0; $i -lt 30; $i++) {
            $paras = $Document.Paragraphs
            if ($paras.Count -le 1) { break }
            $lastPara = $paras.Item($paras.Count)
            $txt = [string]$lastPara.Range.Text
            $normalized = $txt -replace "[\r\n\f\v\t ]", "" -replace [char]7, "" -replace [char]160, ""
            if ($normalized.Length -eq 0) { [void]$lastPara.Range.Delete() }
            else { break }
        }
    } catch { }
}

function Compress-FinalEmptyParagraph($Document) {
    try {
        # Word requires a final paragraph mark after a table. If the table ends near the bottom
        # of the selected last page, that required empty paragraph can be pushed to a new blank page.
        # We keep the document editable and avoid the blank page by shrinking/hiding that final mark.
        if ($Document.Paragraphs.Count -lt 1) { return }
        $lastPara = $Document.Paragraphs.Item($Document.Paragraphs.Count)
        $txt = [string]$lastPara.Range.Text
        $normalized = $txt -replace "[\r\n\f\v\t ]", "" -replace [char]7, "" -replace [char]160, ""

        if ($normalized.Length -eq 0) {
            try { $lastPara.Range.Font.Size = 1 } catch { }
            try { $lastPara.Range.Font.Hidden = $true } catch { }
            try { $lastPara.Range.ParagraphFormat.SpaceBefore = 0 } catch { }
            try { $lastPara.Range.ParagraphFormat.SpaceAfter = 0 } catch { }
            try { $lastPara.Range.ParagraphFormat.LineSpacingRule = 4 } catch { } # wdLineSpaceExactly
            try { $lastPara.Range.ParagraphFormat.LineSpacing = 1 } catch { }
            try { $lastPara.Range.ParagraphFormat.PageBreakBefore = $false } catch { }
            try { $lastPara.Range.ParagraphFormat.KeepWithNext = $false } catch { }
            try { $lastPara.Range.ParagraphFormat.KeepTogether = $false } catch { }
        }
    } catch { }
}

function Remove-TrailingEmptyTableRows($Document) {
    try {
        for ($pass = 0; $pass -lt 80; $pass++) {
            if ($Document.Tables.Count -le 0) { break }
            $lastTable = $Document.Tables.Item($Document.Tables.Count)
            if ($lastTable.Rows.Count -le 0) { break }
            $lastRow = $lastTable.Rows.Item($lastTable.Rows.Count)
            $txt = [string]$lastRow.Range.Text
            $normalized = $txt -replace "[\r\n\f\v\t ]", "" -replace [char]7, "" -replace [char]160, ""
            if ($normalized.Length -eq 0) { [void]$lastRow.Delete() }
            else { break }
        }
    } catch { }
}

function Delete-RangeFromPageToEnd($WordApp, $Document, [int]$StartPage) {
    $Document.Activate()
    $Document.Repaginate()
    $total = Get-PageCount -Document $Document
    if ($StartPage -gt $total) { return }

    # Stronger than deleting \Page one by one:
    # for table-based files, extra pages can be made of empty table rows/cells.
    # Deleting one \Page may leave row/cell structure behind, but deleting the range
    # from the start of the first unwanted page to the document end removes the tail.
    try {
        [void]$WordApp.Selection.GoTo(1, 1, $StartPage)
        $startPos = [int]$WordApp.Selection.Range.Start
        $endPos = [int]$Document.Content.End
        if ($endPos -gt $startPos) {
            $tailRange = $Document.Range($startPos, $endPos)
            [void]$tailRange.Delete()
        }
    }
    catch {
        # Fallback to page-by-page deletion if range deletion fails.
        for ($p = $total; $p -ge $StartPage; $p--) {
            try {
                Delete-PageByNumber -WordApp $WordApp -Document $Document -PageNumber $p
                $Document.Repaginate()
            } catch {
                break
            }
        }
    }
}

function Enforce-ExpectedPageCount($WordApp, $Document, [int]$ExpectedPages) {
    for ($pass = 0; $pass -lt 10; $pass++) {
        $current = Get-PageCount -Document $Document
        if ($current -le $ExpectedPages) { break }

        # Delete everything starting at the first unwanted page.
        Delete-RangeFromPageToEnd -WordApp $WordApp -Document $Document -StartPage ($ExpectedPages + 1)
        Remove-TrailingEmptyTableRows -Document $Document
        Trim-EndingEmptyParagraphs -Document $Document
        $Document.Repaginate()
    }
}

function Create-EditableExtractedDoc($WordApp, [string]$InputPath, [string]$OutPath, [int[]]$PagesToKeep, [int]$OriginalTotalPages) {
    $PagesToKeep = @($PagesToKeep)
    $workDoc = $null
    try {
        Write-Host "Saving full copy first to preserve editable layout..." -ForegroundColor Cyan
        $workDoc = $WordApp.Documents.Open($InputPath, $false, $true)
        $workDoc.Repaginate()
        Save-WordDocx -Document $workDoc -Path $OutPath
        Safe-CloseDoc $workDoc
        $workDoc = $null

        Write-Host "Deleting unselected pages from editable copy..." -ForegroundColor Cyan
        $workDoc = $WordApp.Documents.Open($OutPath, $false, $false)
        $workDoc.Repaginate()
        Keep-Only-PageSet -WordApp $WordApp -Document $workDoc -PagesToKeep $PagesToKeep -OriginalTotalPages $OriginalTotalPages
        $workDoc.Repaginate()

        Remove-TrailingEmptyTableRows -Document $workDoc
        Trim-EndingEmptyParagraphs -Document $workDoc
        Compress-FinalEmptyParagraph -Document $workDoc
        $workDoc.Repaginate()

        Enforce-ExpectedPageCount -WordApp $WordApp -Document $workDoc -ExpectedPages ($PagesToKeep.Count)
        Compress-FinalEmptyParagraph -Document $workDoc
        $workDoc.Repaginate()

        $workDoc.Save()
        $finalPages = Get-PageCount -Document $workDoc
        if ($finalPages -gt $PagesToKeep.Count) {
            # One more pass after saving, because Word can repaginate on save.
            Compress-FinalEmptyParagraph -Document $workDoc
            Enforce-ExpectedPageCount -WordApp $WordApp -Document $workDoc -ExpectedPages ($PagesToKeep.Count)
            Compress-FinalEmptyParagraph -Document $workDoc
            $workDoc.Repaginate()
            $workDoc.Save()
            $finalPages = Get-PageCount -Document $workDoc
        }
        Write-Host "Created: $OutPath" -ForegroundColor Cyan
        Write-Host "Engine: Editable full-copy-delete with final-paragraph compression v3" -ForegroundColor Cyan
        Write-Host "Final pages detected by Word: $finalPages; requested pages: $($PagesToKeep.Count)" -ForegroundColor Cyan
    }
    finally {
        if ($null -ne $workDoc) { Safe-CloseDoc $workDoc }
    }
}

if ([string]::IsNullOrWhiteSpace($InputPath)) {
    $InputPath = Read-NonEmpty "對著你想拆分的檔案按一下右鍵，再對「複製路徑」的選項按一下左鍵。之後回來這裡貼上並按Enter鍵。"
}
$InputPath = [System.IO.Path]::GetFullPath($InputPath.Trim().Trim([char]34))
if (-not (Test-Path -LiteralPath $InputPath)) { throw "File not found: $InputPath" }

if ([string]::IsNullOrWhiteSpace($Ranges)) {
    $Ranges = Read-NonEmpty "輸入想要的範圍之後按Enter鍵，像是：1-3。如果想要不連續的頁碼的話，用英文逗號「,」把他們分開，像是：1-3,5,8-10。"
}
$pageRanges = Parse-PageRanges $Ranges

if ([string]::IsNullOrWhiteSpace($Mode)) { $Mode = Read-Mode }
if ($Mode -eq "Separate") { $Mode = "S" }
if ($Mode -eq "Combined") { $Mode = "C" }
$Mode = $Mode.ToUpperInvariant()

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    $defaultOut = Join-Path (Split-Path -Parent $InputPath) (([System.IO.Path]::GetFileNameWithoutExtension($InputPath)) + "_extracted")
    Write-Host "成果會統一儲存在新建立的資料夾中。輸入想要儲存成果的位置，或是按Enter鍵直接放入原本相同的資料夾，在: $defaultOut" -ForegroundColor Green
    $o = $Host.UI.ReadLine()
    if ([string]::IsNullOrWhiteSpace($o)) { $OutputFolder = $defaultOut } else { $OutputFolder = $o.Trim().Trim([char]34) }
}
$OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$word = $null
$countDoc = $null
try {
    Write-Host "Starting Microsoft Word..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Write-Host "Opening source for page count: $InputPath" -ForegroundColor Cyan
    $countDoc = $word.Documents.Open($InputPath, $false, $true)
    $countDoc.Repaginate()
    $totalPages = Get-PageCount -Document $countDoc
    Write-Host "Total pages detected by Word: $totalPages" -ForegroundColor Cyan
    Write-Host "Tables detected: $($countDoc.Tables.Count); inline shapes detected: $($countDoc.InlineShapes.Count); floating shapes detected: $($countDoc.Shapes.Count)" -ForegroundColor Cyan
    Safe-CloseDoc $countDoc
    $countDoc = $null

    $baseName = Safe-FileName ([System.IO.Path]::GetFileNameWithoutExtension($InputPath))

    if ($Mode -eq "C") {
        $pagesToKeep = Normalize-PageRanges -RangesList $pageRanges -TotalPages $totalPages
        $label = Make-RangeLabel -RangesList $pageRanges -TotalPages $totalPages
        $outPath = Join-Path $OutputFolder ("{0}_{1}_combined.docx" -f $baseName, $label)
        Write-Host "Creating combined editable file containing original pages: $($pagesToKeep -join ", ")" -ForegroundColor Cyan
        Create-EditableExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages
    }
    else {
        foreach ($r in $pageRanges) {
            $singleRange = @($r)
            $pagesToKeep = Normalize-PageRanges -RangesList $singleRange -TotalPages $totalPages
            $label = Make-RangeLabel -RangesList $singleRange -TotalPages $totalPages
            $outPath = Join-Path $OutputFolder ("{0}_{1}.docx" -f $baseName, $label)
            Write-Host "Creating separate editable file containing original pages: $($pagesToKeep -join ", ")" -ForegroundColor Cyan
            Create-EditableExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages
        }
    }

    Write-Host ""
    Write-Host "Done." -ForegroundColor Cyan
    Write-Host "Output folder: $OutputFolder" -ForegroundColor Cyan
    if ($OpenOutputFolder) { Start-Process explorer.exe $OutputFolder }
}
catch {
    Write-Host ""
    Write-Host "Error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    if ($null -ne $countDoc) { Safe-CloseDoc $countDoc }
    if ($null -ne $word) {
        try { $word.Quit() | Out-Null } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
        $word = $null
    }
    try { [GC]::Collect() } catch { }
    try { [GC]::WaitForPendingFinalizers() } catch { }
}
