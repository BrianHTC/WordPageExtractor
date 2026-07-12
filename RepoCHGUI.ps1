<#
.SYNOPSIS
  GUI Word page extractor with editable output and fixed Close button.

.DESCRIPTION
  Lets you select a Word file, choose page range and output mode, and select output folder without typing paths.

  C = Combined file.
  S = Separate files by comma-separated ranges.

  Output remains editable. The engine saves a full copy, deletes unselected pages, removes trailing empty table rows/paragraphs,
  and compresses the required final paragraph mark to reduce blank pages after table-based documents.

  Close button fix:
    - Close always stays enabled.
    - Close calls Close(), Dispose(), and ExitThread() so the GUI exits after extraction.
#>

param(
    [switch]$OpenOutputFolder
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

try {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
        $psExe = (Get-Process -Id $PID).Path
        Start-Process -FilePath $psExe -ArgumentList @("-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-File", $PSCommandPath) | Out-Null
        exit
    }
} catch { }

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
            throw "Invalid range format: $p. Use examples like 1-3,5,8-10."
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
        if ($e -gt $TotalPages) { $e = $TotalPages }
        for ($p = $s; $p -le $e; $p++) { [void]$pages.Add($p) }
    }
    return @($pages | Sort-Object -Unique | ForEach-Object { [int]$_ })
}

function Safe-FileName([string]$Name) {
    foreach ($ch in [System.IO.Path]::GetInvalidFileNameChars()) { $Name = $Name.Replace($ch, "_") }
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
    $keep = @{}
    foreach ($p in @($PagesToKeep)) { $keep[[int]$p] = $true }
    for ($p = $OriginalTotalPages; $p -ge 1; $p--) {
        if (-not $keep.ContainsKey($p)) { Delete-PageByNumber -WordApp $WordApp -Document $Document -PageNumber $p }
    }
}

function Trim-EndingEmptyParagraphs($Document) {
    try {
        for ($i = 0; $i -lt 30; $i++) {
            if ($Document.Paragraphs.Count -le 1) { break }
            $lastPara = $Document.Paragraphs.Item($Document.Paragraphs.Count)
            $txt = [string]$lastPara.Range.Text
            $normalized = $txt -replace "[\r\n\f\v\t ]", "" -replace [char]7, "" -replace [char]160, ""
            if ($normalized.Length -eq 0) { [void]$lastPara.Range.Delete() } else { break }
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
            if ($normalized.Length -eq 0) { [void]$lastRow.Delete() } else { break }
        }
    } catch { }
}

function Compress-FinalEmptyParagraph($Document) {
    try {
        if ($Document.Paragraphs.Count -lt 1) { return }
        $lastPara = $Document.Paragraphs.Item($Document.Paragraphs.Count)
        $txt = [string]$lastPara.Range.Text
        $normalized = $txt -replace "[\r\n\f\v\t ]", "" -replace [char]7, "" -replace [char]160, ""
        if ($normalized.Length -eq 0) {
            try { $lastPara.Range.Font.Size = 1 } catch { }
            try { $lastPara.Range.Font.Hidden = $true } catch { }
            try { $lastPara.Range.ParagraphFormat.SpaceBefore = 0 } catch { }
            try { $lastPara.Range.ParagraphFormat.SpaceAfter = 0 } catch { }
            try { $lastPara.Range.ParagraphFormat.LineSpacingRule = 4 } catch { }
            try { $lastPara.Range.ParagraphFormat.LineSpacing = 1 } catch { }
            try { $lastPara.Range.ParagraphFormat.PageBreakBefore = $false } catch { }
            try { $lastPara.Range.ParagraphFormat.KeepWithNext = $false } catch { }
            try { $lastPara.Range.ParagraphFormat.KeepTogether = $false } catch { }
        }
    } catch { }
}

function Delete-RangeFromPageToEnd($WordApp, $Document, [int]$StartPage) {
    $Document.Activate()
    $Document.Repaginate()
    $total = Get-PageCount -Document $Document
    if ($StartPage -gt $total) { return }
    try {
        [void]$WordApp.Selection.GoTo(1, 1, $StartPage)
        $startPos = [int]$WordApp.Selection.Range.Start
        $endPos = [int]$Document.Content.End
        if ($endPos -gt $startPos) { [void]$Document.Range($startPos, $endPos).Delete() }
    }
    catch {
        for ($p = $total; $p -ge $StartPage; $p--) {
            try { Delete-PageByNumber -WordApp $WordApp -Document $Document -PageNumber $p; $Document.Repaginate() } catch { break }
        }
    }
}

function Enforce-ExpectedPageCount($WordApp, $Document, [int]$ExpectedPages) {
    for ($pass = 0; $pass -lt 10; $pass++) {
        $current = Get-PageCount -Document $Document
        if ($current -le $ExpectedPages) { break }
        Delete-RangeFromPageToEnd -WordApp $WordApp -Document $Document -StartPage ($ExpectedPages + 1)
        Remove-TrailingEmptyTableRows -Document $Document
        Trim-EndingEmptyParagraphs -Document $Document
        Compress-FinalEmptyParagraph -Document $Document
        $Document.Repaginate()
    }
}

function Append-Log([System.Windows.Forms.RichTextBox]$LogBox, [string]$Message) {
    if ($null -ne $LogBox -and -not $LogBox.IsDisposed) {
        $start = $LogBox.TextLength
        $LogBox.AppendText($Message + [Environment]::NewLine)

        # Highlight the Done line in the GUI log.
        if ($Message -eq "擷取完畢。請按「關閉」按鈕。") {
            $LogBox.Select($start, $Message.Length)
            $LogBox.SelectionColor = [System.Drawing.Color]::DarkGreen
            $LogBox.SelectionBackColor = [System.Drawing.Color]::LightGreen
            $LogBox.SelectionFont = New-Object System.Drawing.Font($LogBox.Font, [System.Drawing.FontStyle]::Bold)
            $LogBox.Select($LogBox.TextLength, 0)
            $LogBox.SelectionColor = $LogBox.ForeColor
            $LogBox.SelectionBackColor = $LogBox.BackColor
            $LogBox.SelectionFont = $LogBox.Font
        }

        $LogBox.SelectionStart = $LogBox.Text.Length
        $LogBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Host $Message -ForegroundColor Cyan
}

function Create-EditableExtractedDoc($WordApp, [string]$InputPath, [string]$OutPath, [int[]]$PagesToKeep, [int]$OriginalTotalPages, [System.Windows.Forms.RichTextBox]$LogBox) {
    $workDoc = $null
    try {
        Append-Log $LogBox "Saving full copy first to preserve editable layout..."
        $workDoc = $WordApp.Documents.Open($InputPath, $false, $true)
        $workDoc.Repaginate()
        Save-WordDocx -Document $workDoc -Path $OutPath
        Safe-CloseDoc $workDoc
        $workDoc = $null

        Append-Log $LogBox "Deleting unselected pages from editable copy..."
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
            Compress-FinalEmptyParagraph -Document $workDoc
            Enforce-ExpectedPageCount -WordApp $WordApp -Document $workDoc -ExpectedPages ($PagesToKeep.Count)
            $workDoc.Repaginate()
            $workDoc.Save()
            $finalPages = Get-PageCount -Document $workDoc
        }
        Append-Log $LogBox "Created: $OutPath"
        Append-Log $LogBox "Final pages detected by Word: $finalPages; requested pages: $($PagesToKeep.Count)"
    }
    finally {
        if ($null -ne $workDoc) { Safe-CloseDoc $workDoc }
    }
}

function Run-ExtractionFromGui([string]$InputPathText, [string]$RangesText, [string]$OutputFolderText, [string]$ModeText, [System.Windows.Forms.RichTextBox]$LogBox) {
    $word = $null
    $countDoc = $null
    try {
        $InputPath = [System.IO.Path]::GetFullPath($InputPathText.Trim().Trim([char]34))
        if (-not (Test-Path -LiteralPath $InputPath)) { throw "File not found: $InputPath" }
        $pageRanges = Parse-PageRanges $RangesText
        $Mode = $ModeText.ToUpperInvariant()

        if ([string]::IsNullOrWhiteSpace($OutputFolderText)) {
            $OutputFolder = Join-Path (Split-Path -Parent $InputPath) (([System.IO.Path]::GetFileNameWithoutExtension($InputPath)) + "_extracted")
        } else {
            $OutputFolder = [System.IO.Path]::GetFullPath($OutputFolderText.Trim().Trim([char]34))
        }
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

        Append-Log $LogBox "Starting Microsoft Word..."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        Append-Log $LogBox "Opening source for page count: $InputPath"
        $countDoc = $word.Documents.Open($InputPath, $false, $true)
        $countDoc.Repaginate()
        $totalPages = Get-PageCount -Document $countDoc
        Append-Log $LogBox "Total pages detected by Word: $totalPages"
        Append-Log $LogBox "Tables detected: $($countDoc.Tables.Count); inline shapes detected: $($countDoc.InlineShapes.Count); floating shapes detected: $($countDoc.Shapes.Count)"
        Safe-CloseDoc $countDoc
        $countDoc = $null

        $baseName = Safe-FileName ([System.IO.Path]::GetFileNameWithoutExtension($InputPath))
        if ($Mode -eq "C") {
            $pagesToKeep = Normalize-PageRanges -RangesList $pageRanges -TotalPages $totalPages
            $label = Make-RangeLabel -RangesList $pageRanges -TotalPages $totalPages
            $outPath = Join-Path $OutputFolder ("{0}_{1}_combined.docx" -f $baseName, $label)
            Append-Log $LogBox "Creating combined editable file containing original pages: $($pagesToKeep -join ", ")"
            Create-EditableExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages -LogBox $LogBox
        } else {
            foreach ($r in $pageRanges) {
                $singleRange = @($r)
                $pagesToKeep = Normalize-PageRanges -RangesList $singleRange -TotalPages $totalPages
                $label = Make-RangeLabel -RangesList $singleRange -TotalPages $totalPages
                $outPath = Join-Path $OutputFolder ("{0}_{1}.docx" -f $baseName, $label)
                Append-Log $LogBox "Creating separate editable file containing original pages: $($pagesToKeep -join ", ")"
                Create-EditableExtractedDoc -WordApp $word -InputPath $InputPath -OutPath $outPath -PagesToKeep $pagesToKeep -OriginalTotalPages $totalPages -LogBox $LogBox
            }
        }
        Append-Log $LogBox "擷取完畢。請按「關閉」按鈕。"
        Append-Log $LogBox "Output folder: $OutputFolder"
        try { Start-Process explorer.exe $OutputFolder } catch { }
        # Do not show a modal MessageBox here. On some systems it can appear behind the form
        # and make the GUI look frozen. The log box already shows Done and Output folder.
    }
    catch {
        Append-Log $LogBox ("Error: " + $_.Exception.Message)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
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
}

function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Word Page Extractor - GUI Editable"
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(930, 560)
    $form.MinimumSize = New-Object System.Drawing.Size(930, 560)
    $form.Font = New-Object System.Drawing.Font("Microsoft JhengHei UI", 12)

    $form.Add_FormClosed({
        try { [System.Windows.Forms.Application]::ExitThread() } catch { }
    })

    $lblFile = New-Object System.Windows.Forms.Label
    $lblFile.Text = "Word檔:"
    $lblFile.Location = New-Object System.Drawing.Point(18, 25)
    $lblFile.AutoSize = $true
    $form.Controls.Add($lblFile)

    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = New-Object System.Drawing.Point(130, 22)
    $txtFile.Size = New-Object System.Drawing.Size(470, 28)
    $txtFile.Anchor = "Top,Left,Right"
    $form.Controls.Add($txtFile)

    $btnFile = New-Object System.Windows.Forms.Button
    $btnFile.Text = "瀏覽..."
    $btnFile.Location = New-Object System.Drawing.Point(615, 20)
    $btnFile.Size = New-Object System.Drawing.Size(105, 32)
    $btnFile.Anchor = "Top,Right"
    $form.Controls.Add($btnFile)

    $lblRange = New-Object System.Windows.Forms.Label
    $lblRange.Text = "頁數範圍:"
    $lblRange.Location = New-Object System.Drawing.Point(18, 75)
    $lblRange.AutoSize = $true
    $form.Controls.Add($lblRange)

    $txtRange = New-Object System.Windows.Forms.TextBox
    $txtRange.Location = New-Object System.Drawing.Point(130, 72)
    $txtRange.Size = New-Object System.Drawing.Size(300, 28)
    $txtRange.Text = "1-2"
    $form.Controls.Add($txtRange)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = "範例: 1-2 或 1-3,5,8-10，想要不連續的頁數的話用英文逗號分開"
    $lblHint.Location = New-Object System.Drawing.Point(445, 77)
    $lblHint.AutoSize = $true
    $form.Controls.Add($lblHint)

    $groupMode = New-Object System.Windows.Forms.GroupBox
    $groupMode.Text = "輸出模式"
    $groupMode.Location = New-Object System.Drawing.Point(18, 120)
    $groupMode.Size = New-Object System.Drawing.Size(705, 90)
    $groupMode.Anchor = "Top,Left,Right"
    $form.Controls.Add($groupMode)

    $radioC = New-Object System.Windows.Forms.RadioButton
    $radioC.Text = "C - Combined: 選擇的頁數與範圍將全部輸出至同一檔案"
    $radioC.Location = New-Object System.Drawing.Point(18, 26)
    $radioC.Size = New-Object System.Drawing.Size(650, 24)
    $radioC.Checked = $true
    $groupMode.Controls.Add($radioC)

    $radioS = New-Object System.Windows.Forms.RadioButton
    $radioS.Text = "S - Separate: 用英文逗號分開的頁數範圍將分別輸出至不同檔案"
    $radioS.Location = New-Object System.Drawing.Point(18, 56)
    $radioS.Size = New-Object System.Drawing.Size(650, 24)
    $groupMode.Controls.Add($radioS)

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = "選擇輸出資料夾"
    $lblOut.Location = New-Object System.Drawing.Point(18, 230)
    $lblOut.AutoSize = $true
    $form.Controls.Add($lblOut)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(150, 227)
    $txtOut.Size = New-Object System.Drawing.Size(470, 28)
    $txtOut.Anchor = "Top,Left,Right"
    $form.Controls.Add($txtOut)

    $btnOut = New-Object System.Windows.Forms.Button
    $btnOut.Text = "瀏覽..."
    $btnOut.Location = New-Object System.Drawing.Point(635, 225)
    $btnOut.Size = New-Object System.Drawing.Size(105, 32)
    $btnOut.Anchor = "Top,Right"
    $form.Controls.Add($btnOut)

    $lblOutHint = New-Object System.Windows.Forms.Label
    $lblOutHint.Text = "此格留白將自動輸出至原檔案所在地"
    $lblOutHint.Location = New-Object System.Drawing.Point(150, 262)
    $lblOutHint.AutoSize = $true
    $form.Controls.Add($lblOutHint)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "開始擷取"
    $btnRun.Location = New-Object System.Drawing.Point(130, 300)
    $btnRun.Size = New-Object System.Drawing.Size(160, 38)
    $form.Controls.Add($btnRun)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "關閉"
    $btnClose.Location = New-Object System.Drawing.Point(305, 300)
    $btnClose.Size = New-Object System.Drawing.Size(120, 38)
    $btnClose.Enabled = $true
    $form.Controls.Add($btnClose)

    $txtLog = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Location = New-Object System.Drawing.Point(18, 360)
    $txtLog.Size = New-Object System.Drawing.Size(705, 140)
    $txtLog.Anchor = "Top,Bottom,Left,Right"
    $txtLog.Multiline = $true
    $txtLog.ScrollBars = "Vertical"
    $txtLog.ReadOnly = $true
    $txtLog.BackColor = [System.Drawing.Color]::White
    $txtLog.HideSelection = $false
    $form.Controls.Add($txtLog)

    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = "Word documents (*.docx;*.doc)|*.docx;*.doc|All files (*.*)|*.*"
    $openDialog.Title = "Select Word document"

    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select output folder"

    $btnFile.Add_Click({
        if ($openDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtFile.Text = $openDialog.FileName
            if ([string]::IsNullOrWhiteSpace($txtOut.Text)) {
                $txtOut.Text = Join-Path (Split-Path -Parent $openDialog.FileName) (([System.IO.Path]::GetFileNameWithoutExtension($openDialog.FileName)) + "_extracted")
            }
        }
    })

    $btnOut.Add_Click({
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtOut.Text = $folderDialog.SelectedPath }
    })

    # Close fix: do not just call Close(); also Dispose and ExitThread so the modal GUI exits reliably after extraction.
    $btnClose.Add_Click({
        try { $form.DialogResult = [System.Windows.Forms.DialogResult]::OK } catch { }
        try { $form.Close() } catch { }
        try { $form.Dispose() } catch { }
        try { [System.Windows.Forms.Application]::ExitThread() } catch { }
    })

    $btnRun.Add_Click({
        $modeValue = if ($radioC.Checked) { "C" } else { "S" }
        $btnRun.Enabled = $false
        $btnFile.Enabled = $false
        $btnOut.Enabled = $false
        $btnClose.Enabled = $true
        $txtLog.Clear()
        try { Run-ExtractionFromGui -InputPathText $txtFile.Text -RangesText $txtRange.Text -OutputFolderText $txtOut.Text -ModeText $modeValue -LogBox $txtLog }
        finally {
            if (-not $form.IsDisposed) {
                $btnRun.Enabled = $true
                $btnFile.Enabled = $true
                $btnOut.Enabled = $true
                $btnClose.Enabled = $true
                try { $btnClose.Focus() } catch { }
                try { $form.Activate() } catch { }
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    })

    [void]$form.ShowDialog()
}

Show-Gui
