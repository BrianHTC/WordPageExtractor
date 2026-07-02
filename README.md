# Word Page Extractor

A Windows batch + PowerShell utility for extracting selected pages from Microsoft Word documents while trying to preserve the original look of the pages.

一個基於Windows batch 和 Powershell 指令，無須額外安裝軟體，擷取Word檔中指定頁的工具。此工具旨在盡量保持該頁面的格式及版面配置與原文件一致。

## Files

This project contains two main files that comes in two languages, English and Traditional Chinese:

```text
WordPageExtractor.bat
&
RepoEN.ps1 or RepoGrandmaCH.ps1

OR

點兩下這個.bat
&
RepoGrandmaCH.ps1
```

- `WordPageExtractor.bat` OR `點兩下這個.bat`
  The launcher file. Double-click this file to run the tool.

- `RepoEN.ps1` OR `RepoGrandmaCH.ps1`
  The PowerShell script that performs the Word page extraction.
  

Both files must stay in the same folder.

The Traditional Chinese Version is built to display instructions that are plain and clear for old people. 

## Requirements 需求

- Windows
- Microsoft Word desktop app installed
- PowerShell available on the computer
- Permission to run local PowerShell scripts
- CMD.exe

- Windows系統的電腦
- 安裝並成功啟用的Microsoft Word 應用程式
- 執行Powershell的權限
- CMD.exe

## How to Use

1. Download WordPageExtractor_EN.zip ,and unzip it. Make sure to put `WordPageExtractor.bat` and `RepoEN.ps1` in the same folder.
2. Double-click `WordPageExtractor.bat`.
3. When prompted for the Word file path:
   - Right-click the Word document you want to split.
   - Click **Copy as path** / **複製路徑**.
   - Paste it into the CMD window and press Enter.
4. Enter the page range you want to extract.
5. Choose the output mode.
6. The output folder will open after the task finishes.

## 如何使用

1. 下載WordPageExtractor_GrandmaCH.zip這個檔案並解壓縮。
2. 點兩下 點兩下這個.bat ，若有白色視窗跳出來就選擇「仍要執行」，之後照著黑色視窗跳出來的綠色字指示做。
3. 要分離的頁數可以選擇單頁或多頁，像是：1或是1-3，也可以選擇不連續的頁碼，但要用英文逗號「,」把他們分開，像是：1-3,5,8-10就會選擇第1到3頁、第5頁跟第8到10頁。
4. 選好範圍之後可以選擇需要的檔案數量，按C(大小寫都可以)會選擇將所選頁碼都擷取至同一個Word檔裡，按S(大小寫都可以)會選擇將所選範圍以英文逗號「,」為分界分別擷取至不同Word檔裡。
   - 假設選擇1-3,4：
     - 按C會輸出一個包含1234頁的Word檔
     - 按S會輸出一個包含123頁的Word檔和一個包含第4頁的Word檔
5. 執行完畢後會將成果放在指定位置的新建資料夾中，並自動打開資料夾。此時可以回到黑色視窗按Enter鍵關閉它。

## Page Range Format

You can enter page ranges like this:

```text
1-3
```

This extracts pages 1 to 3.

You can also enter non-continuous ranges separated by English commas:

```text
1-3,5,8-10
```

This extracts pages 1 to 3, page 5, and pages 8 to 10.

## Output Modes

When the script asks for the output mode:

```text
C = one combined file
S = separate files based on comma-separated ranges
```

Examples:

```text
1-3,5,8-10
```

If you choose `C`(lower case or upper case), the script creates one Word file containing all selected pages.

If you choose `S`(lower case or upper case), the script creates separate Word files for each comma-separated range:

```text
1-3
5
8-10
```

## Page Number Preservation

The tool attempts to preserve the original look after extraction.

For example, if you extract:

```text
1-3,8
```

The extracted page that originally came from page 8 should still display page number 8 instead of being renumbered as page 4, and will have the same format and border settings as the oroginal document.

This works best when the source document uses normal Word `PAGE` fields in headers or footers.

## Limitations

- If page numbers are manually typed text, Word cannot automatically renumber them.
- If the document uses complex fields such as `PAGE` of `NUMPAGES`, the `NUMPAGES` value may reflect the new extracted document rather than the original document.
- Very complex Word layouts may produce unexpected page breaks after extraction.
- The displayed CMD font size may not change if the file is opened through Windows Terminal instead of classic Windows Console Host.

## Final Prompt

At the end of execution, the batch file displays:

```text
Press any key to close this window...
```

## Troubleshooting

There will be red prompts if anything didn't work well. 

### The window closes immediately

Make sure the end of the `.bat` file contains:

```bat
powershell.exe -NoLogo -NoProfile -Command "Write-Host 'Press any key to close this window...' -ForegroundColor Cyan"
pause >nul
```

This keeps the window open until you press a key.

### The PowerShell script cannot be found

Make sure these two files are in the same folder:

```text
Run-WordPageExtractor.bat
RepoEN.ps1
```

Also make sure the `.bat` file contains the correct script name:

```bat
set "PS1FILE=%~dp0RepoEN.ps1"
```

### Chinese characters display incorrectly

Make sure the `.bat` file contains:

```bat
chcp 65001 >nul
```

Do not write it as:

```bat
chcp 65001 &gt;nul
```

`&gt;` is HTML-escaped text and will not work correctly in a real batch file.

### The font size does not change

CMD font size changes are not always reliable from inside a script. If Windows Terminal is your default terminal, it may ignore classic CMD font settings. In that case, change the font size manually in Windows Terminal settings or use classic Windows Console Host.

## Suggested Folder Structure

```text
WordPageExtractor/
├─ Run-WordPageExtractor.bat
├─ RepoEN.ps1
└─ README.md
```

## Notes

This tool uses Microsoft Word automation through PowerShell. Do not close Microsoft Word while the script is running.
Please make sure you have your Microsoft Word activated. 

## Disclaimer and Licensing

This tool is created by me under the assitance of Microsoft Copilot AI. Feel free to create your own version of your tool base on this project, as long as you cite me properly. 
I have no intention of maintaining this project. Free free to release your improved version. 
Please report an issue if this tool contains possible security or privacy threat, and feel free to post your solution. 