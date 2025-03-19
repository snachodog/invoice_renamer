# Ensure 'pdftotext.exe' is available in your PATH or specify the full path here
$pdftotextPath = "pdftotext.exe"

# Get all PDF files in the script's current directory
$pdfFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.pdf

foreach ($pdfFile in $pdfFiles) {
    # Temporary text file for extracting content
    $tempTxt = "$env:TEMP\$(New-Guid).txt"

    # Extract text content from PDF
    & $pdftotextPath -layout $pdfFile.FullName $tempTxt

    # Read extracted content line-by-line
    $lines = Get-Content $tempTxt

    # Initialize Bill To name
    $billToName = $null

    # More robust method: look for "Bill to" line and grab the next meaningful line
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Trim() -match '^Bill\s+to$') {
            for ($j = $i + 1; $j -lt $lines.Length; $j++) {
                $nextLine = $lines[$j].Trim()
                if (![string]::IsNullOrWhiteSpace($nextLine)) {
                    $billToName = $nextLine
                    break
                }
            }
            break
        }
    }

    if ($billToName) {
        # Clean name from invalid filename characters
        $cleanName = ($billToName -replace '[\\/:*?"<>|]', '').Trim()

        # Construct safe filename
        $newFileName = "$cleanName 2025 calendar ad invoice.pdf"
        $destPath = Join-Path -Path $pdfFile.DirectoryName -ChildPath $newFileName

        # Avoid conflicts
        if (-not (Test-Path $destPath)) {
            try {
                Rename-Item -Path $pdfFile.FullName -NewName $newFileName -Verbose -ErrorAction Stop
            } catch {
                Write-Warning "Error renaming '$($pdfFile.Name)': $_"
            }
        } else {
            Write-Warning "File '$newFileName' already exists. Skipping rename."
        }
    } else {
        Write-Warning "Could not find 'Bill to' in file '$($pdfFile.Name)'. Skipping file."
    }

    # Cleanup temporary file
    Remove-Item $tempTxt -Force
}
