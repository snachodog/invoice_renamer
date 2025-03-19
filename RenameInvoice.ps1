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

    # Initialize variable to store Bill To name
    $billToName = $null

    # Find the line with 'Bill to' exactly and take the next non-empty line
    for ($i = 0; $i -lt $lines.Length; $i++) {
        if ($lines[$i].Trim() -eq "Bill to") {
            # Find the next non-empty line as the company name
            for ($j = $i + 1; $j -lt $lines.Length; $j++) {
                $nextLine = $lines[$j].Trim()
                if ($nextLine -ne "") {
                    $billToName = $nextLine
                    break
                }
            }
            break
        }
    }

    if ($billToName) {
        # Clean up billToName (remove invalid filename chars)
        $cleanName = ($billToName -replace '[\\/:*?"<>|]', '').Trim()

        # Generate new filename safely
        $newFileName = "$cleanName 2025 calendar ad invoice.pdf"

        # Ensure the filename isn't excessively long
        if ($newFileName.Length -gt 200) {
            $newFileName = $newFileName.Substring(0, 200) + ".pdf"
        }

        # Build full destination path explicitly
        $destPath = Join-Path -Path $pdfFile.DirectoryName -ChildPath $newFileName

        # Check if the file already exists
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

    # Clean up temporary file
    Remove-Item $tempTxt -Force
}
