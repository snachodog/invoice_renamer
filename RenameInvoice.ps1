# Ensure 'pdftotext.exe' is available in your PATH or specify full path here
$pdftotextPath = "pdftotext.exe"

# Get all PDF files in the script's current directory
$pdfFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.pdf

foreach ($pdfFile in $pdfFiles) {
    # Temporary text file for extracting content
    $tempTxt = "$env:TEMP\$(New-Guid).txt"

    # Extract text content from PDF
    & $pdftotextPath -layout $pdfFile.FullName $tempTxt

    # Read extracted content
    $content = Get-Content $tempTxt

    # Initialize variables
    $billToName = $null
    $foundBillTo = $false

    # Search line-by-line for the correct 'Bill to' line
    for ($i = 0; $i -lt $content.Length; $i++) {
        if ($content[$i] -match "Bill to\s*$") {
            # Find the next non-empty line after "Bill to"
            for ($j = $i + 1; $j -lt $content.Length; $j++) {
                if ($content[$j].Trim()) {
                    $billToName = $content[$j].Trim()
                    break
                }
            }
            if ($billToName) { break }
        }
    }

    if ($billToName) {
        # Generate new filename
        $newFileName = "$billToName 2025 calendar ad invoice.pdf"

        # Check if the target file already exists
        if (-not (Test-Path "$($pdfFile.DirectoryName)\$newFileName")) {
            # Rename file
            Rename-Item -Path $pdfFile.FullName -NewName $newFileName -Verbose
        }
        else {
            Write-Warning "File '$newFileName' already exists. Skipping rename."
        }
    }
    else {
        Write-Warning "Could not find 'Bill to' in file '$($pdfFile.Name)'. Skipping file."
    }

    # Clean up temporary file
    Remove-Item $tempTxt -Force
}
