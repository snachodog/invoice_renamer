# Ensure 'pdftotext.exe' is available in your PATH or specify the full path here
$pdftotextPath = "pdftotext.exe"

# Get all PDF files in the script's current directory
$pdfFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.pdf

foreach ($pdfFile in $pdfFiles) {
    # Temporary text file for extracting content
    $tempTxt = "$env:TEMP\$(New-Guid).txt"

    # Extract text content from PDF
    & $pdftotextPath -layout $pdfFile.FullName $tempTxt

    # Read extracted content as a single string
    $content = Get-Content $tempTxt -Raw

    # Robust regex method: finds "Bill to" and captures the next meaningful line even if Bill to is not alone on line
    if ($content -match "Bill to.*?(?:\r?\n|\r|\n)\s*(.+)") {
        $billToName = $matches[1].Trim()

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
