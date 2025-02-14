Write-Host "Starting script execution..."

# Source folder (in Downloads, with a subfolder named BOL-LABEL)
$sourceFolder = "$env:USERPROFILE\Downloads\BOL-LABEL"
if (-not (Test-Path $sourceFolder)) {
    Write-Error "Source folder does not exist: $sourceFolder"
    exit
}
Write-Host "Source folder: $sourceFolder"

# Processed folder (a subfolder within the source folder)
$processedFolder = Join-Path -Path $sourceFolder -ChildPath "Processed"
Write-Host "Processed folder: $processedFolder"

# Path to SumatraPDF executable (update if installed elsewhere)
$pdfViewer = "$env:USERPROFILE\AppData\Local\SumatraPDF\SumatraPDF.exe"
Write-Host "SumatraPDF Path: $pdfViewer"

# Printer names (replace with actual printer names)
$Label_Printer = "Your_Label_Printer"  # For Shipping Labels
$BOL_Printer   = "Your_BOL_Printer"    # For BOL documents

# Ensure the Processed folder exists
if (!(Test-Path $processedFolder)) {
    try {
        New-Item -ItemType Directory -Path $processedFolder -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Error "Error creating Processed folder: $_"
        exit
    }
}

# List PDF files in the source folder
try {
    $files = Get-ChildItem -Path (Resolve-Path $sourceFolder) -File -Filter *.pdf
    Write-Host "Found $($files.Count) PDF files."
}
catch {
    Write-Error "Error listing files in source folder: $_"
    exit
}

# Process each file
foreach ($file in $files) {
    Write-Host "Processing file: $($file.Name)"
    $filePath = $file.FullName
    $fileName = $file.Name
    $orderNumber = "UNKNOWN"
    
    if ($fileName -match "^\d+") {
        $orderNumber = $Matches[0]
    }

    $printerName = ""
    $printCopies = 0

    if ($fileName -like "*BOLAction*") {
        $newFileName = "BOL_$orderNumber.pdf"
        $printerName = $BOL_Printer
        $printCopies = 2
    }
    elseif ($fileName -like "*Shipping Label*") {
        $newFileName = "Label_$orderNumber.pdf"
        $printerName = $Label_Printer
        $printCopies = 1
    }
    else {
        $newFileName = $fileName
    }

    $newPath = Join-Path -Path $processedFolder -ChildPath $newFileName
    $timestamp = Get-Date -Format "yyyyMMddHHmmss"
    if (Test-Path $newPath) {
        $newFileName = [System.IO.Path]::GetFileNameWithoutExtension($newFileName) + "_$timestamp" + [System.IO.Path]::GetExtension($newFileName)
        $newPath = Join-Path -Path $processedFolder -ChildPath $newFileName
    }

    try {
        Move-Item -Path $filePath -Destination $newPath -ErrorAction Stop
    }
    catch {
        Write-Error "Error moving file '$fileName': $_"
        continue
    }

    if ($printerName -ne "") {
        for ($i = 1; $i -le $printCopies; $i++) {
            try {
                $arguments = "-print-to `"$printerName`" -silent `"$newPath`""
                Start-Process -FilePath $pdfViewer -ArgumentList $arguments -NoNewWindow -Wait
            }
            catch {
                Write-Error "Printing failed for '$newFileName' on printer '$printerName': $_"
            }
        }
    }
}

Write-Host "Script execution completed."
