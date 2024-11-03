# Function to search file contents
function Search-FileContent {
    param (
        [string]$filePath,
        [string]$searchTerm
    )
    
    # Check the file type and read its content accordingly
    switch ([System.IO.Path]::GetExtension($filePath).ToLower()) {
        ".docx" {
            try {
                $word = New-Object -ComObject Word.Application
                $doc = $word.Documents.Open($filePath, [ref]$false, [ref]$true)
                $content = $doc.Content.Text
                $doc.Close($false)
                $word.Quit()
                return $content -like "*$searchTerm*"
            } catch {
                Write-Host "Error reading Word file: $filePath"
                return $false
            }
        }
        ".xlsx" {
            try {
                $excel = New-Object -ComObject Excel.Application
                $workbook = $excel.Workbooks.Open($filePath, [ref]$false, [ref]$true)
                $found = $false
                
                foreach ($sheet in $workbook.Sheets) {
                    $usedRange = $sheet.UsedRange
                    for ($row = 1; $row -le $usedRange.Rows.Count; $row++) {
                        for ($col = 1; $col -le $usedRange.Columns.Count; $col++) {
                            $cellValue = $usedRange.Cells.Item($row, $col).Text
                            if ($cellValue -like "*$searchTerm*") {
                                $found = $true
                                break
                            }
                        }
                        if ($found) { break }
                    }
                }
                $workbook.Close($false) 
                $excel.Quit()
                return $found
            } catch {
                Write-Host "Error reading Excel file: $filePath"
                return $false
            }
        }
        ".pptx" {
            try {
                $powerPoint = New-Object -ComObject PowerPoint.Application
                $presentation = $powerPoint.Presentations.Open($filePath, [ref]$false, [ref]$true)
                $found = $false
                
                foreach ($slide in $presentation.Slides) {
                    foreach ($shape in $slide.Shapes) {
                        if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                            $text = $shape.TextFrame.TextRange.Text
                            if ($text -like "*$searchTerm*") {
                                $found = $true
                                break
                            }
                        }
                    }
                    if ($found) { break }
                }
                $presentation.Close()
                $powerPoint.Quit()
                return $found
            } catch {
                Write-Host "Error reading PowerPoint file: $filePath"
                return $false
            }
        }
        default {
            return $false
        }
    }
}

# Ask the user to select a file type by number
Write-Host "Select the file type to search for:"
Write-Host "1. DOCX"
Write-Host "2. XLSX"
Write-Host "3. PPTX"
$fileTypeChoice = Read-Host "Enter the number corresponding to the file type"

switch ($fileTypeChoice) {
    "1" { $fileType = ".docx" }
    "2" { $fileType = ".xlsx" }
    "3" { $fileType = ".pptx" }
    default {
        Write-Host "Invalid choice. Please enter 1, 2, or 3."
        exit
    }
}

# Ask for the keyword to search for within files
$searchTerm = Read-Host "Enter the keyword to search for within the files"

# Define the root directory for the search (entire C drive)
$rootDir = "C:\"

# Main search logic for files
Write-Host "Searching for $fileType files across the entire system..."

# Get all file paths matching search criteria
$files = Get-ChildItem -Path $rootDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { 
    $_.Extension -eq $fileType
}

# Check each file for content
$results = @()
foreach ($file in $files) {
    if (Search-FileContent -filePath $file.FullName -searchTerm $searchTerm) {
        $results += $file.FullName
        Write-Host "Found term '$searchTerm' in file: $($file.FullName)"
    }
}

# Display results
if ($results.Count -eq 0) {
    Write-Host "No files found containing the term '$searchTerm' in $fileType files."
} else {
    Write-Host "Files containing the term '$searchTerm':"
    $results | ForEach-Object { Write-Host $_ }
}

Write-Host "Search completed."
