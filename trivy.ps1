(Invoke-WebRequest -Uri "https://raw.githubusercontent.com/navneettoppo/cloudthat-miscellaneous/refs/heads/main/trivy").Content

# Ask user for the directory to scan
$scanDirectory = Read-Host "Please enter the directory to scan (default: current directory)"
if ([string]::IsNullOrEmpty($scanDirectory)) {
    $scanDirectory = Get-Location  # Default to current directory if none is provided
}

Set-Location -Path $scanDirectory

# Check if Chocolatey is installed
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-Output "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
} else {
    Write-Output "Chocolatey is already installed."
}

# Check if Trivy is installed
if (-not (Get-Command trivy -ErrorAction SilentlyContinue)) {
    Write-Output "Installing Trivy..."
    choco install trivy -y
} else {
    Write-Output "Trivy is already installed."
}

# Get the current directory path and name
$directoryPath = Get-Location
$directoryName = Split-Path -Leaf $directoryPath

# Ask user for the report name
$reportName = Read-Host "Enter the name for the SARIF report file (default: $directoryName-report.sarif)"
if ([string]::IsNullOrEmpty($reportName)) {
    $reportName = "$directoryName-report.sarif"  # Default name
}

# Run the trivy scan on the directory
trivy fs $directoryPath --dependency-tree --detection-priority comprehensive --include-dev-deps
trivy fs $directoryPath --dependency-tree --detection-priority comprehensive --include-dev-deps --format sarif -o $reportName

# Function to convert SARIF to Excel
function Convert-SarifToExcel {
    param (
        [string]$sarifFile,
        [string]$excelFile
    )

    try {
        # Load SARIF file
        $sarifContent = Get-Content -Path $sarifFile -Raw | ConvertFrom-Json

        # Create a new Excel file
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Add headers
        $headers = @("RuleId", "Message", "Level", "File", "StartLine", "EndLine")
        for ($i = 0; $i -lt $headers.Length; $i++) {
            $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
            $worksheet.Cells.Item(1, $i + 1).VerticalAlignment = -4160  # xlVAlignTop
        }

        # Set the column width for the message column (column 2)
        $worksheet.Columns.Item(2).ColumnWidth = 50

        # Add SARIF data to Excel
        $row = 2
        foreach ($run in $sarifContent.runs) {
            foreach ($result in $run.results) {
                $worksheet.Cells.Item($row, 1) = $result.ruleId
                $worksheet.Cells.Item($row, 2) = $result.message.text
                $worksheet.Cells.Item($row, 3) = $result.level
                $worksheet.Cells.Item($row, 4) = $result.locations.physicalLocation.artifactLocation.uri
                $worksheet.Cells.Item($row, 5) = $result.locations.physicalLocation.region.startLine
                $worksheet.Cells.Item($row, 6) = $result.locations.physicalLocation.region.endLine

                # Set vertical alignment to top for data cells
                for ($col = 1; $col -le 6; $col++) {
                    $worksheet.Cells.Item($row, $col).VerticalAlignment = -4160  # xlVAlignTop
                }

                $row++
            }
        }

        # Adjust column widths
        for ($col = 1; $col -le $headers.Length; $col++) {
            $maxLength = 0
            for ($row = 1; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
                $cellValue = $worksheet.Cells.Item($row, $col).Text
                if ($cellValue.Length -gt $maxLength) {
                    $maxLength = $cellValue.Length
                }
            }
            try {
                $worksheet.Columns.Item($col).ColumnWidth = [math]::Min($maxLength + 2, 255)  # Adding a little extra space, max width 255
            } catch {
                Write-Warning "Unable to set the ColumnWidth for column $col"
            }
        }

        # Save the Excel file
        $excelFilePath = Join-Path -Path (Split-Path -Path $sarifFile -Parent) -ChildPath $excelFile
        $workbook.SaveAs($excelFilePath)
        $workbook.Close($false)
        $excel.Quit()

        Write-Output "SARIF report converted to Excel format: $excelFilePath"
    } catch {
        Write-Error "An error occurred: $_"
    }
}



# Ask user for the Excel file name
$excelFileName = Read-Host "Enter the name for the Excel file (default: $directoryName-report.xlsx)"
if ([string]::IsNullOrEmpty($excelFileName)) {
    $excelFileName = "$directoryName-report.xlsx"  # Default name
}

# Convert the SARIF report to Excel
Convert-SarifToExcel -sarifFile $reportName -excelFile $excelFileName
