# Function to analyze prefetch files
function Get-PrefetchFiles {
    param (
        [string]$PrefetchFolderPath,
        [switch]$ExportCsv
    )

    # Suppress error messages from being displayed in the terminal
    $null = $ErrorActionPreference = 'SilentlyContinue'

    # Get all prefetch files in the specified folder
    $prefetchFiles = Get-ChildItem -Path $PrefetchFolderPath -Filter *.pf

    # Initialize an array to store the analysis results
    $results = @()

    # Iterate through each prefetch file
    foreach ($file in $prefetchFiles) {
        # Create a custom object with the analysis results
        $result = [PSCustomObject]@{
            "File" = $file.Name
            "LastExecutionTime" = $file.LastWriteTime
        }

        # Add the result to the array
        $results += $result
    }

    # Sort the results by LastExecutionTime
    $sortedResults = $results | Sort-Object -Property LastExecutionTime

    # Display the sorted results
    $sortedResults | Format-Table -AutoSize

    # Export the results to CSV if the ExportCsv switch is specified
    if ($ExportCsv) {
        $csvPath = Join-Path -Path $PrefetchFolderPath -ChildPath "PrefetchAnalysis.csv"
        $sortedResults | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Analysis results exported to CSV: $csvPath"
    }
}

# Analyze prefetch files in the default prefetch folder
$prefetchFolderPath = "C:\Windows\Prefetch"
Get-PrefetchFiles -PrefetchFolderPath $prefetchFolderPath -ExportCsv
