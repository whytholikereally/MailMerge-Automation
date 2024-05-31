# Parameters assign a file path to a variable we can call later in the script.
param (
    [string]$InputDirectory = "C:\Path\To\Input", # We query for the dated folder, yyyymmddhhmmss, then redefine the $InputDirectory within the script.
    [string]$OutputDirectory = "C:\Path\To\Output", # Location of the MailMerge output. We create a dated folder here 'YYYYMMDD'. Once the script is complete we move all processed .docx and .txt files to the dated folder.
    [string]$CompletedDirectory = "C:\Path\To\Completed", # After the .csv is processed we move it to a dated folder we create within $CompletedDirectory to be archived. 
    [string]$TempDirectory = "C:\Path\To\Temp", # Location where we create all the .docx files before merging them into one .docx. We dump the converted .xlsx into $TempDirectory. After the script completes all contents of the $TempDirectory are deleted.
    [string]$TemplateDirectory = "C:\Path\To\Template", # Location where we store the document templates for the script to reference. The template names are hard coded into the script.
    [string]$LogFilePath = "C:\Path\To\Log\merge_log_$(Get-Date -Format 'yyyy-MM-dd').txt", # Location of the log file for the script.
    [string]$SmtpServer = 'smtp.example.com', # SMTP server for sending emails.
    [string]$EmailFrom = 'Sender <sender@example.com>', # Email sender address.
    [string]$EmailTo = 'Recipient <recipient@example.com>' # Email recipient address.
)

# Function to write logs
function Write-Log { 
    param(
        [string]$LogMessage # Defines the $LogMessage string to store the input value when the Write-Log function is called
    )
    Add-Content -Path $LogFilePath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $LogMessage" # Adds the content of $LogMessage to $LogFilePath with the date and time entered before the string 
}

# Function to get dated folder name
function Get-DatedFolderName { 
    return Get-Date -Format "yyyyMMdd" # Pulls the current date in YYYYMMDD format
}

# Function for dated folder output
function Archive-Output {
    param (
        [string]$OutputDirectory
    )

    # Get the current date for creating the dated folder
    $currentDate = Get-Date -Format "yyyyMMdd"
    $datedFolder = Join-Path -Path $OutputDirectory -ChildPath $currentDate # Stores the dated folder file path

    # Check if the dated folder already exists
    if (-not (Test-Path $datedFolder)) {
        # Create the dated folder if it doesn't exist
        New-Item -Path $datedFolder -ItemType Directory | Out-Null
    }

    # Get output docx files
    $outputFiles = Get-ChildItem -Path $OutputDirectory -Filter *.docx

    # Move output docx files to the dated output folder
    foreach ($file in $outputFiles) {
        Move-Item -Path $file.FullName -Destination $datedFolder -Force
        Write-Log "Moved $($file.Name) to $datedFolder" # Logs the file move
    }
}

# Function to remove blank first or last pages from DOCX files
function Remove-FirstAndLastBlankPages {
    param (
        [string]$filePath
    )
    $doc = $word.Documents.Open($filePath)
    $pagesCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)

    # Remove the first page if it is blank
    $firstPageRange = $doc.Range(0, $doc.Paragraphs[1].Range.End)
    if ($firstPageRange.Text -match '^\s*$') {
        $firstPageRange.Delete()
    }

    # Remove the last page if it is blank
    if ($pagesCount -gt 1) {
        $lastPageStart = $doc.GoTo([Microsoft.Office.Interop.Word.WdGoToItem]::wdGoToPage, [Microsoft.Office.Interop.Word.WdGoToDirection]::wdGoToAbsolute, $pagesCount).Start
        $lastPageRange = $doc.Range($lastPageStart, $doc.Content.End)
        if ($lastPageRange.Text -match '^\s*$') {
            $lastPageRange.Delete()
        }
    }

    $doc.Save()
    $doc.Close()
}

# Initialize Word application
$word = New-Object -ComObject Word.Application # Launches Word COM automation instance to function with the script

# Check if input directory exists
if (-not (Test-Path $InputDirectory)) { 
    Write-Log "Input directory does not exist: $InputDirectory" # If $InputDirectory is not found we log then exit the script
    exit 1
}

# Get the current date for creating the dated folder
$currentDate = Get-Date -Format "yyyyMMdd"
# Query if there's a folder in the input directory named with the current date in YYYYMMDD format
$currentDateFolder = Get-ChildItem -Path $InputDirectory | Where-Object { $_.Name -match "^$currentDate\d{6}$" } 
if ($currentDateFolder -ne $null) { # Check if $currentDateFolder is not null
    if (Test-Path $currentDateFolder.FullName -PathType Container) { 
        # Success: update the $InputDirectory to the dated folder
        $InputDirectory = $currentDateFolder.FullName
    } else {
        # Failure: Output DailyInputMissing.txt to the dated folder in $OutputDirectory
        $datedOutputFolder = Join-Path -Path $OutputDirectory -ChildPath $currentDate
        if (-not (Test-Path $datedOutputFolder)) { 
            New-Item -Path $datedOutputFolder -ItemType Directory | Out-Null 
        }
        $dailyInputMissingFilePath = Join-Path -Path $datedOutputFolder -ChildPath "DailyInputMissing.txt" 
        "Daily input folder missing for $($currentDate)" | Set-Content -Path $dailyInputMissingFilePath 
        Write-Log "Daily input folder missing for $(Get-DatedFolderName). Outputted DailyInputMissing.txt to $datedOutputFolder" 
    }
} else {
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "MailMerge Input Folder Missing" -Body "Dated folder does not exist in $InputDirectory" -SmtpServer $SmtpServer
    Write-Log "No folder found in $InputDirectory matching the naming convention for the current date."
    exit 1
}

# CSV filenames expected to find in $InputDirectory
$expectedCsvFiles = @(
    "File.csv",
    "File2.csv",
    "File3.csv",
    "File4.csv",
    "File5.csv",
    "File6.csv"
)

# Query the input directory for .xlsx files
$inputXlsxFiles = Get-ChildItem -Path $InputDirectory -Filter *.xlsx

# Process each .xlsx file found
foreach ($xlsxFile in $inputXlsxFiles) { # This loops over each Excel file in the $inputXlsxFiles
    # Construct the destination path for the corresponding .csv file
    $csvFileName = [System.IO.Path]::ChangeExtension($xlsxFile.Name, ".csv") # Builds file name with .csv
    $csvFilePath = Join-Path -Path $InputDirectory -ChildPath $csvFileName # Builds file path by combining $InputDirectory and the $csvFileName

    # Read the .xlsx file
    $excelData = Import-Excel -Path $xlsxFile.FullName

    # Convert the data to CSV format and save it
    $excelData | Export-Csv -Path $csvFilePath -NoTypeInformation

    # Wait until the .csv file is saved
    while (-not (Test-Path $csvFilePath)) {
        Start-Sleep -Milliseconds 100 # This was added to prevent the script from outpacing the conversion process during loops
    }

    # Log the conversion
    Write-Log "Converted $($xlsxFile.Name) to CSV: $csvFilePath"

    # Move the .xlsx file to the $TempDirectory
    Move-Item -Path $xlsxFile.FullName -Destination $TempDirectory -Force
}

# Check for missing CSV files
$missingCsvFiles = $expectedCsvFiles | Where-Object { -not (Test-Path (Join-Path -Path $InputDirectory -ChildPath $_)) }

# If any CSV file is missing, output the missing filenames to a text file
if ($missingCsvFiles) { 
    $currentDate = Get-Date -Format "yyyyMMdd" # Stores the current date in $currentDate
    $datedOutputFolder = Join-Path -Path $OutputDirectory -ChildPath $currentDate # Updates the $datedOutputFolder variable to $OutputDirectory adding the dated folder to the path with $currentDate.
    if (-not (Test-Path $datedOutputFolder)) { # Tests if the dated folder exists in the $OutputDirectory
        New-Item -Path $datedOutputFolder -ItemType Directory | Out-Null # Create the dated folder inside of the $OutputDirectory if it does not exist 
    }

    $missingFileNames = @() # Initialize an array to store missing file names

    $missingCsvFiles | ForEach-Object { # Loop for each csv in $missingCsvFiles
        $missingFileName = $_ -replace ".csv", "_NoCsvFound.txt" # Builds file name 
        $missingFilePath = Join-Path -Path $datedOutputFolder -ChildPath $missingFileName # Builds file path 
        $missingMessage = "Missing CSV file: $_" # Defines text added to the $missingFileName_NoCsvFound.txt
        Set-Content -Path $missingFilePath -Value $missingMessage # Adds text to $$missingFileName_NoCsvFound.txt
        Write-Log $missingMessage # Logs $missingMessage
        $missingFileNames += $_ # Add the missing file name to the array
    }

    # Email alert containing the list of missing file names
    $emailBody = "The following CSV files are missing:`r`n"
    $emailBody += ($missingFileNames -join "`r`n") # Join the missing file names with line breaks
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "MailMerge Missing Input CSV Files" -Body $emailBody -SmtpServer $SmtpServer
}

# Get CSV files in input directory
$inputFiles = Get-ChildItem -Path $InputDirectory -Filter *.csv

# Process each CSV file individually
foreach ($csvFile in $inputFiles) {
    # Initialize array to store merged file paths
    $mergedFiles = @()

    # Load the mail merge templates based on the input file name
    $templateNames = @{
        "File.csv" = @("template.docx")
        "File1.csv" = @("template1.docx")
        "File2.csv" = @("template2.docx")
        "File3.csv" = @("template3.docx")
        "File4.csv" = @("template4.docx", "template5.docx") # Outputs two .docx files. One with each template.
        "File5.csv" = @("template4.docx", "template5.docx") # Outputs two .docx files. One with each template.
    }[$csvFile.Name]

    if (-not $templateNames) {
        Write-Log "Unknown file: $($csvFile.Name)"
        continue
    }

    # Read data from the input CSV file
    $data = Import-Csv -Path $csvFile.FullName

    # Process each template separately
    foreach ($templateName in $templateNames) {
        # Clear the array storing merged file paths
        $mergedFiles = @()

        # Loads the template
        $mailMergeDoc = $word.Documents.Open("$TemplateDirectory\$templateName") 

        # Perform mail merge for each row of data
        # This is the heart of the MailMerge process mess with it and it will die... I warned you. Everything else is just data handling specific to the task you are completing.  
        foreach ($entry in $data) {
            # Replace mail merge fields with data from CSV for this row
            $mailMergeDoc.Content.Fields | ForEach-Object {
                $field = $_
                if ($field.Code.Text -match "MERGEFIELD (.+?) ") {
                    $fieldName = $Matches[1].Trim()
                    if ($entry.PSObject.Properties.Name -contains $fieldName) {
                        $field.Result.Text = $entry.$fieldName
                    }
                }
            }

            # Generate a sanitized filename for the merged document
            $sanitizedFileName = ($entry.INSURED -replace '[^\w\s]', '') + "_$($csvFile.BaseName)_$(Get-Date -Format 'yyyyMMdd').docx"

            # Save the merged document with the sanitized filename
            $outputFileName = "$TempDirectory\$sanitizedFileName"
            $mailMergeDoc.SaveAs([ref]$outputFileName)

            # Log the processed file
            Write-Log "Processed: $($entry.INSURED) using $templateName"

            # Add the path to the merged files array
            $mergedFiles += $outputFileName
        }

        # Close the mail merge document
        $mailMergeDoc.Close()

        # Get the current date to include in the final output file name
        $finalOutputFileName = "$OutputDirectory\$(Get-Date -Format 'yyyy-MM-dd')____$($csvFile.BaseName)____$($templateName)_.docx"

        # Merge the content of individual documents into a single document
        # Initialize output document
        $outputDoc = $word.Documents.Add()

        foreach ($file in $mergedFiles) {
            # Open each file
            $mergeFile = $word.Documents.Open($file)

            # Get content of the entire document
            $mergeContent = $mergeFile.Content

            # Copy text and paste into final document
            $outputRange = $outputDoc.Range()
            $outputRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $mergeContent.FormattedText.Copy()
            $outputRange.FormattedText.Paste()

            # Insert page break after pasting content for all files except specified templates
            if ($csvFile.Name -notin ("File1.csv", "File3.csv", "File2.csv")) {
                $outputRange.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
            }

            # Close the file after content is copied
            $mergeFile.Close()
        }

        # Apply formatting to the final output document
        $outputDoc.Content.Font.Name = "Calibri"
        $outputDoc.Content.Font.Size = 10
        $outputDoc.Content.ParagraphFormat.LineSpacing = 12
        $outputDoc.Content.ParagraphFormat.SpaceAfter = 0

        # Save the final merged document as DOCX
        $outputDoc.SaveAs([ref]$finalOutputFileName)

        # Close the final output document
        $outputDoc.Close()

        # Log the final output file
        Write-Log "Final output saved: $finalOutputFileName"
    }

    # Email alert sent after each csv processes. This email is sent after the processed docx files are moved to the dated folder in $OutputDirectory.
    #Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "MailMerge Completed" -Body "MailMerge completed Successfully $($finalOutputFileName)" -SmtpServer $SmtpServer
}

# Move the created missing CSV .txt files into the dated $datedOutputFolder
foreach ($missingFile in $missingCsvFiles) { # Loop
    $missingFileName = $missingFile -replace ".csv", "_NoCsvFound.txt" # Builds file name
    $missingFilePath = Join-Path -Path $datedOutputFolder -ChildPath $missingFileName # Builds file path
    if (Test-Path $missingFilePath) {
        Move-Item -Path $missingFilePath -Destination $datedOutputFolder -Force # Moves the $missingFile txt files to the $datedOutputFolder
        Write-Log "Moved missing CSV file text: $missingFilePath to $datedOutputFolder" # Logs the file move
    }
}

# Function to remove blank first or last pages from all DOCX files in a directory
function Clean-OutputDirectory {
    param (
        [string]$directory
    )
    $docxFiles = Get-ChildItem -Path $directory -Filter *.docx
    foreach ($file in $docxFiles) {
        Remove-FirstAndLastBlankPages -filePath $file.FullName
        Write-Log "Removed blank first or last pages from: $($file.Name)"
    }
}

# Clean the output directory
Clean-OutputDirectory -directory $OutputDirectory

# Close the Word application
$word.Quit()

# Release COM objects from memory
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null # May fail to release if script process is interrupted. Clear stuck Word objects from task manager before running the script. Failure to cleanup stuck objects may result in RPC failures.

# Moves output docx files into the dated output folder 
Archive-Output -OutputDirectory $OutputDirectory

# Email alert sent once the processed docx is moved to the dated output folder
Send-MailMessage -From $EmailFrom -To $EmailTo -Subject "MailMerge Completed" -Body "MailMerge completed Successfully $($OutputDirectory)" -SmtpServer $SmtpServer

# Archiving for processed csv files
# Move processed CSV files to the dated folder within the completed directory and rename them with the date appended
foreach ($csvFile in $inputFiles) {
    # Get the current date for creating the dated folder
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    $datedFolder = Join-Path -Path $CompletedDirectory -ChildPath $currentDate # Stores the dated folder file path

    # Check if the dated folder already exists
    if (-not (Test-Path $datedFolder)) {
        # Create the dated folder if it doesn't exist
        New-Item -Path $datedFolder -ItemType Directory | Out-Null
    }

    # Construct the file path for the archived CSV file
    $archivedCsvFilePath = Join-Path -Path $datedFolder -ChildPath $csvFile.Name

    # Move the processed CSV file to the dated folder
    Move-Item -Path $csvFile.FullName -Destination $archivedCsvFilePath -Force

    # Counter to append to the filename. Added to account for processing multiple csv files with the same name on the same day to prevent overwriting archived csv files.
    $counter = 1
    $newFileName = "{0}_{1}_{2}{3}" -f $csvFile.BaseName, $currentDate, $counter, $csvFile.Extension # Builds file name
    $newFilePath = Join-Path -Path $datedFolder -ChildPath $newFileName # Builds file path

    # Check if the new filename already exists, increment the counter until a unique filename is found
    while (Test-Path $newFilePath) {
        $counter++
        $newFileName = "{0}_{1}_{2}{3}" -f $csvFile.BaseName, $currentDate, $counter, $csvFile.Extension # Builds file name
        $newFilePath = Join-Path -Path $datedFolder -ChildPath $newFileName # Builds file path
    }

    # Rename the file with the updated name
    Rename-Item -Path $archivedCsvFilePath -NewName $newFileName -Force

    # Logs the processed csv was archived to $CompletedDirectory
    Write-Log "Moved $($csvFile.Name) to $datedFolder and renamed to $newFileName"
}

# Delete all contents of the temp directory
Remove-Item "$TempDirectory\*" -Force -Recurse

# Logs $TempDirectory contents deletion
Write-Log "Deleted all contents of $TempDirectory"

# Logs the completion of the script
Write-Log "Script execution completed"
