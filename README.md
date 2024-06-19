# MailMerge Automation Script
This PowerShell script automates the process of mail merging documents. It handles directory management, file processing, merging documents, archiving, and logging activities. It also sends email notifications about the status of the process.

# Features
Converts .xlsx files to .csv.
Performs mail merge using specified templates.
Archives output files into dated folders.
Sends email notifications for the process status.
Logs all actions for monitoring and troubleshooting.
# Prerequisites
PowerShell
Microsoft Word
ImportExcel module
# Parameters
The script uses the following parameters to define file paths and email settings:

$InputDirectory: Path to the input directory (default: "C:\Path\To\Input"). /n
$OutputDirectory: Path to the output directory (default: "C:\Path\To\Output").
$CompletedDirectory: Path to the completed directory (default: "C:\Path\To\Completed").
$TempDirectory: Path to the temporary directory (default: "C:\Path\To\Temp").
$TemplateDirectory: Path to the template directory (default: "C:\Path\To\Template").
$LogFilePath: Path to the log file (default: "C:\Path\To\Log\merge_log_$(Get-Date -Format 'yyyy-MM-dd').txt").
$SmtpServer: SMTP server for sending emails (default: 'smtp.example.com').
$EmailFrom: Email sender address (default: 'Sender <sender@example.com>').
$EmailTo: Email recipient address (default: 'Recipient <recipient@example.com>').
# Script Details
The script performs the following steps:

Initialization: Sets up the Word application and checks for required directories.
Directory Management: Creates dated folders for output and completed files.
File Processing: Converts .xlsx files to .csv and processes each .csv file.
Mail Merge: Merges data into Word templates and handles document formatting.
Archiving: Moves processed files to dated folders.
Cleanup: Deletes temporary files and logs actions.
Notifications: Sends email notifications about the status of the process.
