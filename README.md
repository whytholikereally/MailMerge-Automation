# MailMerge Automation Script
This PowerShell script automates the process of mail merging documents. 

# Features
Converts .xlsx files to .csv. <br>
Performs mail merge using specified templates. <br>
Archives output files into dated folders. <br>
Sends email notifications for the process status. <br>
Logs all actions for monitoring and troubleshooting. <br>
# Prerequisites
PowerShell <br>
Microsoft Word <br>
ImportExcel module <br>
# Parameters
The script uses the following parameters to define file paths and email settings:

InputDirectory: Path to the input directory (default: "C:\Path\To\Input"). <br>
OutputDirectory: Path to the output directory (default: "C:\Path\To\Output"). <br>
CompletedDirectory: Path to the completed directory (default: "C:\Path\To\Completed"). <br>
TempDirectory: Path to the temporary directory (default: "C:\Path\To\Temp"). <br>
TemplateDirectory: Path to the template directory (default: "C:\Path\To\Template"). <br>
LogFilePath: Path to the log file (default: "C:\Path\To\Log\merge_log_$(Get-Date -Format 'yyyy-MM-dd').txt"). <br>
SmtpServer: SMTP server for sending emails (default: 'smtp.example.com'). <br>
EmailFrom: Email sender address (default: 'Sender <sender@example.com>'). <br>
EmailTo: Email recipient address (default: 'Recipient <recipient@example.com>'). <br>
# Script Details
The script performs the following steps:

Initialization: Sets up the Word application and checks for required directories. <br>
Directory Management: Creates dated folders for output and completed files. <br>
File Processing: Converts .xlsx files to .csv and processes each .csv file. <br>
Mail Merge: Merges data into Word templates and handles document formatting. <br>
Archiving: Moves processed files to dated folders. <br>
Cleanup: Deletes temporary files and logs actions. <br>
Notifications: Sends email notifications about the status of the process. <br>
