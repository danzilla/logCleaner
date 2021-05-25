
# Variables
$serverName = hostname
$folderName = $serverName + "_ISSLogFiles"
$logPath = $HOME + "\Documents\TestFolder"
$targetPath = "D:\Archive\" + $folderName  + '\'
$dayToMove = "-30"
$dayToDelete = "-60"

# SMTP user Credentials - This settings been configured for Gmail
# Please changes the settings accordingly 
# Known Error: 
    # Send-MailMessage : The SMTP server requires a secure connection 
    # or the client was not authenticated. The server response was: 5.7.0 Authentication Required
# Its require valid SMTP settings - Google Not allowing unsecure client to send emails
 
$User = "dan.alphonza@gmail.com"
$Password = 'HAHAHA!@#'
$From = "dan.alphonza@gmail.com"
$To = "MidTier@bc.ca"
# SMTP Configuration
$SMTPServer = "smtp.gmail.com"
$SMTPPort = "587"
# Secure Password
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential ($User, $SecurePassword)


# Get list of logs file to move and Delete
function ListItemsAndMoveAndDelete {
    # Check folder exits
    if(!(test-path $logPath)){
        # if not exits 
        $errorMessage = "Invalid Log Path - LogPath not exists " + $logPath
        Write-Error $errorMessage
    } else {

        # Check Target folder exits
        if(!(test-path $targetPath)){
            # if not exits | Create folder $targetPath
            $CreateTargetPath = New-item -ItemType Directory -Path $targetPath
        }

        # CSV file path and Name
        $csvMovedFiles = $targetPath + "\moved_" + $folderName + ".csv"
        $csvDeletedFiles = $targetPath + "\deleted_" + $folderName + ".csv"
        
        # Check CSV files exits | and If exist - Remove
        if((test-path $csvMovedFiles)){
            # if not exits | Create folder $targetPath
            Remove-Item $csvMovedFiles
            Remove-Item $csvDeletedFiles
        }
        
        # Create New csv file to create report | Move items
        $CreateCsvMovedFiles = New-Item $csvMovedFiles -ItemType File
        # Create New csv file to create report | Deleted items
        $CreateCsvDeletedFiles = New-Item $csvDeletedFiles -ItemType File

        # Move Items if they are than older 30 Days
        # Loop through Child | Filter by Day (-30) | Move
        # Filter log path with 30 days to move
        $getMoveItems =  Get-ChildItem -Path $logPath | 
                            Where { $_.LastWriteTime -lt (Get-Date).AddDays($dayToMove) }
        # Object store for Moved logs 
        $MovedObjects = @()
        ForEach($moveItem in $getMoveItems){
            # Move file
            Move-Item $moveItem.FullName $targetPath
            # Record file to csv
            $MovedObjects += [pscustomobject]@{
                Action = "Moved"
                Name = $moveItem.Name
                Date = $moveItem.LastWriteTime
                Directroy = $moveItem.FullName
            }
        }
        # Add object of data to csv to prepare for email attachedment
        $MovedObjects | Export-Csv $csvMovedFiles


        # Delete Items if they are older than 60 Days       
        # Loop through Child | Filter by Day (-60) | Delete from Archive
        # Filter log path with 60 days to Delete from Archive
        $getDeleteItems =  Get-ChildItem -Path $targetPath | 
                            Where { $_.LastWriteTime -lt (Get-Date).AddDays($dayToDelete) }
        # Object store for Delete logs 
        $DeleteObjects = @()
        ForEach($DeleteItem in $getDeleteItems){
            # Delete file
            Remove-Item $DeleteItem.FullName
            # Record file to csv
            $DeleteObjects += [pscustomobject]@{
                Action = "Deleted"
                Name = $DeleteItem.Name
                Date = $DeleteItem.LastWriteTime
                Directroy = $DeleteItem.FullName
            }
        }
        # Add object of data to csv to prepare for email attachedment
        $DeleteObjects | Export-Csv $csvDeletedFiles
      
        # Email Report - Contents
        $EmailSubject = "Report: Scheduled IIS logs cleanup - Server: " + $serverName + " " + (Get-Date)
        $EmailBody = "Please find the attached report for the " + $serverName + " logfile cleanup"
        $EmailAttachement = @($csvMovedFiles)
        $EmailAttachement += ($csvDeletedFiles)
            
        # Send Email
        Send-MailMessage -From $From -to $To -Subject $EmailSubject `
        -Body $EmailBody -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
        -Credential $Cred -Attachments $EmailAttachement

        # Finally console out the detatils 
        Write-Host "End of Cleaning!"
        Write-Host "Subject: " $EmailSubject
        Write-Host "Body: " $EmailBody
        Write-Host "Attachments: " ($EmailAttachement | Format-Table | Out-String)
   }
}

ListItemsAndMoveAndDelete