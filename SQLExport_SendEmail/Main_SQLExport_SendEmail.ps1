# -------------------------------------------------------------------------------
# Script:      Main_SQLExport_SendEmail.ps1 
# Author:      Nishad Bhandarkar
# Date:        22.09.2015
# Keywords:    job, batch, report, email, daily
# Description: This Job script executes the query and sends a mail to the DL
#              IMPORTANT: 
#                 * The script runs the SQL specified in the variables
# Updates: 
#  1. <Date> - <Changed by> - Change Description
# -------------------------------------------------------------------------------

# Loading other functions
. .\cred.ps1
. .\sendEmail.ps1

# ---------------------- VARIABLES ----------------------- #

#optional transcript path below to debug any issues if running through task scheduler
#start-transcript -path D:\Scripts\Main_SQL_Execution_Log\scrptlog.txt

$script_dir = $MyInvocation.MyCommand.Path | Split-Path
$date = get-date
$dateCET = (Get-Date).AddDays(0).AddHours(-3).AddMinutes(-30).ToString("F")
#$ts = New-TimeSpan -Days 0 -Hours 3 -Minutes 30
#$dateCET = $dateCET = $date - $ts



$date_fileformat = Get-Date -format "ddMMyyyy"
$date_fileformat2 = Get-Date -format "dd-MMM-yyyy"
$date_fileformat3 = Get-Date -format "dd-MMM-yyyy HH-mm-ss"

$dist_list = "$script_dir\distribution_list.lst"
$dist_list_cc = "$script_dir\Cc_distribution_list.lst"
$dist_list_error = "$script_dir\distribution_list_error.lst"
$global:log_file = "$script_dir\<FilenameGoesHere>"+ $date_fileformat3 + ".log"
$global:zip_file = "$script_dir\<FilenameGoesHere>" + $date_fileformat + ".zip"

$sql_file1 = "$script_dir\1.sql"
#$sql_file2 = "$script_dir\2.sql"
#$sql_file3 = "$script_dir\3.sql"
#$sql_file4 = "$script_dir\4.sql"

$global:spool_file_name1 = "$script_dir\<FilenameGoesHere>_" + $date_fileformat + ".xls"
#$global:spool_file_name2 = "$script_dir\<Filename2>" + $date_fileformat + ".xls"
#$global:spool_file_name3 = "$script_dir\<Filename3>" + $date_fileformat + ".xls"
#$global:spool_file_name4 = "$script_dir\<Filename4>" + $date_fileformat + ".xls"


$DB_uname = "<DB Username>"
$DB_SID = "<DB SID Value>"

# -------------------------------------------------------- #


# ------------------------------------------------------------- #
# INITIALIZATION
# ------------------------------------------------------------- #

# Writing Header to the Log file
write_log_HF "startup"

log "`n## Deleting all csv and zip files in current folder..." INFO

# Delete any Zip files from the folder before proceeding with new script
# Doing it at the beggining of the script permits sendding the previous execution result
$remove_pattern = $script_dir + "\*.zip"
Remove-Item $remove_pattern

$remove_pattern = $script_dir + "\*.xls"
Remove-Item $remove_pattern

#$remove_pattern = $script_dir + "\*.csv"
#Remove-Item $remove_pattern

$remove_pattern = $script_dir + "\*.txt"
Remove-Item $remove_pattern


# ------------------------------------------------------------- #
# STEP 0 - Check if there is a valid secured password avaialble
# ------------------------------------------------------------- #

    log "`n## 0. Starting password validation..." INFO

    $Credentials = Get_Cred $DB_uname

    if (!$Credentials) {
       log "User credentials not found in PWD scirpt. Calling User setup" WARN 
       
       $result = Reset_Pwd $DB_uname
       
       if ($result ) {
        log "Password reset succesfully" INFO
        log "Please rerun the Job to load the password en memory" WARN
        Write-Host "Please rerun the Script/Job to load the password en memory"
        exit
       } else {
        log "There happened an error setting the password the for user $DB_uname. Please check the Password script log file" ERROR
       }
    } else {
        log "Credentials for user $DB_uname loaded to memory" INFO 
    }


# ------------------------------------------------------------- #
# STEP 1 - Execute the Query
# ------------------------------------------------------------- #

    log "`n## 1. Executing Query 1.sql..." INFO

    $Credentials = Get_Cred $DB_uname
    
#First Query

    $FileExists = (Test-Path $sql_file1 -PathType Leaf)
   
    If (($FileExists)) {
         
         cd $script_dir
         
         Write-Host "Running SQL. May take some seconds..."  
         $output = Invoke-Expression "sqlplus -L -S $($Credentials.GetNetworkCredential().UserName + '/' + $Credentials.GetNetworkCredential().Password + '@' + $DB_SID + ' `@' + $sql_file1) 2>&1"
         
         Write-Host "$output"
         log "Running SQL script: $sql_file1" INFO  
        
         if ($output | select-string -pattern "error") {
                log "ERROR executing SQL command. Please check error: $output" ERROR
                $emp_data = @()
    $file_entries_error = gc $dist_list_error
    $to = $file_entries_error
    

    # Mail to be sent

    $subject = "[Error] My subject line: " + $date_fileformat  
    $body = "Hello Support Team, 
    
    There is an error with the SQL execution. Please check the Main_SQLExport_SendEmail script run.
	
    Regards,
    Sender"
    
 #   $att = $zip_file

 log "`n## 3. Error Mail to be sent is $body..." INFO
 $att = $spool_file_name1
 
    $o = New-Object -com Outlook.Application
    $mail = $o.CreateItem(0)
log "New outlook object created" INFO

#2 = high importance email header
#$mail.importance = 2

$mail.subject = $subject

$mail.body = $body

#$mail.cc = <Email Address to be kept in CC goes here>

$mail.Attachments.Add($att)
#$mail.Attachments.Add($att2)
#$mail.Attachments.Add($att3)
#$mail.Attachments.Add($att4)
log "Attachments added" INFO
#for multiple email, use semi-colon ; to separate
#example: $mail.cc = “User1@test.com; User2@test.com“

foreach ( $email_address in $to ) { $mail.Recipients.Add("$email_address") }
#foreach ( $email_address in $cc ) { $mail.cc.Add("$email_address") }
log "Sleep for 1 second before sending the email" INFO
Start-sleep -Seconds 2
$mail.Send()
log "Mail sent" INFO


#To delete object pointing to Outlook cause sometimes archive folder cannot be used while the script is running
Start-sleep -Seconds 1
#$o.Quit()


[System.Runtime.Interopservices.Marshal]::ReleaseComObject($o)
$o = $null
log "Outlook object deleted" INFO

    #sendMail $from $to $cc $subject $body $att
Start-sleep -Seconds 1
    log "Email sent to distribution list for error" INFO
    
Start-sleep -Seconds 1
                exit
         } else {
                log "SQL executed succesfully" INFO
         }
                
    } else {
        log "ERROR. SQL file not found on the script path: $sql_file1" ERROR
    }


    <# ------------------------------------------------------------- #
# STEP 2 - Execute the Query 2
# ------------------------------------------------------------- #

    log "`n## 2. Executing Query 2.sql for extracting required data..." INFO

    $Credentials = Get_Cred $DB_uname
    
#Second Query

    $FileExists = (Test-Path $sql_file2 -PathType Leaf)
   
    If (($FileExists)) {
         
         cd $script_dir
         
         Write-Host "Running SQL. May take some seconds..."  
         $output = Invoke-Expression "sqlplus -L -S $($Credentials.GetNetworkCredential().UserName + '/' + $Credentials.GetNetworkCredential().Password + '@' + $DB_SID + ' `@' + $sql_file2) 2>&1"
         
         Write-Host "$output"
         log "Running SQL script: $sql_file2" INFO  
        
         if ($output | select-string -pattern "error") {
                log "ERROR executing SQL command. Please check error: $output" ERROR
                exit
         } else {
                log "Update SQL statement executed succesfully. Serial numbers have been updated." INFO
         }
                
    } else {
        log "ERROR. SQL file not found on the script path: $sql_file2" ERROR
    }
Start-sleep -milliseconds 1000

#log "`n## 3. Converting 1st spooled data file to a proper csv..." INFO

#import-csv $spool_file_name1 -Header "Serial Numbers" | export-csv $spool_file_name4 -NoTypeInformation


#log "`n## 4. Converting 2nd spooled data file to a proper csv and adding headers..." INFO

#import-csv $spool_file_name3 -Header "Comments", "Queued Timestamp ", "Reference ID" | export-csv $spool_file_name2 -NoTypeInformation
#Start-sleep -milliseconds 500

# ------------------------------------------------------------- #
# STEP 3 - Execute the Query 3
# ------------------------------------------------------------- #

    log "`n## 1. Executing Query 3.sql..." INFO

    $Credentials = Get_Cred $DB_uname
    
#First Query

    $FileExists = (Test-Path $sql_file3 -PathType Leaf)
   
    If (($FileExists)) {
         
         cd $script_dir
         
         Write-Host "Running SQL. May take some seconds..."  
         $output = Invoke-Expression "sqlplus -L -S $($Credentials.GetNetworkCredential().UserName + '/' + $Credentials.GetNetworkCredential().Password + '@' + $DB_SID + ' `@' + $sql_file3) 2>&1"
         
         Write-Host "$output"
         log "Running SQL script: $sql_file3" INFO  
        
         if ($output | select-string -pattern "error") {
                log "ERROR executing SQL command. Please check error: $output" ERROR
                exit
         } else {
                log "SQL executed succesfully" INFO
         }
                
    } else {
        log "ERROR. SQL file not found on the script path: $sql_file3" ERROR
    }


    # ------------------------------------------------------------- #
# STEP 4 - Execute the Query 4
# ------------------------------------------------------------- #

    log "`n## 2. Executing Query 4.sql for updating serial number records..." INFO

    $Credentials = Get_Cred $DB_uname
    
#Second Query

    $FileExists = (Test-Path $sql_file4 -PathType Leaf)
   
    If (($FileExists)) {
         
         cd $script_dir
         
         Write-Host "Running SQL. May take some seconds..."  
         $output = Invoke-Expression "sqlplus -L -S $($Credentials.GetNetworkCredential().UserName + '/' + $Credentials.GetNetworkCredential().Password + '@' + $DB_SID + ' `@' + $sql_file4) 2>&1"
         
         Write-Host "$output"
         log "Running SQL script: $sql_file4" INFO  
        
         if ($output | select-string -pattern "error") {
                log "ERROR executing SQL command. Please check error: $output" ERROR
                exit
         } else {
                log "2nd Update SQL statement executed succesfully. Serial numbers have been updated." INFO
         }
                
    } else {
        log "ERROR. SQL file not found on the script path: $sql_file4" ERROR
    }
Start-sleep -milliseconds 1000


<#$Count=@()

Import-Csv $spool_file_name4 |`
ForEach-Object {
        $Count += $_."Count"
}    


$EmailBody1 = $Count[0]#>

#$EmailBody2 = import-csv $spool_file_name2


# ------------------------------------------------------------- #
# STEP 7 - Send the file to the DL (Distribution List)
# ------------------------------------------------------------- #
Start-sleep -milliseconds 1000


    log "`n## 2. Sending email..." INFO

    <# $global:bodyTemplate = @"
<html><body><br>Hello All,</body><html>
<html><body><br>The current depth of Retry Queue as on $date : <font color ='red'>$EmailBody1</font>.
<br><br>Kindly take respective actions to clear this.
<br><br><hr>
<br><hr><br></body></html>
"@
#>


    # Load email addresses to send the file to
    $emp_data = @() 
    $file_entries = gc $dist_list
    $file_entries_error = gc $dist_list_error
    #$file_entries_cc = gc $dist_list_cc

    $to = $file_entries
    #$cc = $file_entries_cc
    

    # Mail to be sent

    $subject = "My Email Subject Line : " + $date_fileformat  
    $body = "Hello Recipient, 
    
    Please find attached the required extract.
	
    Regards,
    Sender"
    
 #   $att = $zip_file

 log "`n## 3. Mail to be sent is $body..." INFO
 $att = $spool_file_name1
 #$att2 = $spool_file_name5
 #$att3 = $spool_file_name6
 #$att4 = $spool_file_name7
 
    $o = New-Object -com Outlook.Application
    $mail = $o.CreateItem(0)
log "New outlook object created" INFO

#2 = high importance email header
#$mail.importance = 2

$mail.subject = $subject

$mail.body = $body

#$mail.cc = <Email Address to be kept in CC goes here>

$mail.Attachments.Add($att)
#$mail.Attachments.Add($att2)
#$mail.Attachments.Add($att3)
#$mail.Attachments.Add($att4)
log "Attachments added" INFO
#for multiple email, use semi-colon ; to separate
#example: $mail.cc = “User1@test.com; User2@test.com“

foreach ( $email_address in $to ) { $mail.Recipients.Add("$email_address") }
#foreach ( $email_address in $cc ) { $mail.cc.Add("$email_address") }
log "Sleep for 1 second before sending the email" INFO
Start-sleep -Seconds 2
$mail.Send()
log "Mail sent" INFO


#To delete object pointing to Outlook cause sometimes archive folder cannot be used while the script is running
Start-sleep -Seconds 1
#$o.Quit()


[System.Runtime.Interopservices.Marshal]::ReleaseComObject($o)
$o = $null
log "Outlook object deleted" INFO

    #sendMail $from $to $cc $subject $body $att
Start-sleep -Seconds 1
    log "Email sent to distribution list for error" INFO
    
Start-sleep -Seconds 1
    

# ------------------------------------------------------------- #
# END TASKS
# ------------------------------------------------------------- #

    write_log_HF "footer"

# ------------------------------------------------------------- #
# ------------------------------------------------------------- #
## FUNCTIONS ##
## ------------------------------------------------------------ # 

function log ($string, $level){

    $date_log = get-date -format "HH:mm:ss"
    if ($string | select-string -pattern "`n") {
        "`r`n$date_log $level	$string" >> $log_file
    } else {
        "$date_log $level	$string" >> $log_file
    }
}


function write_log_HF ($header_footer, $special_string){

    $date_log = get-date -format "dd.MM.yyyy HH:mm:ss"
    
    if ($header_footer -eq "startup") {
        # Write header to the log file
        
        $previous_log = $log_file+"_previous.log" 
                
        $FileExists = (Test-Path $log_file -PathType Leaf)
        If (($FileExists)) {
            # Rename the existing log file to previous
           
            move-item -Force $log_file $previous_log
        } 
        
        "# ------------------------------------------------------------------" > $log_file
        "# -------------------------- Script Job ----------------------------" >> $log_file
        "# ------------------------------------------------------------------" >> $log_file
        "# Date: $date_log" >> $log_file
        "# ------------------------------------------------------------------" >> $log_file
        "" >> $log_file
        return
    }
    
    
    if ($header_footer -eq "footer") {
        # Write footer to the log file
        ""  >> $log_file
        "# ------------------------------------------------------------------" >> $log_file
        "# Job Completed: $date_log" >> $log_file
        "" >> $log_file
        return
    }

}
