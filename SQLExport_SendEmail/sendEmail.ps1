# -------------------------------------------------------------------------------
# Script:      sendEmail.ps1(Logging function included) 
# Author:      Nishad Bhandarkar
# Date:        10.11.2014 
# Keywords:    job, batch, report, email, Logging
# Description: Sends an email using Server SmtpClient which is to be configured at server
#              All loggings included here as well
#              Based on code by Zjef van Driel but made dynamic
#              IMPORTANT: Requires at least one Interactive execution to encrypt 
#                         the password
# Updates: 
#  1. <Date> - <Changed by> - Change Description
# -------------------------------------------------------------------------------

$dir_local = $MyInvocation.MyCommand.Path | Split-Path


# ---------------------- VARIABLES ----------------------- #

$global:smtpServer = "<Your_SMPT_Server>"
$global:email_log = "$dir_local\email_log.log"

# -------------------------------------------------------- #

#write_log_HF "header"


# ------------------------------------------------------------- #
## FUNCTIONS ##
## ------------------------------------------------------------ # 

function sendMail ($from, $to_array, $cc_array, $subject, $body, $att_file_path) {

     Write-Host "Sending Email"

     
     #Loading the file into an object
     $att = New-Object Net.Mail.Attachment($att_file_path)
      
     #Creating a Mail object
     $msg = new-object Net.Mail.MailMessage

     #Creating SMTP server object
     $smtp = new-object Net.Mail.SmtpClient($smtpServer)

     #$smtp.UseDefaultCredentials = false

     $smtp.EnableSsl = $True 
     $smtp.Credentials = New-Object System.Net.NetworkCredential("<NetworkUsername>", "<NetworkPassword>"); 



     #Email structure 
     $msg.From = "$from"
     
     # For each to add to the msg Object
     foreach ( $email_address in $to_array ) { $msg.To.Add("$email_address") }
     foreach ( $email_address in $cc_array ) { $msg.Cc.Add("$email_address") }
     $msg.subject = "$subject"
     $msg.body = "$body"
     $msg.Attachments.Add($att)

     #Sending email 
     $smtp.Send($msg)
      
}


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