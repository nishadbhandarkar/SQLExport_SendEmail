
# -------------------------------------------------------------------------------
# Script:      cred.ps1 
# Author:      Nishad Bhandarkar 
# Date:        08.11.2014 
# Keywords:    password, user
# Description: This script is capable to manage users and passwords to a local file
# -------------------------------------------------------------------------------
# Defauting params
param ([string]$pwdreset = "Y")
$dir_local = $MyInvocation.MyCommand.Path | Split-Path
. $dir_local\INI_functions.ps1

# ---------------------- VARIABLES ----------------------- #

$global:cred_file = "$dir_local\cred.ini"

# -------------------------------------------------------- #


Function Reset_Pwd ($user_name)
{
  
    # check if the cred file exists, if not create it with asking the new password
    $FileExists = (Test-Path $cred_file -PathType Leaf)
    
    If (($FileExists)){
        
        # Request to setup a password
        $Spwd = Read-Host -Prompt "Enter password for user $DB_uname" -AsSecureString
        #$SecurePassword = $Spwd | ConvertTo-SecureString -AsPlainText -Force
        $SecureStringAsPlainText = $Spwd | ConvertFrom-SecureString
      
        $result = Set-IniKey $cred_file "USERS" $user_name $SecureStringAsPlainText 1
        
        if ($result) { return $true } else { return $false }
        
    }
  
}

Function Setup_Cred ($user_name, $pwd)
{
    Write-Host “Writing Cred details to file for: $DB_uname”
    
    # check if the cred file exists, if not create it with asking the new password
    $FileExists = (Test-Path $cred_file -PathType Leaf)
    
    If (($FileExists)){
        
        # Request to setup a password
        $SecurePassword = $pwd | ConvertTo-SecureString -AsPlainText -Force
        $SecureStringAsPlainText = $SecurePassword | ConvertFrom-SecureString
     
       Write-Host "Writing to Ini file...." 
       Set-IniKey $cred_file "USERS" $user_name $SecureStringAsPlainText 1
        
    }
  
}

Function Get_Cred ($user_name)
{   

    # check if the cred file exists, if not create it with asking the new password
    $FileExists = (Test-Path $cred_file -PathType Leaf)
    
    If ($FileExists){
    
        #$data = Get-Content "$credentials_file"
        $data = Get-IniKey $cred_file "USERS" $user_name
        
        if ($data) {
            $SecureString = $data | ConvertTo-SecureString
            $Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $DB_uname, $SecureString
            Return $Credentials
        
        } else {
            return
        }
        
     }
  
}




