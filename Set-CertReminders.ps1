#Requires -Modules Microsoft.PowerShell.SecretManagement,iamJenkinsUtils
<#region Comment Header
.SYNOPSIS
    This script is used to send email reminders to certification reviewers who have not modified or signed off on ACTIVE review in more than ?? day.
.DESCRIPTION

.LINK
    redacted
.NOTES
    Author: Jessica Ascanio jascanio@chewy.com, Matthew Wander mwander@chewy.com
    Creation Date: 07/19/2023
    Modified Date: 10/04/2023
    Purpose/Change: 
    

    Version: 0.1.2

endregion #>

#region Param Block - this is not a true param block because of an issue with the Jenkins Powershell plugin
# Specifies which SP environment (prod, sandbox) the script should run against, with an option to hit both.
[string] $SailPointEnvironment = $(if ($env:SailPointEnvironment) { $env:SailPointEnvironment } else { 'putEnvHere' })

# Specifies the Previous Reminder JSON File Path Name
[string] $PreviousRemindersPathName = $(if ($env:PreviousRemindersPathName) { $env:PreviousRemindersPathName } else { $env:Workspace })

# Specifies the Previous Reminder JSON File Name
[string] $PreviousRemindersFileName = $(if ($env:PreviousRemindersFileName) { $env:PreviousRemindersFileName } else { "Default-Reminder-Json.json" })

# Specifies the Report Attachment Save Location
[string] $ReportSaveFile = $(if ($env:ReportSaveFile) { $env:ReportSaveFile } else { "reminderReport.csv" })

# Specifies the path for the Company Logo File
[string] $CompanyLogoPath = $(if ($env:CompanyLogoPath) { $env:CompanyLogoPath } else { "C:\Program Files\PowerShell\7\Modules\CompanyLogo" })

# Specifies the name for the Company Logo File
[string] $CompanyLogoFileName = $(if ($env:CompanyLogoFileName) { $env:CompanyLogoFileName } else { "Company-Banner.png" })

# Specifies whether or not to run the script in Verbose mode
[bool] $Verbose = $(if ($env:Verbose -eq 'True') { $true } else { $false })

# Specifies the sender of the error alert.
[string] $From = $env:From

# Specifies the report recipient
[string[]] $EmailTo = $(if ($env:EmailTo) { $env:EmailTo.split(',').trim() })   

# Specifies the SMTP Server.
[string] $SmtpServer = $env:SmtpServer
#endregion

#region Authenticate to Sailpoint
try {
    $spAccessToken, $spHeaders, $spEnv = Connect-SailpointAPI $SailPointEnvironment
}
catch {
    Write-Host "ERROR: Failed to access $($spEnv) token. Exception: $($_.Exception.Message)" 
    throw $_.Exception
}
#endregion

#region define style, logo, attachments & counter file.
$style = Get-EmailCSS
$companyLogo = "<center><img src=`"cid:$($CompanyLogoFileName)`" alt=`"white Company logo on blue background`" /></center>"
$emailAttachments = @()
$emailAttachments += "$CompanyLogoPath\$CompanyLogoFileName"

# Define the current date to compare with certification dates
$today = Get-Date

# pull in previous reminders tracking file
try {
    if (Test-Path -Path "$PreviousRemindersPathName\$PreviousRemindersFileName" -ErrorAction Stop) { 
        $previousReminders = Get-Content -Path "$PreviousRemindersPathName\$PreviousRemindersFileName" | ConvertFrom-Json
    } 
    else {
        # did not find previous reminders tracking file. Creating empty array
        # this is likely the first time we are running this
        $previousReminders = @()
    }
}
catch {
    Write-Host "ERROR: Failed to check for existing reminders file. Exception: $($_.Exception.Message)"
    throw $_.Exception
}

#endRegion

#region SailPoint
#Sailpoint API call data
$apiPathBeta = "https://$spEnv.api.identitynow.com/beta"
$allCertifications = @()

try {
    $allCertifications = Get-SailpointCertifications -Token $spAccessToken -Environment $spEnv -Headers $spHeaders
    Write-Host "INFO: Found $($allCertifications.count) Certifications in SailPoint."
    Write-Host "==================================="
}
catch {    
    Write-Host "ERROR: Failed to obtain certifications from Sailpoint.  Exception: $($_.Exception.Message)"
    throw $_.Exception
}
#endregion

#region Processing all Certifications found from SailPoint
if ($allCertifications) {
    # Define the arrays being used throughout the script
    $noReminder = 0
    $escalateManager = @()
    $remindUser = @()
    $escalateWhiteGlove = @()
    $errorCampaign = @()

    foreach ($certification in $allCertifications) {

        # Extract relevant certification details
        $certName = $certification.name
        $certDueDate = $certification.due
        $certModifyDate = $certification.modified
        $certCreationDate = $certification.created
        $certOwner = $certification.reviewer.name
        $certDesc = $certification.campaign.description
        $certID = $certification.id
        
        # if we have found a previous reminder
        $reminder = $previousReminders | Where-Object { $_.id -eq $certID }

        # Display certification information
        Write-Host "INFO: Certification Name - $certName"
        Write-Host "INFO: Certification Description - $certDesc"
        Write-Host "INFO: Creation date - $certCreationDate"
        Write-Host "INFO: Modification date - $certModifyDate"
        Write-Host "INFO: Due Date - $certDueDate"
        Write-Host "INFO: Assigned Reviewer - $certOwner"

        if ($reminder) {
            # Display a separator for better readability
            Write-Host "INFO: Object found in JSON file. Processing for new reminder."
            Write-Host "INFO: emailCount is currently set to $($reminder.emailCount) for $certName"
                
            if ($reminder.emailCount -ge '7') {
                $escalateWhiteGlove += $certification
            }
            # This will be true if the due date is less than than today 
            elseif ($certDueDate.date -lt $today.date) {
                $reminder.emailCount++
                $reminder.emailLastSent = $today
                $escalateManager += $reminder
                Write-Host "INFO: Certification requires a manager Escalation."
            }
            #>
            # This will be true if the modify date is less than today and the creation date is greater than 2 days
            elseif ($certModifyDate.date -ne $today.date -and $certCreationDate.AddDays(2).date -le $today.date) {
                $reminder.emailCount++
                $reminder.emailLastSent = $today
                $remindUser += $reminder
                Write-Host "INFO: Certification requires a reminder email."
            }
            # This will be true if the creation date - 2days is less than today
            elseif ($certCreationDate.AddDays(2) -ge $today -or $certModifyDate -eq $today) {
                $noReminder++
                Write-Host "INFO: No reminder or escalation necessary."
            }
            else {
                Write-Host "ERROR: Certificaiton campaign does not fall into a escalation category, Please investigate!"
                $errorCampaign += $certification
            }        
        }
        else {
            Write-Host "INFO: New Certification found...will add relevant information to JSON File."
            $noReminder++

            $params = @{
                Method                  = "GET"
                Token                   = $spAccessToken
                Authentication          = "Bearer"
                Headers                 = $spHeaders
                StatusCodeVariable      = "Status"
                ResponseHeadersVariable = "RHeaderVar"
                ErrorAction             = "STOP"
            }
            
            try {
                # Retrieve the reviewer's details from SailPoint
                $spUserApi = "$apiPathBeta/identities/$($certification.reviewer.id)"
                $spUser = Invoke-RestMethod @params -Uri $spUserApi
            
                # Get the manager's details from SailPoint
                $managerAPI = "$apiPathBeta/identities/$($spUser.managerRef.id)"
                $spUserManager = Invoke-RestMethod @params -Uri $managerAPI
                $certification | Add-Member -MemberType NoteProperty -Name 'ManagerEmail' -Value $spUserManager.emailAddress -Force
                Write-Host "INFO: Manager found $($spUserManager.name) - $($spUserManager.emailAddress)"
            
                $certification | Add-Member -MemberType NoteProperty -Name 'managerName' -Value $spUserManager.name -Force
                $certification | Add-Member -MemberType NoteProperty -Name 'EmailCount' -Value 0 -Force
                $certification | Add-Member -MemberType NoteProperty -Name 'EmailLastSent' -Value '' -Force
                $previousReminders += $certification
            }
            catch {
                Write-Host "ERROR: Unable to retrieve user records from SailPoint. Exception: $($_.Exception.Message)"
                $errorCampaign += $certification
            }
        }
        Write-Host "-----------------------------------"
    }
    #endRegion

    #region Send individual reminder and/or escalation emails
    #saving array of objects to json
    $previousReminders | ConvertTo-Json -Depth 10 | Set-Content -Path "$PreviousRemindersPathName\$PreviousRemindersFileName" -Force
    Write-Host "INFO:  $($remindUser.count) certifications require an email reminder."
    Write-Host "INFO: $($escalateManager.count) certifications that require manager escalation."
    Write-Host "INFO: $noReminder certifications do not require a reminder yet."
    Write-Host "==================================="

    # Send Reminder Email
    if ($remindUser) {
        foreach ($remind in $remindUser) {
            $sendReminderEmailBody += @"
        $style
        $CompanyLogo

        <br>

        <p>Hello $($remind.reviewer.name),<br><br></p>

        BODY OF EMAIL GOES HERE
        BODY OF EMAIL GOES HERE
        BODY OF EMAIL GOES HERE

"@

            $emailSplat = @{
                SmtpServer  = $SmtpServer
                From        = $From
                To          = $remind.reviewer.email
                Subject     = "[ACTION REQUIRED] REMINDER: $($remind.campaign.description) is due!"
                Body        = $sendReminderEmailBody
                BodyAsHtml  = $true
                Attachments = $emailAttachments
                Priority    = "High"
                ErrorAction = 'Stop'
            }
            try {
                Send-MailMessage @emailSplat
                Write-Host "INFO: Sent email report to $($remind.reviewer.email)."
            }
            catch {
                Write-Host "ERROR: Failed to send email. Exception: $($_.Exception.Message)"
                Write-Host $_.Exception
            }
            Remove-Variable sendReminderEmailBody, emailSplat, remind -ErrorAction SilentlyContinue
        }
    }

    # Send Manager Email
    if ($escalateManager) { 
        foreach ($escalation in $escalateManager) {
            $sendManagerEmailBody += @"
        $style
        $CompanyLogo

        <br>

        <p>Hello $($escalation.managerName),<br><br></p>

        <p>You're receiving this email because we need your assistance in ensuring $($escalation.campaign.description) is completed.</p>
        <p>We have reached out to $($escalation.reviewer.name) before escalating to you and now require immediate action as they have passed the due date. If they are out of office, kindly forward this out to Daenerys Targeryon to discuss prompt reassignment. She
        just lost 2 dragons, i'm sure that'll go over well. </p>
        Keep in mind that their part of the audit isn't complete until they click <b style="color:red;">'Sign Off'. </b>This is how the GRC team knows we're good to review your submissions.</p>        

        BODY OF EMAIL GOES HERE
        BODY OF EMAIL GOES HERE
        BODY OF EMAIL GOES HERE
    

        <p>Thanks in advance!<br><br>

        Identity and Access Engineering</p>
"@

            $emailSplat = @{
                SmtpServer  = $SmtpServer
                From        = $From
                To          = $escalation.managerEmail
                CC          = $escalation.reviewer.email
                Subject     = "[ACTION REQUIRED]  ESCALATION: $($escalation.campaign.description) is overdue!"
                Attachments = $emailAttachments
                Body        = $sendManagerEmailBody
                BodyAsHtml  = $true
                Priority    = "High"
                ErrorAction = 'Stop'
                Verbose     = $Verbose
            }

            try {
                Send-MailMessage @emailSplat
                Write-Host "INFO: Sent email report to $($escalation.managerEmail)."
            }
            catch {
                Write-Host "ERROR: Failed to send email. Exception: $($_.Exception.Message)"
                Write-Host $_.Exception
            }

            Remove-Variable sendManagerEmailBody, emailSplat, escalation -ErrorAction SilentlyContinue
        }
    }
    #endregion

    #region Report email section
    $emailBody = $style

    $emailBody += @"
    $companyLogo
Sailpoint Certification Campaign Reminder Report<br/>
<br/>
Found $($allCertifications.count) total active certifications.<br/>
- $noReminder certifications didn't require a reminder yet.<br/>
- $($remindUser.count) reminders sent to campaign reviewers.<br/>
- $($escalateManager.count) escalations sent to campaign reviewers manager. <br/>
- $($escalateWhiteGlove.count) escalations are well over 7 email threshold and require White Glove remediation. <br/>
<br/>
"@

    if ($escalateManager) {
        $attachmentNeeded = $true
        $emailBody += @"
<br/>$($escalateManager.count) reminder emails have been sent to campaign reviewers.<br/><br/>
"@

        $emailBody += $escalateManager | Select-Object @{LABEL = "Campaign Name"; EXPRESSION = { "$($_.campaign.name)" } }, @{LABEL = "Due Date"; EXPRESSION = { "$($_.due)" } }, @{LABEL = "Reviewer Name"; EXPRESSION = { "$($_.reviewer.Name)" } }, @{LABEL = "Reviewer Email"; EXPRESSION = { "$($_.reviewer.email)" } }, @{LABEL = "Decisions Made"; EXPRESSION = { "$($_.decisionsMade)" } }, @{LABEL = "Decisions Total"; EXPRESSION = { "$($_.decisionsTotal)" } }, @{LABEL = "Emails Sent"; EXPRESSION = { "$($_.emailCount)" } }, @{LABEL = "Email Last Sent"; EXPRESSION = { "$($_.emailLastSent)" } } | 
        Get-HtmlTable
    }

    if ($escalateWhiteGlove) {
        $attachmentNeeded = $true
        $emailBody += @"
<br/>$($escalateWhiteGlove.count) campaigns have well exceeded the contact and escalation threshold and need white glove remediation.<br/><br/>
"@

        $emailBody += $escalateWhiteGlove | Select-Object @{LABEL = "Campaign Name"; EXPRESSION = { "$($_.campaign.name)" } }, @{LABEL = "Due Date"; EXPRESSION = { "$($_.due)" } }, @{LABEL = "Reviewer Name"; EXPRESSION = { "$($_.reviewer.Name)" } }, @{LABEL = "Reviewer Email"; EXPRESSION = { "$($_.reviewer.email)" } }, @{LABEL = "Decisions Made"; EXPRESSION = { "$($_.decisionsMade)" } }, @{LABEL = "Decisions Total"; EXPRESSION = { "$($_.decisionsTotal)" } } | 
        Get-HtmlTable
    }
    else {
    }
    
    if ($remindUser) {
        $attachmentNeeded = $true
        $emailBody += @"
<br/>$($remindUser.count) reminder emails have been sent to campaign reviewers.<br/><br/>
"@

        $emailBody += $remindUser | Select-Object @{LABEL = "Campaign Name"; EXPRESSION = { "$($_.campaign.name)" } }, @{LABEL = "Due Date"; EXPRESSION = { "$($_.due)" } }, @{LABEL = "Reviewer Name"; EXPRESSION = { "$($_.reviewer.Name)" } }, @{LABEL = "Reviewer Email"; EXPRESSION = { "$($_.reviewer.email)" } }, @{LABEL = "Decisions Made"; EXPRESSION = { "$($_.decisionsMade)" } }, @{LABEL = "Decisions Total"; EXPRESSION = { "$($_.decisionsTotal)" } }, @{LABEL = "Emails Sent"; EXPRESSION = { "$($_.emailCount)" } }, @{LABEL = "Email Last Sent"; EXPRESSION = { "$($_.emailLastSent)" } } | 
        Get-HtmlTable
    }
    
    $emailBody += @"
    Jenkins build info:
<ul><li>Build Number: $ENV:BUILD_NUMBER</li>
<li>Console Output: <a href = "$ENV:BUILD_URL`console">$ENV:BUILD_URL`console</a></li>
<li>Node Name: $ENV:NODE_NAME</li>
</font>
"@

    if ($attachmentNeeded) {
        $attachmentSavePath = "$PreviousRemindersPathName\$ReportSaveFile"
        # save report csv
        $previousReminders | Select-Object @{LABEL = "Campaign Name"; EXPRESSION = { "$($_.campaign.name)" } }, @{LABEL = "Due Date"; EXPRESSION = { "$($_.due)" } }, @{LABEL = "Reviewer Name"; EXPRESSION = { "$($_.reviewer.Name)" } }, @{LABEL = "Reviewer Email"; EXPRESSION = { "$($_.reviewer.email)" } }, @{LABEL = "Decisions Made"; EXPRESSION = { "$($_.decisionsMade)" } }, @{LABEL = "Decisions Total"; EXPRESSION = { "$($_.decisionsTotal)" } }, @{LABEL = "Emails Sent"; EXPRESSION = { "$($_.emailCount)" } }, @{LABEL = "Email Last Sent"; EXPRESSION = { "$($_.emailLastSent)" } } | 
        Export-Csv -Path $attachmentSavePath -NoTypeInformation

        $emailAttachments += $attachmentSavePath
    }
    else {
        Write-Host "No attachment needed"
    }

    $emailSplat = @{
        SmtpServer  = $SmtpServer
        From        = $From
        To          = $EmailTo
        Subject     = "Sailpoint UAR Reminder Report ($(Get-Date -UFormat '%x %r'))"
        Body        = $emailBody
        BodyAsHtml  = $true
        Attachments = $emailAttachments
        Priority    = "High"
        ErrorAction = 'Stop'
    }

    try {
        Send-MailMessage @emailSplat
        Write-Host "INFO: Sent email report to $EmailTo."
    }
    catch {
        Write-Host "ERROR: Failed to send email. Exception: $($_.Exception.Message)"
        throw $_.Exception
    }
}
else {
    Write-Host "INFO: There were no certifications found today."
}
#endregion