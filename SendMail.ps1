<#
.SYNOPSIS
    Sends an adaptive card email to specified recipients.
.DESCRIPTION
    This script sends an adaptive card email to the specified recipients using Microsoft Graph API. It reads the adaptive card JSON from a file and formats it into an HTML email body. The script requires the Microsoft Graph PowerShell module to be installed and authenticated.
.NOTES
    author: Dong Qi (doqi@microsoft.com)
.PARAMETER recipientEmails
    The email addresses of the recipients.
.EXAMPLE
    .\SendMail.ps1 -recipientEmails "freeman@MngEnvMCAP965703.onmicrosoft.com", "doqi@MngEnvMCAP965703.onmicrosoft.com";
#>
param (
    [string[]]$recipientEmails
)
function Send-AdaptiveCardEmail {
    param (
        [string]$cardFilePath,
        [string]$htmlTemplatePath,
        [string[]]$recipientEmails
    )
    if (-not $recipientEmails -or $recipientEmails.Count -eq 0) {
        Write-Error "Recipient email address is required."
        exit 1
    }
    Connect-MgGraph -Scopes "Mail.Send"
    # Read and validate the card JSON
    $cardJson = Get-Content $cardFilePath -Encoding UTF8
    Write-Host "Card JSON loaded:" -ForegroundColor Green
    Write-Host $cardJson -ForegroundColor Cyan

    # Read and format the HTML template
    $messageBody = Get-Content $htmlTemplatePath -Encoding UTF8
    $messageHtml = [string]($messageBody -replace '\{0\}', $cardJson)

    Write-Host "`nHTML Message Body:" -ForegroundColor Green
    Write-Host $messageHtml -ForegroundColor Cyan

    $senderMail = (Get-MgContext).Account;

    $toRecipients = @()
    foreach ($recipientMail in $recipientEmails) {
        $toRecipients += @{
            EmailAddress = @{
                Address = $recipientMail
            }
        }
    }
    $nowString = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmss')
    $params = @{
        Message         = @{
            Subject      = "Test Adaptive Card at $($nowString)"
            Body         = @{
                ContentType = "HTML"
                Content     = $messageHtml
            }
            ToRecipients = $toRecipients
        }
        SaveToSentItems = $true
    }

    # Send the message
    try {
        Write-Host "`nSending email..." -ForegroundColor Green
        Send-MgUserMail -UserId $senderMail -BodyParameter $params
        Write-Host "Email sent successfully!" -ForegroundColor Green
    
        Write-Host "`nNext steps:" -ForegroundColor Yellow
        Write-Host "1. Check your email for the adaptive card" -ForegroundColor White
        Write-Host "2. If card doesn't appear, install 'Actionable Messages Debugger' add-in" -ForegroundColor White        
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
    }
    finally {
        Disconnect-MgGraph
    }
}

function main {
    param (
        [string[]]$recipientEmails
    )
    $recipientEmails = @($recipientEmails | Where-Object { $_ -match '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$' } | Select-Object -Unique)
    if (-not $recipientEmails -or $recipientEmails.Count -eq 0) {
        Write-Error "Recipient email address is required."
        exit 1
    }
    $folderPath = $PSScriptRoot
    $cardFilePath = Join-Path $folderPath "Card.json";
    $htmlTemplatePath = Join-Path $folderPath "MessageBody.html";
    if ((Test-Path -Path $cardFilePath) -and (Test-Path -Path $htmlTemplatePath)) {
        Send-AdaptiveCardEmail -cardFilePath $cardFilePath -htmlTemplatePath $htmlTemplatePath -recipientEmails $recipientEmails
    }
    else {
        Write-Error "Required files not found. Ensure the file $cardFilePath and $htmlTemplatePath exist."
        exit 1
    }
}

# main script execution
main -recipientEmails $recipientEmails;
