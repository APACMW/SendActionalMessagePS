param (
    [string]$recipientEmail
)
function Send-AdaptiveCardEmail {
    param (
        [string]$cardFilePath,
        [string]$htmlTemplatePath,
        [string]$recipientEmail
    )

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
    $recipientMail = $recipientEmail;

    $nowString = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmss')
    $params = @{
        Message         = @{
            Subject      = "Test Adaptive Card at $($nowString)"
            Body         = @{
                ContentType = "HTML"
                Content     = $messageHtml
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $recipientMail
                    }
                }
            )
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
        Write-Host "3. For sending to others, register at: https://aka.ms/publishoam" -ForegroundColor White
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
    }
    finally {
        Disconnect-MgGraph
    }
}

# main script execution
$scriptFilePath = $MyInvocation.MyCommand.Path;
$folderPath = Split-Path $scriptFilePath;
$cardFilePath = Join-Path $folderPath "Card.json";
$htmlTemplatePath = Join-Path $folderPath "MessageBody.html";

if ((Test-Path -Path $cardFilePath) -and (Test-Path -Path $htmlTemplatePath)) {
    if ($recipientEmail -match '^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') {
        Send-AdaptiveCardEmail -cardFilePath $cardFilePath -htmlTemplatePath $htmlTemplatePath -recipientEmail $recipientEmail;
    }
    else {
        Write-Error "Invalid email address format: $recipientEmail"
        exit 1
    }
}
else {
    Write-Error "Required files not found. Ensure the file $cardFilePath and $htmlTemplatePath exist."
    exit 1
}