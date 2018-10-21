$triggerWords = "Finance", "Cash", "swift", "bank transfer", "swift", "banking", "financial", "@gmail.com", "@yahoo.com", "@yahoo.co.uk", "@mailinator.com", "MS-Charts", "IBAN", "payment", "invoice", "accounts"
$unusualFolders = "RSS Feeds", "RSS Subscriptions", "Deleted Items", "Junk Email", "Drafts", "Junk", "SMS"
$allRules = get-InboxRule
$allMailboxes = get-mailbox -resultSize Unlimited
function Compare-Similar ($obj1, $obj2)
{
    $obj1| % {
        $c = $_
        $obj2| % {
            if ($c -match $_)
            {
                "[$_] : $c"
            }
        }
    }
}

$allRules| ForEach-Object {
    $mbName = $_.MailboxOwnerID
    $ruleName = $_.Name
    $subjectMatch = Compare-Similar $_.SubjectContainsWords $triggerWords
    if ($subjectMatch)
    {
        write-host -BackgroundColor Red -ForegroundColor White "$mbName" -NoNewline
        write-host " - Subject Match: $ruleName : $subjectMatch"
    }

    $bodyMatch = Compare-Similar $_.BodyContainsWords $triggerWords
    if ($bodyMatch)
    {
        write-host -BackgroundColor Red -ForegroundColor White "$mbName" -NoNewline
        write-host " - Body Match: $ruleName : $bodyMatch"
    }

    $subjectbodyMatch = Compare-Similar $_.SubjectOrBodyContainsWords $triggerWords
    if ($subjectbodyMatch)
    {
        write-host -BackgroundColor Red -ForegroundColor White "$mbName" -NoNewline
        write-host " - Subject or Body Match: $ruleName : $subjectbodyMatch"
    }

    if ($_.MoveToFolder)
    {
        $folderMatch = Compare-Similar $_.MoveToFolder $unusualFolders
        if ($folderMatch)
        {
            write-host -BackgroundColor Red -ForegroundColor White "$mbName" -NoNewline
            write-host " - Folder Match: $ruleName : $folderMatch"
        }
    }
    if ($_.ForwardTo)
    {
        $mbName = $_.MailboxOwnerID
        $ruleName = $_.Name
        if ($_.ForwardTo)
        {
            $_.ForwardTo| ForEach-Object {
                if ($_ -like "*SMTP*")
                {
                    $forwarderMatch = $_
                    write-host -BackgroundColor Red -ForegroundColor White "$mbName" -NoNewline
                    write-host " - Forwarding Rule Match: $ruleName : $forwarderMatch"
                }
            }
        }
    }
}
$allMailboxes|Where-Object {$_.ForwardingSmtpAddress} |%{
    write-host -BackgroundColor Red -ForegroundColor White $_.Name -NoNewline
    write-host "$($_.ForwardingSmtpAddress) - $($_.DeliverToMailboxAndForward)"
}