$triggerWords = "Finance", "Cash", "swift", "bank transfer", "swift", "banking", "financial", "@gmail.com", "@yahoo.com", "@yahoo.co.uk", "@mailinator.com", "MS-Charts", "IBAN", "payment", "invoice", "accounts", "Fraud", "Scam", "Spam", "Phishing", "MS-Chart", "MS Chart"
$unusualFolders = "RSS Feeds", "RSS Subscriptions", "Deleted Items", "Junk Email", "Drafts", "Junk", "SMS", "Sent Items", "Trash", "Call Log", "Infected Items"
$allRules = $allMailboxes| foreach-object {get-InboxRule -mailbox $_.Name}
$allMailboxes = get-mailbox -resultSize Unlimited
function Compare-Similar ($obj1, $obj2)
{
    $obj1| foreach-object {
        $objCompare = $_
        $obj2| foreach-object {
            if ($objCompare -match $_)
            {
                "[$_] : $objCompare"
            }
        }
    }
}
$output=@{}
$allRules| ForEach-Object {
    $mbName = $_.MailboxOwnerID
    $ruleName = $_.Name
    $subjectMatch = Compare-Similar $_.SubjectContainsWords $triggerWords
    if ($subjectMatch)
    {
        $output+=@{Mailbox=$mbName
            MatchType="Subject Match"
            Rulename=$ruleName
            Match=$subjectMatch}
    }

    $bodyMatch = Compare-Similar $_.BodyContainsWords $triggerWords
    if ($bodyMatch)
    {
        $output+=@{Mailbox=$mbName
            MatchType="Body Match"
            Rulename=$ruleName
            Match=$bodyMatch}
    }

    $subjectbodyMatch = Compare-Similar $_.SubjectOrBodyContainsWords $triggerWords
    if ($subjectbodyMatch)
    {
        $output+=@{Mailbox=$mbName
            MatchType="Subject or Body Match"
            Rulename=$ruleName
            Match=$subjectbodyMatch}
    }

    if ($_.MoveToFolder)
    {
        $folderMatch = Compare-Similar $_.MoveToFolder $unusualFolders
        if ($folderMatch)
        {
            $output+=@{Mailbox=$mbName
                MatchType="Folder Match"
                Rulename=$ruleName
                Match=$foldermatch}
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
                    $output+=@{Mailbox=$mbName
                        MatchType="Forwarder Match"
                        Rulename=$ruleName
                        Match=$forwardermatch}
                }
            }
        }
    }
}
$allMailboxes|Where-Object {$_.ForwardingSmtpAddress} | foreach-object {
    write-host -BackgroundColor Red -ForegroundColor White $_.Name -NoNewline
    write-host "$($_.ForwardingSmtpAddress) - $($_.DeliverToMailboxAndForward)"
}

$output
#$output|Export-Clixml ./test.xml


