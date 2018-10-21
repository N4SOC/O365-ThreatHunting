$triggerWords = "finance", "cash", "swift", "bank transfer", "swift", "banking", "financial", "@gmail.com", "MS-Charts", "IBAN", "payment", "invoice", "accounts"
$unusualFolders = "RSS Feeds", "RSS Subscriptions", "Deleted Items", "Junk Email"
$allRules = get-InboxRule

$allRules| ForEach-Object {
    $mbName = $_.MailboxOwnerID
    $ruleName = $_.Name
    $subjectMatch = Compare-Similar $_.SubjectContainsWords $triggerWords 
    if ($subjectMatch)
    {
        "Subject Match: $mbName : $ruleName : $subjectMatch"
    }}
$allRules| ForEach-Object {
    $mbName = $_.MailboxOwnerID
    $ruleName = $_.Name
    $bodyMatch = Compare-Similar $_.BodyContainsWords $triggerWords 
    if ($bodyMatch)
    {
        "Body Match: $mbName : $ruleName : $bodyMatch"
    }}
$allRules| ForEach-Object {
    $mbName = $_.MailboxOwnerID
    $ruleName = $_.Name
    $subjectbodyMatch = Compare-Similar $_.SubjectOrBodyContainsWords $triggerWords 
    if ($subjectbodyMatch)
    {
        "Subject or Body Match: $mbName : $ruleName : $subjectbodyMatch"
    }}

$allRules| ForEach-Object {
    if ($_.MoveToFolder)
    {
        $mbName = $_.MailboxOwnerID
        $ruleName = $_.Name
        $folderMatch = Compare-Similar $_.MoveToFolder $unusualFolders 
        if ($folderMatch)
        {
            "Folder Match: $mbName : $ruleName : $folderMatch"
        }
    }}

$allRules| ForEach-Object {
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
                    "Forwarder Match: $mbName : $ruleName : $forwarderMatch"
                }
            }
        }
    }
}

function Compare-Similar ($obj1, $obj2)
{
    $obj1| % {$c = $_; $obj2| % {if ($c -match $_)
            {
                "$c - $_"
            }
        }
    }
}

