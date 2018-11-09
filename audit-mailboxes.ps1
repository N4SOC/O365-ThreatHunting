try {Get-MsolDomain -ErrorAction stop} catch {Write-Error -Category AuthenticationError  -Message "Exchange Online Not Connected" -RecommendedAction "Check readme file for instructions" -TargetObject "Office365 Exchange"  -CategoryTargetName "NotConnected" -CategoryTargetType "Online Services";exit} # Determine if MS Online services is connected
try {Get-Mailbox -resultsize 1 -ErrorAction stop} catch {Write-Error -Category AuthenticationError  -Message "Exchange Online Not Connected" -RecommendedAction "Check readme file for instructions" -TargetObject "Office365 Exchange"  -CategoryTargetName "NotConnected" -CategoryTargetType "Online Services";exit} # Determine if Exchange sessions is connected

$mailboxes = get-mailbox -resultsize unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"} #Get all mailboxes except discovery mailboxes
$mailboxes|Add-Member -MemberType NoteProperty -Name "MFAEnabled" -value "" # Add extra properties to mailbox object
$mailboxes|Add-Member -MemberType NoteProperty -Name "DelegatedAccess" -value ""
$mailboxes|Add-Member -MemberType NoteProperty -Name "Office365Administrator" -value ""
$admins = Get-MsolRoleMember -RoleObjectId (Get-MsolRole -RoleName "Company Administrator").ObjectId
$mb = $mailboxes| ForEach-Object { # For each mailbox
    $upn = $_.UserPrincipalName
    if ($admins|Where-Object {$_.EmailAddress -eq $upn}) # If user matches a global admin then flag as admin
    {
        $_.Office365Administrator = $true
    }
    try
    {
        $user = (Get-MsolUser -UserPrincipalName $_.UserPrincipalName -ErrorAction Stop)
        $_.MFANumber = $user.StrongAuthenticationUserDetails.PhoneNumber
        $_.DelegatedAccess = ($_|Get-MailboxPermission |Where-Object {$_.IsInherited -eq $False -and $_.User -ne "NT AUTHORITY\SELF"}|Group-Object -Property Identity| ForEach-Object {($_.Group.User -join ",")})
    }
    catch
    {
        write-host "No user account for mailbox: " $_.UserPrincipalName
    }
    return $_
}
$mb|select-object UserPrincipalName, MFAEnabled, DelegatedAccess, auditenabled, Office365Administrator|export-excel -IncludePivotTable -PivotRows AuditEnabled -PivotData @{UserPrincipalName = "count"} -IncludePivotChart -ChartType Doughnut -Show -Path ".\out.xlsx"
