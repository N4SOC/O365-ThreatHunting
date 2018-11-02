$mailboxes = get-mailbox -resultsize unlimited -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}
$mailboxes|Add-Member -MemberType NoteProperty -Name "MFANumber" -value ""
$mailboxes|Add-Member -MemberType NoteProperty -Name "DelegatedAccess" -value ""
$mailboxes|Add-Member -MemberType NoteProperty -Name "Office365Administrator" -value ""
$admins = Get-MsolRoleMember -RoleObjectId (Get-MsolRole -RoleName "Company Administrator").ObjectId
$mb = $mailboxes| ForEach-Object {
    $upn = $_.UserPrincipalName
    if ($admins|Where-Object {$_.EmailAddress -eq $upn})
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
$mb|select-object UserPrincipalName, MFANumber, DelegatedAccess, auditenabled, Office365Administrator|export-excel -IncludePivotTable -PivotRows AuditEnabled -PivotData @{UserPrincipalName = "count"} -IncludePivotChart -ChartType Doughnut -Show -Path ".\out.xlsx"
