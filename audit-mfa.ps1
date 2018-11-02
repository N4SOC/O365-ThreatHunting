#connect-msolservice
$users = get-msoluser -All # Get all azure users
# Get users where 2fa number is set and code isn't prefixed with +44
#$users_mfa=$users | where {$_.StrongAuthenticationUserDetails.PhoneNumber -ne $null} | select -expand StrongAuthenticationUserDetails -Property DisplayName | select -Property DisplayName,PhoneNumber
$users_mfa = $users | select -expand StrongAuthenticationUserDetails -Property DisplayName | select -Property DisplayName, PhoneNumber
# Add Country property to object
$users_mfa = $users_mfa |  Add-Member -MemberType NoteProperty -Name "Country" -Value " " -PassThru
$users_mfa | % { # For each account
    if ($_.PhoneNumber)
    {
        $countrycode = ($_.PhoneNumber.split(" ")[0])[1..3] -join '' # Get first section of phone number
        if ($countrycode -eq "44")
        {
            $_.Country = "United Kingdom"
        }
        else
        {
            $_.Country = (Invoke-RestMethod "https://restcountries.eu/rest/v2/callingcode/$countrycode")[0].name # Lookup prefix against country code list
        }
    }
    else
    {
        $_.Country = ""
        
    }
}
$users_mfa|Export-Excel

# -and $_.StrongAuthenticationUserDetails.PhoneNumber.split(" ")[0] -ne "+44"