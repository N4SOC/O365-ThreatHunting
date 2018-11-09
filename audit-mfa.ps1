########################################################
# Modify these lists as required
$expectedCountries="United Kingdom","Ireland","Germany"
$suspiciousCountries="Nigeria","Malaysia","South Africa"
########################################################

$users = get-msoluser -All # Get all azure users
$users_mfa = $users | Select-Object -expand StrongAuthenticationUserDetails -Property DisplayName | select-object -Property DisplayName, PhoneNumber
# Add Country property to object
$users_mfa = $users_mfa |  Add-Member -MemberType NoteProperty -Name "Country" -Value " " -PassThru
$users_mfa | ForEach-Object { # For each account
    if ($_.PhoneNumber) # Only list accounts with MFA attributes set
    {
        $countrycode = ($_.PhoneNumber.split(" ")[0])[1..3] -join '' # Get first section of phone number
        if ($countrycode -eq "44") # Added because the first alphabetic entry for +44 is guernsey
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

$conditionalFormat=($expectedCountries|foreach-object{New-ConditionalText $_ darkblue cyan}) # Highlight expected countries in blue/cyan
$conditionalFormat+=($suspiciousCountries|foreach-object{New-ConditionalText $_ red yellow})# Highlight suspicious countries in red/yellow
$users_mfa|Export-Excel -ConditionalText ($conditionalFormat)
 