
#imports AD module
Import-Module ActiveDirectory
# sets local location, modify if needed 
# Set-Location '.\Desktop'
# tests whether the files already exists from previous iterations of the script
if (Test-Path -Path .\usersPerOU.csv) {
    Write-Output "File exists!"
    Remove-Item -Path .\usersPerOU.csv
} else {
    Write-Output "File does not exist."
}
if (Test-Path -Path .\usersModified.csv) {
    Write-Output "File exists!"
    Remove-Item -Path .\usersModified.csv
} else {
    Write-Output "File does not exist."
}


# variable maps to the specified OU Living
$OU = "OU=exampleOU,DC=exampleDC,DC=domain,DC=com"
$oldDomain = "oldDomain.com"
$newDomain = "newDomain.com"

Try
{
    # checks whether it maps to the defined OU
    if (Get-ADOrganizationalUnit -Identity $OU)
    {
        #Export OU Users with sam, upn and DN from specified OU
        Get-ADUser -Filter * -Properties * -SearchBase $OU | Select-object samaccountname, userprincipalname, distinguishedname | Export-Csv -Path .\usersPerOU.csv
        # Import CSV, modify userprincipalname, and include samaccountname in the output
        Import-Csv -Path ".\usersPerOU.csv" | ForEach-Object {
         # Replace old domain with new domain
        $modifiedUPN = $_.userprincipalname -replace $oldDomain, $newDomain

        # Create a new object that includes both samaccountname and modified userprincipalname
            [PSCustomObject]@{
                samaccountname    = $_.samaccountname
                userprincipalname = $modifiedUPN
            }
            # Exports modified results to csv including samaccountname and upn
        } | Export-Csv -Path ".\usersModified.csv" -NoTypeInformation -Encoding UTF8

        # imports users from modified csv
        $users = Import-Csv -Path ".\usersModified.csv"

        # loops through all users from csv and sets the upn & samaccountname
        foreach($user in $users)
        {
            $sam = $user.samaccountname
            $newUPN = $user.userprincipalname
            # obtains user from AD 
            $adUser = Get-ADUser -Identity $sam -Properties UserPrincipalName
            $oldUPN = $adUser.userprincipalname
            # tests whether the user exists in ad

            if ($adUser)
            {   
                # outputs to the console that it changes the SMTP now
                Write-Host "Updating SMTP"
                # Removes old SMTP entries if they exist
                Set-ADUser -Identity $adUser -Remove @{ProxyAddresses="SMTP:$oldUPN"}
                Set-ADUser -Identity $adUser -Remove @{ProxyAddresses="smtp:$newUPN"}
                # adds new proxy addresses, sets old UPN as alias, newUPN will be the new primary mail SMTP
                Set-ADUser -Identity $adUser -add @{ProxyAddresses="smtp:$oldUPN,SMTP:$newUPN" -split ","}
                # Outputs comment to the console that new UPN is updated 
                Write-Host "Updating $oldUPN UPN to new $newUPN UPN"
                # if user exists set his upn to the new upn
                Set-ADUser -Identity $adUser -UserprincipalName $newUPN
                # sets mail property to updated UPN
                Set-ADUser -Identity $adUser -EmailAddress $newUPN
                Set-ADUser -Identity $adUser -mail $newUPN
            }
            # if user is not found in AD it will display an error
            else 
            {
                Write-Host "$sam not found in directory"
            }
        }
    }
}
catch
{
    # catch statement will display any error information
    Write-Host  -ForegroundColor DarkRed "Error: $($PSItem.ToString())"
} 