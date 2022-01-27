# -- Set csv locations relative to the directory this script is running from.

$csvLocation = $PSScriptRoot + '\AD_Output.csv' 
$csvLocationStudent = $PSScriptRoot + '\AD_Output_Student.csv'
$fullOU = "" # -- This should be the AD student accounts OUs distinguishedName ["OU=Students,OU=Accounts,OU=?,DC=?,DC=?,DC=?" ] 

Import-Module ActiveDirectory

# -- First Output is based off an AD group. in this case all staff members are contained within the DL-Staff AD group.
#    This section pulls all active members of this group and outputs the required fields into a csv file to be used by another script.

$groupname = "DL-Staff" # -- Change based on the group you are trying to pull from
$users = Get-ADGroupMember -Identity $groupname | ? {$_.objectclass -eq "user"}
$result = @()
foreach ($activeusers in $users){
   $result += (Get-ADUser -Identity $activeusers | ? {$_.enabled -eq $true} | Get-ADUser -Properties name, sAMAccountName, description, sn, givenName, department,
mail, title, mobile, physicalDeliveryOfficeName | select name, sAMAccountName, description, sn, givenName, department, mail, title, mobile, physicalDeliveryOfficeName)
}

$result | Export-csv -path $csvLocation -NoTypeInformation # -- Output to AD_output.csv

# -- The second output pulls all active students and outputs the required fields into a csv file.

$users2 = (Get-ADUser -Filter * -SearchBase $fullOU -SearchScope Subtree)
$result2 = @()
foreach ($activeusers in $users2){
   $result2 += (Get-ADUser -Identity $activeusers | ? {$_.enabled -eq $true} | Get-ADUser -Properties name, sAMAccountName, description, sn, givenName, department,
mail, title, mobile, physicalDeliveryOfficeName | select name, sAMAccountName, description, sn, givenName, department, mail, title, mobile, physicalDeliveryOfficeName)
}

$result2 | Export-csv -path $csvLocationStudent -NoTypeInformation # -- Output to AD_Output_Student.csv