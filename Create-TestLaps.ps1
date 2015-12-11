# Use to create a series of test OU's with computer objects, set the password attribute to something on a number of them and then perform a read on the password attributes
# Useful for testing auditing as well as a querying of set passwords in an environment

for($i=100; $i -le 120; $i++){ 

        New-ADOrganizationalUnit -Path "dc=contoso,dc=com" -Name LAPS$i
        for($c=1; $c -le 100; $c++){ 
            new-ADComputer -path "ou=laps$i,dc=contoso,dc=com" -Name LAP$i$c
            }
        for($b=1; $b -le 50; $b++){ 
            {SET-ADcomputer LAPS$i$b –replace @{'ms-Mcs-AdmPwd'=”STUFF”}
            Get-ADComputer LAPS$i$b -Properties ms-Mcs-AdmPwd}
            }
        }

<#
Use the following to delete this test environment if needed

for($i=100; $i -le 120; $i++){ 
            
            Set-ADObject -ProtectedFromAccidentalDeletion $false -Identity "OU=LAPS$i,dc=contoso,dc=com"

            remove-ADOrganizationalUnit -Identity "OU=LAPS$i,dc=contoso,dc=com" -Recursive -Confirm:$false
            
            }
#>