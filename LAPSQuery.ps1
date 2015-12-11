
<#

Script to query various data points for the LAPS password attributes in Active Directory.
By default this queries for all computers in the domain the script is ran from



#>

cd C:\laps
del .\lapsquery.csv

$computers = Get-ADComputer -Filter * -Properties Name,ms-Mcs-AdmPwdExpirationTime,msDS-ReplAttributeMetaData

Foreach ($computer in $computers){
        
               
        $CompObj = New-Object psobject
        
        $CompObj | Add-Member noteproperty Computername $computer.Name
      
        If($computer.'ms-Mcs-AdmPwdExpirationTime' -eq $null){$ExpirationDate = $computer.'ms-Mcs-AdmPwdExpirationTime'}
        Else {$ExpirationDate = [DateTime]::FromFileTime($computer.'ms-Mcs-AdmPwdExpirationTime')} 
        $CompObj | Add-Member noteproperty NextPasswordExpiration $ExpirationDate
        
                
        $xmlAttribute = $computer."msDS-ReplAttributeMetaData"
        $xmlAttribute = “<root>” + $xmlAttribute + “</root>”
        $xmlAttribute = $xmlAttribute.Replace([char]0,” ”)
        $xmlattribute =[xml]$xmlattribute

        $OriginatingDCPwdLastSet = $null
        $OriginatingDCSiteLastSet = $null
        $LastTimePwdModified = $null
        
        foreach ($attribute in $xmlAttribute.root.DS_REPL_ATTR_META_DATA | Where-Object {$_.pszAttributeName -eq "ms-Mcs-AdmPwd"}){ 
                  
                  $OriginatingDCPwdLastSet = ($attribute.pszLastOriginatingDsaDN.Split(','))[1] -replace '^cn='  
                  $OriginatingDCSiteLastSet = ($attribute.pszLastOriginatingDsaDN.Split(','))[3] -replace '^cn=' 
                  $LastTimePwdModified = $attribute.ftimeLastOriginatingChange  
                           

                }
        
        $CompObj | Add-Member noteproperty OriginatingDCLastPwdSet $OriginatingDCPwdLastSet
        $CompObj | Add-Member noteproperty OriginatingDCSiteLastSet $OriginatingDCSiteLastSet
        $CompObj | Add-Member noteproperty LastTimePwdModified $LastTimePwdModified
        
        $ComputerOU = $computer.DistinguishedName -creplace "^[^,]*,",""
        $CompObj | Add-Member NoteProperty ComputerOU $ComputerOU

        $CompObj | Write-Host
        $compobj | Export-Csv .\lapsquery.csv -Append -NoTypeInformation

        }
