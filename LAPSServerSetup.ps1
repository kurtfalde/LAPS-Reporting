<#  Setup script for Subscription Server for LAPS Audit Events and reporting

Microsoft Services - Kurt Falde

#>

#Please change the $DCTOQUERY value to the DC that has been used for Audit Testing
$DCTOQUERY = "ADDC"
cd C:\LAPS

Import-Module ServerManager
If ((Get-WindowsFeature -Name RSAT-AD-PowerShell).Installed -eq $false){
    write-host "This needs the AD Powershell commandlets to proceed please run Install-WindowsFeature -Name RSAT-AD_PowerShell"
    Exit
    }

$rootdse = Get-ADRootDSE
$Guids = Get-ADObject -SearchBase ($rootdse.SchemaNamingContext) -LDAPFilter "(schemaidguid=*)" -Properties lDAPDisplayName,schemaIDGUID
ForEach ($Guid in $Guids){
  If ($guid.lDAPDisplayName -Like "*ms-mcs-admpwd"){
  $SGuid = ([System.GUID]$guid.SchemaIDGuid).Guid
  Write-host -ForegroundColor Cyan "The SchemaIDGuide for ms-mcs-admpwd in this forest is"([System.GUID]$guid.SchemaIDGuid)
  }
 }

$xpath =  "*[EventData[Data[@Name='AccessMask'] and (Data='0x100')]] and *[System[EventID=4662]] and *[EventData[Data[@Name='ObjectType'] and (Data='%{bf967a86-0de6-11d0-a285-00aa003049e2}')]]"
$LAPSAuditEvents = Get-WinEvent -ComputerName $DCTOQUERY -LogName Security -FilterXPath "$xpath" -MaxEvents 200

$xpathdata = @()
:Outer Foreach($LAPSAuditevent in $LAPSAuditEvents){
    [xml]$LAPSAuditEventXML = $LAPSAuditEvent.ToXml()
    $EventXMLNames = $LAPSAuditEventXML.Event.EventData.Data

    #loop through the "Properties" XML Node looking for Events that contain the LAPS ms-mscadmpwd SchemaIDGuid
    #Add all hits to a new Array $xpathdata

        Foreach ($EventXMLName in $EventXMLNames){
            If (($EventXMLName.name -eq "Properties") -and ($EventXMLName.'#text' -match "$SGuid")){
                $LAPSProp = $EventXMLName.'#text'
                If ($xpathdata.count -eq 0){$xpathdata += $LAPSProp}
                ElseIf (($LAPSProp | Select-String -AllMatches $xpathdata) -eq $null){$xpathdata += $LAPSProp   }
                }
        }
}

#Parsing through the $xpathdata array to create the right text needed for the xpath filter
$xpathguid = $null
If ($xpathdata.Count -le 1) {
    Write-host -ForegroundColor Cyan "Either a Single or No Xpath Filters created are you sure you tested LAPS UI and using ADUC at a minium on the DC you are testing? That or change the get-winevent line to a larger number of events"
    $xpathguid = "(Data='$xpathdata')" 
    }

ElseIf ($xpathdata.Count -gt 1) {
    Write-host -ForegroundColor Green "More than a single xpath filter found creating full xpath query output"
    Foreach ($xpaththing in $xpathdata){
        If ($xpaththing -ne $xpathdata[-1]){
        $xpathguid += "(Data='$xpaththing') or "
        }
        Else{$xpathguid += "(Data='$xpaththing')"}
        }
      }
 
#FYI by default we are filtering out local queries to this from the "SYSTEM" account on a DC as this appears to happen with a DC reading it's own account in AD and creates noise
#$xpathlaps = "<QueryList><Query Id=""0"" Path=""Security""><Select Path=""Security"">*[System[EventID=4662]] and *[EventData[Data[@Name='SubjectUserSid'] !='S-1-5-18']] and *[EventData[Data[@Name='ObjectType'] and (Data='%{bf967a86-0de6-11d0-a285-00aa003049e2}')]] and *[EventData[Data[@Name='Properties']]]</Select></Query></QueryList>"
#$xpathlaps = $xpathlaps.Replace("`n", "&#xD;&#xA;").Replace("`t", "&#x09;")

#(gc C:\laps\LAPSSubscription.xml).replace('CHANGETHIS', $xpathlaps) | sc C:\laps\LAPSSubscription.xml -Encoding UTF8

#Setting WinRM Service to automatic start and running quickconfig
Set-Service -Name winrm -StartupType Automatic
winrm quickconfig -quiet

#Set the size of the forwarded events log to 500MB
wevtutil sl forwardedevents /ms:500000000


#Running quickconfig for subscription service
wecutil qc -quiet

#Creating Applocker Subscription from XML files FYI we do delete any existing ones and recreate
If ((wecutil gs "LAPS Audit Events") -ne $NULL) {
    wecutil ds "LAPS Audit Events"
    wecutil cs .\LAPSSubscription.xml
    }
Else {wecutil cs .\LAPSSubscription.xml} 

#Fix up the query in the registry with the proper text with carriage returns replaced.. carriage returns in the xml file imported via wecutil do not make it to the registry key
$xpathguid = $xpathguid.Replace("`n","`r`n")
$xpathlaps = "<QueryList><Query Id=""0"" Path=""Security""><Select Path=""Security"">*[System[EventID=4662]] and *[EventData[Data[@Name='SubjectUserSid'] !='S-1-5-18']] and *[EventData[Data[@Name='ObjectType'] and (Data='%{bf967a86-0de6-11d0-a285-00aa003049e2}')]] and *[EventData[Data[@Name='Properties'] and $xpathguid]]</Select></Query></QueryList>"
Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\EventCollector\Subscriptions\LAPS Audit Events" -Name "Query" -Value $xpathlaps
wecutil rs "LAPS Audit Events"


#FYI if you need to export Subscriptions to fix SIDS or anything in an environment 
#use wecutil gs "%subscriptionname%" /f:xml >>"C:\Temp\%subscriptionname%.xml"

#Creating Task Scheduler Item to restart LAPS parsing script on reboot of system.
schtasks.exe /delete /tn "LAPS Parsing Task" /F
schtasks.exe /create /tn "LAPS Parsing Task" /xml LAPSParsingTask.xml
schtasks.exe /run /tn "LAPS Parsing Task"

#Creating Task Scheduler Item to run LAPS query script on a daily basis.
schtasks.exe /delete /tn "LAPS Query Task" /F
schtasks.exe /create /tn "LAPS Query Task" /xml LAPSQueryTask.xml
schtasks.exe /run /tn "LAPS Query Task"