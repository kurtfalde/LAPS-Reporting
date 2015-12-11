<#
    This Sample Code is provided for the purpose of illustration only and is not 
    intended to be used in a production environment.  THIS SAMPLE CODE AND ANY 
    RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
    EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
    MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a 
    nonexclusive, royalty-free right to use and modify the Sample Code and to 
    reproduce and distribute the object code form of the Sample Code, provided 
    that You agree: 
    (i)            to not use Our name, logo, or trademarks to market Your 
                    software product in which the Sample Code is embedded; 
    (ii)           to include a valid copyright notice on Your software product 
                    in which the Sample Code is embedded; and 
    (iii)          to indemnify, hold harmless, and defend Us and Our suppliers 
                    from and against any claims or lawsuits, including attorneys’ 
                    fees, that arise or result from the use or distribution of 
                    the Sample Code.


    Originally written by Kurt Falde Using PSEventLogWatcher from http://pseventlogwatcher.codeplex.com/ 
    New version for parsing LAPS Audit Log entries 9/20/2015 Kurt Falde MSFT Services

#>



$workingdir = "c:\laps"

cd $workingdir

#The following line was needed on a 2008 R2 setup at one customer It shouldn't hurt having it here however this can be removed in some cases
Add-Type -AssemblyName System.Core

#Following added to unregister any existing SLAMWatcher Event in case script has already been ran
Unregister-Event LAPSWatcher -ErrorAction SilentlyContinue
Remove-Job LAPSWatcher -ErrorAction SilentlyContinue


Import-Module .\EventLogWatcher.psm1

#Verify Bookmark currently exists prior to setting as start point
$TestStream = $null
$ECName = (((Get-Content .\bookmark.stream)[1]) -split "'")[1]
$ERId = (((Get-Content .\bookmark.stream)[1]) -split "'")[3]
$TestStream = Get-WinEvent -LogName $ECName -FilterXPath "*[System[(EventRecordID=$ERID)]]"
If ($TestStream -eq $null) {Remove-Item .\bookmark.stream}

$BookmarkToStartFrom = Get-BookmarkToStartFrom

# The type of events that you want to parse.. in most cases we are parsing the Forwarded Events log with a specific xpath query for the events we want.

$XpathQuery = "*[System[EventID=4662]] and *[EventData[Data[@Name='ObjectType'] and (Data='%{bf967a86-0de6-11d0-a285-00aa003049e2}')]]"
$EventLogQuery = New-EventLogQuery "ForwardedEvents" -Query $XpathQuery


$EventLogWatcher = New-EventLogWatcher $EventLogQuery $BookmarkToStartFrom 

$action = {     
                
                #Following is debug line while developing as it will let you enter the nested prompt/session where you
                #can actually query the $EventRecord / $EventRecordXML etc for troubleshooting
                #$host.EnterNestedPrompt()
                $outfile = "c:\laps\laps.csv"
    

                $EventObj = New-Object psobject
                $EventObj | Add-Member noteproperty DomainController $EventRecord.MachineName
                $EventObj | Add-Member noteproperty User $EventRecordXml.SelectSingleNode("//*[@Name='SubjectUserName']")."#text"
                
                #Adding Date in Short Format to .csv object (Time is nice but it gets very messy when you start putting in Excel/PowerBI to aggregate down to hours/days etc)
                $EventObj | Add-Member noteproperty EventDate $EventRecord.TimeCreated.ToShortDateString()
                
                #Adding EventRecordID field to .csv object for a unique identifier
                $EventObj | Add-Member noteproperty EventRecordID $EventRecordXML.Event.System.EventRecordID
                
                #Adding Target Computer Object after resolving GUID to name
                $ComputerGUID = ($EventRecordXML.SelectSingleNode("//*[@Name='ObjectName']")."#text") -replace "%", "" -replace "{", "" -replace "}", ""
                          $ComputerName = $null
                          Try
                          {
                          $ComputerName = (get-adobject -id $ComputerGUID).Name
                          }
                          Catch
                          {
                          $ComputerName = $ComputerGUID
                          }
               $EventObj | Add-Member noteproperty TargetComputer $ComputerName
                
                          
           If ($Outfile -ne $Null)
            {
                write-host $EventObj
                $EventObj | Convertto-CSV -Outvariable OutData -NoTypeInformation 
                
                $OutPath = $Outfile
                If (Test-Path $OutPath)
                {
                    $Outdata[1..($Outdata.count - 1)] | ForEach-Object {Out-File -InputObject $_ $OutPath -append default}
                } else {
                    Out-File -InputObject $Outdata $OutPath -Encoding default
                }
            }
            
            
            }

Register-EventRecordWrittenEvent $EventLogWatcher -action $action -SourceIdentifier LAPSWatcher

$EventLogWatcher.Enabled = $True