<#
.NOTES
	Name: Monitor_HungWorker.ps1
    Author: Daivd Paulson
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Flow of the Script
1. Throw disclaimer
2. Check Admin & Check for Exchange Snappin 
3. Verify Path of additional Script and Locations - only going to do the local server, the remote server will be done after we get that information 
4. Get Admin Creds 
5. Gather DAG Information and nodes to monitor 
6. Start Data Collections - Verify everything is started 
7. Monitor for the store worker hang event. Sleep 1 second between loops. 
8. Once store worker event occurs, trigger Initial Message to Admins then wait 5 minutes 
9. After 5 minutes, stop the data collection on that server and zip up all relevant data
10. Send another email to the admins, restart the data collection on that server. 
11. Continue to monitor unless the script is stopped

#>

[CmdletBinding()]
param(
[string]$Experfwiz_Script_Name = "Experfwiz.ps1",
[string]$Experfwiz_Directory = ".",
[string]$Experfwiz_Save_Data_Directory = ".",
[int]$Experfwiz_Interval = 3,
[int]$Experfwiz_Data_Maxsize = 5120
)

$ScriptName = "Monitor_HungWorker"
$Disclaimer = @" 

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

"@
cls 
Write-Warning $Disclaimer
Write-Host " " 
Write-Host "By running this script you agree to the license and are aware of the risks"
$y = Read-Host "Please enter 'y' if you want to proceed with the script: "
if($y -ne "y"){exit}



###############################
#
#   Functions 
#
###############################

#Function to test if you are an admin on the server 
Function Is-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

#Function to load the ExShell 
Function Load-ExShell {

    if($exinstall -eq $null){
    $testV14 = Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'
    $testV15 = Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'

    if($testV14){
        $Global:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath	
    }
    elseif ($testV15) {
        $Global:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath	
    }
    else{
        Write-Host "It appears that you are not on an Exchange 2010 or newer server. Sorry I am going to quit."
		exit
    }

    $Global:exbin = $Global:exinstall + "\Bin"

    Write-Host "Loading Exchange PowerShell Module..."
    add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
    }
}

Function Create-Folder {
param(
[Parameter(Mandatory=$true)][string]$Checker,
[Parameter(Mandatory=$false)][switch]$DisplayInfo
)
    if(-not (Test-Path -Path $Checker))
    {
        if($DisplayInfo){Write-Host("Creating Directory {0}" -f $Checker)}
        [System.IO.Directory]::CreateDirectory($Checker) | Out-Null
    }
    else
    {
        if($DisplayInfo){Write-Host("{0} is already created!" -f $Checker)}
    }

}

Function Send-Message {
param(
[Parameter(Mandatory=$true)][string]$SMTP_Sender,
[Parameter(Mandatory=$true)][string]$SMTP_Recipient,
[Parameter(Mandatory=$true)][string]$Message_Subject,
[Parameter(Mandatory=$true)][string]$SMTP_Server,
[Parameter(Mandatory=$false)]$Creds,
[parameter(Mandatory=$false)]$Message_Body = " ",
[Parameter(Mandatory=$false)][int]$Port = 25
)

    $message_error = $false 

    try
    {
        if($Creds -ne $null)
        {
            Send-MailMessage -To $SMTP_Recipient -From $SMTP_Sender -SmtpServer $SMTP_Server -Body $Message_Body -Subject $Message_Subject -Port $Port -Credential $Creds -ErrorAction Stop 
        }
        else
        {
            Send-MailMessage -To $SMTP_Recipient -From $SMTP_Sender -SmtpServer $SMTP_Server -Body $Message_Body -Subject $Message_Subject -Port $Port -ErrorAction Stop 
        }
    }
    catch
    {
        Write-Host "Error occurred when trying to send the message" 
        Write-Host " "
        $display_error = $Error[0].Exception.ToString()
        Write-Warning $display_error 
        $message_error = $true 
    }
    return $message_error
}


Function Get-ExchangeServersInDAG{
param(
[Parameter(Mandatory=$true)][string]$DAG_Name
)
    
    [array]$servers = (Get-DatabaseAvailabilityGroup -Identity $DAG_Name).Servers.Name 
    return $servers
}

Function Get-DAGNameFromMailboxServer {
param(
[Parameter(Mandatory=$true)][string]$MailboxServer_Name
)
    [string]$DAG_Name = (Get-MailboxServer -Identity $MailboxServer_Name).DatabaseAvailabilityGroup.Name
    return $DAG_Name
}

Function Get-WatchingServers {
param(
[parameter(Mandatory=$true)][string]$Server_Name
)
    $watcher_list = (Get-ExchangeServersInDAG -DAG_Name (Get-DAGNameFromMailboxServer -MailboxServer_Name $Server_Name))
    return $watcher_list
}



Function Get-EventLogs{
param(
[Parameter(Mandatory=$true)][string]$ComputerName,
[Parameter(Mandatory=$true)][HashTable]$FilterHashTable
)
    $bErr = $false
    try
    {
        $events = Get-WinEvent -ComputerName $ComputerName -FilterHashtable $FilterHashTable -ErrorAction SilentlyContinue
    }
    catch 
    {
        Write-Host( "Error trying to get event logs from Computer: {0} at {1}" -f $ComputerName,(Get-Date).ToString()) -ForegroundColor Red
        Write-Host("Exception: {0}" -f $Error[0].Exception) -ForegroundColor Red
        Write-Host("Stack Trace: {0}" -f $Error[0].ScriptStackTrace) -ForegroundColor Red
        $bErr = $true
    }

    if($bErr)
    {
        return "Err"
    }
    return $events
}

<#
		$obj | Add-Member -Name ServerName -MemberType NoteProperty -Value $server
		$obj | Add-Member -Name ExperfStatus -MemberType NoteProperty -Value $temp.Status
		$obj | Add-Member -Name ExperfFullPath -MemberType NoteProperty -Value $temp.FullPath
		$obj | Add-Member -Name ExperfRootPath -MemberType NoteProperty -Value $temp.RootPath 
		$obj | Add-Member -Name Circular -MemberType NoteProperty -Value $temp.Circular 
		$obj | Add-Member -Name LastStoreHungTime -MemberType NoteProperty -Value ([System.DateTime]::MinValue)
		$obj | Add-Member -Name LastCheckTime -MemberType NoteProperty -Value ([System.DateTime]::MinValue)
		$obj | Add-Member -Name CheckAgainstTime -MemberType NoteProperty -Value $Script:ScriptStartTime
		$obj | Add-Member -Name PendingWaitTime -MemberType NoteProperty -Value $false
		$obj | Add-Member -Name WaitTime -MemberType NoteProperty -Value ([System.DateTime]::Now)
#>

Function Get-StorageHungEventsManager {
param(
[Parameter(Mandatory=$true)][Array]$Servers_List
)
	foreach($server in $Servers_List)
	{
		#Not going to check the server that is pending the wait time 
		if($server.PendingWaitTime){continue;}
		$storeHungEvents = Get-StoreHungEvents -Start_Time ($server.CheckAgainstTime) -Machine_Name ($server.ServerName)
		$server.LastCheckTime = [System.DateTime]::Now
		if($storeHungEvents -eq $null) {continue; } #Nothing really needs to be done at this point 
		elseif($storeHungEvents -eq $true)
		{
			#Need to update the time to check, because something happened, but it wasn't the issue we were looking for 
			$server.CheckAgainstTime = [System.DateTime]::Now
			continue; 
		}
		else
		{
			#Now an issue has occurred.... 
			#First thing is to update the times... 
			$server.CheckAgainstTime = [System.DateTime]::Now #Not the best way...because more events could occur in between so we are going to update this once we start up experfwiz 
			$server.LastStoreHungTime = [System.DateTime]::Now
			$server.PendingWaitTime = $true
			$server.WaitTime = ([System.DateTime]::Now).AddMinutes(5)
			$Message_Subject = ("Issue Occurred on server {0} @ {1}" -f $server.ServerName,$server.LastStoreHungTime)
			$eventData = ""
			foreach($event in $storeHungEvents)
			{
				$eventData += ("DatabaseGuid: {0} `r`nDatabaseInstanceName: {1} `r`nTag: {2} `r`nComponent: {3}`r`n`r`n" -f $event.event.userdata.eventxml.DatabaseGuid, $event.event.userdata.eventxml.DatabaseInstanceName, $event.event.userdata.eventxml.Tag, $event.event.userdata.eventxml.Component)
			}

			$Message_Body = (@"
An Issue has occurred and we should have the correct data being collected. At {0} the Experfwiz Data will stop and a new one will start. 

Issue Server: {1}
Data File: {2}
Events: 
{3}

"@ -f $server.WaitTime,$server.ServerName,$server.ExperfFullPath,$eventData)
			
			Send-Message -SMTP_Sender $Script:mailObjects.Sender -SMTP_Recipient $Script:mailObjects.Recipient -Message_Subject $Message_Subject -SMTP_Server $Script:mailObjects.Server -Creds $Script:mailObjects.Creds -Message_Body $Message_Body -Port 2525
			
		}

	}

}


Function Get-StoreHungEvents {
param(
[Parameter(Mandatory=$true)][System.DateTime]$Start_Time,
[Parameter(Mandatory=$true)][string]$Machine_Name 
)

    $sStart_Time = $Start_Time.ToString("o") #2017-06-02T13:56:41.7201028-07:00 <-- this is what it changes the Date Time to 
    $hash = @{LogName="Microsoft-Exchange-MailboxDatabaseFailureItems/Operational";StartTime=$sStart_Time}
    $events = Get-EventLogs -ComputerName $Machine_Name -FilterHashTable $hash 

    if($events -ne $null -and
     $events -ne "Err")
    {
        $StoreHangEvents = @()
        $mainIssue = $false
        foreach($event in $events)
        {
            $xml_Event = [xml]$event.Toxml()
            if($xml_Event.Event.UserData.EventXml.Component -eq "StoreService")
            {
                #this should just check for it once vs trying to set it to true over and over again as we are going to always determine an issue has occurred, regardless of the tag type at this point 
                if($mainIssue -eq $false -and 
                $xml_Event.Event.UserData.EventXml.Tag -eq "38")
                {
                    $mainIssue = $true
                }
                
                $StoreHangEvents += $xml_Event
            }
        }

        #if no store service events are found, then we want to return a value to update the start time
        #else we just want to return the events 
        if($StoreHangEvents -eq $null)
        {
            return $true
        }
        else
        {
            return $StoreHangEvents
        }

    }
    #Need to determine how we want to handle the errors for now, just a null return should be fine
    return $null

}


Function Start-Experfwiz {
param(
[Parameter(Mandatory=$true)][string]$Machine_Name,
[Parameter(Mandatory=$true)][string]$Script_Directory, 
[Parameter(Mandatory=$true)][string]$script_Name,
[Parameter(Mandatory=$true)][string]$Directory_to_Save_data,
[Parameter(Mandatory=$true)][int]$interval,
[Parameter(Mandatory=$true)][int]$Max_Size
)
    $org_Location = (Get-Location).Path
    Set-Location $Script_Directory
    $run_me = ".\" + $script_Name + " -filepath " + $Directory_to_Save_data + " -interval " + $interval + " -maxsize " + $Max_Size + " -Server " + $Machine_Name + " -quiet -circular" 
    $results = Invoke-Expression $run_me 
    Set-Location $org_Location
}





Function Get-LogmanData {
param(
[Parameter(Mandatory=$true)][string]$Logman_Name,
[Parameter(Mandatory=$true)][string]$Server_Name
)

    Function Get-LogmanColumnInfo {
    param(
    [Parameter(Mandatory=$true)][array]$rawData,
    [Parameter(Mandatory=$true)][string]$Column_Name
    )
        $i = 0
        while($i -lt $rawData.Count)
        {
            if($rawData[$i].Contains($Column_Name))
            {
                break; 
            }
            $i++
        }
        if($i -ge $rawData.Count)
        {
            return $null
        }

        return $rawData[$i].Substring($rawData[$i].LastIndexOf(" ") + 1)

    }

    $rawResults = logman -s $Server_Name $Logman_Name
    if($rawResults[$rawResults.count -1].Contains("Set was not found.") -or $rawResults[$rawResults.count -1].Contains("The network path was not found."))
    {
        return $null
    }
    else
    {
        $objLogMan = New-Object -TypeName PSObject 
        $objLogMan | Add-Member -Name LogmanName -MemberType NoteProperty -Value $Logman_Name
        $objLogMan | Add-Member -Name Status -MemberType NoteProperty -Value (Get-LogmanColumnInfo -rawData $rawResults -Column_Name "Status:")
        $objLogMan | Add-Member -Name RootPath -MemberType NoteProperty -Value (Get-LogmanColumnInfo -rawData $rawResults -Column_Name "Root Path:")
        $objLogMan | Add-Member -Name FullPath -MemberType NoteProperty -Value (Get-LogmanColumnInfo -rawData $rawResults -Column_Name "Output Location:")
        $objLogMan | Add-Member -Name Circular -MemberType NoteProperty -Value (Get-LogmanColumnInfo -rawData $rawResults -Column_Name "Circular:")
        $objLogMan | Add-Member -Name StartDate -MemberType NoteProperty -Value (Get-LogmanColumnInfo -rawData $rawResults -Column_Name "Start Date:")
        return $objLogMan
    }

}


Function Test-ScriptLoggingDependencies 
{
    if($Experfwiz_Directory -eq "."){$Script:Experfwiz_Directory = (Get-Location).Path}
    $fullExperfwizPath = if($Experfwiz_Directory.EndsWith("\")){$Experfwiz_Directory + $Experfwiz_Script_Name} else{$Experfwiz_Directory + "\" + $Experfwiz_Script_Name}
    if( -not (Test-Path $fullExperfwizPath))
    {
        #script isn't there so we should just exit 
        Write-Warning ("The expefwiz script {0} isn't at the location {1}. Stopping the script" -f $Experfwiz_Script_Name, $Experfwiz_Directory)
        exit
    }
    else
    {
        #We found the script, now to test the save directory location. If doesn't exist, create it 
        if($Experfwiz_Save_Data_Directory -eq "."){$Script:Experfwiz_Save_Data_Directory = (Get-Location).Path}
        Create-Folder -Checker $Experfwiz_Save_Data_Directory -DisplayInfo
    }

}


Function Get-SendMailMessageDependencies 
{
	<#
		Need to get the following 
		>To Address 
		>From Address 
		>Creds 
		>Server that we want to send the mail to
		
	Return object of those values 
	#>
	#The To Address
	$mailObject = New-Object -TypeName PSObject 
	Write-Host ""
	Write-Host "Please provide the SMTP address of the sender you want to use for the script to send off messages as."
	Write-Host "Example: zelda01@contoso.com"
	do{
		$temp_Sender = Read-Host "SMTP Address: "
		try 
		{
			Write-Host "Checking to see if the sender is valid..." -NoNewline
			Get-Mailbox $temp_Sender -ErrorAction Stop | Out-Null
			Write-Host "Looks good" -ForegroundColor Green
			$mailObject | Add-Member -Name Sender -MemberType NoteProperty -Value $temp_Sender
			break;
		}
		catch
		{
			Write-Host "Failed." -ForegroundColor Red
		}
	}while($true)

	#Who we want to send the message to 
	Write-Host ""
	Write-Host "Please provide the SMTP address of the recipient that you want to receive the message that the script sends."
	Write-Host "Example: Admin@contoso.com"

	do{
		$temp_Recipient = Read-Host "SMTP Address: "
		try
		{
			Write-Host "Checking to see if the recipient is valid...." -NoNewline
			Get-Mailbox $temp_Recipient -ErrorAction stop | Out-Null
			Write-Host "Looks good" -ForegroundColor Green
			$mailObject | Add-Member -Name Recipient -MemberType NoteProperty -Value $temp_Recipient
			break;
		}
		catch
		{
			Write-Host "Failed." -ForegroundColor Red
		}
	}while($true)

	#The Exchange Server that we want to send the message to
	Write-Host ""
	Write-Host "Please provide the Exchange 2013/2016 mailbox server that you want to receive the message that the script sends."
	Write-Host "Example: E2K13AIO1.contoso.local"
	do
	{
		$temp_Server = Read-Host "SMTP Server: "
		try
		{
			Write-Host "Checking to see if the Exchange Server is valid..." -NoNewline
			Get-ExchangeServer $temp_Server -ErrorAction Stop | Out-Null
			Write-Host "Looks good" -ForegroundColor Green
			$mailObject | Add-Member -Name Server -MemberType NoteProperty -Value $temp_Server
			break;
		}
		catch
		{
			Write-Host "Failed." -ForegroundColor Red
		}
	}while($true)
	$orgErrorActionPreference = $ErrorActionPreference
	$ErrorActionPreference = "Stop"
	#now we need to test and set the creds 
	$temp_Creds = $null
	Write-Host ""
	Write-Host "Now we need credentials of a user to submit the message."
	do
	{
		Write-Host "Note: this needs to be the creds of: " -NoNewline
		Write-Host $temp_Sender -ForegroundColor Green
		try
		{
			do
			{
				$r = Read-Host "Do you want to restart the script? ('y' or 'n'): "
			}while($r -ne 'y' -and $r -ne 'n')
			if($r -eq 'y'){exit}
			$temp_Creds = Get-Credential
			$mailObject | Add-Member -Name Creds -MemberType NoteProperty -Value $temp_Creds
			break;
		}
		catch
		{
			#Nothing needs to occur
		}
	}while($true)

	Write-Host "Testing the passed information to see if we can send a message..." -NoNewline

	if(Send-Message -SMTP_Sender $temp_Sender -SMTP_Recipient $temp_Recipient -SMTP_Server $temp_Server -Message_Subject "Test Message to see if creds work" -Message_Body "Test Message only" -Creds $temp_Creds -Port 2525)
	{
		Write-Host "Failed." -ForegroundColor Red
		Write-Warning "Looks like we failed to send the message. Review the error to see how you should correct this."
		$ErrorActionPreference = $orgErrorActionPreference
		exit
	}
	else
	{
		Write-Host "Worked!" -ForegroundColor Green
	}
	$ErrorActionPreference = $orgErrorActionPreference
	return $mailObject
}

Function Start-ExperfwizManager {
param(
[Parameter(Mandatory=$true)][Array]$Server_List
)
	foreach($server in $Server_List)
	{
		Start-Experfwiz -Machine_Name $server -Script_Directory $Experfwiz_Directory -script_Name $Experfwiz_Script_Name -Directory_to_Save_data $Experfwiz_Save_Data_Directory -interval $Experfwiz_Interval -Max_Size $Experfwiz_Data_Maxsize 
	}

}

Function Build-MonitorObject {
param(
[Parameter(Mandatory=$true)][Array]$Server_List
)
	$monitorObject = @() 
	foreach($server in $Server_List)
	{
		$obj = New-Object -TypeName PSObject 
		$temp = Get-LogmanData -Logman_Name "Exchange_Perfwiz" -Server_Name $server
		$obj | Add-Member -Name ServerName -MemberType NoteProperty -Value $server
		$obj | Add-Member -Name ExperfStatus -MemberType NoteProperty -Value $temp.Status
		$obj | Add-Member -Name ExperfCheckTime -MemberType NoteProperty -Value ([System.DateTime]::Now)
		$obj | Add-Member -Name ExperfFullPath -MemberType NoteProperty -Value $temp.FullPath
		$obj | Add-Member -Name ExperfRootPath -MemberType NoteProperty -Value $temp.RootPath 
		$obj | Add-Member -Name Circular -MemberType NoteProperty -Value $temp.Circular 
		$obj | Add-Member -Name LastStoreHungTime -MemberType NoteProperty -Value ([System.DateTime]::MinValue)
		$obj | Add-Member -Name LastCheckTime -MemberType NoteProperty -Value ([System.DateTime]::MinValue)
		$obj | Add-Member -Name CheckAgainstTime -MemberType NoteProperty -Value $Script:ScriptStartTime
		$obj | Add-Member -Name PendingWaitTime -MemberType NoteProperty -Value $false
		$obj | Add-Member -Name WaitTime -MemberType NoteProperty -Value ([System.DateTime]::Now)
		$monitorObject += $obj
	}

	return $monitorObject
}

<#
Function to display the information in a quick snapshot to verify that everything is still running smoothly 

Server            Experfwiz Status           Last Check Time             Last Store Hung Event Time 
----------------------------------------------------------------------------------------------------
Wingtip-E13A		Running					6/6/2017					1/1/0001 12:00:00 AM 
Wingtip-E13B		Running					6/6/2017					1/1/0001 12:00:00 AM 
Wingtip-E13C		Running					6/6/2017					1/1/0001 12:00:00 AM 

#>

Function Display-ObjectInformation 
{
param(
	[Parameter(Mandatory=$true)][array]$displayObject
)
	cls
	Write-Host ("{0,-18} {1,18} {2,36} {3,36}" -f "Server", "Experfwiz Check Time","Last Check Time", "Last Store Hung Event Time");	Write-Host "------------------------------------------------------------------------------------------------------------------"
	Foreach($obj in $displayObject)
	{
		if($obj.PendingWaitTime -eq $false)
		{
			Write-Host ("{0,-18} {1,18} {2,36} {3,36}" -f $obj.ServerName, $obj.ExperfCheckTime, $obj.LastCheckTime, $obj.LastStoreHungTime) -ForegroundColor Green
		}
		else
		{
			Write-Host ("{0,-18} {1,18} {2,36} {3,36}" -f $obj.ServerName, $obj.ExperfCheckTime, $obj.LastCheckTime, $obj.LastStoreHungTime) -ForegroundColor Yellow
		}
	}

}

Function Get-StoreHungPendingStop {
param(
[Parameter(Mandatory=$true)][Array]$Server_List
)
	foreach($server in $Server_List)
	{
		if($server.PendingWaitTime -eq $false)
		{
			continue;
		}
		else
		{
			if($server.WaitTime -lt ([System.DateTime]::Now))
			{
				#first get the logman info to be able to send the file to collect 
				$temp = Get-LogmanData -Logman_Name "Exchange_Perfwiz" -Server_Name $server.ServerName
				$message_Subject = ("Performance data ready to be collected on server {0}" -f $server.ServerName)
				$message_body = (@"
				It is okay to collect the file {0} from the server {1}
"@ -f $temp.FullPath, $server.ServerName)
				
				Start-Experfwiz -Machine_Name $server.ServerName -Script_Directory $Experfwiz_Directory -script_Name $Experfwiz_Script_Name -Directory_to_Save_data $Experfwiz_Save_Data_Directory -interval $Experfwiz_Interval -Max_Size $Experfwiz_Data_Maxsize
				Send-Message -SMTP_Sender $Script:mailObjects.Sender -SMTP_Recipient $Script:mailObjects.Recipient -Message_Subject $message_Subject -SMTP_Server $Script:mailObjects.Server -Creds $Script:mailObjects.Creds -Message_Body $message_body -Port 2525
				$temp = Get-LogmanData -Logman_Name "Exchange_Perfwiz" -Server_Name $server.ServerName
				$server.ExperfStatus = $temp.Status
				$server.ExperfFullPath = $temp.FullPath
				$server.ExperfRootPath = $temp.RootPath 
				$server.Circular = $temp.Circular 
				$server.PendingWaitTime = $false
				$server.CheckAgainstTime = [System.DateTime]::Now
				$server.ExperfCheckTime = [System.DateTime]::Now
			}
		}
	}
}

Function Check-ExperfwizRunning {
param(
[Parameter(Mandatory=$true)][Array]$Server_List
)
	foreach($server in $Server_List)
	{
		if($server.PendingWaitTime -eq $false -and
			$server.ExperfCheckTime -lt ([System.DateTime]::Now).AddMinutes(-5))
		{
			$temp =  Get-LogmanData -Logman_Name "Exchange_Perfwiz" -Server_Name $server.ServerName
			if($temp.Status -ne "Running" -and $temp.Status -ne $null)
			{
				Start-Experfwiz -Machine_Name $server.ServerName -Script_Directory $Experfwiz_Directory -script_Name $Experfwiz_Script_Name -Directory_to_Save_data $Experfwiz_Save_Data_Directory -interval $Experfwiz_Interval -Max_Size $Experfwiz_Data_Maxsize
				$temp =  Get-LogmanData -Logman_Name "Exchange_Perfwiz" -Server_Name $server.ServerName
				$server.ExperfStatus = $temp.Status
				$server.ExperfFullPath = $temp.FullPath
				$server.ExperfRootPath = $temp.RootPath 
				$server.Circular = $temp.Circular 
				$server.PendingWaitTime = $false
				$server.CheckAgainstTime = [System.DateTime]::Now
			}
			elseif($temp.FullPath -eq $server.ExperfFullPath)
			{
				$server.ExperfFullPath = $temp.FullPath
			}
			$server.ExperfCheckTime = [System.DateTime]::Now
		}
	}

}


Function Main {
    
    if(-not (Is-Admin))
    {
        Write-Warning "Please run the script as an Administrator"
        exit
    }
    Load-ExShell
	Test-ScriptLoggingDependencies
	$Script:mailObjects = Get-SendMailMessageDependencies

	$Script:WatchServersList = Get-WatchingServers -Server_Name ($env:COMPUTERNAME)
	Start-ExperfwizManager -Server_List $Script:WatchServersList
	$Script:ScriptStartTime = [System.DateTime]::Now
	$Script:mainObject = Build-MonitorObject -Server_List $Script:WatchServersList
	do{
		Get-StorageHungEventsManager -Servers_List $Script:mainObject
		Display-ObjectInformation -displayObject $Script:mainObject
		Get-StoreHungPendingStop -Server_List $Script:mainObject
		Check-ExperfwizRunning -Server_List $Script:mainObject
		sleep 2; #allows for the process to rest so we aren't consuming resources 
	}while($true)
	
}

Main 