<#
.SYNOPSIS
    Invoke citrix health check (Citrix Cloud DaaS)

.DESCRIPTION
    Citrix Healtch check based on REST API of Citrix DaaS, More infos https://developer.cloud.com/citrixworkspace/citrix-daas/citrix-daas-rest-apis/docs/overview

    History:
    - 7.9.2022, Sacha Thomet (sachathomet.ch) First version, (Adoption of Script framework from Jonathan Aebischer from ACE) for new Citrix HealthCheck 

#>

$ErrorActionPreference = "Stop"

#region Variables
# Names of the Delivery Groups for which the tasks are executed
#$deliveryGroups = @("Windows-VDI-pooled","Windows-pesistent-pooled")
$deliveryGroups = @("YourDeliveryGroup1","YourDeliveryGroup2")#PROD
$ShowOnlyErrorVDI = $true # If this set on $true, SingleSession machines only get reported when there is a Problem. 
# Advanced Checks are for an in-deepth check of the VDA which needs connection from where the Script runs to the VDA via WMI, it takes a lot more time as the regular check
$AdvancedCheck = $true
# Define how long a VDA should be up at maximum
$maxUpTimeDays = 60
# Citrix Cloud URL
$citrixCloudBaseUrl = "https://api-eu.cloud.com" # API region
# Citrix Cloud Customer ID
$customerId = "z12345z12abc"
# Citrix Cloud Site ID (if empty, it will be determined via API)
$siteId = "1ab2c456-de78-9f1g-hi12-..."
# Client Credentials ID
$clientId = "a12c34b5-6789-1234-cde5-678a9b333cd"
# The secret is stored in this file. It must only be readable by server admins (+ system account)!
$pathToClientSecretFile = "C:\Scripts\CitrixClientSecret\APIClientSecret.txt"
#$pathToClientSecretFile = "C:\Scripts\ScriptDev\CitrixMaintenance\APIClientSecret.txt"
# A UserAgent must be specified for the query to function. 
$UserAgent = "Mozilla/5.0"
# Path to the logfile
$pathToLogfile = "C:\Scripts\Log\Citrix_HealthCheck.log"
#$pathToLogfile = "C:\temp\ScriptDev\CitrixMaintenance\Log\Citrix_HealthCheck.log"
# Max. Grösse des Logfiles
$maxSizeLogfile = 2MB
# Proxy server (leer lassen = User default)
$proxyServer = ""
# Exit code
$exitCode = 0
#endregion

# email variables
$CheckSendMail = 1 # Turn on or Off the sending of an email (1=on / 0=off)
$emailFrom ="from@domain.com"
$emailTo = "to@domain.com"
$emailCC = "cc@domain.com"
$emailSubject = "Citrix DaaS Report "
$emailPrio = "high"
$smtpServer = "smtp.domain.com"
$smtpEnableSSL = $false


#region HTML Result variables
$ReportDate = (Get-Date -UFormat "%A, %d. %B %Y %R")
$currentDir = Split-Path $MyInvocation.MyCommand.Path
$outputpath = Join-Path $currentDir "" #add here a custom output folder if you wont have it on the same directory
$resultsHTM = Join-Path $outputpath ("Citrix-DaaS-HealthCheck.htm") #add $outputdate in filename if you like


#Header for Table "VDI Checks" Get-BrokerMachine
$VDIfirstheaderName = "VDA"
$VDIHeaderNames = "CatalogName","DeliveryGroup","PowerState", "VDAHealthScore", "MaintMode","Uptime", "Sessions","LastConnect", 	"RegState","VDAVersion","AssociatedUserNames",  "AllocationType", "SessionSupport", "Tags", "HostedOn"
$VDIHeaderWidths = "4",          "4",		    "4",          "4", 	            "4", 		"4", 		 "4",           "4", 		"4",			  "4","4"
$VDItablewidth = 1600

#region delete old HTML file:
Remove-Item $resultsHTM -force -EA SilentlyContinue
#endregion


#region multiple used functions 
#==============================================================================================
Function writeHtmlHeader
{
param($title, $fileName)
$date = $ReportDate
$head = @"
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<title>$title</title>
<STYLE TYPE="text/css">
<!--
td {
font-family: Tahoma;
font-size: 11px;
border-top: 1px solid #999999;
border-right: 1px solid #999999;
border-bottom: 1px solid #999999;
border-left: 1px solid #999999;
padding-top: 0px;
padding-right: 0px;
padding-bottom: 0px;
padding-left: 0px;
overflow: hidden;
}
body {
margin-left: 5px;
margin-top: 5px;
margin-right: 0px;
margin-bottom: 10px;
table {
table-layout:fixed;
border: thin solid #000000;
}
-->
</style>
</head>
<body>
<table width='1600'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='48' align='center' valign="middle">
<font face='tahoma' color='#003399' size='4'>
<strong>$title - $date</strong>
<br>
</font>
</td>

</table>
<table width='1600'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#000000' size='2'>

<strong>Infrastructure HealthScore $([math]::Round($($InfraHealthScore/$VDACounter),2 )) </strong>

</font>
</td>
</table>
</tr>
<br>
</table>
"@
$head | Out-File $fileName
}
  
# ==============================================================================================
Function writeTableHeader
{
param($fileName, $firstheaderName, $headerNames, $headerWidths, $tablewidth)
$tableHeader = @"
  
<table width='$tablewidth'><tbody>
<tr bgcolor=#CCCCCC>
<td width='6%' align='center'><strong>$firstheaderName</strong></td>
"@
  
$i = 0
while ($i -lt $headerNames.count) {
$headerName = $headerNames[$i]
$headerWidth = $headerWidths[$i]
$tableHeader += "<td width='" + $headerWidth + "%' align='center'><strong>$headerName</strong></td>"
$i++
}
  
$tableHeader += "</tr>"
  
$tableHeader | Out-File $fileName -append
}
  
# ==============================================================================================
Function writeTableFooter
{
param($fileName)
"</table><br/>"| Out-File $fileName -append
}
  
#==============================================================================================
Function writeData
{
param($data, $fileName, $headerNames)

$tableEntry  =""  
$data.Keys | Sort-Object | ForEach-Object {
$tableEntry += "<tr>"
$computerName = $_
$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")
#$data.$_.Keys | foreach {
$headerNames | ForEach-Object {
#"$computerName : $_" | LogMe -display
try {
if ($data.$computerName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
elseif ($data.$computerName.$_[0] -eq "WARNING") { $bgcolor = "#FF7700"; $fontColor = "#FFFFFF" }
elseif ($data.$computerName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF" }
else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
$testResult = $data.$computerName.$_[1]
}
catch {
$bgcolor = "#CCCCCC"; $fontColor = "#003399"
$testResult = ""
}
$tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")
}
$tableEntry += "</tr>"
}
$tableEntry | Out-File $fileName -append
}
  
# ==============================================================================================
Function writeHtmlFooter
{
param($fileName)
@"
</table>
<table width='1600'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='left'>
<font face='courier' color='#000000' size='2'>

<strong>Uptime Threshold: </strong> $maxUpTimeDays days <br>
<strong>HypervisorConnectionstate: </strong> $HVCS <br>
<strong>Checked $VDACounter VDA's (some of them are probably not visible in report above becasue all is ok)</strong>

</font>
</td>
</table>
</body>
</html>
"@ | Out-File $FileName -append
}

# ==============================================================================================

function Ping([string]$hostname, [int]$timeout = 200) {
$ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
try { $result = $ping.send($hostname, $timeout).Status.ToString() }
catch { $result = "Failure" }
return $result
}
#==============================================================================================

$wmiOSBlock = {param($computer)
  try { $wmi=Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop }
  catch { $wmi = $null }
  return $wmi
}

#endregion


#region Code Snippet: create client secret
<#
"<ClientSecret>" | Out-File $pathToClientSecretFile -Encoding utf8
$acl = Get-Acl -Path $pathToClientSecretFile
$accessRule_Admins = New-Object System.Security.AccessControl.FileSystemAccessRule("BUILTIN\Administrators","FullControl","Allow")
$accessRule_System = New-Object System.Security.AccessControl.FileSystemAccessRule("NT AUTHORITY\SYSTEM","FullControl","Allow")
# Vererbung deaktivieren
$acl.SetAccessRuleProtection($true,$false)
$acl.SetAccessRule($accessRule_Admins)
$acl.SetAccessRule($accessRule_System)
$acl | Set-Acl -Path $pathToClientSecretFile
#>
#endregion

#region Helpers
function Write-Log {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Text,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Info","Error","Warning","Success")]
        [string]$Type,

        [Parameter(Mandatory=$false)]
        [string]$Logfile,

        [Parameter(Mandatory=$false)]
        [int]$MaxSizeOfLogfile=2MB,

        [Parameter(Mandatory=$false)]
        [switch]$WriteTextToConsole
    )

    $textToLog = $Text

    # Add type
    switch ($Type) {
		"Error" {
			$textToLog = "[error  ] $($textToLog)"
            $foregroundColor = [System.ConsoleColor]::Red
		}

		"Warning" {
			$textToLog = "[warning] $($textToLog)"
            $foregroundColor = [System.ConsoleColor]::Yellow
		}

		"Success" {
			$textToLog = "[success] $($textToLog)"
            $foregroundColor = [System.ConsoleColor]::DarkGreen
		}

		"Info" {
			$textToLog = "[info   ] $($textToLog)"
		}
	}

    # Add date
    $textToLog = "$(Get-Date -format G) - $($textToLog)"

    if (-not [string]::IsNullOrEmpty($Logfile)) {
        if (Test-Path $Logfile) {
            $logfileItem = (Get-Item $Logfile)
            if ($logfileItem.Length -gt $MaxSizeOfLogfile) {
                # Logfile is too large => rename logfile and overwrite existing old logfile
                $pathToOldLogfile = Join-Path $logfileItem.DirectoryName -ChildPath "$($logfileItem.BaseName).old$($logfileItem.Extension)"
                # Rename logfile and overwrite existing old logfile
                Move-Item -Path $Logfile -Destination $pathToOldLogfile -Force
            }
        }

        if (-not (Test-Path $Logfile)) {
            $null = New-Item -Path $Logfile -ItemType File -Force
        }

        $textToLog | Out-File -FilePath $Logfile -Encoding utf8 -Append -Force
    }

    if ($WriteTextToConsole.IsPresent) {
        if($foregroundColor){
            Write-Host -Object $textToLog -ForegroundColor $foregroundColor
        } else {
            Write-Host -Object $textToLog
        }
    }

}

function Remove-IllegalCitrixCharacters {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    $separators = @('/', ';', ':', '#', '.', '*', '?', '=', '<', '>', '|', '[', ']', '(', ')', '"', "'", '\', '`');
    $temp = $Text.Split($separators, [System.StringSplitOptions]::RemoveEmptyEntries)
    return ($temp -join "")
}


#endregion

# region beginning of the actual main script

#region delete old HTML file:
Remove-Item $resultsHTM -force -EA SilentlyContinue
#endregion

try {
    Write-Log -Text "Starting script" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
    
   
    #region Read client secret
    Write-Log -Text "Reading client secret from [$($pathToClientSecretFile)]" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
    $clientSecret = Get-Content -Path $pathToClientSecretFile -Encoding UTF8
    #endregion

    #region Proxy server
    $additionalRestParameters = @{}
    if (-not [string]::IsNullOrEmpty($proxyServer)) {
        $additionalRestParameters += @{
            Proxy = $proxyServer
        }   
    }
    #endregion

    #region Get token
    Write-Log -Text "Getting access token" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
    $tokenUrl = "$($citrixCloudBaseUrl)/cctrustoauth2/root/tokens/clients"
    $response = Invoke-RestMethod $tokenUrl -Method POST @additionalRestParameters -Body @{
      grant_type = "client_credentials"
      client_id = $clientId
      client_secret = $clientSecret
    }
    # http header for api calls
    $headers = @{
        Authorization = "CwsAuth Bearer=$($response.access_token)"
    }
    #endregion

    if ([string]::IsNullOrEmpty($siteId)) {
        #region Get site id
        Write-Log -Text "Getting siteId" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        $headers["Citrix-CustomerId"] = $customerId
        $response = Invoke-RestMethod "$($citrixCloudBaseUrl)/cvad/manage/Me" -Method Get -Headers $headers @additionalRestParameters
        $siteId = ($response.Customers | Where-Object Id -eq $customerId).Sites.Id
        Write-Log -Text "Found siteId: $($siteId)" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        #endregion
    }

    #region Set HTTP headers
    $headers["Citrix-CustomerId"] = $customerId
    $headers["Citrix-InstanceId"] = $siteId
    $headers["User-Agent"] = $UserAgent
    
    #endregion


    $allResults = @{}  #hashtable for all the results which will be put in HTML later
    $VDACounter = 0 # counts all VDAs
    $InfraHealthScore = 0 # A value that should indicated the checked infrastructure health over all checks
   

   #region Process machines in Delivery Groups defined in $deliveryGroups (Line18 .. or so ... )
   foreach ($deliveryGroupName in $deliveryGroups) {
    try {
        Write-Log -Text "$($deliveryGroupName) - Start processing" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        
        #region Get delivery group
        try {
            $deliveryGroup = Invoke-RestMethod "$($citrixCloudBaseUrl)/cvad/manage/DeliveryGroups/$($deliveryGroupName)" -Headers $headers -Method Get @additionalRestParameters
        } catch {
            # Status Code 404 -> Not Found
            if ($_.Exception.Response.StatusCode.value__ -eq 404) {
                throw "Cannot find delivery group [$($deliveryGroupName)]"
            } else {
                throw $_
            }
        }
        #endregion

        #region Machines from delivery group
        Write-Log -Text "$($deliveryGroupName) - Getting machines..." -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        $continuationTokenQueryParameter = ""
        $machines = @()
        do {
            # Per default max. 250 machines are returned per call
            # If more than the specified limit of machines are available the response will have a ContinuationToken that can be passed to get the next batch of results
            $response = Invoke-RestMethod "$($citrixCloudBaseUrl)/cvad/manage/DeliveryGroups/$($deliveryGroup.Id)/Machines$($continuationTokenQueryParameter)" -Headers $headers -Method Get @additionalRestParameters
            if (-not [string]::IsNullOrEmpty($response.ContinuationToken)) {
                $continuationTokenQueryParameter = "?continuationToken=$($response.ContinuationToken)"
            }
            $machines += $response.Items
        } while (-not [string]::IsNullOrEmpty($response.ContinuationToken))
        #endregion

      
        $counter = 0
       
        Write-Log -Text "$($deliveryGroupName) - Found $($machines.Count) machines. Proceed this Delivery Group" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        
        foreach ($machine in ($machines | Sort-Object Name)) {
            $counter++
            $tests = @{} #Array with the collected data of every VDA
            
            $VDAHealthScore = 0
            
            try {
                Write-Progress -Activity "$($deliveryGroupName) - $($machine.Name)" -PercentComplete ($counter / $machines.Count * 100)
                
            
            $singleMachine = Invoke-RestMethod "$($citrixCloudBaseUrl)/cvad/manage/Machines/$($machine.Id)" -Headers $headers -Method Get @additionalRestParameters
            Write-Log -Text "Proceed $($singleMachine.DnsName) " -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole

            if ( $singleMachine.Tags -like "*excludeFromReport*" ) 
            {
                Write-Log -Text "Do nothing with  $($machine.Name) because of one of this Tags: $($machine.Tags)" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole 
            }
            else {

                $VDACounter ++ #count all checked VDAs
                $machineDNS = $singleMachine.DnsName
                Write-Log -Text "$($singleMachine.DnsName) has Powerstate $($singleMachine.PowerState)" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
                
                # Find out if Machine is powered Off or On. All Subchecks only make sense on powered On machines
                
                if ($singleMachine.PowerState -eq "Unknown")
                    {
                    $tests.PowerState =  "WARNING", "$($singleMachine.PowerState)"
                    $VDAHealthScore = $VDAHealthScore -10
                    }

                elseif ($singleMachine.PowerState -eq "Off")
                {
                 $tests.PowerState =  "NEUTRAL", "$($singleMachine.PowerState)"
                
                 }

                    elseif ($singleMachine.PowerState -eq "On")
                    {
                    $tests.PowerState =  "SUCCESS", "$($singleMachine.PowerState)"
                                        
                        # Just check for Registered if the machine is powerd on, otherwise its normal that its unregistered 
                        if ($singleMachine.RegistrationState -eq "Registered") {
                            $tests.RegState = "SUCCESS", "$($singleMachine.RegistrationState)" 

                            #AdvancedCheck starts here, machine has to be up and registered
                            if ($AdvancedCheck -eq $true) {
                                Write-Log -Text "Start AdvancedCheck for $($singleMachine.DnsName) " -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
                                #here later can be added more checks which makes sense on a powered on & registered machine




                                # Column Ping Desktop
                                $result = Ping $machineDNS 100
                                if ($result -eq "SUCCESS") {
                                $tests.Ping = "SUCCESS", $result
                                #==CHECKS OVER WMI============================================================================================
                                # Column Uptime (Query over WMI - only if Ping successfull)
                                $tests.WMI = "ERROR","Error"
                                $job = Start-Job -ScriptBlock $wmiOSBlock -ArgumentList $machineDNS
                                $wmi = Wait-job $job -Timeout 15 | Receive-Job

                                # Perform WMI related checks
                                if ($null -ne $wmi) {
                                    $tests.WMI = "SUCCESS", "Success"
                                    $LBTime=[Management.ManagementDateTimeConverter]::ToDateTime($wmi.Lastbootuptime)
                                    [TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)
                                
                                    if ($uptime.days -gt $maxUpTimeDays) {
                                    Write-Log -Text "reboot warning, last reboot: $($LBTime)" -Type WARNING -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
                                    $tests.Uptime = "WARNING", $uptime.days
                                    
                                    $VDAHealthScore = $VDAHealthScore -5
                                    } 
                                    
                                    else { 
                                    $tests.Uptime = "SUCCESS", $uptime.days 
                                    }
                                } else { 
                                    Write-Log -Text "WMI connection failed - check WMI for corruption" -Type WARNING -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
                                    stop-job $job
                                }
                                }

                                #-----------------
                                                            
                            }

                                   
                            
                            

                        }

                            

                        #if powered On and not registered, this is a evidence for a real problem on this machine
                        else { $tests.RegState = "ERROR", "$($singleMachine.RegistrationState)"
                                $VDAHealthScore = $VDAHealthScore - 30
                            }
                                
                    }
                

                # Maintmode can be a problem, but maybe there is a reasind - so just a warning for Maintmode On
                    if ($singleMachine.InMaintenanceMode) {
                        $tests.MaintMode = "WARNING", "ON" 
                        $VDAHealthScore = $VDAHealthScore -15


                    }
                    else { $tests.MaintMode = "SUCCESS", "OFF" }
                
                # A lot of infos which is provided by Citrix DaaS API 
                $tests.VDAVersion = "NEUTRAL", "$($singleMachine.AgentVersion)"
                $tests.DeliveryGroup = "NEUTRAL", "$($singleMachine.DeliveryGroup.Name)"
                $tests.CatalogName =  "NEUTRAL", "$($singleMachine.MachineCatalog.Name)"
                $tests.LastConnect =  "NEUTRAL", "$($singleMachine.LastConnectionTime)"
                $tests.AssociatedUserNames =  "NEUTRAL", "$($singleMachine.AssociatedUsers.SamName)"
                $tests.Tags =  "NEUTRAL", "$($singleMachine.Tags)"
                $tests.AllocationType =  "NEUTRAL", "$($singleMachine.AllocationType)"
                $tests.HostedOn =  "NEUTRAL", "$($singleMachine.Hosting.HostingServerName)"
                $tests.SessionSupport = "NEUTRAL", "$($singleMachine.SessionSupport)"
                $tests.Sessions = "NEUTRAL", "$($singleMachine.SessionCount)"

                
                # Multiplier for the negative HealthScore
                # if AllocationType is Random* duplicate the negative HealthScore (* = pooled VDIs and vApps Servers, shared ressources)
                if ($singleMachine.AllocationType -eq "Random") {
                    $VDAHealthScore = $($VDAHealthScore*2)
                }

                # if SessionSupport is MultiSession* again a duplication of the HealthScore (*= Win10 AVD Multiuser or vApps Servers)
                 if ($singleMachine.SessionSupport -eq "MultiSession") {
                     $VDAHealthScore = $VDAHealthScore*2
                 }

                # # Count HealthScore, starts with 100 for a VDA, and negative numbers will be added* in case of problems (* yes there is no substraction, you can only add negative values and the world is much simpler)
                $VDAHealthScore = 100 + $VDAHealthScore
                
                
                # Column VDAHealthScore in 3 Colors
                if ($VDAHealthScore -eq "100") {
                    $tests.VDAHealthScore =  "SUCCESS", "$($VDAHealthScore)"
                }
                elseif ($VDAHealthScore -gt 50) {
                    $tests.VDAHealthScore =  "WARNING", "$($VDAHealthScore)"
                } 
                else
                {
                    $tests.VDAHealthScore =  "ERROR", "$($VDAHealthScore)"
                }


                $InfraHealthScore = $InfraHealthScore + $VDAHealthScore

                # Add this array into the hashtable allRestults - one line for each VDA in the table (put the Array in the Hashtable)
                if ($singleMachine.SessionSupport -eq "MultiSession") {
                    # Add this line to allRestults - one line for each Multisession VDA (put the Array in the Hashtable)
                    $allResults.$machineDNS = $tests
                }

                else {
                    if ($ShowOnlyErrorVDI -eq $true -and $VDAHealthScore -eq 100) {
                        Write-Log -Text "$($singleMachine.DnsName) has no Problem only faulty should be showed - so heres nothing added to the table" -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
                    }
                    else {
                        $allResults.$machineDNS = $tests
                    }
                }

                
            }
                #endregion

               
            } catch {
                Write-Log -Text "$($deliveryGroupName) - $($machine.Name) - Failed to process this machine: $($_.Exception.Message)" -Type Error -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
            }
        }




    } catch {
        Write-Log -Text "$($deliveryGroupName) - Failed to process this delivery group: $($_.Exception.Message)" -Type Error -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
        $exitCode = 1
    }    
}
#endregion
} catch {
 Write-Log -Text "A general error occurred: $($_.Exception.Message)" -Type Error -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
 $exitCode = 1
}




Write-Host ("Saving results to html report: " + $resultsHTM)
writeHtmlHeader "Citrix DaaS Report" $resultsHTM

# Write Table with all Desktops

    writeTableHeader $resultsHTM $VDIFirstheaderName $VDIHeaderNames $VDIHeaderWidths $VDItablewidth
    $allResults | ForEach-Object{ writeData $allResults $resultsHTM $VDIHeaderNames }
    writeTableFooter $resultsHTM

    writeHtmlFooter $resultsHTM



    #Only Send Email if Variable in XML file is equal to 1
    if ($CheckSendMail -eq 1){

    #send email
    $emailMessage = New-Object System.Net.Mail.MailMessage
    $emailMessage.From = $emailFrom
    $emailMessage.To.Add( $emailTo )
    $emailMessage.CC.Add( $emailCC )
    $emailMessage.Subject = $emailSubject 
    $emailMessage.IsBodyHtml = $true
    $emailMessage.Body = (Get-Content $resultsHTM) | Out-String
    $emailMessage.Attachments.Add($resultsHTM)
    $emailMessage.Priority = ($emailPrio)
    
    $smtpClient = New-Object System.Net.Mail.SmtpClient( $smtpServer , $smtpServerPort )
    $smtpClient.EnableSsl = $smtpEnableSSL
    
    # If you added username an password, add this to smtpClient
    If ((![string]::IsNullOrEmpty($smtpUser)) -and (![string]::IsNullOrEmpty($smtpPW))){
        $pass = $smtpPW | ConvertTo-SecureString -key $smtpKey
        $cred = New-Object System.Management.Automation.PsCredential($smtpUser,$pass)
    
        $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($cred.Password)
        $smtpUserName = $cred.Username
        $smtpPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
    
        $smtpClient.Credentials = New-Object System.Net.NetworkCredential( $smtpUserName , $smtpPassword );
    }
    
    $smtpClient.Send( $emailMessage )
    
    
    
    }#end of IF CheckSendMail
    else{
    
        Write-Log -Text "Email sending skipped because CheckSendMail = 0"  -Type Info -Logfile $pathToLogfile -MaxSizeOfLogfile $maxSizeLogfile -WriteTextToConsole
    
    }#Skip Send Mail

Exit $exitCode