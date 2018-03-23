
function Get-CxdCallData{

    <#

        .SYNOPSIS

        Retrieves a series of data using Get-CsUserSession including:

        -Rate my Call data (all data is available including surveys closed without a response)
        -MediaLines data relevant to determine what endpoints are being used in a call
        -LastLogonDate which tells us if people have signed in and when
        -Client versions (normalized)
        -Failed calls data
        -Other useful data for understanding usage in the tenant


        .DESCRIPTION

        First, the Get-CxdCallData script will ask the user for credentials to authenticate to Skype for Business Online. If previous credentials are found we will ask to reuse them otherwise we'll ask for new ones. 

        Second, the script will ask if there is a CSV file for input. This file must have a header column of 'UserPrincipalName' as this is what we use to find enabled users in the tenant. You can use the format of 'lastname, firstname' or 'firstname lastname' or 'firstname.lastname@domain.com'. The script will handle searching for users matching the criteria and formatting them into a common UPN format to perform the search.

        Lastly, the script will output (by default) the results to an HTML file in the path you ran it from. The file name will include the current date and time incorporated into the file.

        For the CSV file, the following formats are supported:

        Firstname,Lastname,Userprincipalname
        Jason,Shave,jassha@microsoft.com

        or

        Userprincipalname
        jassha@microsoft.com

        or

        Userprincipalname
        Shave, Jason


        .EXAMPLE 

        Get-CxdCallData -NumberOfDaysToSearch 90 -CsvFileWithUsers C:\users\jason\documents\customerdata\users.csv -ReportSavePath C:\users\jason\documents\customerdata

        .EXAMPLE

        Get-CxdCallData -NumberOfDaysToSearch 180 -ReportSavePath C:\users\jason\documents -Verbose

        .NOTES

        You must be a Skype for Business Administrator in the Office 365 tenant to run the commands in this script. This script is provided without warranty. Although the commands and functions within this script do not make changes to your Office 365 tenant, you agree to use it at your own risk.

        Created by: Jason Shave (jassha@microsoft.com)
    #>

    [cmdletbinding()]
        Param
        (
            [Parameter(ParameterSetName="all",Mandatory=$true)][int]$NumberOfDaysToSearch,
            [Parameter(ParameterSetName="all",Mandatory=$true)][string]$ReportSavePath,

            [Parameter(ParameterSetName="all",Mandatory=$false)][string]$OverrideAdminDomain,
            [Parameter(ParameterSetName="all",Mandatory=$false)][PSCredential]$Credential,
            [Parameter(ParameterSetName="all",Mandatory=$false)][switch]$KeepExistingReports,

            [Parameter(ParameterSetName="all")]
            [Parameter(ParameterSetName="csvUsers",Mandatory=$false)][string]$CsvFileWithUsers,

            [Parameter(ParameterSetName="all")]
            [Parameter(ParameterSetName="userSearch",Mandatory=$false)][string]$UserToSearch,

            [Parameter(ParameterSetName="all")]
            [Parameter(ParameterSetName="csvUsers")]
            [Parameter(ParameterSetName="userSearch")]            
            [Parameter(ParameterSetName="subnet",Mandatory=$false)][switch]$MatchCallsToASubnet,

            [Parameter(ParameterSetName="all")]
            [Parameter(ParameterSetName="csvUsers")]
            [Parameter(ParameterSetName="userSearch")]            
            [Parameter(ParameterSetName="subnet",Mandatory=$true)][Net.IPAddress]$subnetAddress,

            [Parameter(ParameterSetName="all")]
            [Parameter(ParameterSetName="csvUsers")]
            [Parameter(ParameterSetName="userSearch")]            
            [Parameter(ParameterSetName="subnet",Mandatory=$true)][Net.IPAddress]$subnetMask
        )

    begin {
        try{
            Set-WinRMNetworkDelayMS -value "60000" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -ErrorVariable WinRMError
        }catch{
            Write-Warning -Message "Unable to set the WinRM Network Delay. This may be due to User Account Control settings. If you're not encountering timeouts you can safely ignore this warning, otherwise try running the script as an elevated user (Administrator)."
        }

        #remove all files and directories from the save path
        if ($KeepExistingReports){
            Write-Warning -Message "The -KeepExistingReports switch is being deprecated in favor of people removing their own content by other means."
        }

        Write-Host "Check out https://github.com/jasonshave/cxdcalldata for the latest readme and opportunity to improve this module. Community contribution has made this module considerably better so thanks to everyone who has contributed!" -ForegroundColor Green
    }

    process {
        $startTime = (Get-Date).AddDays(-$NumberOfDaysToSearch) #note the .AddDays is subtracting the number of days to search
        $endTime = Get-Date
        $global:sessionStartTime = $null
        $numBuckets = 0
        $numUsers = 0
        $enabledUsers = $null
        $arrNotUsingSkype = $null

        if (!$Credential){
            $Credential = Get-Credential -Message "Authenticate to Skype for Business Online"
        }
        #create initial SFBO Connection
        Invoke-CxdSkypeOnlineConnection

        #define the save path for reports
        $ReportSavePath = Set-CxdSavePath
        
        #process CSV file to strip out invalid entries and make sure formatting is correct
        if ($CsvFileWithUsers){
            $csvValidatedUsers = ProcessCsv
        }

        try{
            [array]$enabledUsers = ProcessSkypeOnlineUsers 
        }catch{
            Write-Warning -Message "Exception encounterd while getting users. Attempting to remove the PowerShell PSSession and will retry..."
            Invoke-CxdSkypeOnlineConnection -RepairPSSession
            [array]$enabledUsers = ProcessSkypeOnlineUsers
        }
        
        #exit if no users to process
        if (!$enabledUsers.Count){
            Write-Warning -Message "We didn't find any users matching your query. Exiting..."
            break
        }

        #get all tenant domains to help parse federated communications report
        $tenantDomains = Get-CsTenant | Select-Object DomainUrlMap
        $tenantDomains = $tenantDomains.DomainUrlMap | Foreach-Object {"*" + $_ + "*"}
        $tenantDomains += "*.lync.com*"
        
        #calculate users and buckets
        $userTotal = $enabledUsers.Count
        $userBuckets = [math]::Ceiling($userTotal /10)
        $arrUserBuckets = @{}

        #create users and put them into buckets
        Write-Verbose -Message "Putting $($enabledUsers.Count) users into buckets..."
        $count = 0
        $enabledUsers | ForEach-Object {
            $arrUserBuckets[$count % $userBuckets] += @($_)
            $count++
        }
        Write-Verbose -Message "Placed $($enabledUsers.Count) users into $($arrUserBuckets.Count) bucket(s)."

        #process all discovered users into buckets
        $pStart = Get-Date
        $arrUserBuckets | ProcessBuckets

        #process federated summary data
        FederatedCommunicationsSummary
        #UserSummary
    }
    
    end{
        Write-Verbose -Message "Removing PowerShell Session..."
        Get-PSSession | Remove-PSSession
    }   
}

function ProcessBuckets{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]$arrUserBuckets
    )
    begin{}
    process{
        foreach ($userItem in $arrUserBuckets.Values){

            #progress bar for total buckets
            $numBuckets++
            $bId = 1
            $bActivity = "Processing " + $enabledUsers.Count + " users into " + $arrUserBuckets.Count + " buckets."
            $bStatus = "Bucket #" + $numBuckets
            [int]$bPercentComplete = ($numBuckets/($arrUserBuckets.Count + 1)) * 100
            $bCurrentOperation = [string]$bPercentComplete + "% Complete"

            if ($bSecondsRemaining){
                Write-Progress -Activity $bActivity -Status $bStatus -Id $bId -PercentComplete $bPercentComplete -CurrentOperation $bCurrentOperation -SecondsRemaining $bSecondsRemaining
            }else{
                Write-Progress -Activity $bActivity -Status $bStatus -Id $bId -PercentComplete $bPercentComplete -CurrentOperation $bCurrentOperation
            }

            $pElapsed = (Get-Date) - $pStart
            $bSecondsRemaining = ($pElapsed.TotalSeconds / $numBuckets) * (($arrUserBuckets.Count + 1) - $numBuckets)
            Write-Progress -Activity $bActivity -Status $bStatus -Id $bId -PercentComplete $bPercentComplete -SecondsRemaining $bSecondsRemaining -CurrentOperation $bCurrentOperation

            #return session data for users in this bucket
            ProcessUsersInBuckets -userItem $userItem
        }
    }
}

function ProcessUsersInBuckets{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]$userItem
    )
    begin{
        #reset numUsers to prevent over 100% being interpreted
        $numUsers = 0
        #before we process this bucket of users we need to check the global session timer to prevent access token expiration
        Invoke-CxdSkypeOnlineConnection
    }
    process{
        foreach($userI in $userItem){
            #set variables
            $numUsers ++
            $uId = 2
            $objLastLogon = $null
            $audioSessions = $null
            $getUserSessionError = $null
            $sipAddress = $userI.SipAddress.Replace("sip:","")
            $uActivity = "Processing users..."
            $uStatus = "Gathering user session data report for " + $sipAddress
            [int]$uPercentComplete = ($numUsers/($userItem.Count + 1)) * 100
            $uCurrentOperation = [string]$uPercentComplete + "% Complete"

            Write-Verbose -Message "Getting the past user session data for $sipAddress"
            Write-Progress -Activity $uActivity -Status $uStatus -Id $uId -PercentComplete $uPercentComplete -CurrentOperation $uCurrentOperation -ParentId $bId

            try{
                [array]$userSession = Get-AllUserSessionData -startTime $startTime -endTime $endTime -sipAddress $sipAddress
            }catch{
                Write-Error -Message "Unable to retrieve user session information. Attempting to repair the Skype Online connection and will try again."
                Invoke-CxdSkypeOnlineConnection -RepairPSSession
                [array]$userSession = Get-AllUserSessionData -startTime $startTime -endTime $endTime -sipAddress $sipAddress
            }

            #MOST PROCESSING/EXPORT HAPPENS HERE#
            if ($userSession){
                RateMyCallFeedback -userSession $userSession
                ClientVersions -userSession $userSession
                UserDevices -userSession $userSession
                FailedCalls -userSession $userSession
                UserSummary -userSession $userSession -UserPrincipalName $sipAddress
                
                if ($MatchCallsToASubnet) {
                    #build array of calls matching the given subnet
                    ProcessSubnetCalls -userSession $userSession
                }

                #need to store in global variable to use in summary report processing
                [array]$global:filteredDomainComms += FederatedCommunications -userSession $userSession
                
            }else{
                UsersNotUsingSkype -userSession $userSession -UserPrincipalName $sipAddress
            }
        }
    }
    end{}
            
}

function Get-AllUserSessionData{
    [cmdletbinding()]
    Param(
        $startTime,
        $endTime,
        $sipAddress
    )

    begin {}
    process {
        try{
            [array]$userSession = Get-CsUserSession -StartTime $startTime -EndTime $endTime -User $sipAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -ErrorVariable $userSessionError
        }catch{
            #there was an exception and the command did not execute
            throw $_
        }finally{
            if ($userSessionError){
                #the command executed but returned an error
                throw $getUserSessionError
            }
        }

        #sort data by oldest to newest to check the array for the oldest call in preparation for recursion
        $userSession = $userSession | Sort-Object -Property EndTime

        #verify record count hasn't been exceeded and perform recursion if necessary to retrieve additional records
        if ($userSession.Count -ge 1000){
            $newStart = $userSession[-1].endTime
            Write-Information -Message "Found more than 1000 records for $sipAddress. Performing recursive query to obtain more data using revised start date of $newStart"
            $userSession += Get-AllUserSessionData -startTime $newStart -endTime $endTime -sipAddress $sipAddress
        }   
    }
    end {
        return $userSession
    }
}

function ProcessSubnetCalls {
    [regex]$subnetRegEx = '(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])'
    #filter all sessions for just calls
    [array]$audioSessions = $userSession | Where-Object MediaTypesDescription -like "*Audio*"

    foreach ($call in $audioSessions) {
        $ipAddress = ($call.QoEReport.MediaLines | Where-Object MediaLineLabelText -eq 'main-audio').FromIPAddr
        if ($ipAddress -match $subnetRegEx) {

            #return true for false based on a match of the IP provided and parameters above (inherited)
            if (SubnetMatch -inputData $Matches[0]) {
                #this is a match of the subnet the user provided
                $dialogId = $call.DialogId -replace ';','%3B'
                [array]$matchedSubnetForCall = [pscustomobject]@{
                    'From' = $call.FromUri
                    'To' = $call.ToUri
                    'Start Time' = $call.StartTime
                    'Dialog Id' = $dialogId
                    'IP Address' = $ipAddress
                    'Subnet Address' = $subnetAddress.IPAddressToString
                    'Subnet Mask' = $subnetMask.IPAddressToString
                    'Call Analytics' = 'https://adminportal.services.skypeforbusiness.com/analytics/user/' + $sipAddress + '/call/' + $dialogId
                }

                #export this item
                Export-CxdData -DataInput $matchedSubnetForCall -ReportType ProcessSubnetCalls
            }
        }
    }
}

function RateMyCallFeedback{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession
    )

    begin {}
    process {
        [array]$audioSessions = $userSession | Where-Object MediaTypesDescription -like "*Audio*"
        [array]$processedFeedback = $audioSessions | Where-Object {$_.QoeReport.FeedbackReports} | ProcessFeedback
    }
    end {
        if ($processedFeedback){
            Export-CxdData -DataInput $processedFeedback -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function UserDevices{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession
    )

    begin {}
    process {
        [array]$audioSessions = $userSession | Where-Object MediaTypesDescription -like "*Audio*"
        [array]$userDevices = $audioSessions | ProcessUserDevices
    }
    end {
        if ($userDevices){
            Export-CxdData -DataInput $userDevices -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function FailedCalls{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession
    )

    begin {}
    process {
        [array]$failedCalls = $audioSessions | Where-Object {$_.ErrorReports.ErrorCategory -eq "UnexpectedFailure"} | ProcessFailedCalls
    }
    end {
        if ($failedCalls){
            Export-CxdData -DataInput $failedCalls -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function UserSummary{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession,
        [Parameter()]$UserPrincipalName
    )

    begin {}
    process {
        [array]$audioSessions = $userSession | Where-Object MediaTypesDescription -like "*Audio*"
        [array]$userSummary = ProcessUserSummary -userSummaryInput $audioSessions -UserPrincipalName $UserPrincipalName
    }
    end {
        if ($userSummary){
            Export-CxdData -DataInput $userSummary -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function FederatedCommunications{
    [cmdletbinding()]
    Param
    (
    [Parameter()]$userSession
    )

    begin {}
    process {
        if ($userSession){
            #gather session information for filtering
            $arrDomainComms = $userSession | Where-Object MediaTypesDescription -ne "[RegisterEvent]" | Select-Object StartTime,EndTime,FromUri,ToUri,MediaTypesDescription,DialogId
            
            #filter out internal to internal communications and remove duplicates
            $domainFilterResult = FilterInternalDomainData -DomainData $arrDomainComms -TenantDomains $tenantDomains -FilterOutFromTo | Select-Object StartTime,EndTime,FromUri,ToUri,MediaTypesDescription -Unique | Where-Object MediaTypesDescription -ne ""
        }
    }
    end {
        if ($domainFilterResult){
            Export-CxdData -DataInput $domainFilterResult -ReportType $PSCmdlet.CommandRuntime -Append
            return $domainFilterResult
        }
        
    }
}

function FederatedCommunicationsSummary{
    [cmdletbinding()]
    Param()

    begin {}
    process {
        #process Federated Domain Summary data
        $objProcessedDomainData = $filteredDomainComms | ProcessDomainCallData
        $objFromDomain = $objProcessedDomainData.FromDomain | Select-Object -Unique
        $objToDomain = $objProcessedDomainData.ToDomain | Select-Object -Unique
        $objAllDomains = $objFromDomain + $objToDomain
        $objFilteredDomains = FilterInternalDomainData -TenantDomains $tenantDomains -DomainData $objAllDomains -FilterOutDomains:$true
        
        $arrFederatedStats = ProcessFederatedData -FederatedComms $filteredDomainComms -FilteredDomains $objFilteredDomains
    }
    end {
        if ($arrFederatedStats){
            Export-CxdData -DataInput $arrFederatedStats -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function ClientVersions{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession
    )
    begin{
 
    }
    process{    
        $clientVersions = $userSession | Where-Object MediaTypesDescription -eq "[RegisterEvent]" | Sort-Object EndTime -Descending | Select-Object @{Name="UserPrincipalName";Expression={$_.FromUri}},@{Name="Version";Expression={$_.FromClientVersion}},@{Name="Date";Expression={$_.StartTime}} | Where-Object Version -ne ""
        $clientVersions = $clientVersions | Select-Object @{Name="User";Expression={$userI.UserPrincipalName}},@{Name="Version";Expression={$_.Version}} -Unique           
    }
    end{
        if ($clientVersions){
            Export-CxdData -DataInput $clientVersions -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }
}

function UsersNotUsingSkype{
    [cmdletbinding()]
    Param(
        [Parameter()]$userSession,
        [Parameter()]$UserPrincipalName
    )
    begin{}
    process{
        if (!$userSession){
            $usersNotUsingSkype = [PSCustomObject]@{
                UsersNotUsingSkype = $UserPrincipalName
            }
        }
    }
    end{
        if ($usersNotUsingSkype){
            Export-CxdData -DataInput $usersNotUsingSkype -ReportType $PSCmdlet.CommandRuntime -Append
        }
    }

}

function ProcessFederatedData{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)]$FederatedComms,   
        [Parameter(Mandatory=$false)]$FilteredDomains
    )
    begin{
        [array]$arrFedResult = @()
    }
    process{
        #count IM Sessions
        foreach ($domainItem in $FilteredDomains){
            $arrFilteredDomainComms = $FederatedComms | Where-Object {$_.ToUri -like "*$($domainItem)" -or $_.fromUri -like "*$($domainItem)"}
            if ($arrFilteredDomainComms){
                $arrFedResult += [PSCustomObject][ordered]@{
                    Domain = $domainItem
                    IM = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[IM]"}).Count
                    Audio = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Audio]"}).Count
                    Video = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Audio][Video]"}).Count
                    Sharing = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Video][AppSharing]"}).Count
                    Conference = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Conference]"}).Count
                    "IM Conference" = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Conference][IM]"}).Count
                    "Sharing Conference" = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Conference][AppSharing]"}).Count
                    "Audio Conference" = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Conference][Audio]"}).Count
                    "Video Conference" = ($arrFilteredDomainComms.MediaTypesDescription | Where-Object {$_ -eq "[Conference][Audio][Video]"}).Count
                }
            }
        }

    }
    end{
        return $arrFedResult
    }
}

function FilterInternalDomainData{
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$false)][array]$TenantDomains,
        [Parameter(Mandatory=$false)]$DomainData,
        [Parameter(Mandatory=$false)][switch]$FilterOutFromTo,
        [Parameter(Mandatory=$false)][switch]$FilterOutDomains
    )
    begin{}
    process{
        #removes objects where to/from are considered internal to the organization
        if($FilterOutFromTo){
            ForEach ($domainItem in $TenantDomains){
                $DomainData = $DomainData | Where-Object {$_.ToUri -and $_.FromUri -notlike $domainItem}
            }
        return $DomainData
        }

        #removes Domain values from the object
        if($FilterOutDomains){
            ForEach ($domainItem in $TenantDomains){
                $DomainData = $DomainData | Where-Object {$_ -notlike $domainItem}
            }
        return $DomainData
        }

    }
    end{}
}

function ProcessDomainCallData{
    [cmdletbinding()]
    Param
    (
    [Parameter(Mandatory=$false,ValueFromPipeline=$true)]
    [object]$Calls
    )
    begin{

    }
    process{
        if (!$calls){
            return
        }
        $tempSplitFrom = $_.FromUri.Split("@")
        $strFromDomain = $tempSplitFrom[1]
        $tempSplitTo = $_.ToUri.Split("@")
        $strToDomain = $tempSplitTo[1]

        $newObject = [PSCustomObject][ordered]@{
            FromUri = $_.FromUri
            ToUri = $_.ToUri
            FromDomain = $strFromDomain
            ToDomain = $strToDomain
            MediaType = $_.MediaTypesDescription
        }

        return $newObject
    }
    end{}
}

function ProcessUserDevices{
    begin{
        [regex]$devicesToExclude = "(Default output device)|(Default input device)|(Remote Audio)|(Built-in Output)"
    }
    process{
        #if no QoE report exists; exit
        if (!$_.QoeReport){return}

        $audioMediaLines = $null
        $toUriSessions = $null
        $fromUriSessions = $null
        $newFromObject = $null
        $newToObject = $null
        
        #break down MediaLines as this can be an array (i.e. multi-part video gives back an array of objects)
        $audioMediaLines = $_.QoEReport.MediaLines | ForEach-Object {$_ | Where-Object MediaLineLabelText -eq "main-audio"}
        
        #get all sessions where the target user matches the called party
        [array]$toUriSessions = $_ | Where-Object {$_.ToUri -eq $UserI.UserPrincipalName}
        
        #get all sessions where the target user matches the calling party
        [array]$fromUriSessions = $_ | Where-Object {$_.FromUri -eq $UserI.UserPrincipalName}# -and $audioMediaLines.FromRenderDev -match $devicesToExclude}

        $newFromObject = [PSCustomObject][ordered]@{
            SipAddress = $fromUriSessions.FromUri
            RenderDevice = $(if ($audioMediaLines.FromRenderDev -match $devicesToExclude -or $audioMediaLines.FromRenderDev -eq ""){$null}else{$audioMediaLines.FromRenderDev})
        }
        $newToObject = [PSCustomObject][ordered]@{
            SipAddress = $ToUriSessions.ToUri
            RenderDevice = $(if ($audioMediaLines.ToRenderDev -match $devicesToExclude -or $audioMediaLines.ToRenderDev -eq ""){$null}else{$audioMediaLines.ToRenderDev})
        }
        
        if ($newFromObject.RenderDevice -and $newFromObject.SipAddress){return $newFromObject}
        if ($newToObject.RenderDevice -and $newFromObject.SipAddress){return $newToObject}
            
    }
    end{}
}

function ProcessFeedback{
    begin{}
    process{
        if ($_.QoeReport.FeedbackReports -eq $null){
            continue; #no feedback report found
        }
        #check for multiple QoE reports in this call (very rare)
        if ($_.QoeReport.FeedbackReports.Count -gt 1){
            $_.QoeReport.FeedbackReports = $_.QoeReport.FeedbackReports | Where-Object FromUserUri -eq $userI.UserPrincipalName
        }
        #process feedback tokens from array into string to present back as an object we can report on
        if ($_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne "0"){
            [array]$arrTokens = $_.QoeReport.FeedbackReports.tokens | Where-Object Value -ne "0" | Select-Object Id #declare an array so we can get the count. if the count is only one, the trim statement preceding won't make sense so we need to handle this.
            if ($arrTokens.Count -gt "1"){
                $arrTokens = $arrTokens.id.trim() -join "," #output would show System.Object[] in the report otherwise so we need to convert them to a string.
            }
        }else{
            $arrTokens = ""
        }
        #process traceroute collection for each call and pull back interesting information.
        $strTraceRoutes = ProcessTraceRoutes -Call $_

        #build array of objects with interesting information about the poor feedback for this given call.
        try{
            $newObject = [PSCustomObject][ordered]@{
                FromUri = $_.FromUri
                ToUri = $_.ToUri
                CaptureTime = $_.QoeReport.FeedbackReports.CaptureTime
                Rating = $_.QoeReport.FeedBackReports.Rating
                FeedbackText = $_.QoeReport.FeedBackReports.FeedbackText
                Tokens = $arrTokens.Id
                MediaType = $_.MediaTypesDescription
                ConferenceUrl = $(if ($_.ConferenceUrl){$_.ConferenceUrl}else{"N/A"})
                FromClientVersion = $_.FromClientVersion
                ToClientVersion = $_.ToClientVersion
                MediaStartTime = $_.QoeReport.Session.MediaStartTime
                MediaEndTime = $_.QoeReport.Session.MediaEndTime
                MediaDurationInSeconds = ($_.QoeReport.Session.MediaEndTime - $_.QoeReport.Session.MediaStartTime).Seconds
                FromOS = $_.QoeReport.Session.FromOS
                ToOS = $_.QoeReport.Session.ToOS
                FromVirtualizationFlag = $_.QoeReport.Session.FromVirtualizationFlag
                ToVirtualizationFlag = $_.QoeReport.Session.ToVirtualizationFlag
                FromConnectivityIce = $_.QoeReport.MediaLines.FromConnectivityIce
                FromRenderDev = $_.QoeReport.medialines.FromRenderDev
                ToRenderDev = $_.QoeReport.MediaLines.ToRenderDev
                FromNetworkConnectionDetail = $_.QoeReport.MediaLines.FromNetworkConnectionDetail
                ToNetworkConnectionDetail = $_.QoeReport.MediaLines.ToNetworkConnectionDetail
                FromIPAddr = $_.QoeReport.MediaLines.FromIPAddr
                ToIPAddr = $_.QoeReport.MediaLines.ToIPAddr
                FromReflexiveLocalIPAddr = $_.QoeReport.MediaLines.FromReflexiveLocalIPAddr
                FromBssid = $_.QoeReport.MediaLines.FromBssid
                FromVPN = $_.QoeReport.MediaLines.FromVPN
                ToVPN = $_.QoeReport.MediaLines.ToVPN
                JitterInterArrival = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].JitterInterArrival}else{"N/A"})
                JitterInterArrivalMax = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].JitterInterArrivalMax}else{"N/A"})
                PacketLossRate = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].PacketLossRate}else{"N/A"})
                PacketLossRateMax = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].PacketLossRateMax}else{"N/A"})
                BurstDensity = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].BurstDensity}else{"N/A"})
                BurstDuration = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].BurstDuration}else{"N/A"})
                BurstGapDensity = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].BurstGapDensity}else{"N/A"})
                BurstGapDuration = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].BurstGapDuration}else{"N/A"})
                PayloadDescription = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].PayloadDescription}else{"N/A"})
                AudioFECUsed = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].AudioFECUsed}else{"N/A"})
                SendListenMOS = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].SendListenMOS}else{"N/A"})
                OverallAvgNetworkMOS = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].OverallAvgNetworkMOS}else{"N/A"})
                NetworkJitterMax = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].NetworkJitterMax}else{"N/A"})
                NetworkJitterAvg = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].NetworkJitterAvg}else{"N/A"})
                StreamDirection = $(if ($_.QoeReport.AudioStreams){$_.QoeReport.AudioStreams[0].StreamDirection}else{"N/A"})
                TraceRoutes = $strTraceRoutes
                DialogId = $_.DialogId
            }
        }catch{
            Write-Warning -Message "Could not process all Rate My Call feedback data correctly for $($userI.UserPrincipalName)."
        }
        return $newObject
    }
    end{
    }
}

function ProcessClientVersions{
    [cmdletbinding()]
        Param
        (
        [Parameter(Mandatory=$false)]
        [int]$ClientVersionCutOffDays,

        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [psobject]$ClientVersionData
        )

    begin {}
    process{
        $strClientVersionSplit = $null
        $strClientVersionSplit = [regex]::Match($ClientVersionData, '(\d{4}\.\d{4})').Captures
        if ($strClientVersionSplit.Success -eq $true){
            $strClientVersionSplit = (($strClientVersionSplit.Groups[1]).Value).Split(".") #need this additional line due to null value possibility along with error: "cannot index null array"
            #need to handle scenario where we only have revision and build data
            if ($strClientVersionSplit.count -eq "2"){
                [array]$objClientVersion = [PSCustomObject][ordered]@{
                    Revision = $strClientVersionSplit[0]
                    Build = $strClientVersionSplit[1]
                }
            #handle full version info here...
            }elseif ($strClientVersionSplit.count -eq "4"){
                [array]$objClientVersion = [PSCustomObject][ordered]@{
                    Major = $strClientVersionSplit[0]
                    Minor = $strClientVersionSplit[1]
                    Revision = $strClientVersionSplit[2]
                    Build = $strClientVersionSplit[3]
                }
            }
        }
        return $objClientVersion        
    }
     end{}
}

function ProcessTraceRoutes{
    <#
        .SYNOPSIS
        Internal function used to process traceroute data from QoeReport.
        
        .DESCRIPTION
        We process this array of traceroute data into a string we can export to CSV/HTML.
        
        .PARAMETER Call
        Object containing the call properties used to build the string we return.
        
        .EXAMPLE
        ProcessTraceRoutes -Call <object>

    #>

    [cmdletbinding()]
        Param
        (
        [Parameter(mandatory=$true, valuefrompipeline=$false)]
        $Call
        )
        begin{
            if ($Call.QoeReport.TraceRoutes -eq $null){
                $networkPath = $null
            }else{
                $call.qoereport.TraceRoutes | ForEach-Object {
                    $networkPath = $networkPath + "Hop:" + $($_.hop) + ",IP:" + $($_.IPAddress) + ",RTT:"+ $($_.RTT) + ";"
                }
                return $networkPath
            }
        }
        end{}

}

function ProcessUserSummary{
    [cmdletbinding()]
        Param(
            [Parameter()]$userSummaryInput,
            [Parameter()]$UserPrincipalName        
        )
    begin{}
    process{
        [int]$intFeedbackTotal = ($userSummaryInput.qoereport | Where-Object FeedbackReports).count
        [int]$intFeedbackAvoided = ($userSummaryInput.QoEReport.FeedbackReports | Where-Object {$_.Rating -like "0*"}).Count #feedback requested but not submitted by the user
        [int]$intFeedbackPoor = ($userSummaryInput.QoEReport.FeedBackReports | Where-Object {$_.Rating -like "1*" -or $_.Rating -like "2*"}).Count #a 1 or a 2 is considered 'poor' feedback in SFBO
        [int]$intFeedbackGood = $intFeedbackTotal - $intFeedbackAvoided - $intFeedbackPoor
        [int]$intFeedbackProvided = $intFeedbackTotal - $intFeedbackAvoided

        if ($intFeedbackProvided -eq "0"){
            $feedbackPercentage = $intFeedbackProvided.ToString("P")
            $feedbackPoorPercentage = $intFeedbackProvided.ToString("P")
        }else{
            $feedbackPercentage = ($intFeedbackProvided / $intFeedbackTotal).ToString("P")
            $feedbackPoorPercentage = ($intFeedbackPoor / $intFeedbackProvided).ToString("P")
        }

        [array]$userSummary = [PSCustomObject][ordered]@{
            UserPrincipalName = $userI.UserPrincipalName
            AudioOnlySessions = ($userSummaryInput | Where-Object MediaTypesDescription -eq "[Audio]").Count
            ConferenceSessions = ($userSummaryInput | Where-Object MediaTypesDescription -like "*[Conference]*").Count
            SessionsWithVideo = ($userSummaryInput | Where-Object MediaTypesDescription -like "*[Video]*").Count
            FeedbackTotal = $intFeedbackTotal
            FeedbackGiven = $intFeedbackProvided
            FeedbackGood = $intFeedbackGood
            FeedbackAvoided = $intFeedbackAvoided
            FeedbackPoor = $intFeedbackPoor
            FeedbackPercentage = $feedbackPercentage
            FeedbackPoorPercentage = $feedbackPoorPercentage
        }
        return $userSummary
    }
    end{}
}

function ProcessFailedCalls{
    begin{}
    process{
        #pull out matching diagnostic codes
        #need to account for multiple objects in the ErrorReports and choose the first one otherwise, accept the default
            if ($_.ErrorReports.Count -gt "1"){
                $requestType = $_.ErrorReports.RequestType[0]
                $diagnosticId = $_.ErrorReports.DiagnosticId[0]
                $diagnosticHeader = $_.ErrorReports.DiagnosticHeader[0]
            }else{
                $requestType = $_.ErrorReports.RequestType
                $diagnosticId = $_.ErrorReports.DiagnosticId
                $diagnosticHeader = $_.ErrorReports.DiagnosticHeader
            }
            #$test = [regex]::Match($diagnosticHeader, '(StatusString:.+?)\;')
            #$test2 = [regex]::Match($diagnosticHeader, '(Reason:.+?)\;')
            [array]$newObject = [PSCustomObject][ordered]@{
                FromUri = $_.FromUri
                ToUri = $_.ToUri
                StartTime = $_.StartTime
                EndTime = $_.EndTime
                CallDurationInSeconds = $(if ($_.EndTime){($_.EndTime - $_.StartTime).TotalSeconds}else{"N/A"})
                MediaType = $_.MediaTypesDescription
                ConferenceUrl = $(if ($_.ConferenceUrl){$_.ConferenceUrl}else{"N/A"})
                RequestType = $requestType
                DiagnosticId = $diagnosticId
                FromClientVersion = $_.FromClientVersion
                ToClientVersion = $_.ToClientVersion
                DiagnosticHeader = $diagnosticHeader
                DialogId = $_.DialogId
            }

        return $newObject
    }
    end{}
}

function Set-CxdSavePath{
    [cmdletbinding()]
    Param(
        #[Parameter(Mandatory=$false)]$KeepExistingReports
    )

        begin{
            #trim trailing backslash in case user entered it
            $ReportSavePath = $ReportSavePath.TrimEnd("\")
        }
        process{
            #create directory
            if (!(Test-Path -Path $ReportSavePath)){
                try{
                    Write-Verbose "Attempting to create base report path: $ReportSavePath"
                    mkdir -Path $ReportSavePath | Out-Null
                }catch{
                    throw
                }
            }

            #create unique folder for test results
            $folderDate = Get-Date -UFormat "%Y-%m-%d"
            [int]$folderCounter = 0
            do {
                $folderCounter ++
                [string]$newSavePath = $ReportSavePath + "\" + $folderDate + "-" + $folderCounter
                Write-Verbose -Message "Testing save path: $newSavePath"
            }until(!(Test-Path -Path $newSavePath))

            try{
                Write-Verbose -Message "Attempting to create unique testing directory: $newSavePath"
                mkdir -Path $newSavePath | Out-Null
            }catch{
                throw
            }            
        }
        end{
            return $newSavePath
        }
}

function Export-CxdData{
    
    [cmdletbinding()]
    Param
    (
        [Parameter(mandatory=$true, valuefrompipeline=$false)][PSObject]$DataInput,
        [Parameter(mandatory=$true, valuefrompipeline=$false)][string]$ReportType,
        [Parameter(mandatory=$false, valuefrompipeline=$false)][switch]$Append
    )

    try{
        $dataInput | Export-Csv -Path ($ReportSavePath + "\" + ($reportType + ".csv")) -NoTypeInformation -Append
    }catch{
        throw
    }

}

function Get-OfficeVersionData{

    <##############################
    #.SYNOPSIS
    #Retrieves HTML data by accessing public web sites containing Office versions, then parsing HTML into a PSObject.
    #
    #.DESCRIPTION
    #We use this internal function to find the table containing the Office version so we can compare it with the user-agent version data in Get-CsUserSession. 
    #
    #.EXAMPLE
    #Get-OfficeVersionData -OfficeVersion 2016 -InstallType C2R -Tag Table -Id 99867asdfg11234r
    #
    #.NOTES
    #The HTML tag and id used to find the right table data may change over time however we'll do our best to maintain this module with the right URI/tag/id fields.
    ##############################>

    [cmdletbinding()]
        Param
        (
        [Parameter(Mandatory=$true)]
        [ValidateSet("2016","2013")]
        [string]$OfficeVersion,

        [Parameter(Mandatory=$true)]
        [ValidateSet("C2R","MSI")]
        [string]$InstallType,

        [Parameter(Mandatory=$true)]
        [string]$Tag,
        
        [Parameter(Mandatory=$true)]
        [string]$Id
        )

    begin{
        switch ($OfficeVersion) {
            2016 {
                if ($InstallType -eq "C2R"){
                    $Uri = "https://support.office.com/en-us/article/Version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7?ui=en-US&rs=en-US&ad=US#"
                }else{
                    $Uri = ""
                }
            }
            2013 {
                if ($InstallType -eq "C2R"){
                    $Uri = ""
                }

            }
        }
    }

    process{
        $webRequest = Invoke-WebRequest -Uri $Uri #initiate web request to URI
        $webObject = @($webRequest.ParsedHtml.getElementsByTagName($tag) | Where-Object id -eq $Id) #find the table matching the parameters given
        $rows = @($webObject.Rows) #extract rows
        $titles = @()
        $skipValueAdd = $null
        $objResult = $null

        foreach($row in $rows){
            $cells = @($row.Cells)
            $resultObject = [Ordered]@{}

            if ($row.rowIndex -eq "0"){
                $titles = @($cells | ForEach-Object {"" + ($_.innerText).Trim()})
                $skipValueAdd = $true #necessary to stop the first row from being values in the first index of the array
            }
        
            for ($counter = 0; $counter -lt $cells.Count; $counter++){
                $title = $titles[$counter]
                $resultObject[$title] = ("" + $cells[$counter].innerText).trim()
            }
            if (!$skipValueAdd){
                [PSCustomObject]$resultObject
                [array][pscustomobject]$objResult += $resultObject
            }
            $skipValueAdd = $null #reset value to add proper rows
        }

        
    }
    end{}
}
function Invoke-CxdSkypeOnlineConnection{
    [cmdletbinding()]
    Param
    (
    [Parameter(mandatory=$false, valuefrompipeline=$false)]
    [switch]$RepairPSSession
    )
    begin{}
    process{
        #calculate session timer to handle access token expiration
        if ($global:sessionStartTime){
            $global:sessionTotalTime = ((Get-Date) - $global:sessionStartTime)
        }

        #determine if Skype for Business PsSession is loaded in memory
        $sessionInfo = Get-PsSession

        #need to loop through each session a user might have opened previously
        foreach ($sessionItem in $sessionInfo){
            #check session timer to know if we need to break the connection in advance of a timeout. Break and make new after 40 minutes.
            if ($sessionItem.ComputerName.Contains(".online.lync.com") -and $sessionItem.State -eq "Opened" -and $global:sessionTotalTime.TotalSeconds -ge "2400"){
                Write-Verbose -Message "The PowerShell session has been running for $($global:sessionTotalTime.TotalMinutes) minutes. We need to shut it down and create a new session due to the access token expiration at 60 minutes."
                $sessionItem | Remove-PSSession
                Start-Sleep -Seconds 2
                $SessionFound = $false
                $global:sessionTotalTime = $null #reset the timer
            }

            #try to repair PSSession
            if ($sessionItem.ComputerName.Contains(".online.lync.com") -and $sessionItem.State -ne "Opened" -and $RepairPSSession){
                Write-Verbose -Message "Attempting to repair broken PowerShell session to Skype for Business Online using cached credential."
                $sessionItem | Remove-PSSession
                Start-Sleep -Seconds 3
                $SessionFound = $false
                $global:sessionTotalTime = $null
            }elseif ($sessionItem.ComputerName.Contains(".online.lync.com") -and $sessionItem.State -eq "Opened"){
                $SessionFound = $true
            }
        }

        if (!$SessionFound){
            Write-Verbose -Message "Creating new Skype Online PowerShell session..."
            try{
                if ($OverrideAdminDomain){
                    #import to global scope
                    $lyncsession = New-CsOnlineSession -Credential $Credential -OverrideAdminDomain $OverrideAdminDomain -ErrorAction SilentlyContinue -ErrorVariable $newOnlineSessionError
                }else{
                    $lyncsession = New-CsOnlineSession -Credential $Credential -ErrorAction SilentlyContinue -ErrorVariable $newOnlineSessionError
                }
            }catch{
                throw;
            }finally{
                if ($newOnlineSessionError){
                    throw $newOnlineSessionError
                }
            }
            Write-Verbose -Message "Importing remote PowerShell session..."
            $global:sessionStartTime = (Get-Date)
            Import-PSSession $lyncsession -AllowClobber | Out-Null
        }
    }
    end{}
}

function ProcessCsv{
    if (!$(Test-Path -Path $CsvFileWithUsers)){
        Write-Error -Message "The file $CsvFileWithUsers cannot be found."
        throw;
    }
        Write-Verbose -Message "Processing users from CSV file."
    try{
        $csvUsers = Import-Csv $CsvFileWithUsers
    }catch{
        throw;
    }
    #check CSV to make sure we have the right header
    if (($csvUsers | Get-Member | ForEach-Object {$_.Name -eq "UserPrincipalName"} | Where-Object {$_ -eq $True}).count -eq "0"){
        Write-Error -Message "The CSV file you've specified does not contain the correct header value of 'UserPrincipalName'. Please correct this and run this function again."
        throw;
    }elseif(($csvUsers | ForEach-Object {[bool]($_.UserPrincipalName -as [System.Net.Mail.MailAddress])} | Where-Object {$_ -eq $False}).count -ge "1"){
        Write-Error -Message "The CSV file you've specified contains an improperly formatted UserPrincipalName in one or more rows. Please correct this and run the function again."
        throw;
    }
    return $csvUsers
}

function ProcessSkypeOnlineUsers{

    process{
        if ($UserToSearch) {
            Write-Verbose -Message "Please wait while we get data for $UserToSearch"
            return Get-CsOnlineUser -Identity $UserToSearch -ErrorAction SilentlyContinue
        }

        if ($csvValidatedUsers) {
            Write-Verbose -Message "Please wait while we find all users in Skype for Business Online using the CSV file provided."
            $enabledUsers = $csvValidatedUsers | ForEach-Object UserPrincipalName | Get-CsOnlineUser -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            return $enabledUsers
        } else {
            Write-Verbose -Message "Please wait while we get all enabled users in Skype for Business Online."
            $enabledUsers = Get-CsOnlineUser -Filter {Enabled -eq $True} -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Return $enabledUsers
        }
    }
}

function Set-WinRMNetworkDelayMS{
    <#
    .SYNOPSIS
    Changes the WinRM setting for the NetworkDelayms property  (WSMan:\localhost\Client\NetworkDelayms).

    .Description
    The Skype For Business Online Connector connects to Skype For Business Remote PowerShell Server in order to establish a remote session. Sometimes this remote session can be broken because of network delay. To help address this problem, Skype For Business recommends setting NetworkDelayms to 30000 milliseconds (30 seconds) instead of the default value of 5000 milliseconds (5 seconds).

    .Parameter value

    Value of the NetworkDelayms property in milliseconds (there are 1000 milliseconds in each second). This must be an integer value.

    .EXAMPLE
    Set-WinRMNetworkDelayMS 30000

    This example sets the NetworkDelayms property (WSMan:\localhost\Client\NetworkDelayms) to 30000 milliseconds (30 seconds).

    #>

  param(
    [Parameter(Mandatory = $false)] [string] $value="30000"
  )

  $networkDelay = Get-Item WSMan:\localhost\Client\NetworkDelayms

  if($networkDelay -eq $null)
  {
    # If cannot get NetworkDelayms due to permission or other reason, just return.
    return
  }

  $oldValue = $networkDelay.Value
  $newValue = $value

  if($newValue -ne $oldValue)
  {
    Set-Item WSMan:\localhost\Client\NetworkDelayms $newValue

    # Warns the user that running this command has changed their client setting.
    Write-Warning "WSMan NetworkDelayms has been set to $newValue milliseconds. The previous value was $oldValue milliseconds."
    Write-Warning "To improve the performance of the Skype For Business Online Connector, it is recommended that the network delay be set to 30000 milliseconds (30 seconds). However, you can use Set-WinRMNetworkDelayMS to change the network delay to any integer value."
  }
}

function SubnetMatch {
    #inherits variables from main function
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true)][Net.IPAddress]$inputData
    )
    ($subnetAddress.Address -band $subnetMask.Address) -eq ($inputData.Address -band $subnetMask.Address)
}


