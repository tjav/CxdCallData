# CxdCallData
## Extract critical usage information from Skype for Business Online and Microsoft Teams

### Created by: [Jason Shave](jason.shave@microsoft.com)

This module is published to [PowerShell Gallery](http://powershellgallery.com) and can be found using:

`Find-Module CxdCallData`

To install the module from the Gallery:

`Install-Module CxdCallData`

To update the module once it's been installed:

`Update-Module CxdCallData`

### Reports
Each report produces a CSV file which can be easily imported into PowerBI or another visualization tool.

1. **ClientVersions** - This report produces a normalized output of all client versions by user. It is particularly helpful if you're trying to understand how many different versions your clients are using in your environment. Having a consistent version to work with across your organization is key to reducing support costs. Use the report to track down older versions or if a particular user has reported a problem, use this data to understand if they're out of date before diving deeper into the problem.

2. **FederatedCommunications** - If you've ever wanted a list of all external sessions by user, this is your report. The list includes the Start/End time, From/To SIPURI, and the media type (IM/Voice/Video/etc.).

3. **FederatedCommunicationsSummary** - This report provides a summary of domains your organization communicates with. It includes a breakdown of communications by type (i.e. Chat, Voice, Video, Meetings, etc.). This is particularly helpful if you want to switch from "open" to "closed" federation and need to know the domains you frequently communicate with.

4. **RateMyCallFeedback** - This list includes all calls including the specific QoE report data for each call where a person rated a 1 or 2 star after their call. Since there is no existing report in Skype for Business Online, this is a highly valuable set of data showing which users are providing poor feedback but also gives context to why the session may have been problematic. Some of the data included in this report comes from the MediaLines object we get back from Get-CsUserSession including a traceroute path and RTT for each hop. You can see if a user was on VPN, if they're wired or wireless, or what headset/audio endpoint they used for the call. 

5. **UserDevices** - This report produces a list of users and the endpoint devices they use for their sessions. You can quickly see who's using an optimized device vs. a built-in audio device like a laptop's built-in speaker/mic.

6. **UsersNotUsingSkype** - Customers use this report to find users who are enabled for Skype for Business Online but haven't signed into the platform for the timeframe given. Admins can use this report to reclaim unused licenses, or more importantly, empower the change management team to approach the user with an adoption plan.

7. **UserSummary** - This report gives a summary of all communications for users over a given time period. Specifically, we provide audio only sessions vs. those with video along with a summary around Rate My Call feedback (i.e. how many times we asked for feedback but the user just closed the window). You can have an informed conversation with a user who is complaining about a poor experience by asking them why they were asked for feedback 40 times but never rated their calls. The business can incentivize staff who provide feedback as an adoption strategy where this report provides the detail to award individuals.

### Release Notes:

For a detailed description see help from 'get-help CxdCallData'

1. Removed functionality in the **-KeepExistingReports** switch as some people were using "D:\" as the path for reports. Omitting this switch would wipe out all files on the drive recursively. 
2. Resolved ParameterSet issue causing errors when some parameters were omitted.
