# CxdCallData
## Extract critical usage information from Skype for Business Online and Microsoft Teams

### Created by: [Jason Shave](jason.shave@microsoft.com)

This module is published to [PowerShell Gallery](powershellgallery.com) and can be found using:

`Find-Module CxdCallData`

To install the module from the Gallery:

`Install-Module CxdCallData`

To update the module once it's been installed:

`Update-Module CxdCallData`

### Release Notes:

For a detailed description see help from 'get-help CxdCallData'

1. Removed functionality in the **-KeepExistingReports** switch as some people were using "D:\" as the path for reports. Omitting this switch would wipe out all files on the drive recursively. 
2. Updated Github MD.
