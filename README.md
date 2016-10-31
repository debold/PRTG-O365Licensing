# PRTG-O365Licensing
## Get Office 365 licensing information into PRTG
This custom sensor can be used with PRTG to retrieve licensing information from Office 365. For each license found
three channels are created:
* Number of available licenses
* Number of used licenses
* Number of consumed licenses
* Number of licenses in warning state

Additionally, the time since last DirSync is displayed, too. If password synchronisation is enabled, this is shown
in a separate channel.