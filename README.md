# PRTG-O365Licensing
## Get Office 365 licensing information into PRTG
This custom sensor can be used with PRTG to retrieve licensing information from Office 365. For each license found
three channels are created:
* Number of available licenses
* Number of used licenses
* Number of consumed licenses
* Number of licenses in warning state
* Time since last DirSync (if applicable)
* Time since last PasswordSync (if applicable)

For detailed information, please visit http://www.team-debold.de/2016/11/05/prtg-office-365-lizenzen-im-blick/