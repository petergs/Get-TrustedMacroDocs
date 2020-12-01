# Get-TrustedMacroDocs
## Description
A PowerShell function that returns a list of the macro-enabled documents that have been trusted in Office.

## Motivations
FireEye released State of the Hack S4E02: Weaponizing Office Documents with VBA Purging while Evading Modern Detection on 19 Nov 2020. They discussed detections for maldocs at the endpoint level and referenced a Trust Record registry key that's changed predictably when macro-enabled documents are trusted. In thinking about how to hunt for this sort of thing, I was curious if it was possible to pull Trust Records at scale via Intune and PowerShell. Even though it isn't a proactive or preventitive measure, it could be utilized to check for instances when potentially malicious documents have been executed in the past.  

### Usage Example
Return trusted macro-enabled docs that were opened from within the Outlook temporary files folder:

    Get-TrustedMacroDocs -Application All | Where -Path Like '%USERPROFILE%/AppData/Local/Microsoft/Windows/Temporary Internet Files/Content.Outlook/*'

## Future Plans
- Example script for Intune deployment and result consolidation in Azure Blob Storage

## Additional Resources
* [Mari DeGrazia](https://twitter.com/maridegrazia) has a good discussion of the registry key utilized by this function on [her blog.](http://az4n6.blogspot.com/2016/02/more-on-trust-records-macros-and.html). She notes that a document **won't** appear under this key if macros have been enabled by default in the Office Trust Center. This is controlled per application, and she highlights a key that can be used to check this setting. 
* [Outflank has an article](https://outflank.nl/blog/2018/01/16/hunting-for-evil-detect-macros-being-executed/) (authored by Pieter Ceelen) that discusses alerting on this key with Sysmon. 
