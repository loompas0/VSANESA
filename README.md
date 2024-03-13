# VsanEsa

**Extract List of Nvme from Vsan ESA DB and load a spreadsheet.**

2 PowerShell scripts are used :

<span class="mark">\* addsheet.ps1</span>

<span class="mark">\* PutVsanNvmeSheet.ps1</span>

------------------------------------------------

**Legal Disclaimer:**

This script is not supported by anyone.  
The creator shall not be liable for any issue or damage.  
All scripts are provided AS IS without a warranty of any kind.  
The author further disclaims all implied warranties including, without
limitation, any implied warranties of merchantability or of fitness for
a particular purpose.  
The entire risk arising out of the use or performance of the sample
scripts and documentation remains with you.  
In no event shall its authors, or anyone else involved in the creation,
production, or delivery of the scripts be liable for any damages
whatsoever  
(Including, without limitation, damages for loss of business profits,
business interruption, loss of business information, or other pecuniary
loss)  
arising out of the use of or inability to use the sample scripts or
documentation, even if the author has been advised of the possibility of
such damage.

------------------------------------------------

## Name : Addsheet.ps1

> What is it :

Create a new Google SpreadSheet using Google API

Set permission to a Google user

> Requirements:

PowerShell 6.5 or higher

Requires A google account to create a Google sheet.

Access to Google API  
(follow the guide at : <https://developers.google.com/sheets/api/guides>
)

Install the PowerShell module UMN-Google from PowerShell gallery to
access the Google API

> More Information :

Some information on manipulating google cloud spreadsheet could be found
at:  
<https://jamesachambers.com/modify-google-sheets-using-powershell/>

Documentation on UMN-Google module could be found at :  
<https://github.com/umn-microsoft-automation/UMN-Google/tree/master/docs>

> Version : 1

03/13/2024

Created by: Rémy Bernheim

## Name : PutVsanNvmeSheet.ps1

> What is it :

- Extract Lists of Nvme from Vmware HCL for Vsan ESA (By default vendor
  is Hewlett Packard Enterprise )

- Put this list to a Google Sheet to be shared and used by HilightThis
  Google Extension

> Requirements:

PowerShell 6.5 or higher

A google account to create a Google sheet.

Access to google API  
(follow the guide at : <https://developers.google.com/sheets/api/guides>
)

Install the powershell module UMN-Google from powershell gallery to
access the Google API

> More Information :

Some information on Highlightthis could be found at :  
<https://highlightthis.net/>

Some information on manipulating Google cloud spreadsheet could be found
at :  
<https://jamesachambers.com/modify-google-sheets-using-powershell/>

Documentation on UMN-Google module could me found at :  
<https://github.com/umn-microsoft-automation/UMN-Google/tree/master/docs>

> Version : 0.9

03/13/2024

Created by: Rémy Bernheim

\-
