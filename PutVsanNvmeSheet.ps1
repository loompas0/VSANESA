#-------------------------------------------------
# 
# Name : PutVsanNvmeSheet
# What :
#   Extract Lists of Nvme from Vmware HCL for Vsan ESA (By default vendor is Hewlett Packard Enterprise )
#   Put this list to a Google Sheet to be shared and used by HilightThis Google Extension
# 
# CopyRight Disclaimer (2024)
#
# Rémy Bernheim, hereby disclaims all copyright interest in the program "PutVsanNvmeSheet.ps1" 
#  written by Rémy Bernheim.
#
#-------------------------------------------------

# Requires PowerCLI 6.5 or higher 
# Requires A google account to create a Google sheet
# Requires Access to google API (follow the guide at : https://developers.google.com/sheets/api/guides)
# Requires to install the powershell module UMN-Google from powershell gallery to access the Google API
# *** Install-Module UMN-Google ***
#  

# Some information on Highlightthis could be found at : 
#   https://highlightthis.net/
# Some information on manipulating google cloud spreadsheet could be found at :
#   https://jamesachambers.com/modify-google-sheets-using-powershell/
# Documentation  on UMN-Google  module could me found at : 
#   https://github.com/umn-microsoft-automation/UMN-Google/tree/master/docs
#

#-------------------------------------------------
# 
# Version : 0.9
# 03/13/2024
# Created by: Rémy Bernheim
#
################################################
# Legal Disclaimer:
# This script is not supported by anyone. 
# The creator shall not be liable for any issue or damage
# All scripts are provided AS IS without warranty of any kind. 
# The author further disclaims all implied warranties including, without limitation, 
# any implied warranties of merchantability or of fitness for a particular purpose. 
# The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
# In no event shall  its authors, or anyone else involved in the creation, production, or delivery 
# of the scripts be liable for any damages whatsoever 
# (including, without limitation, damages for loss of business profits, business interruption, 
# loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if the author has been advised of the possibility of such damages.
################################################
################################################
# License Notice :
# This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License 
# as published by the Free Software Foundation,  version 3 of the License Only.
# This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
# without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
# See the GNU General Public License for more details.
# You should have received a copy of the GNU General Public License along with this program. 
# If not, see https://www.gnu.org/licenses/gpl-3.0.html .
################################################
#-------------------------------------------------   

#
#### This Function is about Retrieving all devices part numbers  from VMWARE VSAN HCL DB ###
# 
# Variable to be used (and modified) are : 
# $DeviceType  which is the type of device to be selected Default : NVME
# $PartnerName which is the company to be selected Default : Hewlett Packard Enterprise
# $VsanSupport which is the type of VSAN used Default : vSAN ESA
# All these Variables are defaulted to my own Google Account (stored in files)


function ExtractNvme {
    param
    (
        [string]$DeviceType = "NVME"
        ,
        [string]$PartnerName = "Hewlett Packard Enterprise"
        ,
        [string]$VsanSupport = "vSAN ESA"
    )
    # URL for list of all vsan validated component from Vmware Hcl
    $uri = "https://partnerweb.vmware.com/service/vsan/all.json"
    # Retrieve in an Object all component selected in an Array
    $vsanHclNVME = Invoke-WebRequest -Uri $uri | ConvertFrom-Json | Select-Object -ExpandProperty Data | Select-object -ExpandProperty Ssd | Where-Object { $_.devicetype -eq $DeviceType -and $_.vsanSupport.mode[0] -eq $VsanSupport -and $_.partnername -eq $PartnerName }
    # Extract partNumber Only in a string format
    $PartNumberString = Out-String -InputObject $vsanHclNVME.partnumber
    # Remove all whitespace and comma and put one part number per line
    $PartNumber = $PartNumberString -replace ',', "`n" -replace " ", "" 
    $PartNumber | Set-Content .\Test.txt -Force
    return $PartNumber
}

#
#### This Function is about Getting a google Auth Token ###
# Variable to be used (and to be modified) are : 
# $certPath which is the certificate downloaded from Google account API
# $iss which is the account name used by Google account API
# $certPswd which is a password used by Google account api
# All these Variables are defaulted to my own Google Account (stored in files)

function GetGoogleToken {
    [CmdletBinding()]
    param 
    (
        [Parameter()]
        $certPath = ".\gogglesheetaccess-8bdfbac8c6f5.p12"
        ,
        [string] $iss = 'script@gogglesheetaccess.iam.gserviceaccount.com',
        $certPswd = 'notasecret'
    )
    # Set security protocol to TLS 1.2 to avoid TLS errors
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # Google API Authozation
    $scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
   
    try {
        $accessToken = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPswd -iss $iss
    }
    catch {
        $err = $_.Exception
        $err | Select-Object -Property *
        "Response: "
        $err.Response
    }
    return $accessToken
}

#
#### This Function is about interting ###
#### All lines contained in a file  called by default ./test.txt ###
#### To a Google Sheet ###
# Variable to be used (and to be modified) are : 
# $SpreadSheetFileName which is the file containing the id of the Google  spreadsheet created before
#   Default is ./spreadsheetId.txt
# $TextFileName which is the file to  store the retrived list of part Number
#   Default is ./Test.txt

function PutToSheet {
    [CmdletBinding()]
    param
    (
        [Parameter()]
        $SpreadSheetFileName = "./spreadsheetId.txt",
        $TextFileName = "./Test.txt"
    )
    # Get the Spreadsheet id from last line of text file
    $Spreadsheets= Get-Content $SpreadSheetFileName
    $NbId = $Spreadsheets.Count
    $SpreadsheetID = $Spreadsheets[$NbId-1]

    $import = New-Object System.Collections.ArrayList($null)    
    # Lets Clear the sheet first
    try {
        Clear-GSheetSheet -accessToken $accessToken -sheetName "Sheet1" -spreadSheetID $SpreadsheetID -Debug 
    }
    catch {
        $err = $_.Exception
        $err | Select-Object -Property *
        "Response: "
        $err.Response
    }
    # Import Txt and build ArrayList
    $import = New-Object System.Collections.ArrayList($null)
    $inputTxt = Get-Content $TextFileName
    $inputTxt | ForEach-Object { 
        $import.Add( @($_)) | Out-Null
    }
    try {
        $SheetResult = Set-GSheetData -accessToken $accessToken -rangeA1 "A1:C$($import.Count)" -sheetName "Sheet1" -spreadSheetID $SpreadsheetID -values $import -Debug 
    }
    catch {
        $err = $_.Exception
        $err | Select-Object -Property *
        "Response: "
        $err.Response
    }

    Write-Host $SheetResult
    return $SheetResult
}
#

# Main Program 

# need to be used for google API
Import-Module UMN-Google

# Call to extract NVME from VMWare VSAN HCL
$PartNumber = ExtractNvme
Write-Host "The Part Numbers are : " -BackgroundColor Blue
Write-Host $PartNumber -ForegroundColor Blue

# Call to get Token from Google API
$accessToken = GetGoogleToken
Write-Host "The Token is : " -BackgroundColor DarkCyan
Write-Host $accessToken -BackgroundColor DarkCyan

# Call to Put all line of the file in the SpreadSheet
$PutToSheet = PutToSheet
$PutToSheet | Format-List

