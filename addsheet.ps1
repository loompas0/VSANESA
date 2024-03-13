#-------------------------------------------------
# 
# Name : Addsheet.ps1
# What :
#   Create a new Google SpreadSheet using Google API
#   Set permission to a Google user
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

# Some information on manipulating google cloud spreadsheet could be found at :
#   https://jamesachambers.com/modify-google-sheets-using-powershell/
# Documentation  on UMN-Google  module could me found at : 
#   https://github.com/umn-microsoft-automation/UMN-Google/tree/master/docs
#

#-------------------------------------------------
# 
# Version : 1
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




Import-Module UMN-Google
#### This section is about Getting a google Auth Token ###
# Set security protocol to TLS 1.2 to avoid TLS errors
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Google API Authozation
$scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
## This section needs to be changed according to your Google setup ##
$certPath = ".\gogglesheetaccess-8bdfbac8c6f5.p12"
$iss = 'script@gogglesheetaccess.iam.gserviceaccount.com'
$certPswd = 'notasecret'
## End of section Change ##
try {
    $accessToken = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPswd -iss $iss
} catch {
    $err = $_.Exception
    $err | Select-Object -Property *
    "Response: "
    $err.Response
}

Write-Host "The access token to Google is :  " -ForegroundColor Yellow
Write-Host $accessToken -ForegroundColor Cyan

#### End of  section  Getting a google Auth Token ###

#### This section is about Creating a new Google SpreadSheet using Google API ###

## This section needs to be changed according to your Google setup ##
$MyEmailAdress ="remybernheim@gmail.com"
## End of section Change ##


$Title = Read-Host "Enter the name of the Worksheet you want to create : " 
# $Sheet = Read-Host "Please Enter the Sheet Name : "
$SpreadsheetID = (New-GSheetSpreadSheet -accessToken $accessToken -title $Title).spreadsheetId

#### End of  section Creating a new Google SpreadSheet using Google API ###

#### This section is about Setting permission to a Google user using Google API ###

Set-GFilePermissions -accessToken $accessToken -fileID $SpreadsheetID -role writer -type user -emailAddress $MyEmailAdress
Write-Host $SpreadsheetID
Add-Content -Path .\spreadsheetId.txt -Value ($SpreadsheetID) -PassThru

#### End of section Setting permission to a Google user using Google API ###
