#
# Script:	    pssService-backup.ps1
# Task:		    Create a backup of your teams SharePoint lists in subfolder Microsoft SharePoint and save a copy of your "Arbeitsordner" on SharePoint
# 
# This file is part of the software solution pobvol Service Solution. 
# pobvol Service Solution is Free Software, delivered as open source. You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. The solution is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with the solution. If not, see <http://www.gnu.org/licenses/>. 
# Copyright © 2025 Volker Pobloth
# Web: https://pobvol.com/
#
# ---------------------------------------------------------------------------------------
Function LogWrite
# ---------------
{
    Param ([string]$strMessage)
    $strCurrentDateTime = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $strValue = $strCurrentDateTime + "  " + $strMessage
    Add-content $strLogfile -value $strValue
}
# ---------------------------------------------------------------------------------------
Function GetSPList
# ----------------
{
    Param ([string]$SPList)
    $fileName = $strPathSharePoint +"\" + $SPList +'.xml'

    $strMessage = "    Downloading definition for " +$SPList
    Write-Host $strMessage
    LogWrite $strMessage

    $strMessage = "    " +$fileName
    Write-Host $strMessage
    LogWrite $strMessage

    # Write list definition to a xml file
    Get-PnPSiteTemplate -Force -Out $fileName -ListsToExtract "$SPList" -ExcludeHandlers ApplicationLifecycleManagement, AuditSettings, ComposedLook, ContentTypes, CustomActions, ExtensibilityProviders, Features, ImageRenditions, Navigation, None, PageContents, Pages, PropertyBagEntries, Publishing, RegionalSettings, SearchSettings, SiteFooter, SiteHeader, SitePolicy, SiteSecurity, SiteSettings, SupportedUILanguages, SyntexModels, Tenant, TermGroups, Theme, WebApiPermissions, WebSettings, Workflows

}
#---------------------------------------------------------------------------------------
Function ExtractSPList
{
    Param ([string]$SPList)
    $fileName = $strPathSharePoint +"\" + $SPList +'.xml'

    $strMessage = "    Downloading data for " +$SPList
    Write-Host $strMessage
    LogWrite $strMessage

    $strMessage = "    " +$fileName
    Write-Host $strMessage
    LogWrite $strMessage

    # Add the data to the xml file
    Add-PnPDataRowsToSiteTemplate -Path $fileName -List "$SPList"

}
#---------------------------------------------------------------------------------------

#Current path
$MyPath = $PSScriptRoot

#Variables
$Environment = ""
$SPPath = "null"
$SPReportsFolder = "null"
$Protokoll = "null"
$strError = ""
#$FlagFound = ""
$SPDomain = "null"
$SPTeam = "null"
#$SPSite = "null"
#$SPSyncFolder = "null"
$Protokoll = "null"
$PnPRocksId = "null"
$strError = ""

#Log-File
$strLogfile = $MyPath + "\backup-log.txt"

$strMessage = "pobvol Service Solution: pssService-backup.ps1 started ...."
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Current folder: " +$MyPath
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Log file: " +$strLogfile
Write-Host $strMessage
LogWrite $strMessage

#Pause

#Start Excel
$strMessage = "    Starting Excel..."
Write-Host $strMessage
LogWrite $strMessage
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #Soll Excel angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?

#Load system settings
$strMessage = "    Reading settings from ..\pobvol-sync\pssService-Settings.xlsx"
Write-Host $strMessage
LogWrite $strMessage

$fileName = $MyPath + '\pobvol-sync\pssService-Settings.xlsx'
$workbook = $excel.Workbooks.Open($fileName)
$sheet = $workbook.worksheets.Item(1)
for ($i=1; $i -le 100; $i++)
{
    #$strMessage = "Reading row from ..\pobvol-sync\pssService-Settings.xlsx: " + $i
    #Write-Host $strMessage
    #LogWrite $strMessage

    # OneDrive folder
    If ($sheet.Cells.Item($i, 2).Text -eq 'OneDrive-Folder') {
        $SPPath = $sheet.Cells.Item($i, 3).Text
    }
    
    # Service reports folder
    If ($sheet.Cells.Item($i, 2).Text -eq 'Service reports are saved in folder') {
        $SPReportsFolder = $sheet.Cells.Item($i, 3).Text
    }
    
    # Full log? true, false
    If ($sheet.Cells.Item($i, 2).Text -eq 'Detailed logging') {
        $Protokoll = $sheet.Cells.Item($i, 3).Text
    }

    # Environment flag: DEV, TST, PROD
    If ($sheet.Cells.Item($i, 2).Text -eq 'pssService Environment') {
        $Environment = $sheet.Cells.Item($i, 3).Text
    }

    # Domain
    If ($sheet.Cells.Item($i, 2).Text -eq 'SharePoint.Domain') {
        $SPDomain = $sheet.Cells.Item($i, 3).Text
    }
    
    # Team
    If ($sheet.Cells.Item($i, 2).Text -eq 'SharePoint.Team') {
        $SPTeam = $sheet.Cells.Item($i, 3).Text
    }

    # PnP Rocks Id
    If ($sheet.Cells.Item($i, 2).Text -eq 'PnP Rocks Id') {
        $PnPRocksId = $sheet.Cells.Item($i, 3).Text
    }

    #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
    #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
}
$workbook.close($false)

$strMessage = "    Parameters found:"
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    SharePoint.Domain: " +$SPDomain
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    SharePoint.Team: " +$SPTeam
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    PnP Rocks Id: " +'found'
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Environment: " +$Environment
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Full log: " +$Protokoll
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    OneDrive sync folder: " +$SPPath
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Service reports saved in sub folder: " +$SPReportsFolder
Write-Host $strMessage
LogWrite $strMessage

#Pause

#Check Parameter
If( $SPDomain -eq "null" -OR $SPDomain -eq "Enter your value" -OR
    $SPTeam -eq "null" -OR $SPTeam -eq "Enter your value" -OR
    $PnPRocksId -eq "null" -OR $PnPRocksId -eq "Enter your value" -OR
    $SPPath -eq "null" -OR $SPPath -eq "Enter your value" -OR
    $SPReportsFolder -eq "null" -OR $SPReportsFolder -eq "Enter your value" -OR
    $Protokoll -eq "null" -OR $Protokoll -eq "Enter your value"){

    $strError = "Error"

    $strMessage = "    Processing service data skipped. Missing parameters!"
    Write-Host $strMessage
    LogWrite $strMessage

    $strMessage = "    Please check your entries in ..\pobvol-sync\pssService-Settings.xlsx"
    Write-Host $strMessage
    LogWrite $strMessage
}

#Remove Excel
$strMessage = "    Removing Excel ..."
Write-Host $strMessage
LogWrite $strMessage
$excel.Quit()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable excel -ErrorAction SilentlyContinue



#Nachsehen, ob die SP-Seite existiert
#If($Environment -ne "DEV" -AND $strError -eq ""){
If($strError -eq ""){
 
    #Provide your SharePoint Online Admin center URL
    #$AdminSiteURL = "https://<Tenant_Name>-admin.sharepoint.com"
    $AdminSiteURL = "https://" + $SPDomain +"-admin.sharepoint.com/"
        
    $strMessage = "    Connecting to SharePoint Admin Site " +$AdminSiteURL
    Write-Host $strMessage
    LogWrite $strMessage

    #Connect-PnPOnline -Url "https://<yourtenant>-admin.sharepoint.com/" -Interactive -ClientId <Your PnP Rocks Id>
    #How to get your own PnP Rocks Id?
    #https://pnp.github.io/powershell/articles/registerapplication
    #https://pnp.github.io/powershell/articles/authentication.html
    
    Connect-PnPOnline -Url $AdminSiteURL -Interactive -ClientId $PnPRocksId
    
    #Get all site collections
    #Thanks to Morgan Tech Space
    #https://morgantechspace.com/2021/10/get-all-sites-and-sub-sites-in-sharepoint-online-using-pnp-powershell.html
    #The below command gets only modern Team & Communication sites

    $SiteAlias = $SPTeam
    $strMessage = "    Searching for team site " +$SiteAlias
    Write-Host $strMessage
    LogWrite $strMessage

    $Sites = Get-PnPTenantSite
    #$TotoalSites = $Sites.Count
    $SPSiteExist = "false"
    ForEach($Site in $Sites)
    {
        #$i++;
        #Write-Host "    Site Name:" $($Site.Title)
        #Write-Host "    Site URL:" $($Site.Url)
        If($($Site.Title) -eq $SiteAlias){
            $SPSiteExist = "true"
            $SPSiteURL = $($Site.Url)
        }
    }
    #Write-Host "Site found:" $SPSiteExist
    #Write-Host "strError:" $strError
    If($SPSiteExist -ne "true"){
        $strMessage = "    Error! Missing the SharePoint-Website " +$SiteAlias
        Write-Host $strMessage
        LogWrite $strMessage
        
        $strError = "Error"
    }    
}


#Einen Backup der SP-Listen im Ordner yyyy-mm-dd speichern
#If($Environment -ne "DEV" -AND $SPSiteExist -eq "true" -AND $strError -eq ""){
If($SPSiteExist -eq "true" -AND $strError -eq ""){

    #Listendefinitionen liegen im Unterordner Microsoft SharePoint
    $Date = Get-Date -Format yyyy-MM-dd
    $strPathSharePoint = $MyPath + '\microsoft-sharepoint\' + $Date
    
    #wenns den Folder nicht gibt,...
    if (!(Test-Path -Path $strPathSharePoint)) {
    
        #Folder anlegen
        $strMessage = "    Creating local backup-folder " +$strPathSharePoint
        Write-Host $strMessage
        LogWrite $strMessage

        New-Item -Path $strPathSharePoint -Type Directory -Force -ErrorAction Stop | Out-Null
    }
    
    #wenns den Folder nicht gibt,...
    if (!(Test-Path -Path $strPathSharePoint)) {
        $strMessage = "    Could not create the local backup-folder " +$strPathSharePoint
        Write-Host $strMessage
        LogWrite $strMessage

        $strError = "Error"
    }

    #die SP-Listen runterladen
    If($strError -eq ""){
        $strMessage = "    Downloading SharePoint lists from " +$SPSiteURL
        Write-Host $strMessage
        LogWrite $strMessage

        Connect-PnPOnline -Url $SPSiteURL -Interactive -ClientId $PnPRocksId

        #Listen exportieren
        $SPList = "ArchivServiceberichte"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ArchivServicevorgaenge"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ArchivServicevorgaengeP"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "BevorzugteSprachen"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Bilder"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Einstellungen"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "EinstellungenBenutzer"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Fahrtbericht"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Kundeninventar"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Leistungsabrechnung"; GetSPList $SPList; ExtractSPList $SPList        
        $SPList = "Serviceauftraege"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ServiceauftraegeP"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Serviceberichte"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Servicekunden"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Servicevertraege"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ServicevertraegeP"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ServicevertraegeAbrechnung"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "Servicevorgaenge"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ServicevorgaengeP"; GetSPList $SPList; ExtractSPList $SPList
        $SPList = "ServicevorgaengeE"; GetSPList $SPList; ExtractSPList $SPList

        $SPList = "Formatbibliothek"; GetSPList $SPList
        $SPList = "Formularvorlagen"; GetSPList $SPList
        $SPList = "Websiteobjekte"; GetSPList $SPList
        $SPList = "Websiteseiten"; GetSPList $SPList
        # $SPList = "Dokumente"; GetSPList $SPList

    }

}

#---------------------------------------------------------------------------
#Save your files on SharePoint 
#---------------------------------------------------------------------------
If($Environment -ne "DEV" -AND $strError -eq ""){
    Write-Host "    " 
    Write-Host "Copy new and changed files to your SharePoint document library"
    $strSource = $MyPath + "\*.*"
    $strTarget = $SPPath + "\Backup\pssService"
    xcopy $strSource $strTarget /S /M /E /V /C /I /G /R /Y 
}

#---------------------------------------------------------------------------

$strMessage = "pobvol Service Solution: pssService-backup.ps1 finished."
Write-Host $strMessage
LogWrite $strMessage

Pause


#---------------------------------------------------------------------------------------
# Ein paar Coding-Hinweise
#-------------------------
# -eq	gleich
# -ne	ungleich
# -lt	kleiner
# -le	kleiner oder gleich
# -gt	groesser
# -ge	groesser oder gleich
# -AND
# -Or
# if( $? ) {
# 	# True, last operation succeeded
#}
# if( !$? ) {
#	# Not True, last operation failed
#}
