#
# Script:	    pssService-setup.ps1
# Task:		    Setup pobvol Service Solution. Create/update the SharePoint team page and the SharePoint lists for the solution.
# 
#This file is part of the software solution pobvol Service Solution. 
#pobvol Service Solution is Free Software, delivered as open source. You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. The solution is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with the solution. If not, see <http://www.gnu.org/licenses/>. 
#Copyright © 2025 Volker Pobloth
#Web: https://pobvol.com/
#
#---------------------------------------------------------------------------------------
Function InvokeSPList
{
    Param ([string]$SPList)
    
    #Gibt es die Liste schon?
    #$SPListExist = "false"
    #$Lists = Get-PnPList
    #foreach($List in $Lists) { 
    #    If($($List.Title) -eq $SPList){
    #        $strMessage = $List.Title + " found"
    #        Write-Host $strMessage
    #        LogWrite $strMessage
    #        $SPListExist = "true"
    #    }
    #}
    
    #Die Liste wird nur erstellt, wenn sie noch nicht vorhanden ist
    #If($SPListExist -eq "false"){

    #If the file exists, process it
    $fileName = $MyPathSharePoint +"\" + $SPList +'.xml'
    $fileNameProcessed = $MyPathSharePoint +"\" + $SPList +'.xml.bak'
    if (Test-Path -Path $fileName -PathType Leaf) {
        try {

            $strMessage = "Creating/updating SharePoint list " + $SPList
            Write-Host $strMessage
            LogWrite $strMessage

            $strMessage = $fileName
            Write-Host $strMessage
            LogWrite $strMessage

            Invoke-PnPSiteTemplate -Path $fileName

            Copy-Item $fileName $fileNameProcessed -Force
            if (Test-Path -Path $fileNameProcessed -PathType Leaf) {
                try {
                    Remove-Item $fileName -Force
                }
                catch {
                    throw $_.Exception.Message
                }
            }
        }
        catch {
            throw $_.Exception.Message
        }
    }

}

Function UpdateSPList
{
    Param ([string]$SPList)
    $fileName = $MyPathSharePoint +"\" + $SPList +'.xml'
    $fileNameProcessed = $MyPathSharePoint +"\" + $SPList +'.xml.bak'
    
    #If the file exists, process it
    if (Test-Path -Path $fileName -PathType Leaf) {
        try {
        
            $strMessage = "Updating SharePoint list " + $SPList
            Write-Host $strMessage
            LogWrite $strMessage
            
            $strMessage = $fileName
            Write-Host $strMessage
            LogWrite $strMessage

            Invoke-PnPSiteTemplate -Path $fileName

            Copy-Item $fileName $fileNameProcessed -Force
            if (Test-Path -Path $fileNameProcessed -PathType Leaf) {
                try {
                    Remove-Item $fileName -Force
                }
                catch {
                    throw $_.Exception.Message
                }
            }
        }
        catch {
            throw $_.Exception.Message
        }
    }
    
}
Function LogWrite
{
    Param ([string]$strMessage)
    $strCurrentDateTime = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $strValue = $strCurrentDateTime + "  " + $strMessage
    Add-content $strLogfile -value $strValue
}
#---------------------------------------------------------------------------------------


#Aktuelles Verzeichnis ermitteln (das ist das Verzeichnis, in welchem das Skript ausgefuehrt wird)
#Current path
$MyPath = $PSScriptRoot
$strLogfile = $MyPath + "\setup-log.txt"

$SPPath = ""
$SPDomain = ""
$SPTeam = ""
#$SPSite = ""
#$SPSyncFolder = ""
$SPReportsFolder = ""
$SPContractsFolder = ""
$Protokoll = ""
$PnPRocksId = ""
$strError = ""
$sourceFilename = ""
$sourceFilenameNew = ""
$fileName = ""
$entryFound = "false"
$lastEntryRow = 0
$Environment = ""

$strMessage = "pobvol Service Solution: pssService-setup.ps1 started ...."
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "Current folder: " + $MyPath 
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "Log file: " + $strLogfile
Write-Host $strMessage
LogWrite $strMessage

#---------------------------------------------------------------------------
# Quit Excel and stop all Excel tasks
$Processes = get-process #Powershell to kill excel and access
ForEach($ProcessName in $Processes){
    If($ProcessName -eq "MSACCESS"){
        Stop-Process -Name "MSACCESS" -Force
    }
    If($ProcessName -eq "EXCEL"){
        Stop-Process -Name "EXCEL" -Force
    }
    If($ProcessName -eq "excel"){
        Stop-Process -Name "excel" -Force
    }
}

#---------------------------------------------------------------------------
# Setup pssService-link-to-customer-db.accdb
#---------------------------------------------------------------------------

# Create file from template
$strSource = $MyPath + "\pobvol-sync\pssService-link-to-customer-db-Template.accdb"
$strTarget = $MyPath + "\pobvol-sync\pssService-link-to-customer-db.accdb"
if (!(Test-Path -Path $strTarget)) {
    #Datei anlegen
    $strMessage = "Creating file " +$strTarget
    Write-Host $strMessage
    LogWrite $strMessage
    try {
        Copy-Item $strSource -Destination $strTarget
    }
    catch {
        #throw $_.Exception.Message
        $strMessage = "    " +$_.Exception.Message
        LogWrite $strMessage
    }
}


#---------------------------------------------------------------------------
# Setup pssService-Articles.xlsx
#---------------------------------------------------------------------------

# Create file from template
$strSource = $MyPath + "\pobvol-sync\pssService-Articles-Template.xlsx"
$strTarget = $MyPath + "\pobvol-sync\pssService-Articles.xlsx"
if (!(Test-Path -Path $strTarget)) {
    #Datei anlegen
    $strMessage = "Creating file " +$strTarget
    Write-Host $strMessage
    LogWrite $strMessage
    try {
        Copy-Item $strSource -Destination $strTarget
    }
    catch {
        #throw $_.Exception.Message
        $strMessage = "    " +$_.Exception.Message
        LogWrite $strMessage
    }
}

#---------------------------------------------------------------------------
# Setup pssService-Checklists.xlsx
#---------------------------------------------------------------------------

# Create file from template
$strSource = $MyPath + "\pobvol-sync\pssService-Checklists-Template.xlsx"
$strTarget = $MyPath + "\pobvol-sync\pssService-Checklists.xlsx"
if (!(Test-Path -Path $strTarget)) {
    #Datei anlegen
    $strMessage = "Creating file " +$strTarget
    Write-Host $strMessage
    LogWrite $strMessage
    try {
        Copy-Item $strSource -Destination $strTarget
    }
    catch {
        #throw $_.Exception.Message
        $strMessage = "    " +$_.Exception.Message
        LogWrite $strMessage
    }
}

#---------------------------------------------------------------------------
# Setup pssService-ZLanguages-Template.xlsx
#---------------------------------------------------------------------------

# Create file from template
$strSource = $MyPath + "\pobvol-sync\pssService-ZLanguages-Template.xlsx"
$strTarget = $MyPath + "\pobvol-sync\pssService-ZLanguages.xlsx"
if (!(Test-Path -Path $strTarget)) {
    #Datei anlegen
    $strMessage = "Creating file " +$strTarget
    Write-Host $strMessage
    LogWrite $strMessage
    try {
        Copy-Item $strSource -Destination $strTarget
    }
    catch {
        #throw $_.Exception.Message
        $strMessage = "    " +$_.Exception.Message
        LogWrite $strMessage
    }
}

#---------------------------------------------------------------------------
# Setup pssService-Settings.xlsx
#---------------------------------------------------------------------------

# Create file from template
$strSource = $MyPath + "\pobvol-sync\pssService-Settings-Template.xlsx"
$strTarget = $MyPath + "\pobvol-sync\pssService-Settings.xlsx"
if (!(Test-Path -Path $strTarget)) {
    #Datei anlegen
    $strMessage = "Creating file " +$strTarget
    Write-Host $strMessage
    LogWrite $strMessage
    try {
        Copy-Item $strSource -Destination $strTarget
    }
    catch {
        #throw $_.Exception.Message
        $strMessage = "    " +$_.Exception.Message
        LogWrite $strMessage
    }
}

#---------------------------------------------------------------------------
# New entries and updates for pssService-Settings.xlsx
#---------------------------------------------------------------------------

$sourceFilename = $MyPath + "\pobvol-sync\setup-tasks\pssService-UpdatesForSystemSettings.xml"
if (Test-Path -Path $sourceFilename) {

    $strMessage = "Processing updates for pssService-Settings.xlsx"
    Write-Host $strMessage
    LogWrite $strMessage

    #Open pssService-Settings.xlsx
    $strTarget = $MyPath + "\pobvol-sync\pssService-Settings.xlsx"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true #Soll Excel angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
    $workbook = $excel.Workbooks.Open($strTarget)
    $sheet = $workbook.worksheets.Item(1)

    #get the last entry in pssService-Settings.xlsx
    $lastEntryRow = 0
    for ($i=1; $i -le 200; $i++)
    {
        If ($sheet.Cells.Item($i, 1).Text -ne "") {
            $lastEntryRow = $i
        }
    }
    If($lastEntryRow -eq 0 -Or $lastEntryRow -eq 200){
        $strError = "Error: last entry is 0 or 200"
        $strMessage = "    "+$strError
        Write-Host $strMessage
        LogWrite $strMessage
    }

    If($strError -eq ""){
        #Read 'xml file with update information'
        $xml = Get-Content $sourceFilename -Raw
        $xmlEntries = [XML]$xml
        foreach($Entry in $xmlEntries.Settings.Entry) { 
            $strMessage = "    Entry and value: "+$Entry.CellB + ": " +$Entry.CellC
            Write-Host $strMessage
            LogWrite $strMessage

            #update entry if alreayd in the list of settings
            $entryFound = "false"
            for ($i=1; $i -le $lastEntryRow; $i++)
            {
                If ($sheet.Cells.Item($i, 1).Text -eq $Entry.CellA -AND
                    $sheet.Cells.Item($i, 2).Text -eq $Entry.CellB) {
                    $entryFound = "true"
                    $sheet.Cells.Item($i, 3) = $Entry.CellC
                    $sheet.Cells.Item($i, 4) = $Entry.CellD
                    $sheet.Cells.Item($i, 5) = $Entry.CellE
                    $i=$lastEntryRow
                }
                #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
                #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
            }
            If($entryFound -eq "false"){
                $lastEntryRow = $lastEntryRow + 1
                $sheet.Cells.Item($lastEntryRow, 1) = $Entry.CellA
                $sheet.Cells.Item($lastEntryRow, 2) = $Entry.CellB
                $sheet.Cells.Item($lastEntryRow, 3) = $Entry.CellC
                $sheet.Cells.Item($lastEntryRow, 4) = $Entry.CellD
                $sheet.Cells.Item($lastEntryRow, 5) = $Entry.CellE
            }
        }
    }

    #Save and close Excel file ($true = save, $false = do not save)
    $workbook.close($true)
    $excel.Quit() #calling quit() on the object makes the GUI of the application disappear but sometimes the process is still running
    Remove-Variable excel -ErrorAction SilentlyContinue
    $Processes = get-process #Powershell to kill excel and access
    ForEach($ProcessName in $Processes){
        If($ProcessName -eq "MSACCESS"){
            Stop-Process -Name "MSACCESS" -Force
        }
        If($ProcessName -eq "EXCEL"){
            Stop-Process -Name "EXCEL" -Force
        }
    }

    #Rename 'xml file with update information' if all entries processed correctly
    If($strError -eq ""){
        $sourceFilenameNew = $MyPath +"\pobvol-sync\setup-tasks\pssService-UpdatesForSystemSettings.xml.bak"
        try {
            Copy-Item $sourceFilename $sourceFilenameNew -Force
            if (Test-Path -Path $sourceFilenameNew -PathType Leaf) {
                try {
                    Remove-Item $sourceFilename -Force
                    $strMessage = "    Saved update file as: "+$sourceFilenameNew
                    Write-Host $strMessage
                    LogWrite $strMessage
                }
                catch {
                    #throw $_.Exception.Message
                    $strMessage = "    " +$_.Exception.Message
                    LogWrite $strMessage
                }
            }
        }
        catch {
            #throw $_.Exception.Message
            $strMessage = "    " +$_.Exception.Message
            LogWrite $strMessage
        }
    }
}

#---------------------------------------------------------------------------
# Load settings from pobvol-sync\pssService-Settings.xlsx
#---------------------------------------------------------------------------

$strMessage = "Reading settings"
Write-Host $strMessage
LogWrite $strMessage

$fileName = $MyPath + '\pobvol-sync\pssService-Settings.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #Soll Excel angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
$workbook = $excel.Workbooks.Open($fileName)
$sheet = $workbook.worksheets.Item(1)
for ($i=1; $i -le 100; $i++)
{
    #$strMessage = "Reading row from pobvol-sync\pssService-Settings.xlsx: " + $i
    #Write-Host $strMessage
    #LogWrite $strMessage

    # OneDrive folder
    If ($sheet.Cells.Item($i, 2).Text -eq 'OneDrive-Folder') {
        $SPPath = $sheet.Cells.Item($i, 3).Text
    }

    # Service Reports folder
    If ($sheet.Cells.Item($i, 2).Text -eq 'Service reports are saved in folder') {
        $SPReportsFolder = $sheet.Cells.Item($i, 3).Text
    }
    
    # Contracts folder
    If ($sheet.Cells.Item($i, 2).Text -eq 'Contracts are saved in folder') {
        $SPContractsFolder = $sheet.Cells.Item($i, 3).Text
    }
    
    # DEV / TST / PROD
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

    # Log
    If ($sheet.Cells.Item($i, 2).Text -eq 'Detailed logging') {
        $Protokoll = $sheet.Cells.Item($i, 3).Text
    }

    # PnP Rock Id
    If ($sheet.Cells.Item($i, 2).Text -eq 'PnP Rocks Id') {
        $PnPRocksId = $sheet.Cells.Item($i, 3).Text
    }

    #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
    #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
}
#Excel-Datei schliessen (nicht speichern)
$workbook.close($false)
$excel.Quit() #calling quit() on the object makes the GUI of the application disappear but sometimes the process is still running
Remove-Variable excel -ErrorAction SilentlyContinue
$Processes = get-process #Powershell to kill excel and access
ForEach($ProcessName in $Processes){
    If($ProcessName -eq "MSACCESS"){
        Stop-Process -Name "MSACCESS" -Force
    }
    If($ProcessName -eq "EXCEL"){
        Stop-Process -Name "EXCEL" -Force
    }
}

$strMessage = "Parameters found: "
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    SharePoint.Domain: " + $SPDomain
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    SharePoint.Team: " + $SPTeam
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Detailed log: " + $Protokoll
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    OneDrive sync folder: " + $SPPath
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Service reports saved in sub folder: " + $SPReportsFolder
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Contracts saved in sub folder: " + $SPContractsFolder
Write-Host $strMessage
LogWrite $strMessage

$strMessage = "    Environment: " + $Environment
Write-Host $strMessage
LogWrite $strMessage

If( $SPDomain -eq $null -OR $SPDomain -eq "" -OR $SPDomain -eq "Enter your value" -OR
    $SPTeam -eq $null -OR $SPTeam -eq "" -OR $SPTeam -eq "Enter your value" -OR
    $PnPRocksId -eq $null -OR $PnPRocksId -eq "" -OR $PnPRocksId -eq "Enter your value" -OR
    $SPReportsFolder -eq $null -OR $SPReportsFolder -eq "" -OR $SPReportsFolder -eq "Enter your value" -OR
    $SPContractsFolder -eq $null -OR $SPContractsFolder -eq "" -OR $SPContractsFolder -eq "Enter your value" -OR
    $Protokoll -eq $null -OR $Protokoll -eq "" -OR $Protokoll -eq "Enter your value"){

    $strError = "Error"

    $strMessage = "    Processing skipped. Missing parameters!"
    Write-Host $strMessage
    LogWrite $strMessage

    $strMessage = "    Please check your entries in " +$MyPath +'\pobvol-sync\pssService-Settings.xlsx'
    Write-Host $strMessage
    LogWrite $strMessage

    #Pause

}

#---------------------------------------------------------------------------
# Processing pobvol-sync\windows-task-scheduler\pssService-check-for-new-data-task.xml
#---------------------------------------------------------------------------

If($strError -eq ""){
    $filename = $PSScriptRoot +"\pobvol-sync\windows-task-scheduler\pssService-check-for-new-data-task.xml"
    $filenameNew = $PSScriptRoot +"\pobvol-sync\windows-task-scheduler\pssService-check-for-new-data-task.xml.bak"
    if (Test-Path -Path $filename) {

        $strMessage = "Processing xml file for the Windows Task Scheduler: " + $filename
        Write-Host $strMessage
        LogWrite $strMessage

        #Open XML file
        $xmldata = New-Object xml
        $xmldata.Load( (Convert-Path $filename) )

        #Replace the URI in the xml file
        #$string = '"' + "\pssService " +$SPTeam +'"'
        #$xmldata.Task.RegistrationInfo.URI = $string
        
        #Replace the start command in the xml file
        $string = '"'+$PSScriptRoot + "\pobvol-sync\pssService-start-check-for-new-data.bat"+'"'
        $xmldata.Task.Actions.Exec.Command = $string

        $strMessage = "    New value for Task.Actions.Exec.Command: "+$string
        Write-Host $strMessage
        LogWrite $strMessage
        

        #Save XML file
        $xmldata.Save((Resolve-Path $filename).Path)
        Remove-Variable xmldata -ErrorAction SilentlyContinue

        #Rename xml to xml.bak ==> file won't be processed again and again
        #try {
        #    Copy-Item $fileName $fileNameNew -Force
        #    if (Test-Path -Path $fileNameNew -PathType Leaf) {
        #        try {
        #            Remove-Item $fileName -Force
        #            $strMessage = "    Saved file as: "+$filenameNew
        #            Write-Host $strMessage
        #            LogWrite $strMessage
        #        }
        #        catch {
        #            #throw $_.Exception.Message
        #            $strMessage = "    " +$_.Exception.Message
        #            LogWrite $strMessage
        #        }
        #    }
        #}
        #catch {
        #    #throw $_.Exception.Message
        #    $strMessage = "    " +$_.Exception.Message
        #    LogWrite $strMessage
        #}
    }
}

#---------------------------------------------------------------------------
# Processing pobvol-sync\windows-task-scheduler\pssService-refresh-pivot-reports-task.xml
#---------------------------------------------------------------------------

If($strError -eq ""){
    $filename = $PSScriptRoot +"\pobvol-sync\windows-task-scheduler\pssService-refresh-pivot-reports-task.xml"
    $filenameNew = $PSScriptRoot +"\pobvol-sync\windows-task-scheduler\pssService-refresh-pivot-reports-task.xml.bak"
    if (Test-Path -Path $filename) {

        $strMessage = "Processing xml file for the Windows Task Scheduler: " + $filename
        Write-Host $strMessage
        LogWrite $strMessage

        #Open XML file
        $xmldata = New-Object xml
        $xmldata.Load( (Convert-Path $filename) )

        #Replace the URI in the xml file
        #$string = '"' + "\pssService " +$SPTeam +'"'
        #$xmldata.Task.RegistrationInfo.URI = $string
        
        #Replace the start command in the xml file
        $string = '"'+$PSScriptRoot + "\pobvol-sync\pssService-start-refresh-pivot-reports.bat"+'"'
        $xmldata.Task.Actions.Exec.Command = $string

        $strMessage = "    New value for Task.Actions.Exec.Command: "+$string
        Write-Host $strMessage
        LogWrite $strMessage

        #Save XML file
        $xmldata.Save((Resolve-Path $filename).Path)
        Remove-Variable xmldata -ErrorAction SilentlyContinue

        #Rename xml to xml.bak ==> file won't be processed again and again
        #try {
        #    Copy-Item $fileName $fileNameNew -Force
        #    if (Test-Path -Path $fileNameNew -PathType Leaf) {
        #        try {
        #            Remove-Item $fileName -Force
        #            $strMessage = "    Saved file as: "+$filenameNew
        #            Write-Host $strMessage
        #            LogWrite $strMessage
        #        }
        #        catch {
        #            #throw $_.Exception.Message
        #            $strMessage = "    " +$_.Exception.Message
        #            LogWrite $strMessage
        #        }
        #    }
        #}
        #catch {
        #    #throw $_.Exception.Message
        #    $strMessage = "    " +$_.Exception.Message
        #    LogWrite $strMessage
        #}
    }
}

#---------------------------------------------------------------------------
# Create SharePoint page and lists
#---------------------------------------------------------------------------

If($strError -eq ""){
    #Provide your SharePoint Online Admin center URL
    #$AdminSiteURL = "https://<Tenant_Name>-admin.sharepoint.com"
    $AdminSiteURL = "https://" + $SPDomain +"-admin.sharepoint.com/"
        
    $strMessage = "Connecting to SharePoint Admin page " + $AdminSiteURL
    Write-Host $strMessage
    LogWrite $strMessage

    #Connect-PnPOnline -Url "https://<yourtenant>-admin.sharepoint.com/" -Interactive -ClientId <Your PnP Rocks Id>
    #How to get your own PnP Rocks Id?
    #https://pnp.github.io/powershell/articles/registerapplication
    Connect-PnPOnline -Url $AdminSiteURL -Interactive -ClientId $PnPRocksId
    
    #Get all site collections
    #Thanks to Morgan Tech Space
    #https://morgantechspace.com/2021/10/get-all-sites-and-sub-sites-in-sharepoint-online-using-pnp-powershell.html
    #The below command gets only modern Team & Communication sites
    $SiteAlias = $SPTeam
    $GroupAlias = $SPTeam

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
        }
    }
    If($SPSiteExist -eq "true"){
        $strMessage = "SharePoint-Website " + $SiteAlias + " exists"
        Write-Host $strMessage
        LogWrite $strMessage
    }else {
        $strMessage = "Creating SharePoint-Website " + $SiteAlias
        Write-Host $strMessage
        LogWrite $strMessage
        #Thanks to Morgang Tech Space
        #https://morgantechspace.com/2022/09/create-a-new-sharepoint-online-site-using-pnp-powershell.html
        New-PnPSite -Type TeamSite -Title $SPTeam -SiteAlias $SiteAlias -Alias $GroupAlias
    }    

    #Wenn die Seite existiert, werden die Listen erstellt
    $SPSiteExist = "false"
    $Sites = Get-PnPTenantSite
    ForEach($Site in $Sites)
    {
        If($($Site.Title) -eq $SiteAlias){
            $SPSiteExist = "true"
            $SPSiteURL = $($Site.Url)
        }
    }
    #Write-Host "SPSiteExist" $SPSiteExist

    #Wenn die SP-Site existiert, die SP-Listen erstellen / aktualisieren
    If($SPSiteExist -eq "true"){
        
        $strMessage = "Creating/updating SharePoint lists on SharePoint page " + $SPSiteURL
        Write-Host $strMessage
        LogWrite $strMessage

	    Connect-PnPOnline -Url $SPSiteURL -Interactive -ClientId $PnPRocksId

        #Listendefinitionen liegen im Unterordner microsoft-sharepoint
        $MyPathSharePoint = $MyPath + '\microsoft-sharepoint'
        #Set-Location -Path $MyPathSharePoint
        
        $SPList = "Dokumente"; InvokeSPList $SPList
        $SPList = "Formatbibliothek"; InvokeSPList $SPList
        $SPList = "Formularvorlagen"; InvokeSPList $SPList
        $SPList = "Websiteobjekte"; InvokeSPList $SPList
        $SPList = "Websiteseiten"; InvokeSPList $SPList

        #Listen erstellen und aktualisieren
        $SPList = "ArchivServiceberichte"; InvokeSPList $SPList
        $SPList = "ArchivServicevorgaenge"; InvokeSPList $SPList
        $SPList = "ArchivServicevorgaengeP"; InvokeSPList $SPList
        $SPList = "Artikel"; InvokeSPList $SPList
        $SPList = "BevorzugteSprachen"; InvokeSPList $SPList
        $SPList = "Bilder"; InvokeSPList $SPList
        $SPList = "Checklisten"; InvokeSPList $SPList
        $SPList = "Einstellungen"; InvokeSPList $SPList
        $SPList = "EinstellungenBenutzer"; InvokeSPList $SPList
        $SPList = "Fahrtbericht"; InvokeSPList $SPList
        $SPList = "Kundeninventar"; InvokeSPList $SPList
        $SPList = "Leistungsabrechnung"; InvokeSPList $SPList
        $SPList = "Serviceauftraege"; InvokeSPList $SPList
        $SPList = "ServiceauftraegeP"; InvokeSPList $SPList
        $SPList = "Serviceberichte"; InvokeSPList $SPList
        $SPList = "Servicekunden"; InvokeSPList $SPList
        $SPList = "Servicevertraege"; InvokeSPList $SPList
        $SPList = "ServicevertraegeP"; InvokeSPList $SPList
        $SPList = "ServicevertraegeAbrechnung"; InvokeSPList $SPList
        $SPList = "Servicevorgaenge"; InvokeSPList $SPList
        $SPList = "ServicevorgaengeP"; InvokeSPList $SPList
        $SPList = "ServicevorgaengeE"; InvokeSPList $SPList

    }

    #Wenn die SP-Site existiert,     
    #1. die SPSite-URL in die Excel-Datei pobvol-sync\pssService-Settings.xlsx eintragen
    #2. die SP-List-Ids in die Excel-Datei pobvol-sync\pssService-Settings.xlsx eintragen
    If($SPSiteExist -eq "true"){

        $fileName = $MyPath + '\pobvol-sync\pssService-Settings.xlsx'
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false #Soll Excel angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
        $workbook = $excel.Workbooks.Open($fileName)

        $strMessage = "Adding SPSiteURL to pobvol-sync\pssService-Settings.xlsx. SpSiteURL: " + $SPSiteURL
        Write-Host $strMessage
        LogWrite $strMessage

        $sheet = $workbook.worksheets.Item(1)
        for ($i=1; $i -le 100; $i++)
        {
        
            If ($sheet.Cells.Item($i, 2).Text -eq 'SharePoint.Site') {
                $sheet.Cells.Item($i, 3) = $SPSiteURL
                $i=100
            }
            #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
            #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
        }

        $strMessage = "Adding SPListIds to pobvol-sync\pssService-Settings.xlsx. SpSiteURL: " + $SPSiteURL
        Write-Host $strMessage
        LogWrite $strMessage

        $sheet = $workbook.worksheets.Item(1)
        $Lists = Get-PnPList
        foreach($List in $Lists) { 

            $SPListTitle = $($List.Title)
            $SPListId = $($List.Id)

            #Den SP-Listeneintrag in der Excel suchen 
            for ($i=1; $i -le 100; $i++)
            {
        
                If ($sheet.Cells.Item($i, 1).Text -eq 'SharePoint.List' -AND
                    $sheet.Cells.Item($i, 2).Text -eq $SPListTitle){
                    
                    $sheet.Cells.Item($i, 3) = "$SPListId"
                    $i=100
                }
                #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
                #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
            }
        
        }

        #Excel-Datei speichern und schliessen
        $workbook.close($true)
        $excel.Quit() #calling quit() on the object makes the GUI of the application disappear but sometimes the process is still running
        Remove-Variable excel -ErrorAction SilentlyContinue
        $Processes = get-process #Powershell to kill excel and access
        ForEach($ProcessName in $Processes){
            If($ProcessName -eq "MSACCESS"){
                Stop-Process -Name "MSACCESS" -Force
            }
            If($ProcessName -eq "EXCEL"){
                Stop-Process -Name "EXCEL" -Force
            }
        }
    }

    #Wenn die SP-Site existiert, Teams hinzufügen
    If($SPSiteExist -eq "true"){

        #Once a new M365 group is added to the classic team site, we can use the below command to create a new Microsoft Teams team (Teamify) using the newly created group which is connected with the site.
        #Provide your team site URL which is connected with a group
        #$SiteURL = "https://contoso.sharepoint.com/sites/salesteamsite"
 
        $strMessage = "Creating Microsoft Teams team (Teamify) for site " + $SPSiteURL
        Write-Host $strMessage
        LogWrite $strMessage

        #Get the id (Group Id) of the connected group
        $SiteInfo = Get-PnPTenantSite -Identity $SPSiteURL
        $GroupId = $SiteInfo.GroupId.Guid
 
        #Teamify - Enable Teams in the group which is connected with the site
        New-PnPTeamsTeam -GroupId $GroupId

    }

    If($SPSiteExist -eq "true"){
        # Setup Access-Datenbanken 
        # Change to sub folder 'pobvol-sync', 
        # open accdb file 'pssService-setup.accdb' and
        # start query 'queryControl'

        $strMessage = "Refreshing links in Access databases"
        Write-Host $strMessage
        LogWrite $strMessage

        $MySyncPath = $MyPath + "\pobvol-sync"
        
        #Set-Location -Path $MySyncPath

        $fileName = $MySyncPath + '\pssService-setup.accdb'
        $access = New-Object -com access.application
        $access.Visible = $true #Soll Access angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
        $access.Application.OpenCurrentDatabase($fileName)
        if( !$? ){
            $strMessage = "Could not open the file " + $fileName
            Write-Host $strMessage
            LogWrite $strMessage
        }Else{
            $access.Application.DoCmd.OpenQuery("queryControl")
            if( !$? ) {
                $strMessage = "Error while executing the query queryControl"
                Write-Host $strMessage
                LogWrite $strMessage
            }
            $access.Application.CloseCurrentDatabase()
            $access.Quit()
        }
        Remove-Variable access -ErrorAction SilentlyContinue
        $Processes = get-process #Powershell to kill excel and access
        ForEach($ProcessName in $Processes){
            If($ProcessName -eq "MSACCESS"){
                Stop-Process -Name "MSACCESS" -Force
            }
            If($ProcessName -eq "EXCEL"){
                Stop-Process -Name "EXCEL" -Force
            }
        }

    }
}

#---------------------------------------------------------------------------
#Wichtige Dateien vom Arbeitsornder in die SharePoint-Bibliothek hochladen
#---------------------------------------------------------------------------

#Aber nur, wenn der OneDrive sync folder ($SPPath) auch schon bekannt ist.
#Bei der Erstinstallation wird SharePoint eingerichtet, da ist der Ordner noch nicht vorhanden.
#Später, nach Start der Synchronisation in OneDrive, ist der Folder dann vorhanden und bekannt.
#Create Service Reports folder if not exists

#---------------------------------------------------------------------------
#Create Service Reports folder if not exists
#---------------------------------------------------------------------------

If($SPPath -ne "" -AND $SPReportsFolder -ne ""){
    $strTarget = $SPPath + "\" + $SPReportsFolder
    #wenns den Folder nicht gibt,...
    if (!(Test-Path -Path $strTarget)) {
        #Folder anlegen
        $strMessage = "Creating folder " +$strTarget
        Write-Host $strMessage
        LogWrite $strMessage
        try {
            New-Item -Path $strTarget -Type Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            #throw $_.Exception.Message
            $strMessage = "    " +$_.Exception.Message
            LogWrite $strMessage
        }
    }
}

#---------------------------------------------------------------------------
#Create Contracts folder if not exists
#---------------------------------------------------------------------------

If($SPPath -ne "" -AND $SPContractsFolder -ne ""){
    $strTarget = $SPPath + "\" + $SPContractsFolder
    #wenns den Folder nicht gibt,...
    if (!(Test-Path -Path $strTarget)) {
        #Folder anlegen
        $strMessage = "Creating folder " +$strTarget
        Write-Host $strMessage
        LogWrite $strMessage
        try {
            New-Item -Path $strTarget -Type Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            #throw $_.Exception.Message
            $strMessage = "    " +$_.Exception.Message
            LogWrite $strMessage
        }
    }
}

#---------------------------------------------------------------------------
# Copy png-files from Reports-folder 
# /M Copies only files with the archive attribute set, turns off the archive attribute.
#---------------------------------------------------------------------------

If($SPPath -ne "" -AND $SPReportsFolder -ne ""){
    Write-Host "    " 
    Write-Host "Copy png files to the SharePoint document library"
    $strSource = $MyPath + "\pobvol-sync\reports\*.png"
    $strTarget = $SPPath + "\" + $SPReportsFolder
    xcopy $strSource $strTarget /M /V /C /I /G /R /Y 
}

#---------------------------------------------------------------------------
# Quit Excel and stop all Excel tasks
$Processes = get-process #Powershell to kill excel and access
ForEach($ProcessName in $Processes){
    If($ProcessName -eq "MSACCESS"){
        Stop-Process -Name "MSACCESS" -Force
    }
    If($ProcessName -eq "EXCEL"){
        Stop-Process -Name "EXCEL" -Force
    }
    If($ProcessName -eq "excel"){
        Stop-Process -Name "excel" -Force
    }
}
#---------------------------------------------------------------------------
$strMessage = "pobvol Service Solution: pssService-setup.ps1 finished."
Write-Host $strMessage
LogWrite $strMessage

#Pause

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
