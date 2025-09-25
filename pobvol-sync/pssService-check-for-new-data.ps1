#
# Script:	    pssService-check-for-new-data.ps1
# Task:		    Processing new and changed data
# 
# This file is part of the software solution pobvol Service Solution. 
# pobvol Service Solution is Free Software, delivered as open source. You can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. The solution is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with the solution. If not, see <http://www.gnu.org/licenses/>. 
# Copyright © 2025 Volker Pobloth
# Web: https://pobvol.com/
#
#---------------------------------------------------------------------------------------
Function LogWrite
# ---------------
{
    Param ([string]$strMessage)
    $strCurrentDateTime = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $strValue = $strCurrentDateTime + "  " + $strMessage
    Add-content $strLogfile -value $strValue
}
# ---------------------------------------------------------------------------------------

#Current path
$MyPath = $PSScriptRoot

#Variables
$Environment = ""
$SPPath = "null"
$SPReportsFolder = "null"
$SPContractsFolder = "null"
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


Write-Host "pobvol Service Solution: pssService-check-for-new-data.ps1 started ...."
Write-Host "Current folder: " $MyPath 

#Log-File
$strLogfile = $MyPath + "\log.txt"
Write-Host "Log file: " $strLogfile


#Load system settings
$strMessage = "    Reading settings from pssService-Settings.xlsx"
Write-Host $strMessage
$fileName = $MyPath + '\pssService-Settings.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false #Soll Excel angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
$workbook = $excel.Workbooks.Open($fileName)
$sheet = $workbook.worksheets.Item(1)
for ($i=1; $i -le 100; $i++)
{
    #$strMessage = "Reading row from SystemSettings.xlsx: " + $i
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

    # true / false
    If ($sheet.Cells.Item($i, 2).Text -eq 'Detailed logging') {
        $Protokoll = $sheet.Cells.Item($i, 3).Text
    }

    # DEV / TST / PROD
    If ($sheet.Cells.Item($i, 2).Text -eq 'pssService Environment') {
        $Environment = $sheet.Cells.Item($i, 3).Text
    }

    If ($sheet.Cells.Item($i, 2).Text -eq 'SharePoint.Domain') {
        $SPDomain = $sheet.Cells.Item($i, 3).Text
    }
    
    If ($sheet.Cells.Item($i, 2).Text -eq 'SharePoint.Team') {
        $SPTeam = $sheet.Cells.Item($i, 3).Text
    }

    If ($sheet.Cells.Item($i, 2).Text -eq 'PnP Rocks Id') {
        $PnPRocksId = $sheet.Cells.Item($i, 3).Text
    }

    #$strProtokoll = $sheet.Cells.Item(18, 1).Text          #Wert aus Zelle A18 lesen (Zeile, Spalte)
    #$sheet.cells.item(1,1) = "Test"                        #Test in die Zelle A1 schreiben (Zeile, Spalte)
}
$workbook.close($false)
$excel.Quit() #calling quit() on the object makes the GUI of the application disappear but sometimes the process is still running
$Processes = get-process #Powershell to kill excel and access
ForEach($ProcessName in $Processes){
    If($ProcessName -eq "MSACCESS"){
        Stop-Process -Name "MSACCESS" -Force
    }
    If($ProcessName -eq "EXCEL"){
        Stop-Process -Name "EXCEL" -Force
    }
}

Write-Host "Parameters found:"
Write-Host "    SharePoint.Domain:" $SPDomain
Write-Host "    SharePoint.Team:" $SPTeam
Write-Host "    PnP Rocks Id:" $PnPRocksId
Write-Host "    Environment:" $Environment
Write-Host "    Full log:" $Protokoll
Write-Host "    OneDrive sync folder:" $SPPath
Write-Host "    Service reports saved in sub folder:" $SPReportsFolder
Write-Host "    Contracts saved in sub folder:" $SPContractsFolder

#Check Parameter
If( $SPDomain -eq "null" -OR $SPDomain -eq "Enter your value" -OR
    $SPTeam -eq "null" -OR $SPTeam -eq "Enter your value" -OR
    $PnPRocksId -eq "null" -OR $PnPRocksId -eq "Enter your value" -OR
    $SPPath -eq "null" -OR $SPPath -eq "Enter your value" -OR
    $SPReportsFolder -eq "null" -OR $SPReportsFolder -eq "Enter your value" -OR
    $SPContractsFolder -eq "null" -OR $SPContractsFolder -eq "Enter your value" -OR
    $Protokoll -eq "null" -OR $Protokoll -eq "Enter your value"){

    $strError = "Error"

    $strMessage = "pobvol Service Solution: pssService-check-for-new-data.ps1"
    LogWrite $strMessage
    $strMessage = "    Processing service data skipped. Missing parameters!"
    Write-Host $strMessage
    LogWrite $strMessage
    $strMessage = "    Please check your entries in " +$MyPath +'\pssService-Settings.xlsx'
    Write-Host $strMessage
    LogWrite $strMessage
}


#Create Service Reports folder if not exists
If($strError -eq ""){
    $strTarget = $SPPath + "\" + $SPReportsFolder
    #wenns den Folder nicht gibt,...
    if (!(Test-Path -Path $strTarget)) {
        #Folder anlegen
        $strMessage = "Creating  folder " +$strTarget
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

#Create Contracts folder if not exists
If($strError -eq ""){
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

#Process linked master data
If($strError -eq ""){
    # Setup Access-Datenbanken 
    $strMessage = "Processing linked master data"
    Write-Host $strMessage
    LogWrite $strMessage
    $fileName = $MyPath + '\' + "pssService-link-to-customer-db.accdb"
    $access = New-Object -com access.application
    $access.Visible = $false #Soll Access angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
    $access.Application.OpenCurrentDatabase($fileName)
    if( !$? ){
        $strMessage = "Konnte die Datei $fileName nicht öffnen."
        Write-Host $strMessage
        LogWrite $strMessage
    }Else{
        $access.Application.DoCmd.OpenQuery("queryControl")
        if( !$? ) {
            $strMessage = "Fehler bei Ausführung der Abfrage queryControl"
            Write-Host $strMessage
            LogWrite $strMessage
        }
        $access.Application.CloseCurrentDatabase()
        $access.Quit()
    }

    #get-process | where {$_.StartTime -lt $startTimeLimit -and ($_.path -like "*access.exe" -or $_.path -like "*excel.exe") } |  foreach { Stop-Process -Name "MSACCESS" -Force }

    #$TotalProcesses = $Processes.Count
    #If($TotalProcesses -gt 0){
    #    Stop-Process -Name "MSACCESS" -Force
    #}

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

#Process service data
If($strError -eq ""){
    # Setup Access-Datenbanken 
    $strMessage = "Processing service data"
    Write-Host $strMessage
    LogWrite $strMessage
    $fileName = $MyPath + '\' + "pssService-sp-connector.accdb"
    $access = New-Object -com access.application
    $access.Visible = $false #Soll Access angezeigt werden ($true) oder unsichtbar im Hintergrund laufen ($false)?
    $access.Application.OpenCurrentDatabase($fileName)
    if( !$? ){
        $strMessage = "Konnte die Datei $fileName nicht öffnen."
        Write-Host $strMessage
        LogWrite $strMessage
    }Else{
        $access.Application.DoCmd.OpenQuery("queryControl")
        if( !$? ) {
            $strMessage = "Fehler bei Ausführung der Abfrage queryControl"
            Write-Host $strMessage
            LogWrite $strMessage
        }
        $access.Application.CloseCurrentDatabase()
        $access.Quit()
    }

    #get-process | where {$_.StartTime -lt $startTimeLimit -and ($_.path -like "*access.exe" -or $_.path -like "*excel.exe") } |  foreach { Stop-Process -Name "MSACCESS" -Force }

    #$TotalProcesses = $Processes.Count
    #If($TotalProcesses -gt 0){
    #    Stop-Process -Name "MSACCESS" -Force
    #}

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

# ---------------------------------------------------------------------------
# Copy log file 
# ---------------------------------------------------------------------------
Write-Host "    " 
Write-Host "Copy changed log file to your SharePoint document library"
$strSource = $MyPath + "\log.txt"
$strTarget = $SPPath + "\" + $SPReportsFolder
xcopy $strSource $strTarget /M /S /V /C /I /G /R /Y 


# Powershell to kill excel
$Processes = get-process excel #Powershell to kill excel
ForEach($ProcessName in $Processes){Stop-Process -Name "EXCEL" -Force}
Remove-Variable excel -ErrorAction SilentlyContinue


Write-Host "    " 
Write-Host "pobvol Service Solution: pssService-check-for-new-data.ps1 finished."

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
