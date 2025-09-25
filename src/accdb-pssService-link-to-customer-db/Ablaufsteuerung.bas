Attribute VB_Name = "Ablaufsteuerung"
Option Compare Database
Option Explicit

Public Function ProcedureExecute()
'On Error Resume Next
On Error GoTo Fin


    Dim dbs As Database
    
    'Dim rst As Recordset
    'Dim qdf As QueryDef
    'Dim strSQL As String
    
    Dim strMyPath As String
    Dim ErrNumber As Long
    Dim strProtokoll As String
    Dim strfilename As String
    Dim strSPActive As String
    Dim strSPPDFPath As String
    Dim strSPTeam As String
    Dim strSPSite As String
   
    Dim app As Object
    Dim wbk As Object
    
    Set dbs = CurrentDb
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass True

    strMyPath = Application.CurrentProject.Path
    Call add_entry_to_log_always(strProtokoll, strMyPath, "     Running scripts in pssService-link-to-customer-db.accdb")

    'SystemSettings
    strProtokoll = "true"
    strSPPDFPath = ""
    strSPTeam = ""
    strSPSite = ""
    strfilename = strMyPath & "\pssService-Settings.xlsx"
    If Dir(strfilename, vbNormal) <> "" Then
        'Start Excel
        Set app = VBA.CreateObject("Excel.Application")
        app.Visible = False
        'Open Excel file
        Set wbk = app.Workbooks.Open(strfilename)
        
        'Konfiguration lesen
        'strProtokoll = wbk.worksheets("Config").Range("A18").Value
        'strSPPDFPath = wbk.worksheets("Config").Range("B9").Value
        'strSPTeam = wbk.worksheets("Config").Range("C9").Value
        'strSPSite = wbk.worksheets("Config").Range("D9").Value
        
        'lastRow = wbk.worksheets("SystemSettings").Range("B" & Rows.Count).End(xlUp).Row
        Dim Zelle
        Dim intRow
        Dim intColumn
        For Each Zelle In wbk.worksheets("SystemSettings").Range("B2:B100")
            
            intRow = Zelle.Row
            intColumn = Zelle.Column
            
            If (Zelle.Value = "Detailed logging") Then
                strProtokoll = wbk.worksheets("SystemSettings").Range("C" & Zelle.Row).Value
            End If
            If (Zelle.Value = "OneDrive-Folder") Then
                strSPPDFPath = wbk.worksheets("SystemSettings").Range("C" & Zelle.Row).Value
            End If
            If (Zelle.Value = "SharePoint.Team") Then
                strSPTeam = wbk.worksheets("SystemSettings").Range("C" & Zelle.Row).Value
            End If
            If (Zelle.Value = "SharePoint.Site") Then
                strSPSite = wbk.worksheets("SystemSettings").Range("C" & Zelle.Row).Value
            End If
        Next Zelle
        
        'Close Excel file
        wbk.Close savechanges:=False
        'Quit Excel
        app.Quit
        Set wbk = Nothing
        Set app = Nothing
    Else
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: File " & strfilename & " not found!")
        GoTo Fin
    End If
    
    'Check paramters
    If (strSPPDFPath = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: Parameter for the OneDrive folder not found in " & strfilename)
        GoTo Fin
    End If
    If (strSPTeam = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Team not found in " & strfilename)
        GoTo Fin
    End If
    If (strSPSite = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Site not found in " & strfilename)
        GoTo Fin
    End If
    
    'Execute queries
    Call add_entry_to_log(strProtokoll, strMyPath, "     Processing service customers")
    dbs.Execute ("01 Append KundenDB_Kunden")
    dbs.Execute ("03 Update GPSLocation")
    dbs.Execute ("03 Update Kunde")
    dbs.Execute ("03 Update Kunde Demodatenflag")
    dbs.Execute ("03 Update Kundenort")
    dbs.Execute ("03 Update Land")
    dbs.Execute ("03 Update Plz")
    dbs.Execute ("03 Update Strasse")
    dbs.Execute ("04 Update Ansprechpartner")
    dbs.Execute ("99 Delete Kunden")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Processing customer devices")
    dbs.Execute ("02 Append KundenDB_Inventar")
    dbs.Execute ("05 Update Inventar Baujahr")
    dbs.Execute ("05 Update Inventar Bemerkungen")
    dbs.Execute ("05 Update Inventar Demodaten")
    dbs.Execute ("05 Update Inventar Geraeteart")
    dbs.Execute ("05 Update Inventar Geraetetyp")
    dbs.Execute ("05 Update Inventar GPSLocation")
    dbs.Execute ("05 Update Inventar Hersteller")
    dbs.Execute ("05 Update Inventar Kundeninventarnummer")
    dbs.Execute ("05 Update Inventar SerienNr")
    dbs.Execute ("05 Update Inventar Serviceintervall")
    dbs.Execute ("05 Update Inventar Standort")
    dbs.Execute ("99 Delete Inventar")

Fin:
    DoCmd.SetWarnings True
    DoCmd.Hourglass False
    Set dbs = Nothing
    If Err.Number <> 0 Then
        'ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If
        
End Function

