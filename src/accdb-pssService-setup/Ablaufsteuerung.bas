Attribute VB_Name = "Ablaufsteuerung"
Option Compare Database
Option Explicit

Public Function ProcedureExecute()
'On Error Resume Next
On Error GoTo Fin


    'Dim dbs As Database
    
    'Dim rst As Recordset
    'Dim qdf As QueryDef
    'Dim strSQL As String
    
    Dim lastRow As Integer
    Dim strMyPath As String
    Dim ErrNumber As Long
    Dim strProtokoll As String
    Dim strfilename As String
    'Dim strSPActive As String
    Dim strSPPDFPath As String
    Dim strSPTeam As String
    Dim strSPSite As String
   
    'Access Excel files
    Dim app As Object
    Dim wbk As Object
        
    'Use the current db
    'Set dbs = CurrentDb

    'Set Warnings off
    'DoCmd.SetWarnings False
    
    'Set Hourglass on
    'DoCmd.Hourglass True
   
    strMyPath = Application.CurrentProject.Path
    Call add_entry_to_log_always(strProtokoll, strMyPath, "     Running scripts in Setup.accdb")
   
    'Call ProcessUpdates(strMyPath, "Ja", "C:\Users\Volker\pobvol Software Services\Service - Documents", "Service", "https://pobvol.sharepoint.com/sites/Service", ErrNumber)
    'GoTo Fin

    'Check for file pssService-Settings.xlsx?
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
    
    'Check parameters
    If (strSPTeam = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Team not found in " & strfilename)
        GoTo Fin
    End If
    If (strSPSite = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Site not found in " & strfilename)
        GoTo Fin
    End If
    
    'Check for updates
    Call ProcessUpdates(strMyPath, strProtokoll, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
    

Fin:
    'Set Warnings on
    'DoCmd.SetWarnings True
    
    'Set Hourglass off
    'DoCmd.Hourglass False
    
    If Err.Number <> 0 Then
        'ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If
    
        
End Function


