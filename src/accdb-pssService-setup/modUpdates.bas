Attribute VB_Name = "modUpdates"

Public Function ProcessUpdates(strMyPath, strProtokoll, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
'On Error Resume Next
On Error GoTo Fin

    'Define all we need for processing the accdb commands from \setup-tasks\setup_accdb*.xlsx
    Dim oFSO As Object
    Dim oOrdner As Object
    Dim oDatei As Object
    Dim strfilename As String
    Dim strfilenamebak As String
    Dim Zeile As Long
    Dim ZeileEnd As Long
    Dim strDB As String
    Dim strDoCmd As String
    Dim strObjectType As String
    Dim strObjectName As String
    Dim strSQL As String
    Dim strSourceName As String
    Dim strSourceWS As String
    Dim strDBPath As String
    Dim a As Long
    Dim i As Long
    'Dim WaitUntil As Variant
    
    'Excel
    Dim app As Object
    Dim wbk As Object
    
    'Start Excel
    Set app = VBA.CreateObject("Excel.Application")
    app.Visible = False
    
    'Check for command files (setup_accdb*.xlsx) in \setup-tasks
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oOrdner = oFSO.GetFolder(strMyPath & "\setup-tasks")
    For Each oDatei In oOrdner.Files
        
        If (Left(oDatei.Name, 11) = "setup_accdb" And Right(oDatei.Name, 5) = ".xlsx") Then
            
            strfilename = strMyPath & "\setup-tasks\" & oDatei.Name
            strfilenamebak = strMyPath & "\setup-tasks\" & oDatei.Name & ".bak"
            Call add_entry_to_log(strProtokoll, strMyPath, "     Processing commands from file " & strfilename)
            
            'Open Excel file
            Set wbk = app.Workbooks.Open(strfilename)
            
            'Anzahl der Zeilen ermitteln
            'ZeileEnd = wbk.worksheets("Commands").Range("A65536").End(xlUp).Row ' funkt nicht aus Access heraus
            ZeileEnd = 65536
            
            'Processing line by line
            For Zeile = 2 To ZeileEnd
                
                'Kommando einlesen
                strDB = wbk.worksheets("Commands").Cells(Zeile, 1).Value
                strDoCmd = wbk.worksheets("Commands").Cells(Zeile, 2).Value
                strObjectType = wbk.worksheets("Commands").Cells(Zeile, 3).Value
                strObjectName = wbk.worksheets("Commands").Cells(Zeile, 4).Value
                strSQL = wbk.worksheets("Commands").Cells(Zeile, 5).Value
                strSourceName = wbk.worksheets("Commands").Cells(Zeile, 6).Value
                strSourceWS = wbk.worksheets("Commands").Cells(Zeile, 7).Value
                
                'DB file
                strDBPath = strMyPath & "\" & strDB
                
                'Entry found?
                If (strDB = "") Then
                    Zeile = ZeileEnd
                End If
                    
                'DB file found?
                If (strDB > "" And Dir$(strDBPath, vbHidden + vbReadOnly + vbSystem) > "") Then

                    If strDoCmd = "DeleteObject" Then
                        Call subDeleteObject(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "RunSQL" Then
                        Call subRunSQL(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "Execute" Then
                        Call subExecute(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "CreateQueryDef" Then
                        Call subCreateQueryDef(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "TransferSpreadsheet" Then
                        Call subTransferSpreadsheet(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "CreateTableDef" Then
                        Call subCreateTableDef(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "TransferSharePointList" Then
                        Call subTransferSharePointList(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    ElseIf strDoCmd = "LinkSharePointList" Then
                        Call subLinkSharePointList(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
                    
                    'Else
                    '    Bonus = 0
                    End If
                
                End If
            
            Next Zeile
            
            'Close Excel file
            wbk.Close savechanges:=False
            
            'Rename command file (*.bak)
            Name strfilename As strfilenamebak
            
         
        End If
    
    Next oDatei
                    
    'Example: wait a second
    'WaitUntil = Now + TimeValue("00:00:01")
    'Do
    '    DoEvents
    'Loop Until Now >= WaitUntil
    
    'Example: delete an external database file (.accdb)
    'Kill ("c:\temp\TestDB.accdb")

    'Example: close Access
    'DoCmd.Quit


Fin:
    
    'Quit Excel
    app.Quit
    Set wbk = Nothing
    Set app = Nothing
    
    'Quit file system
    Set oFSO = Nothing
    Set oOrdner = Nothing
    
    If Err.Number <> 0 Then
        'ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If
    
        
End Function


