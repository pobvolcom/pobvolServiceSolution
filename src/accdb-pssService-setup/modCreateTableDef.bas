Attribute VB_Name = "modCreateTableDef"

Sub subCreateTableDef(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim DB As DAO.Database
    Dim RecordSet As Object
    Dim tbl As TableDef
    Dim strSource As String

    Call add_entry_to_log(strProtokoll, strMyPath, "     CreateTableDef (" & strDB & "-->" & strObjectName & ")")
    
    'Externe Datenbank öffnen
    Call add_entry_to_log(strProtokoll, strMyPath, "        OpenDatabase (" & strDBPath & ")")
    Set DB = OpenDatabase(strDBPath)
    
    'Tabelle löschen?
    For Each RecordSet In DB.TableDefs
        If (RecordSet.Name = strObjectName) Then
            strSource = RecordSet.Connect
            Call add_entry_to_log(strProtokoll, strMyPath, "        DB.TableDefs.Delete (" & strObjectName & ")")
            DB.TableDefs.Delete strObjectName 'delete linked table
            Exit For
        End If
    Next RecordSet
    
    'Set table
    Call add_entry_to_log(strProtokoll, strMyPath, "        DB.CreateTableDef (" & strObjectName & ")")
    Set tbl = DB.CreateTableDef(strObjectName)
    
    If (strObjectType = 10) Then
        'Link auf eine Excel-Datei
        strSource = strMyPath & "\" & strSourceName
        'tbl.Connect = "Excel 12.0;DATABASE=" & strSource & ""
        tbl.Connect = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & strSource & ""
        Call add_entry_to_log(strProtokoll, strMyPath, "        tbl.Connect =" & tbl.Connect)
     
    ElseIf (strObjectType = 11) Then
        'Link auf eine Tabelle in einer Access-Datenbank
        strSource = strMyPath & "\" & strSourceName
        tbl.Connect = ";DATABASE=" & strSource & ""
        Call add_entry_to_log(strProtokoll, strMyPath, "        tbl.Connect =" & tbl.Connect)
    
    ElseIf (strObjectType = 12) Then
        'Link auf eine SharePoint-Liste
        tbl.Connect = "WSS;HDR=NO;IMEX=2;ACCDB=YES;DATABASE=" & strSPSite & ";LIST=" & strSourceName & ";VIEW=;RetrieveIds=No"
        Call add_entry_to_log(strProtokoll, strMyPath, "        tbl.Connect =" & tbl.Connect)
    
    End If
    
    'Set source sheet
    tbl.SourceTableName = strSourceWS
    Call add_entry_to_log(strProtokoll, strMyPath, "        tbl.SourceTableName =" & tbl.SourceTableName)
    
    'Append table
    Call add_entry_to_log(strProtokoll, strMyPath, "        TableDefs.Append")
    DB.TableDefs.Append tbl
    
    'Refresh list
    'Call RefreshLinkOutput(DB, strObjectName)
  
Fin:
    
    'Close external database
    Call add_entry_to_log(strProtokoll, strMyPath, "        Close")
    DB.Close
    Set DB = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub


