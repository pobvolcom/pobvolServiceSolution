Attribute VB_Name = "modLinkSharePointList"

Sub subLinkSharePointList(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim DB As DAO.Database
    Dim RecordSet As Object
    Dim tbl As TableDef
    Dim strSource As String
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     LinkSharePointList (" & strDB & "-->" & strObjectName & ")")

    'Externe Datenbank öffnen
    Set DB = OpenDatabase(strDBPath)
    
    'Tabelle löschen?
    For Each RecordSet In DB.TableDefs
        If (RecordSet.Name = strObjectName) Then
            'strSource = RecordSet.Connect
            DB.TableDefs.Delete strObjectName 'delete linked table
            Exit For
        End If
    Next RecordSet
    
    'Set table
    Set tbl = DB.CreateTableDef(strObjectName)
    'Connect to SharePoint list
    tbl.Connect = _
    "WSS;HDR=NO;IMEX=2;ACCDB=YES;DATABASE=" & strSPSite & ";LIST=" & strSourceName & ";VIEW=;RetrieveIds=No;ListDisplayName=" & strObjectName & ";"
    'Set source sheet
    tbl.SourceTableName = strSourceName
    'Append table
    DB.TableDefs.Append tbl
    'Refresh list
    Call RefreshLinkOutput(DB, strObjectName)
   
Fin:
    
    'Close external database
    DB.Close
    Set DB = Nothing
    Set tbl = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub

