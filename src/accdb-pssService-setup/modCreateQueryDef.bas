Attribute VB_Name = "modCreateQueryDef"

Sub subCreateQueryDef(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim DB As DAO.Database
    Dim RecordSet As Object
    Dim qdf As QueryDef

    Call add_entry_to_log(strProtokoll, strMyPath, "     CreateQueryDef (" & strDB & "-->" & strObjectName & ")")
    
    'Open db
    Set DB = OpenDatabase(strDBPath)
    
    For Each RecordSet In DB.QueryDefs
        If (RecordSet.Name = strObjectName) Then
            DB.QueryDefs.Delete strObjectName ' Delete QueryDef
            Exit For
        End If
    Next RecordSet
    
    Set qdf = DB.CreateQueryDef(strObjectName, strSQL) ' Create permanent QueryDef
    GetrstTemp qdf ' Open Recordset and print report to refresh definition and data
    
Fin:
    
    'Close external database
    DB.Close
    Set DB = Nothing
    Set qdf = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub




