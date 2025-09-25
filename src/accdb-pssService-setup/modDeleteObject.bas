Attribute VB_Name = "modDeleteObject"

Sub subDeleteObject(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim DB As DAO.Database
    Dim RecordSet As Object

    Call add_entry_to_log(strProtokoll, strMyPath, "     DeleteObject (" & strDB & "-->" & strObjectName & ")")
    
    'Externe Datenbank öffnen
    Set DB = OpenDatabase(strDBPath)
    
    'Tabelle löschen?
    If (strObjectType = 0) Then
        For Each RecordSet In DB.TableDefs
            If (RecordSet.Name = strObjectName) Then
                DB.Execute strSQL ' Delete table
                Exit For
            End If
        Next RecordSet
    End If
    
    'Abfrage löschen?
    If (strObjectType = 1) Then
        For Each RecordSet In DB.QueryDefs
            If (RecordSet.Name = strObjectName) Then
                DB.QueryDefs.Delete strObjectName ' Delete QueryDef
                Exit For
            End If
        Next RecordSet
    End If

Fin:
    
    'Close external database
    DB.Close
    Set DB = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub

