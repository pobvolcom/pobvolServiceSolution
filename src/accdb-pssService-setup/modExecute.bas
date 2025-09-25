Attribute VB_Name = "modExecute"

Sub subExecute(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim DB As DAO.Database

    Call add_entry_to_log(strProtokoll, strMyPath, "     Execute (" & strDB & "-->" & strSQL & ")")
    
    'Externe Datenbank öffnen
    Set DB = OpenDatabase(strDBPath)
    
    DB.Execute strSQL
    
Fin:
    
    'Close external database
    DB.Close
    Set DB = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub



