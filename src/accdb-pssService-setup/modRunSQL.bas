Attribute VB_Name = "modRunSQL"

Sub subRunSQL(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim oAcc As Access.Application
    'Dim DB As DAO.Database
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     DoCmd.RunSQL (" & strDB & "-->" & strSQL & ")")
    
    'Access App starten
    Set oAcc = CreateObject("Access.Application")
    
    'Externe Datenbank öffnen
    oAcc.OpenCurrentDatabase (strDBPath)
        
    'Sichtbar?
    oAcc.Visible = False
    
    'SQL ausführen
    oAcc.DoCmd.RunSQL strSQL
    
Fin:
    
    'Datenbank schliessen
    oAcc.CloseCurrentDatabase
    
    'Access App schliessen
    oAcc.Quit
    Set oAcc = Nothing
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If

End Sub


