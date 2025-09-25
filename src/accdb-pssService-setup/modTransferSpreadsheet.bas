Attribute VB_Name = "modTransferSpreadsheet"

Sub subTransferSpreadsheet(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim oAcc As Access.Application
    'Dim DB As DAO.Database
    'Dim RecordSet As Object
    Dim strSource As String

    Call add_entry_to_log(strProtokoll, strMyPath, "     DoCmd.TransferSpreadsheet (" & strDB & "-->" & strObjectName & ")")
    
    'Access App starten
    Set oAcc = CreateObject("Access.Application")
    
    'Externe Datenbank öffnen
    oAcc.OpenCurrentDatabase (strDBPath)
        
    'Sichtbar?
    oAcc.Visible = False

    'verlinkte Tabelle löschen
    For Each RecordSet In oAcc.CurrentDb.TableDefs
        If (RecordSet.Name = strObjectName) Then
            oAcc.CurrentDb.TableDefs.Delete strObjectName 'delete linked table
            Exit For
        End If
    Next RecordSet

    'Aktualisiere verlinkte Dateien
    'DoCmd.TransferSpreadsheet-Methode (Access)
    'https://learn.microsoft.com/de-DE/office/vba/api/Access.DoCmd.TransferSpreadsheet
    strSource = strMyPath & "\" & strSourceName
    oAcc.DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, strObjectName, strSource, True
    Call RefreshLinkOutput(oAcc.CurrentDb, strObjectName)
    
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

