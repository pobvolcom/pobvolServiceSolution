Attribute VB_Name = "modTransferSharePointList"
Option Compare Database

Sub subTransferSharePointList(strProtokoll, strMyPath, strDB, strDBPath, strObjectType, strObjectName, strSQL, strSourceName, strSourceWS, strSPPDFPath, strSPTeam, strSPSite, ErrNumber)
On Error GoTo Fin

    'Zugriff aus Access-Objekte
    Dim oAcc As Access.Application
    Dim DB As DAO.Database
    Dim RecordSet As Object

    Call add_entry_to_log(strProtokoll, strMyPath, "     DoCmd.TransferSharePointList (" & strDB & "-->" & strObjectName & ")")
    
    'Externe Datenbank öffnen
    Set DB = OpenDatabase(strDBPath)
    
    'Tabelle löschen?
    For Each RecordSet In DB.TableDefs
        If (RecordSet.Name = strObjectName) Then
            DB.TableDefs.Delete strObjectName 'delete linked table
            Exit For
        End If
    Next RecordSet
    
    'Close external database
    DB.Close
    Set DB = Nothing

    'Access App starten
    Set oAcc = CreateObject("Access.Application")
    
    'Sichtbar?
    oAcc.Visible = False
    
    'Externe Datenbank öffnen
    oAcc.OpenCurrentDatabase (strDBPath)
    
    'Aktualisiere verlinkte SharePoint-Liste
    'DoCmd.TransferSharePointList method (Access)
    'https://learn.microsoft.com/en-us/office/vba/api/access.docmd.transfersharepointlist
    oAcc.DoCmd.TransferSharePointList acLinkSharePointList, strSPSite, strSourceName, , strObjectName
    
Fin:
    
    'Reset DB
    Set DB = Nothing
    
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


