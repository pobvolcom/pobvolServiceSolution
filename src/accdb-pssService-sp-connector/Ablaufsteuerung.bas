Attribute VB_Name = "Ablaufsteuerung"
Option Compare Database
Option Explicit

Public Function ProcedureExecute()
'On Error Resume Next
On Error GoTo Fin

    Dim dbs As Database
    
    'Dim rst As Recordset
    'Dim qdf As QueryDef
    'Dim strSQL As String
    
    Dim strMyPath As String
    Dim ErrNumber As Long
    Dim strProtokoll As String
    Dim strfilename As String
    Dim strSPActive As String
    Dim strSPPDFPath As String
    Dim strSPTeam As String
    Dim strSPSite As String
   
    Dim app As Object
    Dim wbk As Object
    
    Set dbs = CurrentDb
    
    DoCmd.SetWarnings False
    DoCmd.Hourglass False
    
    strMyPath = Application.CurrentProject.Path
    Call add_entry_to_log_always(strProtokoll, strMyPath, "     Running scripts in pssService-sp-connector.accdb")
   
    'SystemSettings
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
                strProtokoll = LCase(wbk.worksheets("SystemSettings").Range("C" & Zelle.Row).Value)
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
    
    'Check parameter
    If (strSPPDFPath = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: Parameter for the OneDrive folder not found in " & strfilename)
        GoTo Fin
    End If
    If (strSPTeam = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Team not found in " & strfilename)
        GoTo Fin
    End If
    If (strSPSite = "") Then
        Call add_entry_to_log_always(strProtokoll, strMyPath, "     Error: SharePoint Site not found in " & strfilename)
        GoTo Fin
    End If
   
    'Execute queries
    
    dbs.Execute ("03 Update Berichtsflag in Serviceauftraegen")
    dbs.Execute ("03 Update Genehmigtflag in Serviceauftraegen")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed customers to SharePoint")
    dbs.Execute ("04 Append Servicekunden (Demodaten)") 
    dbs.Execute ("04 Append Servicekunden")
    dbs.Execute ("12 Update Servicekunden")
    dbs.Execute ("12 Update Servicekunden Bemerkungen")
    dbs.Execute ("12 Update Servicekunden GPS")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed contacts to SharePoint")
    dbs.Execute ("13 Update Servicekunden Ansprechpartner")
    dbs.Execute ("13 Update Servicekunden EMail")
    dbs.Execute ("13 Update Servicekunden Sprache")
    dbs.Execute ("13 Update Servicekunden Telefon")

    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed customer devices to SharePoint")
    dbs.Execute ("04 Append Kundeninventar (Demodaten)")
    dbs.Execute ("04 Append Kundeninventar")
    dbs.Execute ("16 Update Kundeninventar Artikelnummer")
    dbs.Execute ("16 Update Kundeninventar Baujahr")
    dbs.Execute ("16 Update Kundeninventar Bemerkungen")
    dbs.Execute ("16 Update Kundeninventar Geraeteart")
    dbs.Execute ("16 Update Kundeninventar Geraetetyp")
    dbs.Execute ("16 Update Kundeninventar GPSLocation")
    dbs.Execute ("16 Update Kundeninventar Hersteller")
    dbs.Execute ("16 Update Kundeninventar Kunde Teil1")
    dbs.Execute ("16 Update Kundeninventar Kunde Teil2")
    dbs.Execute ("16 Update Kundeninventar Kundeninventarnummer")
    dbs.Execute ("16 Update Kundeninventar SammelINVNR")
    dbs.Execute ("16 Update Kundeninventar SerienNr")
    dbs.Execute ("16 Update Kundeninventar Serviceintervall")
    dbs.Execute ("16 Update Kundeninventar Standort")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed checklists to SharePoint")
    dbs.Execute ("10 Append Checklisten")
    dbs.Execute ("11 Delete Checklisten")
    dbs.Execute ("11 Update Checklisten ArticleNo")
    dbs.Execute ("11 Update Checklisten Artikelnummer")
    dbs.Execute ("11 Update Checklisten Baujahr")
    dbs.Execute ("11 Update Checklisten Betriebsstunden")
    dbs.Execute ("11 Update Checklisten Checkliste")
    dbs.Execute ("11 Update Checklisten ChecklisteText")
    dbs.Execute ("11 Update Checklisten Checkpunkt")
    dbs.Execute ("11 Update Checklisten CheckpunktText")
    dbs.Execute ("11 Update Checklisten Code")
    dbs.Execute ("11 Update Checklisten CustInvNo")
    dbs.Execute ("11 Update Checklisten Default")
    dbs.Execute ("11 Update Checklisten DefectClass")
    dbs.Execute ("11 Update Checklisten DeviceLocation")
    dbs.Execute ("11 Update Checklisten DeviceNo")
    dbs.Execute ("11 Update Checklisten Eingangsdatum")
    dbs.Execute ("11 Update Checklisten Flat rate for transport service")
    dbs.Execute ("11 Update Checklisten FlatRateFor24HoursSystemTest")
    dbs.Execute ("11 Update Checklisten FlatRateForCleaningAndDisinfection")
    dbs.Execute ("11 Update Checklisten FlexForm")
    dbs.Execute ("11 Update Checklisten Garantie")
    dbs.Execute ("11 Update Checklisten Geraetetyp")
    dbs.Execute ("11 Update Checklisten GeraetetypText")
    dbs.Execute ("11 Update Checklisten Hardwarestand")
    dbs.Execute ("11 Update Checklisten Icon")
    dbs.Execute ("11 Update Checklisten IconPlakette")
    dbs.Execute ("11 Update Checklisten KEYFORMAT")
    dbs.Execute ("11 Update Checklisten LanguageTag")
    dbs.Execute ("11 Update Checklisten Modulinformation")
    dbs.Execute ("11 Update Checklisten NaechstePruefung")
    dbs.Execute ("11 Update Checklisten OnSiteFlag")
    dbs.Execute ("11 Update Checklisten Plakette")
    dbs.Execute ("11 Update Checklisten Pos")
    dbs.Execute ("11 Update Checklisten PossibleStatus")
    dbs.Execute ("11 Update Checklisten Required")
    dbs.Execute ("11 Update Checklisten SecurityRelatedComments")
    dbs.Execute ("11 Update Checklisten Serviceart")
    dbs.Execute ("11 Update Checklisten ServiceartText")
    dbs.Execute ("11 Update Checklisten Softwareversion")
        
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed service article to SharePoint")
    dbs.Execute ("17 Append Artikel")
    dbs.Execute ("18 Update Artikel Artikelart")
    dbs.Execute ("18 Update Artikel BeschreibungDE")
    dbs.Execute ("18 Update Artikel BeschreibungEN")
    dbs.Execute ("18 Update Artikel BeschreibungXX")
    dbs.Execute ("18 Update Artikel Geraetetyp")
    dbs.Execute ("18 Update Artikel Hersteller")
    dbs.Execute ("18 Update Artikel MwstSatz")
    dbs.Execute ("18 Update Artikel Netto")
    dbs.Execute ("18 Update Artikel PreisProStunde")
    dbs.Execute ("18 Update Artikel Serviceart")
    dbs.Execute ("18 Update Artikel StdDauerMin")
    dbs.Execute ("18 Update Artikel StdDauerMin")
    dbs.Execute ("19 Delete Artikel")
        
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed system settings to SharePoint")
    dbs.Execute ("20 Append SystemSettings")
    dbs.Execute ("21 Update SystemSettings")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed languages to SharePoint")
    dbs.Execute ("28 Append Sprachen")
    dbs.Execute ("29 Update Sprachen")
    dbs.Execute ("28 Append Woerter from ZSprachen")
    dbs.Execute ("29 Update ZSprachen")
    dbs.Execute ("29 Delete Woerter")
    dbs.Execute ("29 Delete Woerter from ZSprachen")
    
    Call add_entry_to_log(strProtokoll, strMyPath, "     Uploading new and changed app version information to SharePoint")
    dbs.Execute ("30 Append Versionen")
    dbs.Execute ("31 Update Versionen")

Fin:
    DoCmd.SetWarnings False
    DoCmd.Hourglass False
    Set dbs = Nothing
    If Err.Number <> 0 Then
        'ErrNumber = Err.Number
        Call add_entry_to_log_always(strProtokoll, strMyPath, "Error: " & Err.Number & " " & Err.Description)
    End If
        
End Function

