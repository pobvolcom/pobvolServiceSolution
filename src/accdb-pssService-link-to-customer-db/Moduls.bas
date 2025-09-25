Attribute VB_Name = "Moduls"
Sub add_entry_to_log_always(strProtokoll, strMyPath, strMessage)
    Dim x As Long
    Dim TB, Pfad As String, Datei As String
    
        Datei = strMyPath & "\log.txt"
        Close #1
        If Dir(Datei) = "" Then
            Open Datei For Output As 1
        Else
            Open Datei For Append As 1
        End If
        Print #1, Format(Now, "YYYY_MM_DD") & "/" & Format(Now, "hh:mm:ss") & " " & strMessage
        Close #1
End Sub

Sub add_entry_to_log(strProtokoll, strMyPath, strMessage)
    Dim x As Long
    Dim TB, Pfad As String, Datei As String
    
    If LCase(strProtokoll) = "true" Then
        Datei = strMyPath & "\log.txt"
        Close #1
        If Dir(Datei) = "" Then
            Open Datei For Output As 1
        Else
            Open Datei For Append As 1
        End If
        Print #1, Format(Now, "YYYY_MM_DD") & "/" & Format(Now, "hh:mm:ss") & " " & strMessage
        Close #1
    End If
End Sub


Function GetrstTemp(qdf As QueryDef)
 
   Dim rstTemp As Recordset
 
   With qdf
      Debug.Print .Name
      Debug.Print "  " & .SQL
      ' Open Recordset from QueryDef.
      Set rstTemp = .OpenRecordset(dbOpenSnapshot)
 
      With rstTemp
         ' Populate Recordset and print number of records.
         .MoveLast
         Debug.Print "  Number of records = " & _
            .RecordCount
         Debug.Print
         .Close
      End With
 
   End With
 
End Function

Sub RefreshLinkOutput(dbs, strObjectName)
 
    Dim rstRemote As Recordset
    Dim intCount As Integer
    
    ' Open linked table.
    Set rstRemote = dbs.OpenRecordset(strObjectName)
    
    intCount = 0
    
    ' Enumerate Recordset object, but stop at 50 records.
    With rstRemote
       Do While Not .EOF And intCount < 50
           Debug.Print , .Fields(0), .Fields(1)
           intCount = intCount + 1
           .MoveNext
       Loop
       If Not .EOF Then Debug.Print , "[more records]"
    End With
 
End Sub


'INV3: IIf([Expr16-5];GetItem([Expr16-5];";";2);"")
Public Function GetItem(strText As String, strDelimiter As String, part As Integer)
On Error Resume Next
    Dim LArray() As String
    GetItem = ""
    'GetItem = LTrim(Split(strText, strDelimiter)(part))
    LArray = Split(strText, strDelimiter)
    For i = LBound(LArray) To UBound(LArray)
        If part = i And LArray(i) > "" Then GetItem = LTrim(LArray(i))
    Next i

End Function

' Ersetzt alle nicht zulässigen Zeichen im angegebenen Dateinamen
Public Function CleanFilename(ByVal sFilename As String, Optional ByVal sChar As String = "") As String
  Dim oRegExp As RegExp
  Set oRegExp = New RegExp
  With oRegExp
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    .Pattern = "[\\/:?*^""<>|]"
 
    ' alle nicht zulässigen Zeichen ersetzen
    CleanFilename = .Replace(sFilename, sChar)
  End With
  Set oRegExp = Nothing
End Function





