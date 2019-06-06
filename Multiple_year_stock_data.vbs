' Chris Glessner

' List of columns.  Delete when done
' 1         2        3        4        5     6        7
' <ticker>  <date>  <open>  <high>  <low>   <close> <vol>

Sub Summarize_Sheets()

   Dim ws As Worksheet
   Dim starting_ws As Worksheet

   ' Remember which worksheet is active in the beginning
   Set starting_ws = ActiveSheet

   For Each ws In ThisWorkbook.Worksheets
      ws.Activate

      ' What is the last row on this sheet?
      Dim lastRow As Long
      lastRow = ws.Range("A1").CurrentRegion.Rows.Count

      ' What is the year supposed to be on this sheet?
      Dim sheetName As String
      sheetName = ws.Name
      If (Not Len(sheetName) = 4) Then
         MsgBox ("Unexpected length for name of sheet: " + sheetName + "." + vbCr + "Expecting year.")
         Exit Sub
      End If

      ' Read through every row in the
      For row_num = 2 To 5 'lastRow
         ' Does the date value make sense?
         If ((Not Left(ws.Cells(row_num, 2), 4) = sheetName) Or (Not Len(ws.Cells(row_num, 2)) = 8)) Then
            MsgBox ("Unexpected value (" + ws.Cells(row_num, 2) + ") on sheet " + sheetName + " at row " & row_num & "." + vbCr + "Was expecting date.")
         End If
      Next row_num

      'do whatever you need
      MsgBox ("Last row is: " & lastRow)
   Next ws

   ' Activate the worksheet that was originally active
   starting_ws.Activate

End Sub
