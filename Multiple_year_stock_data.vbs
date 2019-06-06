' Chris Glessner

' 1         2       3       4       5       6       7
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
      lastRow = ws.Range("A:A").End(xlDown).Row
      ' Check that code above is working correctly
      ' MsgBox ("Last row is: " & lastRow)

      ' What is the year supposed to be on this sheet?
      Dim sheetName As String
      sheetName = ws.Name
      If (Not Len(sheetName) = 4) Then
         MsgBox ("Unexpected length for name of sheet: " + sheetName + "." + vbCr + "Expecting year.")
         Exit Sub
      End If

      ' NOTE: The code below assumes the data are sorted by <ticker> then <date>
      ' For proper function, sort the data prior to execution
      ' I tried implementing this programatically, but was sorry I did

      ' Create summary section headers
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      
      Dim outRow As Integer
      outRow = 1
      
      ' Value needed to define yearly change
      Dim firstOpen As Double

      ' Read through every row in the
      For row_num = 2 To lastRow
         ' Does the date value make sense?
         If ((Not Left(ws.Cells(row_num, 2), 4) = sheetName) Or (Not Len(ws.Cells(row_num, 2)) = 8)) Then
            MsgBox ("Unexpected value (" + ws.Cells(row_num, 2) + ") on sheet " + sheetName + " at row " & row_num & "." + vbCr + "Was expecting date.")
         Else
            ' First record of a new type (or just first record)
            If ws.Cells(row_num, 1).Value <> ws.Cells(row_num - 1, 1).Value Then
            
                ' Fill out the summary values for the previous index
                If row_num <> 2 Then
                    ws.Cells(outRow, 10) = ws.Cells(row_num - 1, 6).Value - firstOpen
                    If firstOpen <> 0 Then
                        ws.Cells(outRow, 11).Value = ws.Cells(outRow, 10).Value / firstOpen
                        ws.Cells(outRow, 11).NumberFormat = "0.00%"
                        If ws.Cells(outRow, 10).Value > 0 Then
                           ws.Cells(outRow, 10).Interior.Color = RGB(0, 255, 0)
                        Else
                            ws.Cells(outRow, 10).Interior.Color = RGB(255, 0, 0)
                        End If
                    Else
                        ws.Cells(outRow, 11).Value = "(Undefined)"
                    End If
                End If
                
                ' Get output row for summary values for current index and start entry
                outRow = outRow + 1
                ws.Cells(outRow, 9).Value = ws.Cells(row_num, 1).Value
                
                ' Retain opening stock value for yearly change computation
                firstOpen = ws.Cells(row_num, 3).Value
                
                ' Initialize Total Stock Volume column to first volume value
                ws.Cells(outRow, 12).Value = ws.Cells(row_num, 7).Value
            
            ' Not the first record, not the last record
            ElseIf row_num <> lastRow Then
                ws.Cells(outRow, 12).Value = ws.Cells(outRow, 12).Value + ws.Cells(row_num, 7).Value
                
            ' Last record
            Else
                ws.Cells(outRow, 12).Value = ws.Cells(outRow, 12).Value + ws.Cells(row_num, 7).Value
                ws.Cells(outRow, 10) = ws.Cells(row_num, 6).Value - firstOpen
                If firstOpen <> 0 Then
                    ws.Cells(outRow, 11).Value = ws.Cells(outRow, 10).Value / firstOpen
                    ws.Cells(outRow, 11).NumberFormat = "0.00%"
                    If ws.Cells(outRow, 10).Value > 1 Then
                       ws.Cells(outRow, 10).Interior.Color = RGB(0, 255, 0)
                    Else
                        ws.Cells(outRow, 10).Interior.Color = RGB(255, 0, 0)
                    End If
                Else
                    ws.Cells(outRow, 11).Value = "(Undefined)"
                End If
            End If
         End If
      Next row_num
      
      
      ' Overall summay
      
      ' What is the last row of the summary section?
      ws.Range("O1").Value = "Ticker"
      ws.Range("P1").Value = "Value"
      
      ws.Range("N2").Value = "Greatest % Increase"
      ws.Range("N3").Value = "Greatest % Decrease"
      ws.Range("N4").Value = "Greatest Total Volume"
     
      ' Initialize to values corresponding to first row
      ws.Range("O2").Value = ws.Cells(2, 9).Value
      ws.Range("O3").Value = ws.Cells(2, 9).Value
      ws.Range("O4").Value = ws.Cells(2, 9).Value
     
      ws.Range("P2").Value = ws.Cells(2, 11).Value
      ws.Range("P3").Value = ws.Cells(2, 11).Value
      ws.Range("P4").Value = ws.Cells(2, 12).Value
      
      Dim lastRowSum As Long
      lastRowSum = ws.Range("I:I").End(xlDown).Row
      
      ' Loop through data and update values in summary as needed
      For sumrow_num = 2 To lastRowSum
         
         If IsNumeric(ws.Cells(sumrow_num, 11)) Then
            If ws.Cells(sumrow_num, 11).Value > ws.Range("P2").Value Then
               ws.Range("O2").Value = ws.Cells(sumrow_num, 9).Value
               ws.Range("P2").Value = ws.Cells(sumrow_num, 11).Value
            End If
            
            If ws.Cells(sumrow_num, 11).Value < ws.Range("P3").Value Then
               ws.Range("O3").Value = ws.Cells(sumrow_num, 9).Value
               ws.Range("P3").Value = ws.Cells(sumrow_num, 11).Value
            End If
        End If
            
        If ws.Cells(sumrow_num, 12).Value > ws.Range("P4").Value Then
           ws.Range("O4").Value = ws.Cells(sumrow_num, 9).Value
           ws.Range("P4").Value = ws.Cells(sumrow_num, 12).Value
        End If

        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"

      Next sumrow_num
      
   Next ws

   ' Activate the worksheet that was originally active
   starting_ws.Activate

End Sub
