Sub Yearly_analysis()

' Defining Variables
Dim ticker As String
Dim lastRow As Double
Dim lastRow2 As Double
Dim open_value As Double
Dim close_value As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_volume As Double
Dim volume As Double
Dim volume_variable As Double
Dim ws As Worksheet
Dim row_tab As Double
Dim date_start As String
Dim date_end As String
Dim datevar As String
Dim yc As Double
Dim Change_percent As Double
Dim max_value As Double
Dim min_value As Double
Dim max_volume As Double
Dim i As Double
Dim j As Double
Dim max_ticker As String
Dim min_ticker As String
Dim max_volume_ticker As String
Dim response As Variant
Dim Rng As Variant
Dim Rng2 As Variant
Dim Rng3 As Variant

'Worksheet creation list

For Each ws In Worksheets
' Initial Variables
    row_tab = 2
    j = 2
    i = 2
    max_value = 0
    min_value = 0
    max_volume = 0
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    date_start = "0102"
    date_end = "1231"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
  If j <> lastRow + 1 Then
    For j = 2 To lastRow
        datevar = Right(Str(ws.Cells(j, 2).Value), 4)
            If datevar = date_start Then
                total_volume = 0
                ticker = ws.Cells(j, 1).Value
                 ws.Cells(row_tab, 9).Value = ticker
                 open_value = ws.Cells(j, 3).Value
                 volume = ws.Cells(j, 7).Value
                 total_volume = total_volume + volume
   
            ElseIf datevar = date_end Then
                close_value = ws.Cells(j, 6).Value
                yearly_change = close_value - open_value
                percent_change = ((close_value - open_value) / open_value)
                ws.Cells(row_tab, 10).Value = yearly_change
                ws.Cells(row_tab, 11).Value = percent_change
                volume = ws.Cells(j, 7).Value
                total_volume = total_volume + volume
                ws.Cells(row_tab, 12).Value = total_volume
                row_tab = row_tab + 1
     
            Else
                volume = ws.Cells(j, 7).Value
                total_volume = total_volume + volume
        
            End If

    Next j
        
  End If
  
 'Worksheet analysis
   response = MsgBox("Would you like to analyze this data set " & ws.Name & "?", vbYesNo)
   If response = vbYes Then
   
    lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To lastRow2
     yc = ws.Cells(i, 10).Value
     Change_percent = ws.Cells(i, 11).Value
     volume_variable = ws.Cells(i, 12).Value
     If volume_variable > max_volume Then
            max_volume = volume_variable
            max_volume_ticker = ws.Cells(i, 9).Value
     End If
     
     If yc < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3

            If Change_percent < min_value Then
                min_value = Change_percent
                min_ticker = ws.Cells(i, 9).Value
            End If
            
     ElseIf yc > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
         If Change_percent > max_value Then
            max_value = Change_percent
            max_ticker = ws.Cells(i, 9).Value
         End If
        
     End If

    Next i

    ws.Cells(2, 16).Value = max_ticker  'Greatest % Increase name
    ws.Cells(2, 17).Value = max_value   'Greatest % Increase value
    ws.Cells(3, 16).Value = min_ticker  'Greatest % Decrease name
    ws.Cells(3, 17).Value = min_value   'Greatest % Decrease value
    ws.Cells(4, 16).Value = max_volume_ticker  'Greatest % stock volume name
    ws.Cells(4, 17).Value = max_volume         'Greatest % stock volume value
    lastRow3 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    Set Rng = ws.Range("K2:K" & lastRow2)
    Rng.NumberFormat = "0.00%"
    Set Rng2 = ws.Range("Q2:Q3")
    Rng2.NumberFormat = "0.00%"

   Else
   lastRow3 = ws.Cells(Rows.Count, 9).End(xlUp).Row
   Set Rng3 = ws.Range("K2:K" & lastRow3)
   Rng3.NumberFormat = "0.00%"

   End If

Next ws

End Sub

