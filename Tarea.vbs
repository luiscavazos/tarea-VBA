Sub Tarea()

'Loop through the sheets
Dim ws As Worksheet
For Each ws In Sheets
  
' Set an initial variable for holding the Ticker
  Dim Ticker_Name As String

  ' Set an initial variables for holding the totals
  Dim Open_Total As Double
  Open_Total = 0
  Dim Close_Total As Double
  Close_Total = 0
  Dim Stock_Total As Double
  Stock_Total = 0
  

  ' Keep track of the location for each Ticker
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop
  Dim i As Long
  For i = 2 To ws.Range("A" & Rows.Count).End(xlUp).Row
  '70927 is final row in A
  

    ' Check if we are still within the same ticker and then check what to do if its not
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Brand name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Add to the Brand Total
      Open_Total = Open_Total + ws.Cells(i, 3).Value
      Close_Total = Close_Total + ws.Cells(i, 6).Value
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Yearly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Close_Total - Open_Total
      
      ' Color format the Cell Green if positive change and Red if negative change
      If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If
      
      'In order to avoid the PNLT Ticker error
        If Open_Total = 0 And Close_Total = 0 Then
            ws.Range("K" & Summary_Table_Row).Value = 0
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("L" & Summary_Table_Row).Value = Stock_Total
        End If
        
      'Print the Percent Change in the Summary Table and in Percentage Format
      ws.Range("K" & Summary_Table_Row).Value = (Close_Total - Open_Total) / Open_Total
      'This to avoid crash in case an amount is divided by 0
      On Error Resume Next
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Print The Total Stock Volume
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Totals
      Open_Total = 0
      Close_Total = 0
      Stock_Total = 0

    ' If the cell immediately following a row is the same ticker
    Else

      ' Add to the totals
      Open_Total = Open_Total + ws.Cells(i, 3).Value
      Close_Total = Close_Total + ws.Cells(i, 6).Value
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub

