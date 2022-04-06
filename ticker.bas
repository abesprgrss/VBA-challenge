Attribute VB_Name = "Module1"

Sub tickerloop()

Cells(3, 9) = "Ticker"
Cells(3, 10) = "Yearly Change"
Cells(3, 11) = "Percent Change"
Cells(3, 12) = "Total Stock Volume"

'Dim ticker As String

'ticker = 0

'For each ws in Worksheet

'lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'For r = 2 To lastrow
 ' If Cells(r, 1) = "A" Then
 ' ticker = ticker + 1
  
  Next r
 '

  
'End If
End Sub
Sub CCexample()


Dim WS_Count As Integer
Dim sht As Integer
WS_Count = ActiveWorkbook.Worksheets.Count
 Dim WS_Count As Integer
 Dim I As Integer
 
 WS_Count = ActiveWorkbook.Worksheets.Count
 
 For I = 1 To WS_Count
 
   
   
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

  ' Set an initial variable for holding the brand name
  Dim ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim volume As Double
  volume = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all credit card purchases
  For I = 2 To lastrow

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Brand name
      ticker = Cells(I, 1).Value

      ' Add to the Brand Total
      volume = volume + Cells(I, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker

      ' Print the Brand Amount to the Summary Table
      Range("L" & Summary_Table_Row).Value = volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      volume = volume + Cells(I, 7).Value

    End If

  Next I
  Next sht


End Sub

Sub Worksheetloop()
 Dim WS_Count As Integer
 Dim I As Integer
 
 WS_Count = ActiveWorkbook.Worksheets.Count
 
 For I = 1 To WS_Count
 
   MsgBox (ActiveWorkbook.Worksheets(I).Name)
   
Next I
 
 
End Sub


