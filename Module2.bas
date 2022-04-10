Attribute VB_Name = "Module2"
Sub SheetCycle()

 Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()
 

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percentage Change"
Cells(1, 12) = "Stock Volume"


Dim Ticker As String
Dim volume As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim change As Double
Dim output As Double
Dim total As Double
Dim percent As Double
volume = 0
result = 2
close1 = 2

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  
  For r = 2 To LastRow
    
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
      
      Cells(result, 10).Value = change
      percent = Cells(result, 10).Value / Cells(close1, 3).Value
      Cells(result, 11).NumberFormat = "0.00%"
      change = Cells(r, 6) - Cells(close1, 6)
      close1 = r + 1
      
      Cells(result, 11).Value = percent
      Cells(result, 12).Value = total
      result = result + 1
      total = 0
      
      Ticker = Cells(r, 1).Value
      volume = volume + Cells(r, 7).Value
      Range("I" & Summary_Table_Row).Value = Ticker
      Range("L" & Summary_Table_Row).Value = volume
      Summary_Table_Row = Summary_Table_Row + 1
      volume = 0
      
  
    Else
   
      volume = volume + Cells(r, 7).Value

End If
Next r


    
        
End Sub
