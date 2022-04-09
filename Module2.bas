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
Dim open1 As Double
Dim close1 As Double
Dim beginning As Integer
Dim change As Integer
Dim ticker As String
Dim volume As Double
volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
For c = 2 To 2
  
  For r = 2 To lastrow
    
    If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
      
      ticker = Cells(r, 1).Value
    
      volume = volume + Cells(r, 7).Value
      
      Range("I" & Summary_Table_Row).Value = ticker
      
      Range("L" & Summary_Table_Row).Value = volume
     
      Summary_Table_Row = Summary_Table_Row + 1
            
      volume = 0
    
    Else

      volume = volume + Cells(r, 7).Value

    End If


  Next r
  
    
    
    If Cells(c, 2) = "20160101" Then
      
      Cells(c, 6).Value = open1
      
      
    End If
  
    If Cells(c, 2) = "20161231" Then
      
      Cells(rr, 3).Value = close1
     
      
    End If
    
    
     Cells(r, 10) = close1 - open1
     
Next c
        
        
  
  


End Sub
