Attribute VB_Name = "Module3"
    
Sub High_low_change():

Dim open1 As Integer
Dim close1 As Integer
Dim change As Integer
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

 
 
 For c = 2 To Lastrow
  
  change = 0
  open1 = 0
  close1 = 0
  
  If Cells(c, 2) = 20160101 Then
      
     open1 = Cells(c, 6).Value
      
      
    End If
  
    If Cells(c, 2) = 20161231 Then
      
     close1 = Cells(c, 3).Value

      
    End If
    
    change = close1 - open1
    
    If change <> 0 Then
    Cells(c, 10).Value = change
    End If
    
    

 Next c






End Sub
    
  
