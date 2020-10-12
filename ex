
Public Sub Funwithloops()

Range("A1").Select

Do While ActiveCell.Value <> ""

ActiveCell.Font.Bold = True

ActiveCell.Offset(1, 0).Select

Loop

End Sub






__________________________________________


Public Sub Funwithloops()
 Dim i As Integer
 
 i = 1
 
 Do While i <= 10
    If ActiveCell.Value > 10 Then
    ActiveCell.Interior.Color = RGB(255, 0, 0)
    
    End If
    
    ActiveCell.Offset(1, 0).Select
    
   
     
     i = i + 1
     
 
 Loop
 

End Sub




__________________________________________
