Attribute VB_Name = "Module1"
Sub Easy():

    Dim ticker As String
    Dim vol As Double
    Dim Summary_Table_Row As Integer
    
        vol = 0
        Cells(1, 9).Value = "ticker"
        Cells(1, 10).Value = "Total Stock Volume"
   
        Summary_Table_Row = 2
        
' Getting the value for the last row in the sheet

         x = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
        For i = 2 To x
        
     If Cells(i - 1, 1).Value = Cells(i, 1).Value And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           ticker = Cells(i, 1).Value
           vol = vol + Cells(i, 7).Value
           Range("I" & Summary_Table_Row).Value = ticker
           Range("J" & Summary_Table_Row).Value = vol
           
           Summary_Table_Row = Summary_Table_Row + 1
           vol = 0
      Else

            vol = vol + Cells(i, 7).Value


      End If


        Next i
 
 End Sub
