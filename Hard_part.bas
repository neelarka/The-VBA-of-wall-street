Attribute VB_Name = "Module9"
Sub Hard():
    Dim ticker As String
    Dim vol As Double
    Dim Summary_Table_Row As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim year_percent As Double
    Dim x As Long
    Dim y As Integer
    Dim Row1 As String
    Dim Row2 As String
    Dim Row3 As String
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
        vol = 0
      
        
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly change"
        ws.Cells(1, 11).Value = "Percentage change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Summary_Table_Row = 2
        
' Getting the value for the last row in the sheet

         x = ws.Cells(Rows.Count, 1).End(xlUp).Row
          
' Intialising the year_open value for the the first iteration

        year_open = ws.Cells(2, 3).Value
        
        For i = 2 To x
        
       
        If ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           ticker = ws.Cells(i, 1).Value
           vol = vol + ws.Cells(i, 7).Value
 
           ws.Range("I" & Summary_Table_Row).Value = ticker
           
' To avoid the division by zero condition
            If year_open = 0 Then
        
            year_open = ws.Cells(i, 3).Value
            
            End If
            
           year_close = ws.Cells(i, 6).Value
           yearly_change = year_close - year_open
           
' To avoid division by zero condition when both year_open and year_close is zero
            If year_change = 0 And year_open = 0 Then
        
            year_percent = 0
            
             Else
           
           year_percent = yearly_change / year_open
           
            End If
            
            If year_percent > 0 Then
            
             ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
            Else
            
             ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
            End If
           
               year_open = ws.Cells(i + 1, 3).Value
           
           
           ws.Range("J" & Summary_Table_Row).Value = yearly_change
           ws.Range("K" & Summary_Table_Row).Value = year_percent
           ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
           ws.Range("L" & Summary_Table_Row).Value = vol
           Summary_Table_Row = Summary_Table_Row + 1
           vol = 0
         
         Else
           
            vol = vol + ws.Cells(i, 7).Value
           

      End If


                Next i
                
        ws.Range("P" & 1).Value = "ticker"
        ws.Range("Q" & 1).Value = "Value"
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        
        
 ' Code to find the greatest Total stock volume
            
               
                 ws.Range("Q" & 4).Value = Application.WorksheetFunction.Max(ws.Columns("L"))
             
                 
'Greatest % Increase
                 
                 ws.Range("Q" & 2).Value = Application.WorksheetFunction.Max(ws.Columns("K"))
                 
                  ws.Cells(2, 17).NumberFormat = "0.00%"
                 
'Greatest % Decrease
                 
                 ws.Range("Q" & 3).Value = Application.WorksheetFunction.Min(ws.Columns("K"))
                
                 ws.Cells(3, 17).NumberFormat = "0.00%"
                      
                 

                     
' To get the ticker with Greatest Total
                     y = ws.Cells(Rows.Count, 9).End(xlUp).Row


                      For j = 2 To y
                        If (ws.Cells(j, 12).Value - ws.Cells(4, 17).Value) = 0 Then
                             Row2 = ws.Cells(j, 9).Value
                             
                        
                    Exit For
                        End If
                    Next j
                    
                     ws.Cells(4, 16).Value = Row2
                     
                     
' To get the ticker with Greatest % Increase

                 
                    For j = 2 To y
                        If ws.Cells(j, 11).Value = ws.Cells(2, 17).Value Then
                             Row1 = ws.Cells(j, 9).Value
                             
                             MsgBox ("The ticker with greatest total" & Row1)
                 
                        
                    Exit For
                       End If
                       
                    Next j
                    
                     ws.Cells(2, 16).Value = Row1


                    
' To get the ticker with Greatest % Decrease


                     For j = 2 To y
                        If ws.Cells(j, 11).Value = ws.Cells(3, 17).Value Then
                             Row3 = ws.Cells(j, 9).Value
                        
                    Exit For
                        End If
                    Next j
                    
                     ws.Cells(3, 16).Value = Row3
    
    
    Next ws

End Sub



