Attribute VB_Name = "Module1"
Sub Stocks()
Dim I As Long
Dim ticker As String
Dim yearly_change As Double
Dim percentage As Double
Dim total_volume As Double
total_volume = 0
Dim x As Double
Dim y As Double


Dim WS As Worksheet

         ' Loop through all of the worksheets in the active workbook.
        For Each WS In Worksheets
         
            N_Rows = WS.Cells(Rows.Count, "A").End(xlUp).Row
        
        
        
            WS.Cells(1, 9).Value = "ticker"
            WS.Cells(1, 10).Value = "yearly change"
            WS.Cells(1, 11).Value = "percentage change"
            WS.Cells(1, 12).Value = "total volume"
            x = 2
            y = 2
        
            For I = 2 To N_Rows
                If WS.Cells(I + 1, 1).Value <> WS.Cells(I, 1).Value Then
       
                    total_volume = total_volume + WS.Cells(I, 7).Value
        
                    yearly_change = WS.Cells(I, 6).Value - WS.Cells(x, 3).Value
                    percentage = yearly_change / WS.Cells(x, 3).Value
        
                    WS.Cells(y, 9).Value = WS.Cells(I, 1).Value
                    WS.Cells(y, 10).Value = yearly_change
                    WS.Cells(y, 10).NumberFormat = "0.00"
                    WS.Cells(y, 11).Value = percentage
                    WS.Cells(y, 11).NumberFormat = "0.00%"
                    WS.Cells(y, 12).Value = total_volume
        
        
                    If WS.Cells(y, 10).Value > 0 Then
                        WS.Cells(y, 10).Interior.ColorIndex = 4
                    Else
                        WS.Cells(y, 10).Interior.ColorIndex = 3
                    End If
                    y = y + 1
                    x = I + 1
                    
                    

      Else
            total_volume = total_volume + WS.Cells(I, 7).Value
    End If
        
        Next I

            ' Insert your code here.
            ' This line displays the worksheet name in a message box.
           'MsgBox WS.Name
         Next WS


End Sub

