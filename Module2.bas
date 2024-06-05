Attribute VB_Name = "Module2"
Sub bonus()

    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_total_volume As Double
    Dim ticker_increase As String
    Dim ticker_decrease As String
    Dim ticker_stock As String
    
    lastRow = Cells(Rows.Count, 11).End(xlUp).Row
    lastRowTotal = Cells(Rows.Count, 12).End(xlUp).Row
    greatest_increase = 0
    greatest_decrease = 0
    greatest_total_volume = 0
    
    
    For i = 2 To lastRow
    
        
        If Cells(i, 11) > greatest_increase Then
            greatest_increase = Cells(i, 11).value
            ticker_increase = Cells(i, 9).value
        End If
        
        If Cells(i, 11) < greatest_decrease Then
            greatest_decrease = Cells(i, 11).value
            ticker_decrease = Cells(i, 9).value
        End If
                
    Next i
    
    Cells(2, 17).value = greatest_increase
    Cells(2, 16).value = ticker_increase
    Cells(3, 17).value = greatest_decrease
    Cells(3, 16).value = ticker_decrease
    
    For i = 2 To lastRowTotal
        
        If Cells(i, 12).value > greatest_total_volume Then
            greatest_total_volume = Cells(i, 12).value
            ticker_stock = Cells(i, 9).value
        End If
    Next i
    
    Cells(4, 17).value = greatest_total_volume
    Cells(4, 16).value = ticker_stock
    
End Sub
