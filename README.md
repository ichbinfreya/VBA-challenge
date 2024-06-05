# VBA-challenge
Visual Basic for Applications

What is the definition of VBA?

VBA is a programming language that interacts directly with Excel worksheets and cells, allowing us to write scripts to automate simple tasks. It also allows us to write complicated scripts to perform complex analyses, which we wouldn't be able to do with Excel alone. And in developer parlance, automated tasks are called macros. 

How to access the VBA?

The Developer ribbon can lead you to VBA and marcos.

For this case, we are going to sort out the dataset of multiple tickers.

In the VBA code, there are couples of function I used, which are "For" loops and "If" statement
I fomatted the color by using Cells(x,y).Interior.ColorIndex rather than conditional formmating in Excel.


Here is what I scripted in a VBA:

'Create a subroutine
Sub tickerName()
    'Declare the variables
    Dim i As Long
    Dim j As Long
    Dim ticker_name As String
    Dim open_price As Double
    Dim close_price As Double
    Dim stock_volume As Double
    Dim quarterly_change As Double
    Dim percent_change As Double

    'I just make the first row title as default, so each time I applied the marcos, the default title can appear.
    Cells(1, 9).value = "Ticker"
    Cells(1, 10).value = "Quarterly Change"
    Cells(1, 11).value = "Percent Change"
    Cells(1, 12).value = "Total Stock Volume"
    
    'Define output row variable
    j = 2
    'Define ticker_name variable
    ticker_name = Cells(2, 1).value
    open_price = Cells(2, 3).value
    stock_volume = Cells(2, 7).value
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 3 To lastRow
    ' the value starts from 3rd row as I have already set the first value.
        
        'Total stock volume added up as this value is a dependent variable. 
        stock_volume = stock_volume + Cells(i, 7).value
        
        If Cells(i, 1).value <> Cells(i + 1, 1).value Then
            
            close_price = Cells(i, 6).value
            quarterly_change = close_price - open_price
            percent_change = quarterly_change / open_price
        
            'MsgBox (Cells(i + 1 , 1).Value)
            Cells(j, 9).value = ticker_name
            Cells(j, 10).value = quarterly_change
            
            If Cells(j, 10).value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            Cells(j, 11).value = percent_change
            Cells(j, 12).value = stock_volume
            
            ticker_name = Cells(i + 1, 1).value
            j = j + 1
            open_price = Cells(i + 1, 3).value
            
            stock_volume = 0
            
        End If
    
    Next i
    
    'MsgBox ("Ticker's names have been sorted.")
    'I just wnant to text if each row of data has been rotated in the loop. I can always make it a comment or guidance.
End Sub
