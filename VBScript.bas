Attribute VB_Name = "Module1"
Sub tickersymbol()

For Each ws In Worksheets

Dim tickersymbol As String

Dim tickervolume As Double

'Dim openprice As Double

'Dim closeprice As Double

Dim yearlychg As Double

Dim percentchg As Double

Dim lastrow As Long
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

'Dim ws As Worksheet
'Set ws = Worksheets("A")

'tickersymbol = ws.Cells(irow, 1).Value
openprice = ws.Cells(2, 3).Value
'closeprice = ws.Cells(2, 6).Value
tickervolume = 0
output_row = 2

'column header
ws.Cells(1, 11).Value = "Ticker"
ws.Cells(1, 12).Value = "Yearly Change"
ws.Cells(1, 13).Value = "Percentage Change"
ws.Cells(1, 14).Value = "Total Stock Volume"


    For irow = 2 To lastrow
       tickersymbol = ws.Cells(irow, 1).Value
    If ws.Cells(irow + 1, 1).Value <> tickersymbol Then
     
    tickervolume = tickervolume + ws.Cells(irow, 7).Value
          closeprice = ws.Cells(irow, 6).Value
        yearlychg = closeprice - openprice
        If yearlychg > 0 Then
        ws.Cells(output_row, 12).Interior.ColorIndex = 4
        Else
        ws.Cells(output_row, 12).Interior.ColorIndex = 3
        End If
        
        'place yearlychg output here
        
        If openprice = 0 Then
        percentchg = 0
        
        Else
        percentchg = (yearlychg / openprice)
        
        End If
        
        
        
             'place ticker output here
        ws.Cells(output_row, 11).Value = tickersymbol
        ws.Cells(output_row, 12).Value = yearlychg
      
        'place percentage change outout here
        ws.Cells(output_row, 13).Value = percentchg
        ws.Cells(output_row, 13).NumberFormat = "0.00%"
        ws.Cells(output_row, 14).Value = tickervolume
        
        'incrementing output_row
        output_row = output_row + 1
        'reset total volume so that it will count for each ticker
       tickervolume = 0
       'finding the opening price for each ticker
       openprice = ws.Cells(irow + 1, 3).Value
        'place tickervolume output here
       Else
       tickervolume = tickervolume + ws.Cells(irow, 7).Value
       yearlychg = closeprice - openprice
        

    End If
        Next irow
        Next ws
 
End Sub
