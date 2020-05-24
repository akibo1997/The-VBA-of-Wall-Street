Attribute VB_Name = "Module1"
Sub Stock_Ticker()
    ' Establish worksheets and active worksheets
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        ' Calculate the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Headings for each worksheet
        Cells(1, "L").Value = "Total Stock Volume"
        Cells(1, "I").Value = "Ticker"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "J").Value = "Yearly Change"
        
        'Creation of variables
        Dim yropen As Double
        Dim yrclose As Double
        Dim pctchange As Double
        ' initialize total volume as zero
        
        Dim totalvolume As Double
        totalvolume = 0
        Dim yrchange As Double
        Dim ticker As String
        
        Dim stocktickerrow As Double
        stocktickerrow = 2
        Dim column As Integer
        column = 1
        Dim i As Long
        
        ' The opening price as the year open price
        yropen = Cells(2, column + 2).Value
        
        
        For i = 2 To lastrow
         ' If the next ticker value is not equal to the next one, then calculate these values
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                
                ' Set Ticker name
                ticker = Cells(i, column).Value
                Cells(stocktickerrow, column + 8).Value = ticker
                
                ' Set Close Price
                yrclose = Cells(i, 5).Value
                
                ' summate yearly change
                yrchange = yrclose - yropen
                Cells(stocktickerrow, 9).Value = yrchange
                Cells(stocktickerrow, 9).NumberFormat = "0000.00"
                
                ' Calculate percent change
                
                If (yropen = 0 And yrclose = 0) Then
                    pctchange = 0
                ElseIf (yropen = 0 And yrclose <> 0) Then
                    pctchange = 1
                Else
                    pctchange = yrchange / yropen
                    Cells(stocktickerrow, "K").Value = pctchange
                    Cells(stocktickerrow, "K").NumberFormat = "0.00%"
                End If
                ' Add total volume, then go to the next row
                totalvolume = Volume + Cells(i, 6).Value
                Cells(stocktickerrow, column + 11).Value = totalvolume
                
                ' Go to the next row in the new table
                stocktickerrow = stocktickerrow + 1
                
                ' Reset year opening price
                yropen = Cells(i + 1, column + 2)
                
                ' Reset the total volume
                totalvolume = 0
            'if cells are the same ticker
            Else
                totalvolume = totalvolume + Cells(i, 6).Value
            End If
        Next i
        
        ' Determine the last row for every worksheet
        YCLastRow = ws.Cells(Rows.Count, column + 8).End(xlUp).Row
    
    Next ws
End Sub
