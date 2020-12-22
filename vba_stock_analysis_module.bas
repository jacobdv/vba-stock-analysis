Attribute VB_Name = "Module2"
Sub stockAnalysis()
    Dim rowCount As LongLong
    Dim columnCount As Long
    Dim stockCount As Long
    Dim volTotal As LongLong
    Dim i As LongLong
    Dim conditionalRange As Range
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearlyChange As Double
    
    ' Defines default values for some variables.
    stockCount = 1
    volTotal = 0

    ' Counts number of columns and rows with data.
    columnCount = Cells(1, Columns.Count).End(xlToLeft).Column
    rowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Sets column headers for analysis columns.
    Cells(1, columnCount + 2).Value = "Ticker"
    Cells(1, columnCount + 3).Value = "Yearly Change"
    Cells(1, columnCount + 4).Value = "Percent Change"
    Cells(1, columnCount + 5).Value = "Total Stock Volume"

    ' Iterates through stocks, noting when it switches names(increases stockCount).
    For i = 2 To rowCount + 1
        ' If statement for "Ticker" column.
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            stockCount = stockCount + 1
            Cells(stockCount, columnCount + 2).Value = Cells(i, 1).Value
        End If
        
        
        
        ' If statement for "Yearly Change" column.
        ' COMMENTED OUT BECAUSE YEAR CLOSE LINE ISN'T WORKING AND I NEED TO DO SOMETHING ELSE
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            yearOpen = Cells(i, 3).Value
            yearClose = Cells(i - 1, 6).Value
            yearlyChange = yearClose - yearOpen
            Cells(stockCount, columnCount + 3).Value = yearlyChange
        End If

        If i <> 2 Then
            ' If statement for "Total Stock Volume" column.
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                Cells(stockCount - 1, columnCount + 5).Value = volTotal
                volTotal = 0
            Else
                volTotal = volTotal + Cells(i, 7).Value
            End If
        End If
    Next i
    
    ' Conditional formatting for "Yearly Change" column.
    ' Set conditionalRange = Range(Cells(2, columnCount + 3), Cells(stockCount, columnCount + 3))
    
End Sub
