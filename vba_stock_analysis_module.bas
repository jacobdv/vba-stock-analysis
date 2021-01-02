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
    Dim changePercentage As Double
    Dim ws As Worksheet
    
    For Each ws In Sheets
        ws.Activate
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
                ' This evaluation for "i" not equaling two counteracts first row problems.
                If i <> 2 Then
                    yearClose = Cells(i - 1, 6).Value
                    yearlyChange = yearClose - yearOpen
                    
                    ' Prevents div by 0 error.
                    If yearOpen = 0 Then
                        changePercentage = 0
                    Else
                        changePercentage = (yearlyChange / yearOpen)
                    End If
                    
                Cells(stockCount - 1, columnCount + 3).Value = yearlyChange
                Cells(stockCount - 1, columnCount + 4).Value = changePercentage
                Cells(stockCount - 1, columnCount + 4).NumberFormat = "0.00%"
                yearOpen = Cells(i, 3).Value
                Cells(stockCount - 1, columnCount + 5).Value = volTotal
                volTotal = 0
                Else
                volTotal = volTotal + Cells(i, 7).Value
                End If
            Else
                yearOpen = Cells(i, 3).Value
            End If
        Next i
                
        For i = 2 To stockCount
            ' Conditional formatting for "Yearly Change" column.
            If Cells(i, columnCount + 3) > 0 Then
                Cells(i, columnCount + 3).Interior.ColorIndex = 4
            ElseIf Cells(i, columnCount + 3) < 0 Then
                Cells(i, columnCount + 3).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub
