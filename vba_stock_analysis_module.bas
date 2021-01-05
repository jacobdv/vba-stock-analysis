Attribute VB_Name = "Module2"
Sub stockAnalysis()
    ' Full variable assignments. Some as longs or longlongs because I was running into overflow errors.
    Dim rowCount As LongLong
    Dim columnCount As Long
    Dim stockCount As Integer
    Dim volTotal As LongLong
    Dim i As LongLong
    Dim conditionalRange As Range
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim yearlyChange As Double
    Dim changePercentage As Double
    Dim ws As Worksheet
    
    ' Challenge Variables
    Dim greatestPercentageIncrease As Double
    Dim gpi As String
    Dim greatestPercentageDecrease As Double
    Dim gpd As String
    Dim volTotalHigh As LongLong
    Dim vthStock As String
    
    ' FOR LOOP: Full loop to iterate through all the sheets.
    For Each ws In Sheets
        ws.Activate
        
        ' Defines default values for variables that reset for each sheet.
        stockCount = 0
        volTotal = 0
        yearOpen = 0
    
        ' Counts number of columns and rows with data.
        columnCount = Cells(1, Columns.Count).End(xlToLeft).Column
        rowCount = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Sets column headers for analysis columns.
        Cells(1, columnCount + 2).Value = "Ticker"
        Cells(1, columnCount + 3).Value = "Yearly Change"
        Cells(1, columnCount + 4).Value = "Percent Change"
        Cells(1, columnCount + 5).Value = "Total Stock Volume"
        ' Challenge headers.
        Cells(1, columnCount + 8).Value = "Ticker"
        Cells(1, columnCount + 9).Value = "Value"
        Cells(2, columnCount + 7).Value = "Greatest % Increase"
        Cells(3, columnCount + 7).Value = "Greatest % Decrease"
        Cells(4, columnCount + 7).Value = "Greatest Total Volume"
        
        ' FOR LOOP: Iterates through stocks, row by row.
        For i = 2 To rowCount + 1
        
            ' IF STATEMENT: Checks if stock has changed when switching to a new row.
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                ' If the stock HAS changed:
                ' - increase the count of stocks and add the new stock to the list of stocks.
                ' - calculate the changePercentage
                ' - assign values to the columns for yearlyChange and changePercentage (format as %).
                ' - pull yearOpen data for new stock.
                ' - assigns volTotal to its column and begins the new volTotal.
                
                ' Increase stock count and add new stock to list.
                stockCount = stockCount + 1
                Cells(stockCount + 1, columnCount + 2).Value = Cells(i, 1).Value
                
                ' IF STATEMENT: For the first line of each sheet, don't try to pull the previous closing price because it doesn't exist.
                If i = 2 Then
                Else
                    yearClose = Cells(i - 1, 6).Value
                    yearlyChange = yearClose - yearOpen
                    changePercentage = (yearlyChange / yearOpen)
                    ' Simple cell assignments for yearlyChange and changePercentage.
                    Cells(stockCount, columnCount + 3).Value = yearlyChange
                    Cells(stockCount, columnCount + 4).Value = changePercentage
                    Cells(stockCount, columnCount + 4).NumberFormat = "0.00%"
                    
                    Cells(stockCount, columnCount + 5).Value = volTotal
                End If
                If Cells(i, 3).Value = 0 Then
                Else
                    ' Pulls open value for a new stock.
                    yearOpen = Cells(i, 3).Value
                End If
                
                ' Assigns volTotal to previous stock and sets it to the first value for the new stock.
                volTotal = Cells(i, 7).Value
                
            ' ELSE: When it's not a new stock, just increase the volTotal variable.
            Else
                volTotal = volTotal + Cells(i, 7).Value
            End If
        Next i
                
        ' FOR LOOP / IF STATEMENT: Does color coding for yearlyChange column.
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 0
        volTotalHigh = 0
        For i = 2 To stockCount + 1
            If Cells(i, columnCount + 3) > 0 Then
                Cells(i, columnCount + 3).Interior.ColorIndex = 4
            ElseIf Cells(i, columnCount + 3) < 0 Then
                Cells(i, columnCount + 3).Interior.ColorIndex = 3
            End If
            
            ' Challenges
            If Cells(i, columnCount + 4).Value > 0 Then
                If Cells(i, columnCount + 4).Value > greatestPercentageIncrease Then
                    greatestPercentageIncrease = Cells(i, columnCount + 4).Value
                    gpi = Cells(i, 1).Value
                End If
            ElseIf Cells(i, columnCount + 4).Value < 0 Then
                If Cells(i, columnCount + 4).Value < greatestPercentageDecrease Then
                    greatestPercentageDecrease = Cells(i, columnCount + 4).Value
                    gpd = Cells(i, 1).Value
                End If
            End If
            
            If Cells(i, columnCount + 5).Value > volTotalHigh Then
                volTotalHigh = Cells(i, columnCount + 5).Value
                vthStock = Cells(i, 1).Value
            End If
        Next i
        Cells(2, columnCount + 8).Value = gpi
        Cells(2, columnCount + 9).Value = greatestPercentageIncrease
        Cells(2, columnCount + 9).NumberFormat = "0.00%"
        Cells(3, columnCount + 8).Value = gpd
        Cells(3, columnCount + 9).Value = greatestPercentageDecrease
        Cells(3, columnCount + 9).NumberFormat = "0.00%"
        Cells(4, columnCount + 8).Value = vthStock
        Cells(4, columnCount + 9).Value = volTotalHigh
        
        Columns("I:P").AutoFit
    Next ws
End Sub
