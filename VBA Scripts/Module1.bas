Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim tickerDict As Object
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
            
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Create a dictionary to store unique ticker values and their corresponding data
        Set tickerDict = CreateObject("Scripting.Dictionary")
                    
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
        ' Start outputting data from row 2
        outputRow = 2
            
        ' Add column headers to the output sheet
        ws.Cells(1, 10).Value = "Ticker Symbol"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
            
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Get the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' Check if the ticker symbol is already processed
            If Not tickerDict.Exists(ticker) Then
                ' Find the range of rows for the current ticker
                Dim startRow As Long
                Dim endRow As Long
                startRow = i
            
                ' Find the last row for the current ticker
                Do While ws.Cells(i + 1, 1).Value = ticker
                    i = i + 1
                Loop
                endRow = i
            
                ' Get the opening price at the beginning of the quarter (first row of the range)
                openingPrice = ws.Cells(startRow, 3).Value ' Assuming the opening price is in column C
            
                ' Get the closing price at the end of the quarter (last row of the range)
                closingPrice = ws.Cells(endRow, 6).Value ' Assuming the closing price is in column F
            
                ' Calculate the quarterly change
                quarterlyChange = closingPrice - openingPrice
            
                ' Calculate the percentage change
                If openingPrice <> 0 Then
                    percentageChange = quarterlyChange / openingPrice
                Else
                    percentageChange = 0
                End If
            
                ' Get the total stock volume
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7))) ' Assuming the total stock volume is in column G
            
                ' Add the ticker symbol and its corresponding data to the dictionary
                tickerDict.Add ticker, Array(quarterlyChange, percentageChange, totalVolume)
            
                ' Output the information to a new range of cells
                ws.Cells(outputRow, 10).Value = ticker ' Ticker Symbol
                ws.Cells(outputRow, 11).Value = quarterlyChange ' Quarterly Change
                ws.Cells(outputRow, 12).Value = percentageChange ' Percentage Change
                ws.Cells(outputRow, 13).Value = totalVolume ' Total Stock Volume
            
                ' Apply conditional formatting to the Quarterly Change column (Column K)
                With ws.Range("K" & outputRow)
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
                    .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End With
            
                ' Apply conditional formatting to the Percentage Change column (Column L)
                With ws.Range("L" & outputRow)
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
                    .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
                    .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                End With
            
                ' Move to the next row for output
                outputRow = outputRow + 1
            End If
        Next i
    Next ws
End Sub
