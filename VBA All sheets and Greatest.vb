Sub StockAnalysis()
    ' Constants learned from the tutor.
    Const START_ROW As Integer = 2
    Const TICKER_COLUMN As Integer = 1
    Const OPEN_COLUMN As Integer = 3
    Const CLOSE_COLUMN As Integer = 6
    Const VOLUME_COLUMN As Integer = 7
    Const OUTPUT_START_COLUMN As Integer = 9
    
    ' Variables
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim OutputRows As Integer
    
    ' Variables for Greatest values
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    ' Find the last row
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Start of the row
    OutputRows = START_ROW
    
    ' Names to first row in columns I to L
    Cells(1, OUTPUT_START_COLUMN).Value = "Ticker"
    Cells(1, OUTPUT_START_COLUMN + 1).Value = "Year Change"
    Cells(1, OUTPUT_START_COLUMN + 2).Value = "Percent Change"
    Cells(1, OUTPUT_START_COLUMN + 3).Value = "Total Volume"
    
    ' Loop through each row of data
    For i = START_ROW To lastRow
        ' Checking for a new ticker that does not match the above ticker
        If Cells(i + 1, TICKER_COLUMN).Value <> Cells(i, TICKER_COLUMN).Value Then
            ' Ticker symbol
            ticker = Cells(i, TICKER_COLUMN).Value
            
            ' Closing price
            closingPrice = Cells(i, CLOSE_COLUMN).Value
            
            ' Total stock volume
            totalVolume = totalVolume + Cells(i, VOLUME_COLUMN).Value
            
            ' Yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculating the percentageChange
            percentChange = yearlyChange / openingPrice
            
            ' Output the data points
            Cells(OutputRows, OUTPUT_START_COLUMN).Value = ticker
            Cells(OutputRows, OUTPUT_START_COLUMN + 1).Value = yearlyChange
            Cells(OutputRows, OUTPUT_START_COLUMN + 2).Value = percentChange
            Cells(OutputRows, OUTPUT_START_COLUMN + 3).Value = totalVolume
            
            ' Update Greatest values
            If greatestIncreaseTicker = "" Or percentChange > greatestIncrease Then
                greatestIncreaseTicker = ticker
                greatestIncrease = percentChange
            End If
            
            If greatestDecreaseTicker = "" Or percentChange < greatestDecrease Then
                greatestDecreaseTicker = ticker
                greatestDecrease = percentChange
            End If
            
            If greatestVolumeTicker = "" Or totalVolume > greatestVolume Then
                greatestVolumeTicker = ticker
                greatestVolume = totalVolume
            End If
            
            ' Conditional formatting to the "Yearly Change" column
            ApplyConditionalFormatting Cells(OutputRows, OUTPUT_START_COLUMN + 1)
            
            ' Reset: Reset values for the next ticker
            openingPrice = 0
            totalVolume = 0
            OutputRows = OutputRows + 1
        Else
            ' Calculation: Accumulate total stock volume for the same ticker
            totalVolume = totalVolume + Cells(i, VOLUME_COLUMN).Value
            
            ' Calculation: Set the opening price for the current ticker
            If openingPrice = 0 Then
                openingPrice = Cells(i, OPEN_COLUMN).Value
            End If
        End If
    Next i
    
        ' Output Greatest values
        Cells(1, OUTPUT_START_COLUMN + 10).Value = "Greatest % Increase"
        Cells(1, OUTPUT_START_COLUMN + 11).Value = greatestIncreaseTicker
        Cells(1, OUTPUT_START_COLUMN + 12).Value = greatestIncrease
        
        Cells(1 + 1, OUTPUT_START_COLUMN + 10).Value = "Greatest % Decrease"
        Cells(1 + 1, OUTPUT_START_COLUMN + 11).Value = greatestDecreaseTicker
        Cells(1 + 1, OUTPUT_START_COLUMN + 12).Value = greatestDecrease
        
        Cells(1 + 2, OUTPUT_START_COLUMN + 10).Value = "Greatest Total Volume"
        Cells(1 + 2, OUTPUT_START_COLUMN + 11).Value = greatestVolumeTicker
        Cells(1 + 2, OUTPUT_START_COLUMN + 12).Value = greatestVolume

End Sub



Sub ApplyConditionalFormatting(rng As Range)
    ' Specified range
    Dim ws As Worksheet
    Set ws = rng.Worksheet
    
    ' Clear existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Define the formatting rules
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="1")
        .Interior.Color = RGB(255, 0, 0) ' Red
    End With
    
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="-1")
        .Interior.Color = RGB(255, 255, 0) ' Yellow
    End With
End Sub
