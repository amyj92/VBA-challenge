Attribute VB_Name = "Module1"
' ------------------------------------------------------------------
' LOOP THROUGH ALL WORKSHEETS
' ------------------------------------------------------------------
    
Sub WSLoop()

    ' Define variable for the worksheets
    Dim ws As Worksheet
    
    ' Loop through all worksheets data
    For Each ws In ThisWorkbook.Worksheets
    
        Call StockData(ws)
    
    Next ws
    
End Sub



' ------------------------------------------------------------------
' ANALYZE STOCK DATA
' ------------------------------------------------------------------

Sub StockData(ws As Worksheet)


    ' ------------------------------------------------------------------
    ' DEFINE VARIABLES
    ' ------------------------------------------------------------------

    ' Define for loop variables
    Dim a As Long 'used in for loop for ticker summary table
    Dim b As Long 'used in for loop for greatest values table
    Dim c As Long 'used for looping through for formatting quarterly change color
 
    ' Define OpeningPrice variable and set the value of the first row
    Dim OpeningPrice As Long
        OpeningPrice = 2

    ' Define ticker variable
    Dim ticker As String
    
    ' Define the row variable for the summary table to list ticker information and set the value of the first row
    Dim SummaryTableRow As Long
        SummaryTableRow = 2
    
    ' Define quarterly change variable
    Dim QuarterlyChange As Double
    
    ' Define percent change variable
    Dim PercentChange As Double
    
    ' Define a variable to hold the total stock volume for each ticker and set the value
    Dim VolumeTotal As Double
        VolumeTotal = 0
        
    ' Define last row variable and set the value
    Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Add column headers for ticker summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Add labels to greatest % and volume table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Autofit columns and rows
    ws.Columns("A:Q").AutoFit
    

    
    ' ------------------------------------------------------------------
    ' LOOP THROUGH TICKERS TO CREATE TICKER SUMMARY TABLE
    ' ------------------------------------------------------------------
    
    ' Loop through all rows to list tickers
    For a = 2 To LastRow
    
        ' Check to see if we are still within the same ticker (search for when the next cell value is different)
        If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
        
            ' Set the ticker value
            ticker = ws.Cells(a, 1).Value
            
            ' Create list of tickers
            ws.Range("I" & SummaryTableRow).Value = ticker
            
            ' Calculate the quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter and list the value for each ticker
            QuarterlyChange = ws.Cells(a, 6).Value - ws.Cells(OpeningPrice, 3).Value
            ws.Range("J" & SummaryTableRow).Value = QuarterlyChange
            
            ' Calculate the percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter and list the value for each ticker in % format
            PercentChange = ws.Cells(a, 6).Value / ws.Cells(OpeningPrice, 3).Value - 1
            ws.Range("K" & SummaryTableRow).Value = PercentChange
            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
            
            ' Calculate the total stock volume of the stock and list the value for each ticker
            VolumeTotal = VolumeTotal + ws.Cells(a, 7).Value
            ws.Range("L" & SummaryTableRow).Value = VolumeTotal
    
            ' Move to the next ticker's opening price
            OpeningPrice = a + 1
            
            ' Move to the next row for the next ticker's determine value
            SummaryTableRow = SummaryTableRow + 1
            
            ' Reset the volume total
            VolumeTotal = 0
         
        Else
    
            VolumeTotal = VolumeTotal + ws.Cells(a, 7).Value
    
        End If
    
    Next a
        


' -----------------------------------------------------------------------------------------------------
    
    ' ------------------------------------------------------------------
    ' LOOP THROUGH TICKER SUMMARY TABLE TO DETERMINE GREATEST VALUES
    ' ------------------------------------------------------------------
    
    ' Loop through all percent changes to find max and min and to find the max total volume
    Dim LastRow2 As Long
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Define greatest % increase variable
    Dim GreatestIncrease As Double
        GreatestIncrease = WorksheetFunction.Max(ws.Range("K:K").Value)
    
    ' Define greatest % decrease variable
    Dim GreatestDecrease As Double
        GreatestDecrease = WorksheetFunction.Min(ws.Range("K:K").Value)
    
    ' Define greatest total volume variable
    Dim GreatestVolumeTotal As Double
        GreatestVolumeTotal = WorksheetFunction.Max(ws.Range("L:L").Value)

    For b = 2 To LastRow2
    
        If ws.Cells(b, 11).Value = GreatestIncrease Then
        
            ws.Range("P2").Value = ws.Cells(b, 9)
            ws.Range("Q2").Value = GreatestIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            
        ElseIf ws.Cells(b, 11).Value = GreatestDecrease Then
        
            ws.Range("P3").Value = ws.Cells(b, 9)
            ws.Range("Q3").Value = GreatestDecrease
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ElseIf ws.Cells(b, 12).Value = GreatestVolumeTotal Then
        
            ws.Range("P4").Value = ws.Cells(b, 9)
            ws.Range("Q4").Value = GreatestVolumeTotal
            
        End If
    
    Next b
    
    
   
    ' ------------------------------------------------------------------
    ' LOOP THROUGH QUARTERLY CHANGE VALUES TO FORMAT COLOR
    ' ------------------------------------------------------------------
    
    For c = 2 To LastRow2
    
        If ws.Cells(c, 10).Value > 0 Then
            
            ws.Cells(c, 10).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(c, 10).Value < 0 Then
    
            ws.Cells(c, 10).Interior.ColorIndex = 3
            
        End If
    
    Next c



End Sub


