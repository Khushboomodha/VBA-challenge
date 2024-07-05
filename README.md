Purpose:
The VBA script "Stock_Analysis" is designed to analyze stock data across multiple worksheets in an Excel workbook. It calculates quarterly metrics for each stock, including the quarterly price change, percent change, and total stock volume. Additionally, it identifies and highlights certain metrics such as greatest percentage increase, greatest percentage decrease, and greatest total volume. The results are summarized in a structured table format on each worksheet.
Instructions:
1.	Ensure Data Structure:
o	The script assumes that each worksheet contains stock data in columns with the following structure:
	Column A: Ticker Symbols
	Column C: Opening Price
	Column F: Closing Price
	Column G: Volume
2.	Implementation:
o	Open Excel and the workbook containing the stock data.
o	Press Alt + F11 to open the VBA Editor.
o	Insert a new module (if not already present) and paste the provided VBA script into the module.
3.	Execution:
o	To execute the script, run the macro Stock_Analysis:
	This macro iterates through each worksheet, calculates the necessary metrics, and populates summary tables and formatting.
4.	Additional Macros:
o	Color Formatting (colorformat):
	This macro formats the percentage change (Column L) in the summary table based on positive (green) or negative (red) values.
	Run the macro colorformat after running Stock_Analysis.
o	Summary of Top Metrics (tickersummary):
	This macro identifies and displays the top metrics (greatest % increase, greatest % decrease, greatest total volume) in the specified cells.
	Run the macro tickersummary after running Stock_Analysis.
o	Ticker Symbols for Top Metrics (Tickersymbols):
	This macro identifies the ticker symbols corresponding to the top metrics and displays them in the specified cells.
	Run the macro Tickersymbols after running tickersummary.
5.	Output:
o	After running the macros sequentially (Stock_Analysis, colorformat, tickersummary, Tickersymbols), each worksheet will have:
	A summary table with ticker symbols, quarterly changes, percent changes, and total stock volume.
	Color formatting applied to highlight positive and negative percentage changes.
	Cells populated with the top metrics (percent increase, percent decrease, total volume) and their corresponding ticker symbols.
6.	Adjustments:
o	Modify ranges or formatting instructions in the VBA code (colorformat, tickersummary, Tickersymbols) if column locations or output requirements change in your dataset.

VBA Coding

Sub Stock_Analysis()
    Dim ws As Worksheet
    Dim Tickersymbols As String
    Dim TotalStockvolume As Double
    Dim Summarytablerow As Long
    Dim Quarter_open As Double
    Dim Quarter_close As Double
    Dim Quarterly_pricechange As Double
    Dim Percent_change As Double
    Dim LastRow As Long
    Dim start As Long

    For Each ws In Worksheets
        TotalStockvolume = 0
        Summarytablerow = 2
        Quarterly_pricechange = 0
        Quarter_open = 0
        Quarter_close = 0

        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        start = 2
        
        ' Assigning titles to summarytable (move outside the loop)
        ws.Range("J1").Value = "Tickersymbols"
        ws.Range("K1").Value = "Quarterly_Change"
        ws.Range("L1").Value = "Percent_Change"
        ws.Range("M1").Value = "TotalStockvolume"
        ws.Range("P1").Value = "Tickersymbols"
        ws.Range("Q1").Value = "TotalStockvolume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Tickersymbols = ws.Cells(i, 1).Value
                Quarter_close = ws.Cells(i, 6).Value
                Quarter_open = ws.Cells(start, 3).Value
                start = i + 1

                TotalStockvolume = TotalStockvolume + ws.Cells(i, 7).Value
                Quarterly_pricechange = Quarter_close - Quarter_open
                Percent_change = Quarterly_pricechange / Quarter_open
                ws.Range("L" & Summarytablerow).Value = Percent_change
                ws.Range("L" & Summarytablerow).NumberFormat = "0.00%"

                ws.Range("J" & Summarytablerow).Value = Tickersymbols
                ws.Range("K" & Summarytablerow).Value = Quarterly_pricechange
                ws.Range("M" & Summarytablerow).Value = TotalStockvolume

                Summarytablerow = Summarytablerow + 1
                TotalStockvolume = 0
            Else
                TotalStockvolume = TotalStockvolume + ws.Cells(i, 7).Value
            End If
        Next i

        ws.Columns("A:R").AutoFit
    Next ws
End Sub

Sub colorformat()

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

If ws.Cells(i, 11).Value > 0 Then
ws.Cells(i, 11).Interior.ColorIndex = 4

ElseIf ws.Cells(i, 11).Value < 0 Then
ws.Cells(i, 11).Interior.ColorIndex = 3

End If

Next i

Next ws
End Sub





Sub tickersummary()


For Each ws In Worksheets

ws.Range("Q2") = WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3") = WorksheetFunction.Min(ws.Range("L:L"))
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("Q4") = WorksheetFunction.Max(ws.Range("M:M"))

ws.Columns("Q:Q").AutoFit

Next ws

End Sub

Sub Tickersymbols()

For Each ws In Worksheets


Dim Ticker1 As String
Dim Ticker2 As String
Dim Ticker3 As String
Dim Summarytablerow As Integer

Summarytablerow = 2

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow


If ws.Cells(i, 12).Value = ws.Range("Q2").Value Then

Ticker1 = ws.Cells(i, 10).Value

ws.Range("P" & Summarytablerow).Value = Ticker1

Summarytablerow = Summarytablerow + 1



End If

If ws.Cells(i, 12).Value = ws.Range("Q3").Value Then

Ticker2 = ws.Cells(i, 10).Value

ws.Range("P3").Value = Ticker2




End If

If ws.Cells(i, 13).Value = ws.Range("Q4").Value Then

Ticker3 = ws.Cells(i, 10).Value

ws.Range("P4").Value = Ticker3


End If

Next i

Next ws


End Sub
