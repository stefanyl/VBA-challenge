'Script must loop through one year
'Must read/store variables for ticker symbol, volume of stock,
'open price, and open price.

Sub stock_script()
    'Define initial variables
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim year_change As Double 
    Dim stock_vol As Double
    Dim perct_change As Double
    Dim summary_row As Integer
 

    'Declare worksheet 
    Dim ws As Worksheet
    'ws.Activate
    
    'Loop through worksheet
    For Each ws In Worksheets
        'Column headers using ranges
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'Integers
        summary_row = 2
        previous_i = 1
        stock_vol = 0

        'Start with last row of Column A
        EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            'Another for loop, 
            For i = 2 To EndRow
                If ws.Cells(i+1, 1).Value <> ws.Cells(i, 1).Value Then

                ticker = ws.Cells(i, 1).Value
                previous_i = previous_i + 1

                'Get the first value from day open and last day close from nect column
                year_open = ws.Cells(previous_i, 3).Value
                year_close = ws.Cells(i, 6).Value

                'Another for loop (volume)
                For j = previous_i To i
                    stock_vol = stock_vol + ws.Cells(j, 7).Value

                Next j 
                
                'If loop gets zero 
                If year_open = 0 Then
                    perct_change = year_close

                Else
                    year_change = year_close - year_open'
                    perct_change = year_change / year_open

                End If

                'Populate the summary section
                ws.Cells(summary_row, 9).Value = ticker
                ws.Cells(summary_row, 10).Value = year_change
                ws.Cells(summary_row, 11).Value = perct_change
                ws.Cells(start_data, 11).NumberFormat = "#.##%"
                ws.Cells(start_data, 12).Value = stock_vol

                'Got to the next row and rest variables
                summary_row = summary_row + 1
                stock_vol = 0
                year_change = 0
                perct_change = 0
                previous_i = i 
            End If
        Next i 

        'conditional formatting (ColorIndex)
        jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            For j = 2 To jEndRow
                If ws.Cells(j, 10) > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                Else 
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            Next j
        
    Next ws  

End Sub