Attribute VB_Name = "Module1"
Sub Main()

    For sheet_index = 1 To ThisWorkbook.Sheets.Count
        Sheets(sheet_index).Activate
    
        With ActiveSheet
            Call Calculate
            Call Format_results
            Call Highlights
        End With
    
    Next sheet_index
End Sub

Sub Calculate()


'getting sheet dimensions'
    Dim row_count, column_count As Long
    row_count = Cells(Rows.Count, 1).End(xlUp).Row
    column_count = Cells(1, Columns.Count).End(xlToLeft).Column
  
    
'Placing Column headers for first results'
    Cells(1, 9).Value = "TickerXX"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

'declaring variables'
'new_Ticker is variable to store Ticker name on the row being analyzed'
'last_Ticker is variable to store Ticker name in thr row analyzed just before the current row'
'Unique_Ticker_index kees track of the row number wherein results for the stock will be printed after its all entries are analyzed'
'start_Price and end_Price variables holds the closing Price at begining of year and end of year respectively'
'stock_Volume is holds the Stock trading volume as we iterate through all the rows belonging to a stock. Initialized when we encounter first row for any stock'

    Dim new_Ticker, last_Ticker As String
    Dim Unique_Ticker_index, stock_Volume As Long
    Dim start_Price, end_Price As Double

'Initializing variables
    start_Price = 0
    end_Price = 0
    'Unique_Ticker_index initial value would be row number 2, wher we will store the results for first stock in the list'
    Unique_Ticker_index = 2
    last_Ticker = ""

'looping through the sheet'
    For I = 2 To row_count
        new_Ticker = Cells(I, 1).Value
        If new_Ticker = last_Ticker Then
                end_Price = Cells(I, 6).Value
                Volume = Volume + Cells(I, 7).Value
        Else
            If I = 2 Then
            'set variables for first iteration of the loop'
                Volume = Cells(2, 7).Value
                start_Price = Cells(2, 6).Value
                last_Ticker = new_Ticker
                
            Else
            'publish results for last stock when encounter row with a new ticker name'
                Cells(Unique_Ticker_index, 9).Value = Cells(I - 1, 1).Value
                Cells(Unique_Ticker_index, 10).Value = (Cells(I - 1, 6) - start_Price)
                    'check for zero division error before calculating Percentage change'
                        If start_Price <> 0 Then
                        Cells(Unique_Ticker_index, 11).Value = (Cells(Unique_Ticker_index, 10) / start_Price)
                        End If
                Cells(Unique_Ticker_index, 12).Value = Volume
                
                
            'reset variables for new Ticker'
                last_Ticker = new_Ticker
                start_Price = Cells(I, 6).Value
                Volume = Cells(I, 7).Value
                Unique_Ticker_index = Unique_Ticker_index + 1
            End If
        End If
    Next I
End Sub

Sub Format_results()
'formatting our results so far'

    'setting Autowidth for result columns'
        Columns("I:L").AutoFit
    'showing % sign in percentage change column'
        Columns("K").NumberFormat = "0.00%"
    
    'count number of Unique Ticker rows in our results for conditional formating loop'
        c = Cells(Rows.Count, 9).End(xlUp).Row
        'compare values under Percentage Change Column with a target of 0 (0%)'
        target = 0
        
    'loop for conditional formatting'
        For j = 2 To c
        Set testcell = Cells(j, 10)
        Select Case testcell
            Case Is > target
                With testcell
                    .Interior.Color = vbGreen
                End With
            Case Is < target
                With testcell
                    .Interior.Color = vbRed
                End With
            Case Is = target
                With testcell
                    .Interior.Color = vbWhite
                End With
           End Select
        Next j
End Sub

Sub Highlights()
'Print summary of stocks with greatest increase/decrease/trade volume'

    'Populate Column headers'
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"

    'set variables to store 3 summary results and their row indexes'
        Dim greatest_increase, greatest_decrease, greatest_volume As Double
        Dim greatest_increase_index, greatest_decrease_index, greatest_volume_index As Long

    'Initialize variables'
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0

    'Looping through the results to search for Tickers with greatest increse, greatest decrease and highest trade volume'
    
    'count number of Unique Ticker rows for this loop'
        c = Cells(Rows.Count, 9).End(xlUp).Row
        
    'Executing loop'
        For k = 2 To c
            If Cells(k, 11).Value > greatest_increase Then
                greatest_increase = Cells(k, 11).Value
                greatest_increase_index = k
            End If
            
            If Cells(k, 11).Value < greatest_decrease Then
                greatest_decrease = Cells(k, 11).Value
                greatest_decrease_index = k
            End If
            
            If Cells(k, 12).Value > greatest_volume Then
                greatest_volume = Cells(k, 12).Value
                greatest_volume_index = k
            End If
            
        Next k
        
    'Publishing Highlights'
        Range("P2").Value = Cells(greatest_increase_index, 9).Value
        Range("Q2").Value = Cells(greatest_increase_index, 11).Value
        
        Range("P3").Value = Cells(greatest_decrease_index, 9).Value
        Range("Q3").Value = Cells(greatest_decrease_index, 11).Value
        
        Range("P4").Value = Cells(greatest_volume_index, 9).Value
        Range("Q4").Value = Cells(greatest_volume_index, 12).Value
    
    'Format Highlights'
        Range("Q2:Q3").NumberFormat = "0.00%"
        Columns("O:Q").AutoFit
End Sub

