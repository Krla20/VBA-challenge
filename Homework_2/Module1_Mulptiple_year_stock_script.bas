Attribute VB_Name = "Module1"
Sub stockmarket1()

'Loop through all the stocks for all years (Run all workbook at once).
'Add the single ticker symbol.
'Yearly change from openning price at the begining to closing price at year end -> LAST ROW of A since it's the last day of the year F-C.
'Percentage change from opening price at beg. year to closing price at end year F-C/C.
'Total stock volum of the stock. SUM TOTAL AMOUNT per TICKER SYMBOL.
'Conditional formating positive change in green and negative change in red.
'Stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
'Submission in Github: a screenshot for each year on the Multi Year Stock Data, VBA scripts, READ.ME file, and, submit link to BCS spot.
   
    
       
'Set an initial variable for holding the Ticker symbol
Dim ticker As String
        
'Set an initial variable to keep track of number of tickers for each worksheet
Dim number_tickers As Integer

'Set an initial variable to keep track of the last row in each worksheet
Dim lastRowState As Long
    
'Set an initial variable to keep track of the opening price for specific year
Dim opening_price As Double
        
'Set an initial variable to keep track of closing price for specific year
Dim closing_price As Double
    
'Set an initial variable to keep track of yearly change
Dim yearly_change As Double
        
'Set an initial variable to keep track of percentage change
Dim percent_change As Double
        
'Set an initial variable to keep track of total sock volume
Dim total_stock_volume As Double
        
'Set an initial variable to keep track of the greatest percentage increase value for specific year
Dim greatest_percent_increase As Double
    
'Set an initial variable to keep track of the ticker that has the greatest percent increase
Dim greatest_percent_increase_ticker As String
        
'Set an initial variable to keep track of the greatest percent decrease value for specific year
Dim greatest_percent_decrease As Double
    
'Set an initial variable to keep track of the greastest percent decrease percent for specific year
Dim greatest_percent_decrease_ticker As String
    
'Set an initial variable to keep track of the greatest stock volume value for specific year
Dim greatest_stock_volume As Double
        
'Set an initial variable to keep track of the ticker that has the greatest stock volume
Dim greatest_stock_volume_ticker As String
    
'Loop through every worksheet in the workbook
 For Each ws In Worksheets
    
    'Make the worksheet active
    ws.Activate
  
    ' Determine the Last Row
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set up columns and rows width
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.EntireRow.AutoFit
    
    
    'Add the column names for part 1
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
       
    'Start set up variables for each worksheet
        number_tickers = 0
        ticker = ""
        yearly_change = 0
        opening_price = 0
        percent_change = 0
        total_stock_volume = 0
    
'Skip the header row, for loop the list of tickers
For I = 2 To lastRowState
        
        'Set the Ticker symbol
        ticker = Cells(I, 1).Value
        
        'Set up the start of the year opening price for the ticker
        If opening_price = 0 Then
            opening_price = Cells(I, 3).Value
        End If
        
        'Sum the total stock volume values for a ticker
        total_stock_volume = total_stock_volume + Cells(I, 7).Value
        
            'If it reaches a different ticker name in the list then do this
            If Cells(I + 1, 1).Value <> ticker Then
                'Increment the number of tickers when it reaches a different ticker name in the list
                number_tickers = number_tickers + 1
                Cells(number_tickers + 1, 9) = ticker
                
                'Get to the end of the year closing price for ticker
                closing_price = Cells(I, 6)
                
                'Get the yearly change value
                yearly_change = closing_price - opening_price
                
                'Add yearly change value to the respective ticker/cell in each woorksheet
                Cells(number_tickers + 1, 10).Value = yearly_change
                
                'Color change, if the yearly change is higher than 0 then color that cell green (positive)
                If yearly_change > 0 Then
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
                
                'Color change, if the yearly change is less than 0 then color that cell red (negative)
                ElseIf yearly_change < 0 Then
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            End If
                
            
            'Calculate percent change for the value of the tickers
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
        'Format the percent_change value as percent
         Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
           
        'Set opening price back to 0 when there is a different ticker name in the list
         opening_price = 0
            
        'Add total stock volume value to the respective cell in each worksheet
        Cells(number_tickers + 1, 12).Value = total_stock_volume
        
        'Set total stock volume back to 0 when there is a different ticker name on name on the list
        total_stock_volume = 0
    End If
     
Next I

'Add the column names for part 2, rows "Greatest % increase", "Greatest % decrease" and "Greatest total volume" and columns "Ticker", "Value" for each year
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
            
    'Set the last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row

'Initialize variables and set up values to the first row in the list
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value

'Skip the header row, loop through the list of tickers
For I = 2 To lastRowState

    'Find the ticker with the greatest percent increase
    If Cells(I, 11).Value > greatest_percent_increase Then
        greatest_percent_increase = Cells(I, 11).Value
        greatest_percent_increase_ticker = Cells(I, 9).Value
    End If

    'Find the ticker with the greatest percent decrease
    If Cells(I, 11).Value < greatest_percent_decrease Then
        greatest_percent_decrease = Cells(I, 11).Value
        greastest_percent_decrease_ticker = Cells(I, 9).Value
    End If
    
    'Find the ticker with the greatest stock volume
    If Cells(I, 12).Value > greatest_stock_volume Then
        greatest_stock_volume = Cells(I, 12).Value
        greatest_stock_volume_ticker = Cells(I, 9).Value
    End If
    
Next I
    
  'Add the values for the greatest percent increase, greatest percent decrease, and stock volume to each worksheet in the workbook
  Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
  Range("Q2").Value = Format(greatest_percent_increase, "Percent")
  Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
  Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
  Range("P4").Value = greatest_stock_volume_ticker
  Range("Q4").Value = greatest_stock_volume

    'Set up columns and rows width
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    ActiveSheet.UsedRange.EntireRow.AutoFit
    
Next ws
   
' MSG BOX MACRO COMPLETE
    MsgBox ("Data Analysis Complete")


End Sub






