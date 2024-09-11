Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()

    'Declaring a variable to hold the counter
    Dim i As Long
    Dim ws As Worksheet
    
    'Loop through all sheets
    For Each ws In Worksheets
    
        'Declaring an initial variable to hold the ticker name
        Dim ticker_name As String
    
        'Declaring an initail variable to hold the opening and closing price
        Dim opening_price As Double
        Dim closing_price As Double
        Dim quarterly_change As Double
        Dim percentage_change As Double
        Dim total_stock_volume As Double
     
        opening_price = 0
        closing_price = 0
        quarterly_change = 0
        percentage_change = 0
        total_stock_volume = 0
        
        
        'Declaring variables to track the greatest percentage increase, decrease, and total volume
        Dim max_percentage As Double
        Dim min_percentage As Double
        Dim max_volume As Double
        
        Dim max_percentage_ticker As String
        Dim min_percentage_ticker As String
        Dim max_volume_ticker As String
        
        
        max_percentage = 0
        min_percentage = 0
        max_volume = 0
    
        'Keeping track of the location of each tricker in a summary table
        Dim summary_table As Integer
        summary_table = 2
    
        'Declaring a variable that counts the number of rows
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Create headers for the output data
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Optimize code: Disable screen updating and calculations during the process
        'Googled how to make the loading data faster because my excel was timing out
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    
    
        'Loop through the data
        For i = 2 To LastRow
    
            'Check if we are still within the same tricker name, if we are not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the tricker name
                ticker_name = ws.Cells(i, 1).Value
            
                'setting the closing price
                closing_price = ws.Cells(i, 6).Value
            
                'calculating the quarterly change
                quarterly_change = closing_price - opening_price
                
                'calculating the percentage change
                percentage_change = quarterly_change / opening_price
            
                'calculating the total stock volume
                total_stock_volume = total_stock_volume + Cells(i, 7).Value
            
                'Print the tricker name in the Summary Table
                ws.Range("I" & summary_table).Value = ticker_name
            
                'Print the quaterly change calculation in the Summary Table
                ws.Range("J" & summary_table).Value = quarterly_change
            
                'Checking if quarterly change is greater than 0
                If quarterly_change > 0 Then
            
                    'Setting the Cell Color to Green
                    ws.Range("J" & summary_table).Interior.Color = vbGreen
                
                'Checking if quarterly change is less than 0
                ElseIf quarterly_change < 0 Then
            
                    'Setting the Cell Color to Red
                    ws.Range("J" & summary_table).Interior.Color = vbRed
                
                End If
            
            
                'Print the percentage change in the Summary Table
                ws.Range("K" & summary_table).Value = percentage_change
            
                'Displaying the result as a percentage with two decimal places
                ws.Range("K" & summary_table).NumberFormat = "0.00%"
            
                'Print the total stock volume calculation in the Summary Table
                ws.Range("L" & summary_table).Value = total_stock_volume
                
                
                'Check if we have a new greatest percentage increase
                If percentage_change > max_percentage Then
                    
                    max_percentage = percentage_change
                    max_percentage_ticker = ticker_name
                    
                End If
                
                
                'Check if we have a new greatest percentage decrease
                If percentage_change < min_percentage Then
                    
                    min_percentage = percentage_change
                    min_percentage_ticker = ticker_name
                    
                End If
                
                
                'Check if we have a new greatest total volume
                If total_stock_volume > max_volume Then
                    
                    max_volume = total_stock_volume
                    max_volume_ticker = ticker_name
                    
                End If
                              
            
                'Adding one to the summary table row
                summary_table = summary_table + 1
            
                'Reseting all variables that are used for calculation
                closing_price = 0
                opening_price = 0
                quarterly_change = 0
                total_stock_volume = 0
            
            'Gets the first opening price of each tricker name
            ElseIf opening_price = 0 Then
        
                'setting the opening price
                opening_price = ws.Cells(i, 3).Value
            
            'If the cell immediately following a row is the same tricker name...
            Else
        
            'Adding to the volume rows together
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            End If

        Next i
        
        'Displaying the result as a greatest percentage increase
        ws.Range("P2").Value = max_percentage_ticker
        ws.Range("Q2").Value = max_percentage
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Displaying the result as a greatest percentage decrease
        ws.Range("P3").Value = min_percentage_ticker
        ws.Range("Q3").Value = min_percentage
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Displaying the result as a greatest total volume
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
        
    Next ws
    
    'Re-enable screen updating and calculations
    'Googled how to make the loading data faster because my excel was timing out
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub resetTable()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        ws.Range("I1:L1501").Value = ""
        ws.Range("J1:J1501").Interior.ColorIndex = 0
        ws.Range("O1:Q4").Value = ""
        
    Next ws

End Sub

