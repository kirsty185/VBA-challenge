Attribute VB_Name = "Module1"
Sub ChallengeVBA():

'Declare variables
Dim Ticker As String
Dim year_open As Double
Dim year_close As Double
Dim Yearly_Change As Double
Dim Total_Stock_Volume As Double
Dim Percent_Change As Double

'Declare worksheet and variable locations
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'loop to start
    start_data = 2
    previous_i = 1
    Total_Stock_Volume = 0
    
'start
'Dim start_data As Integer


    
    'loop end is last row  of column A - for loop
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'For each Ticker work out the yearly change, percent change and total volume
        For i = 2 To LastRow
            
            'Tickersymbol must match
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'locate Ticker
            Ticker = ws.Cells(i, 1).Value
            
            'Move to the next Ticker
            previous_i = previous_i + 1
            
            'Work out the open price and the end price
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value
            
            'Sum total volume using a loop
            For j = previous_i To i
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
            Next j
            'If 0
            If year_open = 0 Then
                Percent_Change = year_close
            Else
                Yearly_Change = year_close - year_open
                Percent_Change = Yearly_Change / year_open
                
            End If
                  
            'Ticker, year change and percent change table
            ws.Cells(start_data, 9).Value = Ticker
            ws.Cells(start_data, 10).Value = Yearly_Change
            ws.Cells(start_data, 11).Value = Percent_Change
            
            'Format column K/11 and L/12
            ws.Cells(start_data, 11).NumberFormat = "0.00%"
            ws.Cells(start_data, 12).Value = Total_Stock_Volume
            
                      
            start_data = start_data + 1
            
            'Reset
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
                       
            previous_i = i
        
        End If
    Next i
    
    
    'last row of column k
    kEndRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    'Define start
    Increase = 0
    Decrease = 0
    Greatest = 0
    
        'find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow
        
            'Define previous increment to check
            last_k = k - 1
                        
            'Define current row for percentage
            current_k = ws.Cells(k, 11).Value
            
            'Define Previous row for percentage
            prevous_k = ws.Cells(last_k, 11).Value
            
            'greatest total volume row
            volume = ws.Cells(k, 12).Value
            
            'Prevous greatest volume row
            prevous_vol = ws.Cells(last_k, 12).Value
            
   
            
            'Find the increase
            If Increase > current_k And Increase > prevous_k Then
                
                Increase = Increase
                'define name for increase percentage
                'increase_name = ws.Cells(k, 9).Value
                
            ElseIf current_k > Increase And current_k > prevous_k Then
                Increase = current_k
                'define name for increase percentage
                increase_name = ws.Cells(k, 9).Value
                
            ElseIf prevous_k > Increase And prevous_k > current_k Then
                Increase = prevous_k
                'define name for increase percentage
                increase_name = ws.Cells(last_k, 9).Value
            End If
      
      
            'Find the decrease
            If Decrease < current_k And Decrease < prevous_k Then
                
                'Define decrease as decrease
                Decrease = Decrease
                
                'Define name for increase percentage
            ElseIf current_k < Increase And current_k < prevous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
            ElseIf prevous_k < Increase And prevous_k < current_k Then
                Decrease = prevous_k
                decrease_name = ws.Cells(last_k, 9).Value
            End If
            
     
           'Find the greatest volume
            If Greatest > volume And Greatest > prevous_vol Then
                Greatest = Greatest
                'define name for greatest volume
                'greatest_name = ws.Cells(k, 9).Value
            
            ElseIf volume > Greatest And volume > prevous_vol Then
            
                Greatest = volume
                
                'define name for greatest volume
                greatest_name = ws.Cells(k, 9).Value
                
            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                
                Greatest = prevous_vol
                
                'define name for greatest volume
                greatest_name = ws.Cells(last_k, 9).Value
            End If
        Next k

    'Table column names
    ws.Range("O1").Value = "Column Name"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker Name"
    ws.Range("Q1").Value = "Value"
    
    'not sure about getting values but they will need to be formatted as a %
    
    


'Format colours for column J
    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

        For j = 2 To jEndRow
            'if greater than or less than zero
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    
'Excute to next worksheet
Next ws
End Sub

