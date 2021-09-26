Sub stock()

'loop through all worksheets
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    ws.Cells(1, 1) = 1
        
    'define variables: Lastrow, Ticker, Open, Close
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly As Double
    Dim Percent As Double
    Dim Stock_Volume As Double
    Stock_Volume = 0
    'Declare Max variables
    Dim Max_Volume As Double
    Dim Max_Increase As Double
    Dim Max_Decrease As Double
                
    'establish table for variables to go
    Dim Summary_Table As Integer
    Summary_Table = 2
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    'Cells(2, 15).Value = "Greatest % Increase"
    'Cells(3, 15).Value = "Greatest % Decrease"
    'Cells(4, 15).Value = "Greatest Total Volume"
    'Cells(1, 16).Value = "Ticker"
    'Cells(1, 17).Value = "Value"
    
         'get the variables; begin loop
         For I = 2 To lastrow
            'Get the ticker and stock volume
             If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
             Ticker = (Cells(I, 1).Value)
             Stock_Volume = Stock_Volume + Cells(I, 7).Value
             
                      
             'Condition to define Op and C and calculate Yearly
                    'Find Close_Price
                    If Cells(I + 1, 1).Value <> Ticker And Cells(I, 1).Value = Ticker Then
                    Close_Price = Cells(I, 6).Value
                    End If
                    'Check Close_Price
                    'Range("U" & Summary_Table).Value = Close_Price
                    
                    'Find Open_Price with Loop
                    For j = 2 To lastrow
                    If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
                    Open_Price = Cells(j, 3).Value
                    End If
                    If Open_Price = 0 Then
                    Open_Price = Cells(j+1,3).Value
                    End If 
                    Next j
                    'Check Open_Price
                    'Range("T" & Summary_Table).Value = Open_Price
                
                    
                    'Calculate Yearly and Percent
                    Yearly = Close_Price - Open_Price
                    Percent = (Yearly / Open_Price)
               
 'conditional formatting that will highlight positive change in green and negative change in red
                                                      
            'Put variables above in Summary table
             Range("J" & Summary_Table).Value = Ticker
             Range("K" & Summary_Table).Value = Yearly
             Range("L" & Summary_Table).Value = Percent
             Range("M" & Summary_Table).Value = Stock_Volume
             'Row Count for K
             RowCount = Cells(Rows.Count, 11).End(xlUp).Row
             'Loop for Coloring Change
             For K = 2 To RowCount
                If Cells(K, 11).Value < 0 Then
                    Cells(K, 11).Interior.ColorIndex = 3
                    Else
                    Cells(K, 11).Interior.ColorIndex = 4
                End If
             Next K
             'Update formatting for Percent
             Range("L:L").NumberFormat = "0.00%"
             
             Summary_Table = Summary_Table + 1
             Stock_Volume = 0
             
             Else
             ' summing volume
             Stock_Volume = Stock_Volume + Cells(I, 7).Value
              
             End If
   
            
         Next I

'Find Max Voume
    'Dim Ticker_Final As String
      
    'For J = 2 To lastrow
    'Ticker_Final = Cells(J, 10).Value
    'Max_Volume = Cells(J, 13).Value
    
             'If Cells(J, 13).Value > Cells(J + 1, 13).Value And Cells(J, 13) = Max_Volume Then
             'Cells(4, 17).Value = Max_Volume And Cells(4, 16).Value = Ticker_Final
             'Else
             'Max_Volume = Cells(J, 13).Value
             'End If
        
    'Next J
    
 'BONUS return greatest % increase, greatest % decrease, and greatest total volume
             
             
             'Max_Increase = WorksheetFunction.Max(Range("L:L").Value)
             'If Cells(I,12).Value=Max_Increase Then
             'Cells(2,17).Value=Max_Increase AND Cells(2,16).Value=Ticker
             'End If
             'Max_Decrease=WorksheetFunction.Min(Range("L:L").Value)
             'If Cells(I,12).Value=Max_Decrease Then
             'Cells(3,17).Value=Max_Decrease AND Cells(3,16).Value=Ticker
             'End If
             

Next

End Sub
