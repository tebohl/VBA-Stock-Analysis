Attribute VB_Name = "Module1"
Sub Stocks()

'Run sub routine for each worksheet
Dim ws As Worksheet
For Each ws In Worksheets

'Set column and row labels
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Set variable to hold ticker name
Dim Ticker As String

'Set variables for yearly change for each ticker name
Dim Year_Open As Double
Dim Year_Close As Double
Dim Yearly_Change As Double

'Set variable to hold percent change
Dim Percent_Change As Double

'Set variable to hold total stock volume starting at zero
Dim Total_Volume As Double
Total_Volume = 0

'Keep track of location for table
Dim Stock_Table_Row As Double
Stock_Table_Row = 2

'Count rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'Loop through rows
 For i = 2 To lastrow
 
    'Find year open price at beginning of loop
    Dim Stock_Price As Boolean
    
    If Stock_Price = False Then
    Year_Open = ws.Cells(i, 3).Value
    
    'Set condition to find year close price when ticker name changes
    Stock_Price = True
    
    End If
 
        'Check to see when Ticker name changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set ticker name
            Ticker = ws.Cells(i, 1).Value
    
            'Print ticker name in table
            ws.Range("I" & Stock_Table_Row).Value = Ticker
        
             'Add to total stock volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
            'Print total volume amount in table
            ws.Range("L" & Stock_Table_Row).Value = Total_Volume
     
            'Reset total stock volume
            Total_Volume = 0
     
            'Find year close price and set condition to find year open on next loop
            Year_Close = ws.Cells(i, 6).Value
            Stock_Price = False
    
            'Calculate yearly change
            Yearly_Change = Year_Close - Year_Open
        
            'Print yearly change in table
            ws.Range("J" & Stock_Table_Row).Value = Yearly_Change
            
                'Set colors for yearly change
                'cells with negative values red
                If ws.Range("J" & Stock_Table_Row).Value < 0 Then
                    ws.Range("J" & Stock_Table_Row).Interior.ColorIndex = 3
            
                'Set positive values green
                ElseIf ws.Range("J" & Stock_Table_Row).Value >= 0 Then
                    ws.Range("J" & Stock_Table_Row).Interior.ColorIndex = 4
   
                End If
     
            'Calculate percent change
            Percent_Change = (Yearly_Change / Year_Open)
    
            'Print percent change in table
            ws.Range("K" & Stock_Table_Row).Value = FormatPercent(Percent_Change)
     
            'Add one to table row
            Stock_Table_Row = Stock_Table_Row + 1
 
         'If the ticker name hasn't changed yet
        Else

            'Add to total stock volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value

        End If

    Next i
              
 'Now that we have completed the summary table, find the values within it for the bonus table
         
 'Print min and max percent change and max volume in the bonus table
   ws.Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(Columns("K")))
   ws.Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(Columns("K")))
   ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(Columns("L"))
          
      'Print ticker for values above in the bonus table
      For i = 2 To lastrow
          
           'Ticker for max percent change
            If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
                ws.Cells(2, 16) = ws.Cells(i, 9).Value
            
           'Ticker for min percent change
            ElseIf ws.Cells(i, 11) = ws.Cells(3, 17).Value Then
                ws.Cells(3, 16) = ws.Cells(i, 9).Value
            
            'Ticker for max volume
            ElseIf ws.Cells(i, 12) = ws.Cells(4, 17).Value Then
                ws.Cells(4, 16) = ws.Cells(i, 9).Value
                    
           End If
          
      Next i
 
 Next ws

End Sub
