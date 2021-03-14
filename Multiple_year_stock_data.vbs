VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stockmarket():

    'Define variables
    Dim i As Long
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Double
    Dim ws As Worksheet
    Dim Last_Row As Long
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    'Loop through all worksheets
    For Each ws In ActiveWorkbook.Worksheets
    
    'Create Summary_Table
    ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        
   'Set starting points
    Total_Stock_Volume = 0
    Yearly_Change = 0
    Summary_Table_Row = 1
        
    'Determine Last_Row
       Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through Dataset
       For i = 2 To Last_Row
            
            'Check for change in Ticker and set new starting values
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                 'Add one to Summary_table_row for next entry
                 Summary_Table_Row = Summary_Table_Row + 1
                 Ticker = ws.Cells(i, 1).Value
                 Opening_Price = ws.Cells(i, 3).Value
                 Total_Stock_Volume = 0
                  
            'If next cell is different, calculate sum of ticker stock volume and determine final closing_price for ticker
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                Closing_Price = ws.Cells(i, 6).Value
                
                'Calculate yearly_change value
                Yearly_Change = Closing_Price - Opening_Price
                
                'Account for zero and then determine percent_change
                If Yearly_Change = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = (Yearly_Change / Opening_Price)
                End If
                
                'Percent_Change casted to Double and multiply by 100
                ws.Range("K" & Summary_Table_Row).Value = CDbl(Percent_Range) * 100
                          
                'Print Ticker, Yearly_change, Percent_change, and Total_Stock_Volume in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'Format Percent_Change column
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            Else
                'Add stock volumes if ticker is the same
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'Account for zero for opening_price
                If Opening_Price = 0 And ws.Cells(i, 3).Value <> 0 Then
                    Opening_Price = ws.Cells(i, 3).Value
                
                End If
                
            End If
        
        Next i
    
        'Format summary table
       For i = 2 To Last_Row

            'Yearly_Change column Green for positive and Red for negative
            If ws.Cells(i, 10) > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
              Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
              
                End If
            
            'No color change if cell is empty
            If ws.Cells(i, 10) = "" Then
                ws.Cells(i, 10).Interior.ColorIndex = Null
            
                End If
            
          Next i
          
                   
Next ws
    
End Sub


