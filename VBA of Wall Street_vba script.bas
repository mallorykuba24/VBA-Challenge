Attribute VB_Name = "Module1"
Sub VBAChallenge()

'Loop through all sheets

Dim WS As Worksheet

For Each WS In Worksheets
 
'Label column titles
WS.Range("I1") = "Ticker Symbol"
WS.Range("J1") = "Yearly Change"
WS.Range("K1") = "Percent Change"
WS.Range("L1") = "Total stock volume"

'Define variables
Dim Ticker_Symbol As String
Dim Yearly_change As Double
Yearly_change = 0
Dim Percent_change As Double
Percent_change = 0
Dim Total_stock_volume As Double
Total_stock_volume = 0
Dim LastRow As Long
Dim Counter As Integer
Dim ClosePrice As Double
ClosePrice = 0
Dim OpenPrice As Double
OpenPrice = 0

'Keep track of location in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
   
'Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
       
'Loop
    
For i = 2 To LastRow
    
    'Conditional to determine open_price
    
    If WS.Cells(i, 1).Value <> WS.Cells(i - 1, 1).Value Then

    OpenPrice = WS.Cells(i, 3).Value

    End If
                                    
        'Check if we are still the same ticker value, if it is not...
        If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
       
            'Set ticker symbol
            Ticker_Symbol = WS.Cells(i, 1).Value
           
            'Add to total stock volume
            Total_stock_volume = Total_stock_volume + WS.Cells(i, 7).Value
           
            'Calculate Yearly_change
            ClosePrice = WS.Cells(i, 6).Value
                      
            Yearly_change = ClosePrice - OpenPrice
                      
            'Conditional to determine Percent_change
                               
                If OpenPrice = 0 And ClosePrice = 0 Then
                    'Cannot divide by zero
                    Percent_change = 0
                    WS.Cells(Summary_Table_Row, 11).Value = Percent_change
                    WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                ElseIf OpenPrice = 0 Then
                    'If a stock starts at zero and increases,we need to look at actual price increase by dollar amount
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    WS.Cells(Summary_Table_Row, 11).Value = Percent_change
                Else
                    Percent_change = Yearly_change / OpenPrice
                    WS.Cells(Summary_Table_Row, 11).Value = Percent_change
                    WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                End If
            
            'Print the percent_change in the summary table
             
            WS.Range("K" & Summary_Table_Row).Value = Percent_change
        
            'Print the yearly_change in the summary table
            WS.Range("J" & Summary_Table_Row).Value = Yearly_change
           
            'Print the ticker symbol in the summary table
            WS.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
           
            'Print the total stock volume to the summary table
            WS.Range("L" & Summary_Table_Row).Value = Total_stock_volume
                   
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
           
            'Rest values
            Total_stock_volume = 0
            Percent_change = 0
            Yearly_change = 0
            ClosePrice = 0
            OpenPrice = 0
           
        'If cell immediately following a row is the same ticker...
        Else
   
            'Add to the total stock volume
            Total_stock_volume = Total_stock_volume + WS.Cells(i, 7).Value
                                           
                      
        End If
       
        Next i

    'Convert format from number to percent for Percent_Change
        WS.Range("K1:K" & LastRow).NumberFormat = "0.00%"
   
    'Color code the yearly_change box to green
       
        For j = 2 To LastRow
       
        If WS.Cells(j, 10).Value > 0 Then
     
        WS.Cells(j, 10).Interior.ColorIndex = 4
       
        End If
       
        Next j
   
    'Color code the yearly_change box to red
       
        For k = 2 To LastRow
       
        If WS.Cells(k, 10).Value < 0 Then
     
        WS.Cells(k, 10).Interior.ColorIndex = 3
       
        End If
       
        Next k

WS.Activate

Next WS

End Sub

