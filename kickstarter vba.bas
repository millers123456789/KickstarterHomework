Attribute VB_Name = "Module3"
Sub Multiple_year_stock_data()

'Set an initial varaible for holding the Ticker name
Dim Ticker As String
'Set an inital variable for holding yearly change
Dim Yearly_Change As Double
'Set an inital variable for Percent Change
Dim Percent_Change As Double
'Set an inital variable for Greatest % Increase
Dim Greatest_Percent_Increase As Double
Dim p As Integer
Dim I As Double
Dim Greatest_Ticker As Double
Dim Greates_Volume As Double





'Set an inital variable for Greatest % Decrease
Dim Greatest_Percent_Decrease As Double

Dim k As Integer


'Set an inital variable for holding the total volume

Dim Total_Volume As Double
Total_Volume = 0

'Keep track of the location for each tacker name
Yearly_Change = 0
Dim Summary_Table_Row As Integer

    Summary_Table_Row = 2
    
    'Loop through  all credit card purchases
    
    For I = 2 To 5000
        'Check if we are still within the same Ticker name, if not....
        
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        'Calculate yearly change by subtracting closing at end of the year by opening at beginning
        
            Yearly_Change = Cells(I, 6).Value - Cells(I - 260, 3)
            
            'Calculate the Percent change
            Percent_Change = (Cells(I, 6).Value / Cells(I - 260, 3)) - 1
            'Print Percent_Change in row in summary table
            Range("L" & Summary_Table_Row).Value = Percent_Change
            
            'Set Ticker name
            Ticker = Cells(I, 1).Value
            
            'Add to the Volume Total
            Total_Volume = Total_Volume + Cells(I, 7).Value
            'Print the Ticker name in the Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker
            
            'Print Yearly_Change in row in summary table row
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            'Print the Total Volume
            Range("M" & Summary_Table_Row).Value = Total_Volume
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset Total Volume
            Total_Volume = 0
            'Reset yearly change
            Yearly_Change = 0
            'Reset Percent Change
            Percent_Change = 0
            Range("L2:L20").NumberFormat = "0.00%"
            
            'If the cell immediately following a row is the same Ticker
            
            
            Else
            'Add to the Total Volume
                Total_Volume = Total_Volume + Cells(I, 7).Value
                
            End If
            
            If Range("K" & Summary_Table_Row).Value < 0 Then
             Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
             
             Else: Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
             
             End If
             
            
            
        Next I
   'Set iteration for loop to compare percent changes
   'Set holding place for percentage Greatest Change
   Range("Q2:Q3").NumberFormat = "0.00%"
   Greatest_Percent_Increase = 0
   
        For k = 2 To 20
            If Cells(k, 12).Value > Greatest_Percent_Increase Then
                Greatest_Percent_Increase = Cells(k, 12).Value
                Range("P2").Value = Cells(k, 10).Value
                
                
            
            
            End If
            
            
        Next k
        Greatest_Percent_Decrease = 0
        'Iterate through percent change cells
        For p = 2 To 20
            If Cells(p, 12).Value < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Cells(p, 12).Value
                Range("P3").Value = Cells(p, 10).Value
                
                
            
            
            End If
        Next p
        
        Range("Q2").Value = Greatest_Percent_Increase
        Range("Q3").Value = Greatest_Percent_Decrease
        For p = 2 To 20
            If Cells(p, 12).Value < Greatest_Percent_Decrease Then
                Greatest_Percent_Decrease = Cells(p, 12).Value
                Range("P2:P3").Value = Cells(p, 10).Value
                
            
            
            End If
           
        Next p
        Greatest_Volume = 0
        For k = 2 To 20
            If Cells(k, 13).Value > Greatest_Volume Then
                Greatest_Volume = Cells(k, 13).Value
                Range("P4").Value = Cells(k, 10).Value
                Range("Q4").Value = Cells(k, 13).Value
                
            
            
            End If
        Next k

End Sub
