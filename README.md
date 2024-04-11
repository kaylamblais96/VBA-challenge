# VBA-challenge
Sub Create_Headers()


'create a column named ticker
    Cells(1, 9).Value = "Ticker"
    
'create a column named yearly change
    Cells(1, 10).Value = "Yearly Change"

'create a column named percent change
    Cells(1, 11).Value = "Percent Change"

'create a column named total stock volume
    Cells(1, 12).Value = "Total Stock Volume"
    
'label cell O2 "greatest % increase"
    Cells(2, 15).Value = "Greatest % Increase"
    
'label cell O3 "greatest % decrease"
    Cells(3, 15).Value = "Greatest % Decrease"
    
'label cell O4 "greatest total volume"
    Cells(4, 15).Value = "Greatest Total Volume"
    
'label cell P1 "Ticker
    Cells(1, 16).Value = "Ticker"
    
'label cell Q1 "Value"
    Cells(1, 17).Value = "Values"
    
    
End Sub
    
  
    
    
Sub Calculate_Total_Stock_Volume()

    
'set an initial variable for holding the ticker
Dim ticker As String

'set an initial variable for holding the total stock volume
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'create a list of all the tickers down column i
Dim ticker_row As Integer
ticker_row = 2

'Loop through all tickers
For i = 2 To 753001

    'check if we are still withing the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'set the ticker
        ticker = Cells(i, 1).Value
        
        'add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
        'print the ticker names down column I
        Range("I" & ticker_row) = ticker
               
        'print the total stock volumes down column L
        Range("L" & ticker_row).Value = Total_Stock_Volume
        
        'add one to the ticker name rows
        ticker_row = ticker_row + 1
        
        'reset the total stock volume total
        Total_Stock_Volume = 0
        
        
        'if the immediate cell descending is the same ticker:
        Else
        
            'add to the total stock volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
        End If   
 Next i
 
 End Sub
 
 

 
Sub Calculate_Yearly_and_Percent_Change()

'set initial variables for yearly and percent change
Dim yearly_change As Double
Dim percent_change As Double

'set initial variables for opening and closing values
Dim opening_value As Double
Dim closing_value As Double

'create a list of all the tickers down column i
Dim ticker_row As Integer
ticker_row = 2


'Loop through all tickers
For i = 2 To 753001

    'if the cell below 'i' does not match, then
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'closing count is the value in that row in column "F"
    closing_count = Cells(i, 6).Value
    
    yearly_change = (closing_count - opening_count)
    
    percent_change = ((yearly_change / opening_count) * 100)
    
    'display the yearly change date in the "J" column
    Range("J" & ticker_row).Value = yearly_change
    
    'display the percent change date in the "K" column
    Range("K" & ticker_row).Value = Round(percent_change, 2)
        
    
    ticker_row = ticker_row + 1
         
    End If
    
    
    'if the cell above 'i' does not match, then
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
    'opening count is the value in that row in column "C"
    opening_count = Cells(i, 3).Value
    
    End If

Next i
    
End Sub




Sub Conditional_Color_Fill()


'if the yearly change is positive, fill the cell with green; if not, fill the cell with red

    For i = 2 To 22771
    
    If Cells(i, 10).Value >= 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        
    Else
        Cells(i, 10).Interior.ColorIndex = 3
        
    End If
    
    Next i
    

End Sub

Sub Max_and_Min_in_Dataset()

'set initial variable for Max Total Stock Volume
Dim Max_Volume As Double
Max_Volume = Application.Max(Range("L2:L3001"))

'Find Max Total Stock Volume
Application.Max (Range("L2:L3001"))

'display Max in Q4
Cells(4, 17).Value = Max_Volume



'set initial variable for Greatest % Increase
Dim Greatest_Increase As Double
Greatest_Increase = Application.Max(Range("K2:K3001"))

'Find Greatest % Increase
Application.Max (Range("K2:K3001"))

'display greatest increase in Q2
Cells(2, 17).Value = Greatest_Increase



'set initial variable for Greatest Decrease
Dim Greatest_Decrease As Double
Greatest_Decrease = Application.Min(Range("K2:K3001"))

'Find Greatest % Decrease
Application.Min (Range("K2:K3001"))

'display greatest decrease in Q3
Cells(3, 17).Value = Greatest_Decrease



'find the ticker associated with Max Total Stock Volume and display it in cell P4
Cells(4, 16).Value = Cells(Range("L2:L3001").Find(Max_Volume).Row, 9)

'find the ticker associated with Greatest % Increase and display it in cell P2
Cells(2, 16).Value = Cells(Range("K2:K3001").Find(Greatest_Increase).Row, 9)

'find the ticker associated with Greatest % Decrease and display it in cell P3
Cells(3, 16).Value = Cells(Range("K2:K3001").Find(Greatest_Decrease).Row, 9)

End Sub
