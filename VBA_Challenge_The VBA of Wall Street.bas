Attribute VB_Name = "Module1"
'Procedure name created as VBA_Wall_Street

Public Sub VBA_Wall_Street()

'Define variables using available data
    
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Stock_Volume As Long
    
'Define required variables for the analysis (New table (#1) to be generated)
    Dim Ticker_Count As Long
    Dim Total_Open_Price As Double
    Dim Total_Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
  
'Define variable to read the last row of each column which helps to apply this VBA program to any file with exsisting format
    Dim LastRow As Long
    
'Use For loop that can be use for all the worksheets in the same excel file
    For Each ws In Worksheets

'Activate all the sheets within the same file
    ws.Activate

'Define the last row using standard VBA functions: Starts from Column "A" or it can be represented as column 1
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Assign column/field names to the New table#1 to be generated
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total_Stock_Volume"
  
'Set up the counts and values for newly defined variables for analysis
    Ticker = ""
    Ticker_Count = 0
    Open_Price = 0
    Yearly_Change = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
    
'USe For loop to read rows from 2 to last row of each sheet. Starting i = 2 because first row is the field/column name. Therefore skipped the first row.
  For i = 2 To LastRow
  
'Define the Ticker_count for which represents in the first column
    Ticker = Cells(i, 1).Value
    
'Use If conditional to initiate the Open_price and define the value which is third column.
    If Open_Price = 0 Then
        Open_Price = Cells(i, 3).Value
    End If
    
'Set up Total_Stock_Volume count to add the values which is seventh column.
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
  
' Use If conditional to get the unmatching Tickers and get their count
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker_Count = Ticker_Count + 1
        Cells(Ticker_Count + 1, 9).Value = Ticker
        Close_Price = Cells(i, 6).Value
  
        Yearly_Change = Close_Price - Open_Price
        Cells(Ticker_Count + 1, 10).Value = Yearly_Change
        
 'Use If conditional to perform conditional formatting for color coding: green for positive Yearly_change and Red for Negative Yearly_change
    If Yearly_Change > 0 Then
            Cells(Ticker_Count + 1, 10).Interior.Color = vbGreen
    Else
            Cells(Ticker_Count + 1, 10).Interior.Color = vbRed
          
    End If
    
 'Use If conditional to generate the Total_open _count and the Yearly_percent change
    If Open_Price = 0 Then
        Percent_Change = 0
        
    Else
        Percent_Change = (Yearly_Change / Open_Price)
        Cells(Ticker_Count + 1, 11).Value = Format(Percent_Change, "Percent") 'percent format change?
    End If
    
 'Use If condition to perform conditional formatting for color coding
    If Percent_Change > 0 Then
            Cells(Ticker_Count + 1, 11).Interior.Color = vbGreen
    Else
        Cells(Ticker_Count + 1, 11).Interior.Color = vbRed
          
    End If
        
        Open_Price = 0
        Cells(Ticker_Count + 1, 12).Value = Total_Stock_Volume
        Total_Stock_Volume = 0
    End If
 Next i
'Bonus section
'Assign columns for the new variables: column O, P, Q
'Leave O1 Blank because no field name there
Range("O2").Value = "Greatest_Percent_Increase"
Range("O3").Value = "Greatest_Percent_Decrease"
Range("O4").Value = "Greatest_Total_Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

'Define the last row using standard VBA functions: Starts from Column "I" or it can be represented as 9. Because the newset table that is going to be generated is based on the previous table generated.
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Define the variables
Greatest_Percent_Increase = Cells(i, 11).Value
Greatest_Percent_Decrease = Cells(i, 11).Value
Greatest_Total_Volume = Cells(i, 12).Value
Greatest_Percent_Increase_Ticker = Cells(i, 9).Value
Greatest_Percent_Decrease_Ticker = Cells(i, 9).Value
Greatest_Total_Volume_Ticker = Cells(i, 9).Value

'Use for loop to create data looking for
For i = 2 To LastRow

If Cells(i, 11).Value > Greatest_Percent_Increase Then
    Greatest_Percent_Increase = Cells(i, 11).Value
    Greatest_Percent_Increase_Ticker = Cells(i, 9).Value
End If


If Cells(i, 11).Value < Greatest_Percent_Decrease Then
    Greatest_Percent_Decrease = Cells(i, 11).Value
    Greatest_Percent_Decrease_Ticker = Cells(i, 9).Value
End If


If Cells(i, 12).Value > Greatest_Total_Volume Then
    Greatest_Total_Volume = Cells(i, 12).Value
    Greatest_Total_Volume_Ticker = Cells(i, 9).Value
End If
Next i

'Format changes and assign values
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Range("P2").Value = Greatest_Percent_Increase_Ticker
Range("P3").Value = Greatest_Percent_Decrease_Ticker
Range("P4").Value = Greatest_Total_Volume_Ticker

Range("Q2").Value = Format(Greatest_Percent_Increase, "percent")
Range("Q3").Value = Format(Greatest_Percent_Decrease, "percent")
Range("Q4").Value = Greatest_Total_Volume

Next ws

End Sub
    
