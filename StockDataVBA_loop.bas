Attribute VB_Name = "Module1"

Sub StockData()

'set up loop to cycle through worksheets
'Worksheet loop syntax found on udemy

Dim ws As Worksheet

For Each ws In Worksheets

Worksheets(ws.Name).Select
    
    

'Initial Loop for finding unique Ticker Names and Total Stock volume per unique Ticker

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

  ' Set an initial variable for ticker name
  Dim Ticker_Name As String
  
  ' Set an initial variable for holding the total stock volume value
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0 'to identify starting value of stock volume

  ' Keep track of the location for each new ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock data
  For i = 2 To 759001 'because there is 753001 rows of stock data

    ' Check if we are still within the ticker name
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Total Stok Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      ' Print the each unique Ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Total Stock Volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      ' Add one row to the summary table
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      Total_Stock_Volume = 0

    ' If cell does not match above condition or aka new ticker name is not identified
    Else

      ' Add to the Brand Total
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i

'syntax found on udemy
Columns("J").ColumnWidth = 15
Columns("K").ColumnWidth = 15
Columns("L").ColumnWidth = 17




'Finding Opening and Closing Values for Future Calculations


'Set Headers
Range("N1").Value = "Opening Value"
Range("O1").Value = "Closing Value"

'Identify Variables
Dim Date_Ref As String
Dim Opening_Value As Double
Dim Closing_Value As Double

'Set the the row
'Dim Summary_Table_Row As Integer (do not need to reclassify variable so commented out for my reference)
Summary_Table_Row = 2 'resetting the starting place of the table row

'Begin For Loop to find opening and closing values of all unique ticker values
    For i = 2 To 759001
        
        'First conditional statement for finding the opening value
        If (Right(Cells(i, 2).Value, 4) = "0102") Then
    
            Opening_Value = Cells(i, 3).Value
    
            Range("N" & Summary_Table_Row).Value = Opening_Value
        
        'Second conditional statement for finding the opening value
        ElseIf (Right(Cells(i, 2).Value, 4) = "1231") Then
    
            Closing_Value = Cells(i, 6).Value
    
            Range("O" & Summary_Table_Row).Value = Closing_Value
            
            Summary_Table_Row = Summary_Table_Row + 1
    
        Else
    
        End If
        
    Next i

Columns("N").ColumnWidth = 14
Columns("O").ColumnWidth = 14





'Yearly and Percent Change


'now create a loop to find the yearly and percent change

Dim Yearly_Change As Double
Dim Percent_Change As Double


    For i = 2 To 3001
    
        Opening_Value = Cells(i, 14).Value
        
        Closing_Value = Cells(i, 15).Value
        
        Yearly_Change = Closing_Value - Opening_Value
        
        'Print the Yearly Change
        Range("J" & i).Value = Format(Yearly_Change, "0.00#")
        
        Percent_Change = (Yearly_Change / Opening_Value)
        
        ' Print the Percent Change and Format as a Percent
        Range("K" & i).Value = FormatPercent(Percent_Change)
        
        
            'If function for conditional formatting
            If (Yearly_Change >= 0) Then
            
            'if the yearly change is 0 or above then interior color is green
            Range("J" & i).Interior.ColorIndex = 4
            
            Else
            
            'if the yearly change is below 0 then interior color is red
            Range("J" & i).Interior.ColorIndex = 3
            
            End If
            
                 
    Next i



'Create Final Summary Table

'this table with include Greatest % Increase and Decrease, and the Great Total Volume

'Make room for our summary table, asked ChatGPT for this syntax
Columns("M:O").Hidden = True

'set labels for summary table
Range("R1").Value = "Ticker"
Range("S1").Value = "Value"
Range("Q2").Value = "Greatest % Increase"
Range("Q3").Value = "Greatest % Decrease"
Range("Q4").Value = "Greatest Total Volume"

'idenify variables
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Increase_Ticker As String
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Greatest_Total_Volume As Double
Dim Greatest_Total_Volume_Ticker As String

'clean up column widths for better presentation
Columns("Q").ColumnWidth = 20
Columns("R").ColumnWidth = 11
Columns("S").ColumnWidth = 11


    'Find the Greatest Percent Increase my indentifying the max value in the percent change column
    '(Max and Min formula found on learn.microsoft.com)
    Greatest_Percent_Increase = WorksheetFunction.Max(Range("K2:K3001"))
    Cells(2, 19) = (FormatPercent(Greatest_Percent_Increase))
 
    'Find the Greatest Percent Decrease my indentifying the min value in the percent change column
    Greatest_Percent_Decrease = WorksheetFunction.Min(Range("K2:K3001"))
    Cells(3, 19) = (FormatPercent(Greatest_Percent_Decrease))
    
    'Find the Great Total Stock Volume my indentifying the max value in the Total Stock Volume Column
    Greatest_Total_Volume = WorksheetFunction.Max(Range("L2:L3001"))
    Cells(4, 19) = Greatest_Total_Volume
 
    
    ' Creating a for loop and if conditionals to match the ticker name
    For i = 2 To 3001
        
        'identifying the ticker with the greatest percent increase
        If (Greatest_Percent_Increase = Cells(i, 11).Value) Then
        
            Greatest_Percent_Increase_Ticker = Cells(i, 9).Value
        
            Cells(2, 18) = Greatest_Percent_Increase_Ticker
        
        'identifying the ticker with the greatest percent decrease
        ElseIf (Greatest_Percent_Decrease = Cells(i, 11).Value) Then
        
            Greatest_Percent_Decrease_Ticker = Cells(i, 9).Value
            
            Cells(3, 18) = Greatest_Percent_Decrease_Ticker
        
        'identifying the ticker with the greatest total stock volume
        ElseIf (Greatest_Total_Volume = Cells(i, 12).Value) Then
        
            Greatest_Total_Volume_Ticker = Cells(i, 9).Value
            
            Cells(4, 18) = Greatest_Total_Volume_Ticker
                       
        'else to search in the next row if none of the above conditions are met
        Else
        
        End If
        
    Next i


'to next worksheet

Next ws

End Sub

