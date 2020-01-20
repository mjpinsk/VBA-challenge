

' Creates a table for each company stock ticker - the yearly dollar change, the yearly
' percentage change, and the yearly total stock volume is calculated
' Also the ticker for company with the best stock performance, the worst stock
' performance, and the greatest yearly total volume is also listed

Sub stock_performance()

  ' Set an initial variable for holding the company ticker
  Dim Comp_Tic As String

  ' Set an initial variable for the items in the summary table
  Dim Stock_Volume_Total As Double
  Dim Open_Price As Double
  Dim Close_Price As Double
  Dim Yearly_Dollar_Change As Double
  Dim Yearly_Percentage_Change As Double
  Dim Greatest_Percent_Increase As Double
  Dim Greatest_Percent_Decrease As Double
  Dim Greatest_Stock_Volume As Double
  
  ' declare a Worsheet variable for the Workbook
  Dim ws As Worksheet
  
' Cycle through each Worksheet in the Workbook
For Each ws In Worksheets
  
  ' Initialize these variable to 0
  Greatest_Percent_Increase = 0
  Greatest_Percent_Decrease = 0
  Yearly_Percentage_Change = 0
  Greatest_Stock_Volume = 0
  Stock_Volume_Total = 0
  
  ' Set these variables to these specific RGB colors
  GreenColor = RGB(0, 255, 0)
  RedColor = RGB(255, 0, 0)

  ' Keep track of the location for each company in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
  ' Increment place holder for each day for each ticker
  Dim i As Long
  i = 2

  ' Set the headers in the tables with the appropriate labels
  ws.Range("I" & 1).Value = "Ticker"
  ws.Range("J" & 1).Value = "Yearly Change"
  ws.Range("K" & 1).Value = "Percent Change"
  ws.Range("L" & 1).Value = "Total Stock Volume"
  ws.Range("P" & 1).Value = "Ticker"
  ws.Range("Q" & 1).Value = "Value"
  ws.Range("O" & 2).Value = "Greatest % Increase"
  ws.Range("O" & 3).Value = "Greatest % Decrease"
  ws.Range("O" & 4).Value = "Greatest Total Volume"
  
  
  ' Set the opening stock price of the first company
  Open_Price = ws.Cells(2, 3).Value

  ' Loop through the company tickers for each market day of the year
   While IsEmpty(ws.Cells(i, 1).Value) = False
 
    ' Check if we are still within the same company ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the company ticker symbol
      Comp_Tic = ws.Cells(i, 1).Value
      
      ' Set the closing price of the last market day for each company ticker
      Close_Price = ws.Cells(i, 6).Value
      
      ' Caculate the yearly dollar change for each symbol
      Yearly_Dollar_Change = Close_Price - Open_Price
     
      ' Caculate the yearly percentage change for each symbol, but first check if opening price is not 0
      If Open_Price <> 0 Then
        Yearly_Percentage_Change = (Close_Price - Open_Price) / Open_Price
      Else
       ' Only should occur if both open price and closing price is 0
        Yearly_Percentage_Change = 0
      End If

      ' Add to the stock volume total
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

      ' Print the company ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Comp_Tic
      
      ' Print the yearly dollar change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Dollar_Change
      
      ' Print the yearly percentage change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Format(Yearly_Percentage_Change, "Percent")
      
      ' Print the total stock volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
      
      ' Change the cell background color to red if negative and green if positive
       If (ws.Cells(Summary_Table_Row, "J").Value < 0) Then
          ws.Cells(Summary_Table_Row, "J").Interior.Color = RedColor
       ElseIf (ws.Cells(Summary_Table_Row, "J").Value > 0) Then
          ws.Cells(Summary_Table_Row, "J").Interior.Color = GreenColor
       End If
       
      ' Check to see if this company has the greatest percentage increase if so print it in the table
       If ws.Range("K" & Summary_Table_Row).Value > Greatest_Percent_Increase Then
          Greatest_Percent_Increase = ws.Range("K" & Summary_Table_Row).Value
          ws.Range("P" & 2).Value = Comp_Tic
          ws.Range("Q" & 2).Value = Format(Greatest_Percent_Increase, "Percent")
       End If
      
      ' Check to see if this company has the greatest percentage decrease if so print it in the table
       If ws.Range("K" & Summary_Table_Row).Value < Greatest_Percent_Decrease Then
          Greatest_Percent_Decrease = ws.Range("K" & Summary_Table_Row).Value
          ws.Range("P" & 3).Value = Comp_Tic
          ws.Range("Q" & 3).Value = Format(Greatest_Percent_Decrease, "Percent")
      End If
      
      ' Check to see if this company has the greatest stock volume if so print it in the table
      If ws.Range("L" & Summary_Table_Row).Value > Greatest_Stock_Volume Then
         Greatest_Stock_Volume = ws.Range("L" & Summary_Table_Row).Value
         ws.Range("P" & 4).Value = Comp_Tic
         ws.Range("Q" & 4).Value = Greatest_Stock_Volume
      End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset the Stock_Volume_Total for new ticker
      Stock_Volume_Total = 0
      
      ' Reset opening price of new ticker
      Open_Price = ws.Cells(i + 1, 3).Value


    ' If the cell immediately following a row is the same company ticker
    Else
    
    ' Check to see if the stock has opened with a price in the current year
       If Open_Price = 0 Then
         Open_Price = ws.Cells(i, 3).Value
       End If
       
      ' Add to the stock volume total
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

    End If

  ' Increment i
   i = i + 1
 Wend

   ' Activate the current worksheet
    ws.Activate
Next ws

End Sub

