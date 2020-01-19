Attribute VB_Name = "Module1"
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
  
  ' Initialize these variable to 0
  Greatest_Percent_Increase = 0
  Greatest_Percent_Decrease = 0
  Yearly_Percentage_Change = 0
  Greatest_Stock_Volume = 0
  Stock_Volume_Total = 0

  ' Keep track of the location for each company in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Set the opening stock price of the first company
  Open_Price = Cells(2, 3).Value

  ' Loop through the company tickers for each market day of the year
  For i = 2 To 797711

    ' Check if we are still within the same company ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the company ticker symbol
      Comp_Tic = Cells(i, 1).Value
      
      ' Set the closing price of the last market day for each company ticker
      Close_Price = Cells(i, 6).Value
      
      ' Caculate the yearly dollar change for each symbol
      Yearly_Dollar_Change = Close_Price - Open_Price
      
      ' Caculate the yearly percentage change for each symbol
      If Open_Price <> 0 Then
        Yearly_Percentage_Change = (Close_Price - Open_Price) / Open_Price
      Else
        Yearly_Percentage_Change = 0
      End If

      ' Add to the stock volume total
      Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

      ' Print the company ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Comp_Tic
      
      ' Print the yearly dollar change in the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_Dollar_Change
      
      ' Print the yearly percentage change in the Summary Table
      Range("K" & Summary_Table_Row).Value = Format(Yearly_Percentage_Change, "Percent")
      
      ' Print the total stock volume to the Summary Table
      Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
      
      ' Check to see if this company has the greatest percentage increase if so print it in the table
       If Range("K" & Summary_Table_Row).Value > Greatest_Percent_Increase Then
          Greatest_Percent_Increase = Range("K" & Summary_Table_Row).Value
          Range("P" & 2).Value = Comp_Tic
          Range("Q" & 2).Value = Format(Greatest_Percent_Increase, "Percent")
      End If
      
      ' Check to see if this company has the greatest percentage decrease if so print it in the table
      If Range("K" & Summary_Table_Row).Value < Greatest_Percent_Decrease Then
         Greatest_Percent_Decrease = Range("K" & Summary_Table_Row).Value
         Range("P" & 3).Value = Comp_Tic
         Range("Q" & 3).Value = Format(Greatest_Percent_Decrease, "Percent")
      End If
      
      ' Check to see if this company has the greatest stock volume if so print it in the table
      If Range("L" & Summary_Table_Row).Value > Greatest_Stock_Volume Then
         Greatest_Stock_Volume = Range("L" & Summary_Table_Row).Value
         Range("P" & 4).Value = Comp_Tic
         Range("Q" & 4).Value = Greatest_Stock_Volume
      End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset the Stock_Volume_Total for new ticker
      Stock_Volume_Total = 0
      
      ' Reset opening price of new ticker
      Open_Price = Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same company ticker
    Else

      ' Add to the stock volume total
      Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

