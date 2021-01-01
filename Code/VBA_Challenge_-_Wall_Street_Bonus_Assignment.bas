Attribute VB_Name = "Module2"
Sub wall_street_bonus()
  ' --------------------------------------------
  ' --- VBA Wall Street Bonus Assignment     ---
  ' --------------------------------------------

  ' --------------------------------------------
  ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
  For Each ws In Worksheets


  ' --------------------------------------------
  ' INITIALIZATION
  ' --------------------------------------------
  ' Create and Set variable values for the set of ticker symbols
  Dim Great_Perc_Inc_Ticker_Symbol As String
  Dim Great_Perc_Dec_Ticker_Symbol As String
  Dim Great_Total_Vol_Ticker_Symbol As String
  
  ' Setting each ticker value to the first line value as a starting point
  Great_Perc_Inc_Ticker_Symbol = ws.Cells(2, "I").Value
  Great_Perc_Dec_Ticker_Symbol = ws.Cells(2, "I").Value
  Great_Total_Vol_Ticker_Symbol = ws.Cells(2, "I").Value
    
  ' Create and Set initial variable values for holding the associated values per ticker symbol
  Dim Great_Perc_Inc_Value As Double
  Dim Great_Perc_Dec_Value As Double
  Dim Great_Total_Vol_Value As Double
    
  ' Setting each variable to the first line value as a starting point
  Great_Perc_Inc_Value = ws.Cells(2, "K").Value
  Great_Perc_Dec_Value = ws.Cells(2, "K").Value
  Great_Total_Vol_Value = ws.Cells(2, "L").Value

  ' Output Formatted Headers for the Summary table
  ' ** Add the year for readability ***
  ws.Range("O" & "1").Value = "Analysis Year = " & ws.Name
  ws.Range("P" & "1").Value = "Ticker"
  ws.Range("Q" & "1").Value = "Value"
  ws.Range("O1:Q1").Interior.ColorIndex = 15
  ws.Range("O1:Q1").Font.FontStyle = "Bold"
  ws.Range("O1:Q1").HorizontalAlignment = xlCenter
  
  ' Put boxes around the headers
  Dim rng As Range
  
  ' Define range
  Set rng = ws.Range("O1:Q1")

    With rng.Borders
      .LineStyle = xlContinuous
      .Weight = xlThick
      .ColorIndex = 0
      .TintAndShade = 0
    End With
  
  ' Determine the Last Row in the Summary Table so we know the extent of the search loop
  Dim LastRow As Long
  LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

  ' ------------------------------------------------------
  ' --- Main Code for VBA Wall Street Bonus Assignment ---
  ' ------------------------------------------------------

  ' Loop over the Summary Table created in Module 1
  
  For i = 2 To LastRow
  
  ' For Each Value: Compare to next value and if needed revise values and Ticker Symbol for each item
  
  
    ' For Greatest Percent: Compare to next value and if needed revise value and Ticker Symbol
    If ws.Cells(i + 1, "K").Value > Great_Perc_Inc_Value Then

      ' Set the Percent Increase Ticker Symbol and Value to the new values
      Great_Perc_Inc_Ticker_Symbol = ws.Cells(i + 1, "I").Value
      Great_Perc_Inc_Value = ws.Cells(i + 1, "K").Value
    
    End If
    
    ' For Least percent: Compare to next value and if needed revise value and Ticker Symbol
    If ws.Cells(i + 1, "K").Value < Great_Perc_Dec_Value Then

      ' Set the Percent Decrease Ticker Symbol and Value to the new values
      Great_Perc_Dec_Ticker_Symbol = ws.Cells(i + 1, "I").Value
      Great_Perc_Dec_Value = ws.Cells(i + 1, "K").Value
    End If
    
    ' For Greatest Volume: Compare to next value and if needed revise value and Ticker Symbol
    If ws.Cells(i + 1, "L").Value > Great_Total_Vol_Value Then

      ' Set the Volume Ticker Symbol and Value to the new values
      Great_Total_Vol_Ticker_Symbol = ws.Cells(i + 1, "I").Value
      Great_Total_Vol_Value = ws.Cells(i + 1, "L").Value
    End If
      
  Next i
  
  ' ---------------------------
  ' --- Finish and Tidy up  ---
  ' ---------------------------
  
  ' Output the Final Values and Format the Values appropriately
  ws.Range("O" & 2).Value = "Greatest % Increase"
  ws.Range("P" & 2).Value = Great_Perc_Inc_Ticker_Symbol
  ws.Range("Q" & 2).Value = Great_Perc_Inc_Value
  ws.Range("Q" & 2).NumberFormat = "0.00%"
  
  ws.Range("O" & 3).Value = "Greatest % Decrease"
  ws.Range("P" & 3).Value = Great_Perc_Dec_Ticker_Symbol
  ws.Range("Q" & 3).Value = Great_Perc_Dec_Value
  ws.Range("Q" & 3).NumberFormat = "0.00%"
     
  ws.Range("O" & 4).Value = "Greatest Total Volume"
  ws.Range("P" & 4).Value = Great_Total_Vol_Ticker_Symbol
  ws.Range("Q" & 4).Value = Great_Total_Vol_Value
  ws.Range("Q" & 4).NumberFormat = "0,000"
   
  
  ' Reset the columns so everything looks nice
  ws.Columns("O:Q").AutoFit
  Set rng = ws.Range("O2", "Q4")

  ' Put boxes around the output
  With rng.Borders
     .LineStyle = xlContinuous
     .Weight = xlThick
     .ColorIndex = 0
     .TintAndShade = 0
  End With
      
  ' --------------------------------------------
  ' Go to the next worksheet
  ' --------------------------------------------
  Next ws

End Sub

