Attribute VB_Name = "Module3"
Sub wall_street_bonus_with_dup_search()
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
 
  '*****************************************************************
  '*** EXTRA - In Case of Duplicates - set second set of Tickers ***
  '*****************************************************************
  ' Create variable values for the duplicate set of ticker symbols
  Dim Great_Perc_Inc_Ticker_Symbol_Dup As String
  Dim Great_Perc_Dec_Ticker_Symbol_Dup As String
  Dim Great_Total_Vol_Ticker_Symbol_Dup As String
  
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

  '*****************************************************************
  '*** EXTRA - In Case of Duplicates - set second set of values  ***
  '*****************************************************************
  ' Create variable values for the set of duplicate values
  Dim Great_Perc_Inc_Value_Dup As Double
  Dim Great_Perc_Dec_Value_Dup As Double
  Dim Great_Total_Vol_Value_Dup As Double

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
    
  '*****************************************************************
  '*** EXTRA - Clear Duplicate if new base value set             ***
  '*****************************************************************
      ' Set the Duplicate Percent Increase Ticker Symbol and Value to Null
      Great_Perc_Inc_Ticker_Symbol_Dup = ""
      Great_Perc_Inc_Value_Dup = 0
        
  '*****************************************************************
  '*** EXTRA - If Duplicate, capture it                          ***
  '*****************************************************************
    ElseIf ws.Cells(i + 1, "K").Value = Great_Perc_Inc_Value Then
      ' Set the Dup Percent Increase Ticker Symbol and Value to the new values
      Great_Perc_Inc_Ticker_Symbol_Dup = ws.Cells(i + 1, "I").Value
      Great_Perc_Inc_Value_Dup = ws.Cells(i + 1, "K").Value
                
    End If
        
    ' For Least percent: Compare to next value and if needed revise value and Ticker Symbol
    If ws.Cells(i + 1, "K").Value < Great_Perc_Dec_Value Then

      ' Set the Percent Decrease Ticker Symbol and Value to the new values
      Great_Perc_Dec_Ticker_Symbol = ws.Cells(i + 1, "I").Value
      Great_Perc_Dec_Value = ws.Cells(i + 1, "K").Value
    
  '*****************************************************************
  '*** EXTRA - Clear Duplicate if new base value set             ***
  '*****************************************************************
      ' Set the Duplicate Percent Decrease Ticker Symbol and Value to Null
      Great_Perc_Dec_Ticker_Symbol_Dup = ""
      Great_Perc_Dec_Value_Dup = 0
        
  '*****************************************************************
  '*** EXTRA - If Duplicate, capture it                          ***
  '*****************************************************************
    ElseIf ws.Cells(i + 1, "K").Value = Great_Perc_Dec_Value Then
      ' Set the Dup Percent Decrease Ticker Symbol and Value to the new values
      Great_Perc_Dec_Ticker_Symbol_Dup = ws.Cells(i + 1, "I").Value
      Great_Perc_Dec_Value_Dup = ws.Cells(i + 1, "K").Value
                  
    End If
    
    ' For Greatest Volume: Compare to next value and if needed revise value and Ticker Symbol
    If ws.Cells(i + 1, "L").Value > Great_Total_Vol_Value Then

      ' Set the Volume Ticker Symbol and Value to the new values
      Great_Total_Vol_Ticker_Symbol = ws.Cells(i + 1, "I").Value
      Great_Total_Vol_Value = ws.Cells(i + 1, "L").Value
  
  '*****************************************************************
  '*** EXTRA - Clear Duplicate if new base value set             ***
  '*****************************************************************
   ' Set the Duplicate Percent Decrease Ticker Symbol and Value to Null
      Great_Total_Vol_Ticker_Symbol_Dup = ""
      Great_Total_Vol_Value_Dup = 0
        
  '*****************************************************************
  '*** EXTRA - If Duplicate, capture it                          ***
  '*****************************************************************
    ElseIf ws.Cells(i + 1, "L").Value = Great_Total_Vol_Value Then
      ' Set the Dup Percent Decrease Ticker Symbol and Value to the new values
      Great_Total_Vol_Ticker_Symbol_Dup = ws.Cells(i + 1, "I").Value
      Great_Total_Vol_Value_Dup = ws.Cells(i + 1, "L").Value
    
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
   
  '*****************************************************************
  '*** EXTRA - If Duplicates, Print them                         ***
  '*****************************************************************
  If Great_Perc_Inc_Ticker_Symbol_Dup <> "" Then
  ws.Range("R" & 2).Value = "Greatest % Increase (Dup)"
  ws.Range("S" & 2).Value = Great_Perc_Inc_Ticker_Symbol_Dup
  ws.Range("T" & 2).Value = Great_Perc_Inc_Value_Dup
  ws.Range("T" & 2).NumberFormat = "0.00%"
  End If
  
  If Great_Perc_Dec_Ticker_Symbol_Dup <> "" Then
  ws.Range("R" & 3).Value = "Greatest % Decrease (Dup)"
  ws.Range("S" & 3).Value = Great_Perc_Dec_Ticker_Symbol_Dup
  ws.Range("T" & 3).Value = Great_Perc_Dec_Value_Dup
  ws.Range("T" & 3).NumberFormat = "0.00%"
  End If
     
  If Great_Total_Vol_Ticker_Symbol_Dup <> "" Then
  ws.Range("R" & 4).Value = "Greatest Total Volume (Dup)"
  ws.Range("S" & 4).Value = Great_Total_Vol_Ticker_Symbol_Dup
  ws.Range("T" & 4).Value = Great_Total_Vol_Value_Dup
  ws.Range("T" & 4).NumberFormat = "0,000"
  End If
       
  If (Great_Perc_Inc_Ticker_Symbol_Dup <> "") Or (Great_Perc_Dec_Ticker_Symbol_Dup <> "") Or (Great_Total_Vol_Ticker_Symbol_Dup <> "") Then
  ' If Duplicates then Output Formatted Headers for the Summary table
  ' ** Add the year for readability ***
  ws.Range("R" & "1").Value = "Duplicate Type"
  ws.Range("S" & "1").Value = "Dup Ticker"
  ws.Range("T" & "1").Value = "Dup Value"
  ws.Range("R1:T1").Interior.ColorIndex = 15
  ws.Range("R1:T1").Font.FontStyle = "Bold"
  ws.Range("R1:T1").HorizontalAlignment = xlCenter
    
  ' Define range
  Set rng = ws.Range("R1:T1")

    With rng.Borders
      .LineStyle = xlContinuous
      .Weight = xlThick
      .ColorIndex = 0
      .TintAndShade = 0
    End With
  
  Set rng = ws.Range("O2", "T4")
  
  Else: Set rng = ws.Range("O2", "Q4")
  End If

  ' Put boxes around the output
  With rng.Borders
     .LineStyle = xlContinuous
     .Weight = xlThick
     .ColorIndex = 0
     .TintAndShade = 0
  End With
      
  ' Reset the columns so everything looks nice
  ws.Columns("O:T").AutoFit
     
      
  ' --------------------------------------------
  ' Go to the next worksheet
  ' --------------------------------------------
  Next ws

End Sub

