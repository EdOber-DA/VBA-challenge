Attribute VB_Name = "Module1"
Sub wall_street()
  ' --------------------------------------------
  ' --- VBA Wall Street Primary Assignment   ---
  ' --------------------------------------------

  ' --------------------------------------------
  ' LOOP THROUGH ALL SHEETS
  ' --------------------------------------------
  For Each ws In Worksheets


  ' --------------------------------------------
  ' INITIALIZATION
  ' --------------------------------------------
  
  ' --------------------------------------------
  ' Prepare for the search of the Input Table data rows
  ' --------------------------------------------
  
  ' Set an initial variable for ticker symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the opening price per ticker symbol
  ' Initialize it to the opening day value for the first ticker symbol
  Dim Opening_Price As Double
  Opening_Price = ws.Cells(2, "C").Value

  ' Set an initial variable for holding the total volume per ticker symbol
  ' Initialize it to the opening day volume for the first ticker symbol
  Dim Total_Volume As Double
  Total_Volume = ws.Cells(2, "G").Value

  ' Determine the Last Row in the Input Table so we know when to end the loop
  Dim LastRow As Long
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ' Output Check Range("I" & "2").Value = LastRow


  ' --------------------------------------------
  ' Prepare for the output to the Summary table
  ' --------------------------------------------

  ' Set up variable for outputting each ticker symbol summary line in the Summary Table
  ' This will be used to increment our row as we output each line
  ' Note: the first row is headers, so we start printing in row 2
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Output Headers for the Summary table
  ws.Range("I" & "1").Value = "Ticker"
  ws.Range("J" & "1").Value = "Yearly Change"
  ws.Range("K" & "1").Value = "Percent Change"
  ws.Range("L" & "1").Value = "Total Stock Volume"
  ws.Range("I1:L1").Interior.ColorIndex = 15
  ws.Range("I1:L1").Font.FontStyle = "Bold"
  ws.Range("I1:L1").HorizontalAlignment = xlCenter
  
  ' Put some boxes around the header cells and format them
  Dim rng As Range
    
    ' Define range
    Set rng = ws.Range("I1:L1")

    With rng.Borders
      .LineStyle = xlContinuous
      .Weight = xlThick
      .ColorIndex = 0
      .TintAndShade = 0
    End With
  

  ' -----------------------------------------------------
  ' --- Main Code for VBA Homework Primary Assignment ---
  ' -----------------------------------------------------

  ' Loop through all Ticker Symbols
  For i = 2 To LastRow
  
    ' Check if we are still within the same Ticker Symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Time to Output since the next Line is a new Ticker symbol
      
      ' Set the Ticker Symbol to the current line
      ' Note: this could be removed, but I had this as a trace variable when debugging
      Ticker_Symbol = ws.Cells(i, 1).Value

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
      
      ' Print the yearly change in price in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = ws.Cells(i, "F").Value - Opening_Price
      
      ' Print the yearly percentage change in price in the Summary Table and format as a percentage
      ' But Need to be sure it was not "0" to start to avoid a division by zero
      If Opening_Price <> 0 Then
        ws.Range("K" & Summary_Table_Row).Value = (ws.Cells(i, "F").Value / Opening_Price) - 1
        Else: ws.Range("K" & Summary_Table_Row).Value = "0"
      End If
             
      ' Format the Percentage
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
      ' Print the Total Volume in the Summary Table and format it
      ws.Range("L" & Summary_Table_Row).Value = Total_Volume
      ws.Range("L" & Summary_Table_Row).NumberFormat = "0,000"
      
      ' Conditional Format the Cells based on Positive or Negative
      ' Assign color based on the price change being positive or negative
      If ws.Range("J" & Summary_Table_Row).Value < 0 Then
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
         Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
        
      ' -----------------------------------------------------
      ' --- Output Completed, Reset for the next Ticker   ---
      ' -----------------------------------------------------
        
      ' Increment Summary Table Row number so we are point to the next row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume
      Total_Volume = 0

      ' Reset the Opening Price to the new Ticker's Opening Price
      Opening_Price = ws.Cells(i + 1, "C").Value


    ' If the cell immediately following a row is the Ticker Symbol...then process
    Else

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
  ' ---------------------------
  ' --- Finish and Tidy up  ---
  ' ---------------------------
  ' Reset the columns so everything looks nice
  ws.Columns("I:L").AutoFit
  Set rng = ws.Range("I2", "L" & Summary_Table_Row - 1)

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

