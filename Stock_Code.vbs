  Sub Stock_Moderate()

    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
    
        ws.Range("J1").Value = "Yearly Change"
        
        ws.Range("K1").Value = "Percent Change"
        
       ws.Range("L1").Value = "Total Stock Volume"
    
    Next ws
    
    For Each ws In Worksheets

  ' Set initial variables
    Dim Ticker As String
    Dim Yearly_Change As Double
        Yearly_Change = 0
    Dim Percent_Change As Double
        Percent_Change = 0
    Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    Dim opening_price As Double
        opening_price = ws.Cells(2, 3).Value
    Dim closing_price As Double
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
  
  For i = 2 To lastrow

    ' Check if we are still within the same Ticker symbol, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        Ticker = ws.Cells(i, 1).Value
    
        closing_price = ws.Cells(i, 6).Value
        
        Yearly_Change = closing_price - opening_price
        
        
            If (opening_price = 0 And closing_price = 0) Then
                Percent_Change = 0
                
                ElseIf (opening_price = 0 And closing_price <> 0) Then
                    Percent_Change = 100
                    
                Else: Percent_Change = (Yearly_Change / opening_price)
                ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
                
            End If
                

        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        'Assign Value locations in columns
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        'Reset
        Summary_Table_Row = Summary_Table_Row + 1
        
        Total_Stock_Volume = 0
        
        opening_price = ws.Cells(i + 1, 3).Value
        
        Else
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
    End If

Next i

   ' Determine the Last Row of Yearly Change per WS
        Color_Format = ws.Cells(Rows.Count, 10).End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To Color_Format
            If (ws.Cells(j, 10).Value >= 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
    Next j
    
  Next ws
  
  Dim wrksht As Worksheet
    For Each wrksht In Worksheets
       
    wrksht.Select
    Cells.EntireColumn.AutoFit
       
    Next wrksht

End Sub