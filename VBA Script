Sub Worksheet_loop()

    
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Module2
    Next
    Application.ScreenUpdating = True
    
End Sub


Sub Module2()

    Dim Ticker As String
    Dim Stock_Volume As LongPtr
    Stock_Volume = 0
    Dim Summary_table_row As Integer
    Summary_table_row = 2
    Dim Open_Price As Double
    Open_Price = Cells(2, 3).Value
    Dim Closing_Price As Double
    Dim Year_Change As Double
    Dim Percent_Change As Double
    Dim Yearly_Change_CF As Range
    Dim Lastrow As Long
    Dim ws As Worksheet
    
    
    
    Lastrow = ActiveSheet.Range("a1").CurrentRegion.Rows.Count
    
    Set Yearly_Change_CF = Range("k2:K91")
    
    Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Stock Volume")
    
    
    For i = 2 To Lastrow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
        
        Range("I" & Summary_table_row).Value = Ticker
        
        Range("L" & Summary_table_row).Value = Stock_Volume
        
        Closing_Price = Cells(i, 6).Value
        
        Year_Change = (Closing_Price - Open_Price)
        
        Range("J" & Summary_table_row).Value = Year_Change
            
            If (Open_Price = 0) Then
                
                Percent_Change = 0
                
            Else
            
                Percent_Change = Year_Change / Open_Price
                
            End If
            
        Range("K" & Summary_table_row).Value = Percent_Change
        Range("K" & Summary_table_row).NumberFormat = "0.00%"
        
        Summary_table_row = Summary_table_row + 1
        
        Ticker = 0
        
        Open_Price = Cells(i + 1, 3)
        
    Else
    
        Stock_Volume = Stock_Volume + Cells(i, 7).Value
        
    End If
    
Next i

last_row_table = ActiveSheet.Range("j1").CurrentRegion.Rows.Count

For i = 2 To last_row_table

    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    
    Else: Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    
    End If
    
Next i
    
End Sub


