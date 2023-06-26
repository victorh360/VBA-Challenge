Attribute VB_Name = "Module1"
Sub MainLoop()

    For Each Worksheet In ThisWorkbook.Sheets
        
        Worksheet.Activate
        
        Call TickerSummary
        
        Call Greatest_Increase_Decrease
    
    Next Worksheet
    
End Sub

Sub TickerSummary()

    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    
    Columns("A:Q").AutoFit
    
    Dim Ticker As String

    Dim Total_Volume As Double
    Total_Volume = 0
    
    Dim Opening_Value As Double
        
    Dim Closing_Value As Double
    
    Dim Percent_Change As Double
    
    Dim Yearly_Change As Double
                
    Dim Table_Row As Integer
    Table_Row = 2
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
        Opening_Value = Cells(i, 3).Value
    
    End If
        
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        
        Closing_Value = Cells(i, 6).Value
        
        Yearly_Change = Closing_Value - Opening_Value
        
        Percent_Change = (Closing_Value - Opening_Value) / Opening_Value
    
        
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
        Range("I" & Table_Row).Value = Ticker
        
        Range("J" & Table_Row).Value = Yearly_Change
            
            
            If Yearly_Change > 0 Then
                
                Range("J" & Table_Row).Interior.ColorIndex = 4
            
            Else
                
                Range("J" & Table_Row).Interior.ColorIndex = 3
            
            End If
        
        Range("K" & Table_Row).Value = FormatPercent(Percent_Change)
        
        Range("L" & Table_Row).Value = Total_Volume
        
        Table_Row = Table_Row + 1
        
        Total_Volume = 0
        
        Closing_Value = 0
                
                    
    Else
    
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
        
    End If
    
    Next i
    
    
    

End Sub

Sub Greatest_Increase_Decrease()

    [O2] = "Greatest Percent Increase"
    [O3] = "Greatest Percent Decrease"
    [O4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [Q1] = "Value"

    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
    
    Dim Greatest_Volume As Double
    Greatest_Volume = 0
        
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
    
        If Cells(i, 11).Value > Greatest_Increase Then
               
            Greatest_Increase = Cells(i, 11).Value
        
            Range("P2").Value = Cells(i, 9).Value
                
            Range("Q2").Value = FormatPercent(Greatest_Increase)
                     
        End If

        If Cells(i, 11).Value < Greatest_Decrease Then
               
            Greatest_Decrease = Cells(i, 11).Value
        
            Range("P3").Value = Cells(i, 9).Value
                
            Range("Q3").Value = FormatPercent(Greatest_Decrease)
                     
        End If
        
        If Cells(i, 12).Value > Greatest_Volume Then
               
            Greatest_Volume = Cells(i, 12).Value
        
            Range("P4").Value = Cells(i, 9).Value
                
            Range("Q4").Value = Greatest_Volume
                     
        End If
    
    Next i
          

End Sub


