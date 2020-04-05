Sub Stocks()

Dim i As Double
        Dim LastRow As Double
        Dim Ticker As String
        Dim Summary_Table_Row As Long
        Dim Yearly_Change As Double
        Dim Total_Volume As Double
        Dim Open_Value As Double
        Dim Closing_value As Double
        Dim Percentage_Change As Double
        'Dim Rows_Summary_Table As Double
        Dim x As Integer
        Dim Max As Double
        Dim Tick As String

    ' loop thru each worksheet
   For Each ws In Worksheets
        Range("I1").Value = "ticker"
        Range("J1").Value = "yearly_change"
        Range("K1").Value = "percentage_change"
        Range("L1").Value = "total_volume"
'        Range("M1").Value = "open_value"
'        Range("N1").Value = "closed_value"
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
'        ' initializing vars
       LastRow = ws.Range("A1").End(xlDown).Row
'        
       
       Summary_Table_Row = 2
        x = 1
        ' looping thru each row 
       
        For i = 2 To LastRow
            If x = 1 Then
            
                Open_Value = Cells(i, 3).Value
                x = 0
                Total_Volume = 0
            End If
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                x = 1
                Ticker = Cells(i, 1).Value
                Closing_value = Cells(i, 6).Value
                Yearly_Change = Closing_value - Open_Value
                    If Open_Value = 0 Then
                        Percentage_Change = 0
                    Else
                Percentage_Change = Yearly_Change / Open_Value
                    End If
                Range("I" & Summary_Table_Row).Value = Ticker
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                Range("K" & Summary_Table_Row).Value = Percentage_Change
                Range("L" & Summary_Table_Row).Value = Total_Volume
'                Range("M" & Summary_Table_Row).Value = Open_Value
'                Range("N" & Summary_Table_Row).Value = Closing_value
                Summary_Table_Row = Summary_Table_Row + 1
                
            Else

                Total_Volume = Total_Volume + Cells(i + 1, 7).Value

            End If
                
            Next i
          'Finding last row of Summary Table
          Rows_Summary_Table = Range("I1").End(xlDown).Row
        'Looping to apply conditional formatting
        For j = 2 To Rows_Summary_Table
            If Cells(j, 10) >= 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10) < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

    
        Max = 0 'set "inital" max
        
        For i = 2 To Rows_Summary_Table
           If Cells(i, 11).Value > Max Then 'if a value is larger than the old max,
            Max = Cells(i, 11).Value ' store it as the new max!
            Tick = Cells(i, 9).Value
          
           
            Range("Q2").Value = Max
            Range("P2").Value = Tick
           End If
        Next i
        
        'set "inital" min
        Min = 0
           
        For i = 2 To Rows_Summary_Table
           If Cells(i, 11).Value < Min Then 'if a value is larger than the old max,
            Min = Cells(i, 11).Value ' store it as the new max!
            Tick = Cells(i, 9).Value
          
           
            Range("Q3").Value = Min
            Range("P3").Value = Tick
           End If
        Next i
        
        Max_Vol = 0 'set "inital" max Vol

        For i = 2 To Rows_Summary_Table
           If Cells(i, 12).Value > Max_Vol Then 'if a value is larger than the old max,
            Max_Vol = Cells(i, 12).Value ' store it as the new max!
            Tick = Cells(i, 9).Value
          
           
            Range("Q4").Value = Max_Vol
            Range("P4").Value = Tick
           End If
        Next i
  Next ws

    
End Sub