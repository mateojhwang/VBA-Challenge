Attribute VB_Name = "Module1"
Sub stocks():

Dim count As Integer
Dim total_volume As LongLong
Dim yearly_change As Double
Dim close_num As Double
Dim open_num As Double

row_num = Cells(Rows.count, "A").End(xlUp).Row

count = 2
total_volume = 0

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"


For i = 2 To row_num
    If Cells(i, 1) <> Cells(i + 1, 1) Then
        Cells(count, 9) = Cells(i + 1, 1)
        
        'total stock volume
        
        total_volume = total_volume + Cells(i, 7)
        Cells(count, 11) = total_volume
        total_volume = 0
    
        'yearly change
    
        open_num = Cells(i + 1, 3)
        close_num = Cells(i + 1, 6)
    
        yearly_change = close_num - open_num
        Cells(count, 10) = yearly_change
            If yearly_change < 0 Then
                Cells(count, 10).Interior.ColorIndex = 3
            Else
                Cells(count, 10).Interior.ColorIndex = 4
            End If
        
    
        'percent change
    
        Cells(count, 11) = (yearly_change - 1 / open_num) * 100
        
        'counter
        count = count + 1
    
        'total stock volume
    
    ElseIf Cells(i, 1) = Cells(i + 1, 1) Then
        total_volume = total_volume + Cells(i, 7)
        
    End If
    
Next i

biggest_increase = 0
biggest_decrease = 0

For i = 2 To 500

    If biggest_increase < Cells(i, 11) Then
        biggest_increase = Cells(i, 11)
        Range("Q2") = Cells(i, 11)
        Range("P2") = Cells(i, 9)
        
    ElseIf biggest_decrease > Cells(i, 11) Then
        biggest_decrease = Cells(i, 11)
        Range("Q3") = Cells(i, 11)
        Range("P3") = Cells(i, 9)
        
    End If
    
Next i

biggest_volume = 0

For i = 2 To 500

    If biggest_volume < Cells(i, 12) Then
        biggest_volume = Cells(i, 12)
        Range("Q4") = Cells(i, 12)
        Range("P4") = Cells(i, 9)
        
    End If
    
Next i


End Sub

