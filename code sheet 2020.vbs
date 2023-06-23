# VBA-challenge

# This is the Version 1.0 of my challenge it includes the code for the first sheet the 2020.

Sub stockmkt()

    Dim ticker As String
    Dim ws_num As String
    Dim volume_total As Double
    Dim percent_change, yearly_change As Double
    
    Dim i, LR, LR2 As Long
    
    
     
    ' LR is last row, used to detect empty cell on a row
    
    Dim open_price As Double
    Dim close_price As Double
   
    
    volume_total = 0
    yearly_change = 0
    percent_change = 0
    Row = 2
    LR = 0
    
    open_price = 0
    close_price = 0
   
    
    
    
    Worksheets("2020").Cells(1, 9) = "Ticker Symbol"
    Worksheets("2020").Cells(1, 10) = "Yearly change"
    Worksheets("2020").Cells(1, 11) = "Percent change"
    Worksheets("2020").Cells(1, 12) = "Total Stock Volume"
    Worksheets("2020").Cells(1, 13) = "Close price"
    Worksheets("2020").Cells(1, 14) = "Open price"
    
    
    

' 1. Keep track on Ticker Symbol until it changes.

    LR = Range("A1").End(xlDown).Row
    
    'For i = 2 To 22771
    For i = 2 To LR
          
        If i = 2 Then
            open_price = Cells(i, 3).Value
            Range("N" & Row).Value = open_price
            yearly_change = Range("M" & i).Value - Range("N" & i).Value
            Range("J" & Row).Value = yearly_change
            volume_total = volume_total + Range("G" & i).Value

        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          
            ticker = Cells(i, 1).Value
            Range("I" & Row).Value = ticker
            
            volume_total = volume_total + Range("G" & i).Value
            Range("L" & Row).Value = volume_total
            
            close_price = Cells(i, 6).Value
            Range("M" & Row).Value = close_price
            
            open_price = Range("C" & (i + 1)).Value
            Range("N" & (Row + 1)).Value = open_price

            
           percent_change = Range("M" & Row).Value / Range("N" & Row).Value * 100 - 100
           Range("K" & Row).Value = percent_change
        
            Row = Row + 1
            volume_total = 0

        
       Else
        
           ' LR2 = Range("I1").End(xlDown).Row
            
            volume_total = volume_total + Range("G" & i).Value
            yearly_change = Range("M" & Row).Value - Range("N" & Row).Value
            Range("J" & Row).Value = yearly_change
            

                   
        End If
        
        
Next i


    Worksheets("2020").Cells(2, 16) = "Greatest % Increase"
    Worksheets("2020").Cells(3, 16) = "Greatest % Decrease"
    Worksheets("2020").Cells(4, 16) = "Greatest Total Volume"
    Worksheets("2020").Cells(1, 17) = "Ticker"
    Worksheets("2020").Cells(1, 18) = "Value"
    
    'LR2 = Range("I1").End(xlDown).Row
    
    ' a,b,c= contadores, k= rows
    a = 0
    k = 2
    While Cells(k, 11) <> ""
        If k = 2 Then
            a = Cells(k, 11)
            t1 = Cells(k, 9)
        End If
        If a < Cells(k, 11) Then
            a = Cells(k, 11)
            t1 = Cells(k, 9)
        End If
        k = k + 1
    Wend
    Cells(2, 17) = t1
    Cells(2, 18) = a
    Cells(2, 18).NumberFormat = "0.00%"
    
    b = 0
    k = 2
    While Cells(k, 11) <> ""
        If k = 2 Then
            b = Cells(k, 11)
            t2 = Cells(k, 9)
        End If
        If b > Cells(k, 11) Then
            b = Cells(k, 11)
            t2 = Cells(k, 9)
        End If
        k = k + 1
    Wend
    Cells(3, 17) = t2
    Cells(3, 18) = b
    Cells(3, 18).NumberFormat = "0.00%"
    
    c = 0
    k = 2
    While Cells(k, 12) <> ""
        If k = 2 Then
            c = Cells(k, 12)
            t3 = Cells(k, 9)
        End If
        If c < Cells(k, 12) Then
            c = Cells(k, 12)
            t3 = Cells(k, 9)
        End If
        k = k + 1
    Wend
    Cells(4, 17) = t3
    Cells(4, 18) = c
    Cells(4, 18).NumberFormat = "##0.00E+0"

    


End Sub




