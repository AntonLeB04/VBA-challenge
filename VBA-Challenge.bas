Attribute VB_Name = "Module1"
Sub Stock()

  Dim Ticker As String
  
  Cells(1, 9).Value = "Ticker"
  
  Cells(1, 10).Value = "Yearly Change"
  
  Cells(1, 11).Value = "Percent Change"
  
  Cells(1, 12).Value = "Total Stock Volume"

  Cells(1, 16).Value = "Ticker"
  
  Cells(1, 17).Value = "Value"
  
  Cells(2, 15).Value = "Greatest % Increase"
  
  Cells(3, 15).Value = "Greatest % Decrease"
  
  Cells(4, 15).Value = "Greatest Total Volume"

  Dim TotalVolume As Double
  
  TotalVolume = 0
  
  Dim PriceOpen As Double
  
  Dim PriceClose As Double
  
  Dim YearChange As Double
  
  Dim PercentChange As Double

  Dim Summary_Table_Row As Integer
  
  Dim lastrow As Long
  
  Min = 1E+16
  Max = -1E+15
  Max2 = -10000000000000#
  rw = 2
  
  Dim k As Integer
  
  Summary_Table_Row = 2
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  
For i = 2 To lastrow

    
    If Cells(i + 1, 1) <> Cells(i, 1) Then
      
      Ticker = Cells(i, 1).Value
      
      PriceOpen = Cells(i - 250, 3).Value
      
      PriceClose = Cells(i, 6).Value
      
      YearChange = PriceClose - PriceOpen
    
      PercentChange = ((PriceClose - PriceOpen) / PriceClose)
    
      TotalVolume = TotalVolume + Cells(i, 7).Value
      
      Range("I" & Summary_Table_Row).Value = Ticker
      
      Range("J" & Summary_Table_Row).Value = YearChange
      
      Range("K" & Summary_Table_Row).Value = Format(PercentChange, "00.0" + "%")

      Range("L" & Summary_Table_Row).Value = TotalVolume

      Summary_Table_Row = Summary_Table_Row + 1
      
      TotalVolume = 0
      
      PriceOpen = 0
      
      PriceClose = 0
    
    Else

      TotalVolume = TotalVolume + Cells(i, 7).Value

    End If

  Next i
  

For i = 2 To lastrow
    If Max2 < Cells(i, 12) Then
        Max2 = Cells(i, 12).Value
        Cells(4, 16).Value = Cells(i, 9).Value
        
        Cells(4, 17).Value = Max2
    End If
Next i
    
    For i = 2 To lastrow
        If Cells(i, 10).Value > 0 Then
        
        Cells(i, 10).Interior.ColorIndex = 4
       
        ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        
        ElseIf Cells(i, 10).Value = "" Then
        Cells(i, 10).Interior.ColorIndex = 2
        
        ElseIf Cells(i, 10).Value = 0 Then
        Cells(i, 10).Interior.ColorIndex = 6
        
        Else

        
    End If
Next i
    
Do While Cells(rw, 11) <> ""
    If Min > Cells(rw, 11) Then
        Min = Cells(rw, 11).Value
        Cells(3, 16).Value = Cells(rw, 9).Value
    End If
    
    If Max < Cells(rw, 11) Then
        Max = Cells(rw, 11).Value
        Cells(2, 16).Value = Cells(rw, 9).Value
    End If
    
    rw = rw + 1
    
    Cells(2, 17) = Max
    Cells(2, 17).Value = Format(Max, "00.0" + "%")
    
    Cells(3, 17) = Min
    Cells(3, 17).Value = Format(Min, "00.0" + "%")
    
Loop


End Sub
