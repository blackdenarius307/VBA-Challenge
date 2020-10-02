Attribute VB_Name = "Module1"
Sub LetsTryThis()

'Define All Variables

Dim Length As Double
Dim TickerSymbol As String
Dim Stocktotal As Double
Dim SummaryRows As Integer
Dim ws As Integer
Dim OpenVal As Double
Dim Closeval As Double
Dim Change As Double
Dim PercentChange As Double
Dim Challenge As Double

'Set Worksheet Variable so it can run
    
ws = Application.WorkSheets.Count

  
' The Sheet Loop
  For i = 1 To ws
    WorkSheets(i).Activate

'Set Variables and rows for each iterated Sheet that have to be hard set
    Length = Range("A2", Range("A2").End(xlDown)).Rows.Count + 1
    Stocktotal = 0
    SummaryRows = 2
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    OpenVal = Range("C2").Value
    Closeval = 0
    Change = 0
    PercentChange = 0
    
    
'Inner Sheet Loop
    For j = 2 To Length

    ' Same Ticker Symbol? No? Then...
    If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
      
        'Account for any division by 0
        
        If OpenVal = 0 Then
            TickerSymbol = Cells(j, 1).Value
            
             'Print the Ticker in the Table
            Range("I" & SummaryRows).Value = TickerSymbol

            ' Print the Stock Total to the Table
            Range("L" & SummaryRows).Value = 0
        
            'Print the Change
            Range("J" & SummaryRows).Value = 0
            
            'Print Percent Change
            Range("K" & SummaryRows).Value = 0
        
            ' Add one to the Table
            SummaryRows = SummaryRows + 1
      
            ' Reset the Volume total
            Stocktotal = 0
        
            'Set Next Open Value
            OpenVal = Cells(j + 1, 3).Value
        
        
        Else:
        ' Set the Ticker
        TickerSymbol = Cells(j, 1).Value
        
        ' Set The CloseValue
        Closeval = Cells(j, 6).Value

        ' Add to the Stock Volume
        Stocktotal = Stocktotal + Cells(j, 7).Value
        
        'Subtract Open from Close
        Change = Closeval - OpenVal
        
        'Find Percent Change
        PercentChange = Change / OpenVal * 100
        
        ' Print the Ticker in the Table
        Range("I" & SummaryRows).Value = TickerSymbol

        ' Print the Stock Total to the Table
        Range("L" & SummaryRows).Value = Stocktotal
        
        'Print the Change
        Range("J" & SummaryRows).Value = Change
            
            'Positive or Negative? Color Appropriately
            If Range("J" & SummaryRows).Value > 0 Then
                Range("J" & SummaryRows).Interior.ColorIndex = 4
            Else
                Range("J" & SummaryRows).Interior.ColorIndex = 3
            End If
            
        'Print Percent Change
        Range("K" & SummaryRows).Value = PercentChange & "%"

        ' Add one to the Table
        SummaryRows = SummaryRows + 1
      
        ' Reset the Volume total
        Stocktotal = 0
        
        'Set Next Open Value
        OpenVal = Cells(j + 1, 3).Value
        
        End If
        

    ' Same Ticker? Then
    Else

      ' Add to Volume Total
      Stocktotal = Stocktotal + Cells(j, 7).Value

    End If

  'Challenge iteration definitions.
  Next j
        Challenge = Range("I2", Range("I2").End(xlDown)).Rows.Count + 1
        LargesttickerValue = 0
        Largestticker = 0
        SmallesttickerValue = 0
        Smallestticker = 0
        VolumeValue = 0
        Volumeticker = 0
    
    For k = 2 To Challenge
        'Comparison
        If Cells(k, 11) > LargesttickerValue Then
            LargesttickerValue = Cells(k, 11).Value
            Largestticker = Cells(k, 9)
            Range("P2") = Largestticker
            Range("Q2") = LargesttickerValue * 100
        Else
            Range("P2") = Largestticker
            Range("Q2") = FormatPercent(LargesttickerValue, 7)
        End If
        
    Next k
    
    For l = 2 To Challenge
        'Comparison
        If Cells(l, 11) < SmallesttickerValue Then
            SmallesttickerValue = Cells(l, 11).Value
            Smallestticker = Cells(l, 9)
            Range("P3") = Smallestticker
            Range("Q3") = SmallesttickerValue * 100
        Else
            Range("P3") = Smallestticker
            Range("Q3") = FormatPercent(SmallesttickerValue, 7)
        End If
        
    Next l
    
    For m = 2 To Challenge
        'Comparison
        If Cells(m, 12) > VolumeValue Then
            VolumeValue = Cells(m, 12).Value
            Volumeticker = Cells(m, 9)
            Range("P4") = Volumeticker
            Range("Q4") = VolumeValue
        Else
            Range("P4") = Volumeticker
            Range("Q4") = VolumeValue
        End If
        
    Next m

Next i

End Sub


