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

  Next j

Next i

End Sub


