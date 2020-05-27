Sub VBAStocks()

'Declare each sheet

For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        Sheets(ws.Name).Select
        
    Columns("I:Q").Select
    Selection.Clear
        
'Define variables

Dim DateMinOpen As Variant
Dim DateMaxClose As Variant
Dim i As Double
Dim j As Double

i = 2
j = 2

'Column Headers

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Volume"
    
Cells(j, 9).Value = Cells(j, 1).Value
DateMinOpen = Cells(i, 3).Value
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    'Find the unique ticker symbol & find variance
    If Cells(i, 1).Value = Cells(j, 9).Value Then
        TotalVal = TotalVal + Cells(i, 7).Value
        DateMaxClose = Cells(i, 6).Value
    Else
        Cells(j, 10).Value = DateMaxClose - DateMinOpen
        
    'Percent
    If DateMaxClose <= 0 Then
        Cells(j, 11).Value = 0
    Else
        If DateMinOpen <= 0 Then
        Cells(j, 11).Value = 0
        Else
        Cells(j, 11).Value = (DateMaxClose / DateMinOpen) - 1
        Cells(j, 11).Style = "Percent"
        End If
    End If
    
    'Conditional Formatting
    If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    End If
    Cells(j, 12).Value = TotalVal
    
    
    DateMinOpen = Cells(i, 3).Value
    TotalVal = Cells(i, 7).Value

    j = j + 1
    Cells(j, 9).Value = Cells(i, 1).Value
    
    End If
    
Next i
    Cells(j, 10).Value = DateMaxClose - DateMinOpen

    If DateMaxClose <= 0 Then
        Cells(j, 11).Value = 0
    Else
        If DateMinOpen <= 0 Then
        Cells(j, 11).Value = 0
        Else
        Cells(j, 11).Value = (DateMaxClose / DateMinOpen) - 1
        Cells(j, 11).Style = "Percent"
        End If
    End If
    
        
    If Cells(j, 10).Value >= 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    Else
        Cells(j, 10).Interior.ColorIndex = 3
    End If
             
    Cells(j, 12).Value = TotalVal
    
    'Calculations for % changes
        VolumeInc = 2
        VolumeDec = 2
        TickerInc = 2
        TickerDec = 2
                
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To LastRow
        If Cells(j, 11).Value > VolumeInc Then
            
            TickerInc = Cells(j, 9).Value
            VolumeInc = Cells(j, 11).Value
        
        End If
        
        If Cells(j, 11).Value < VolumeDec Then
            
            TickerDec = Cells(j, 9).Value
            VolumeDec = Cells(j, 11).Value
        
        End If
        If Cells(j, 12).Value > VolumeInc Then
            
            TickerTV = Cells(j, 9).Value
            VolumeTV = Cells(j, 12).Value
        
        End If
        Next j
        
Cells(2, 16).Value = TickerInc
Cells(2, 17).Value = VolumeInc
Cells(2, 17).Style = "Percent"
Cells(3, 16).Value = TickerDec
Cells(3, 17).Value = VolumeDec
Cells(3, 17).Style = "Percent"
Cells(4, 16).Value = TickerTV
Cells(4, 17).Value = VolumeTV
        
        
    
    Next ws

End Sub

