Sub StockProject()

    'Setting variables
    Dim i As Long
    Dim j As Integer
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim InitialOpen As Double
    Dim FinalClose As Double
    Dim Percent_Change As Double
    Dim VolumeTotal As LongLong
    Dim GreatIn As Double
    Dim GreatDec As Double
    Dim GreatVolume As Double
    Dim GreatInTick As String
    Dim GreatDecTick As String
    Dim GreatVolumeTick As String
    
    'Set range for all worksheets
    For Each ws In Worksheets

    'Set title row
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Set initial values
    j = 1
    InitialOpen = ws.Cells(2, 3).Value
    VolumeTotal = 0

    'Loop through all ticker names
    For i = 2 To LastRow
    
        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        
            'Set Volume Total Formula
            VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
            
        Else
        
            ' Set the Ticker
            Ticker = ws.Cells(i, 1).Value
          
        ' Add one to the summary table row
            j = j + 1
    
            'Set values for open and close
            
            FinalClose = ws.Cells(i, 6).Value
                
            'Set formula for yearly change
            Yearly_Change = FinalClose - InitialOpen
                
            'Set formula for percent change
            Percent_Change = (Yearly_Change / InitialOpen)
                
            InitialOpen = ws.Cells(i + 1, 3).Value
            
            'Find max & min percentage, and max volume
            GreatIn = Application.WorksheetFunction.Max(ws.Range("K:K"))
            GreatDec = Application.WorksheetFunction.Min(ws.Range("K:K"))
            GreatVolume = Application.WorksheetFunction.Max(ws.Range("L:L"))
            
                
            ' Print results
            ws.Range("I" & j).Value = Ticker
            ws.Range("J" & j).Value = Yearly_Change
            ws.Range("K" & j).Value = Percent_Change
            ws.Range("K" & j).NumberFormat = "0.00%"
            ws.Range("L" & j).Value = VolumeTotal
            VolumeTotal = 0
            ws.Range("Q" & 2).Value = GreatIn
            ws.Range("Q" & 2).NumberFormat = "0.00%"
            ws.Range("Q" & 3).Value = GreatDec
            ws.Range("Q" & 3).NumberFormat = "0.00%"
            ws.Range("Q" & 4).Value = GreatVolume
              
        End If
        
        'Set parameters for green color
        If Cells(i, 10).Value >= 0 Then

            ' Color the positive green
            Cells(i, 10).Interior.ColorIndex = 4
        
        Else
            'Color the negatives red
            Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    
    Next i
    
    Next ws

End Sub

