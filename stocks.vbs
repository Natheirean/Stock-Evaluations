Sub stocks():

outputrow = 2
volumesum = 0
Dim greatestincreasename As String
Dim greatestdecreasename As String
Dim greatestvolumename As String

worksheetcount = ActiveWorkbook.Worksheets.Count ' find max number of worksheets

For k = 1 To worksheetcount ' loop through all worksheets
    stockname = Worksheets(k).Cells(2, 1).Value 'initial sheet stock ticker
    stockstart = Worksheets(k).Cells(2, 6).Value ' initial sheet stock start
    endrow = Worksheets(k).Cells(Rows.Count, 2).End(xlUp).Row ' end row of sheet
    ' Create headers for output
    Worksheets(k).Cells(1, 9) = "Ticker"
    Worksheets(k).Cells(1, 10) = "Yearly Change"
    Worksheets(k).Cells(1, 11) = "Percent Change"
    Worksheets(k).Cells(1, 12) = "Total Stock Volume"
    Worksheets(k).Cells(1, 15) = "Ticker"
    Worksheets(k).Cells(1, 16) = "Value"
    Worksheets(k).Cells(2, 14) = "Greatest % Increase"
    Worksheets(k).Cells(3, 14) = "Greatest % Decrease"
    Worksheets(k).Cells(4, 14) = "Greatest Total Volume"

    greatestincrease = 0
    greatestdecrease = 0
    greatestvolume = 0

    For i = 2 To endrow ' loop from row 2 to end of data
    ' if stock name matches
        If Worksheets(k).Cells(i, 1) = stockname Then
            volumesum = volumesum + Worksheets(k).Cells(i, 7).Value
        ' if stock name does not match
        Else
            stockend = Worksheets(k).Cells(i - 1, 6).Value
            Worksheets(k).Cells(outputrow, 9).Value = stockname
            Worksheets(k).Cells(outputrow, 10).Value = (stockend - stockstart)
            If Worksheets(k).Cells(outputrow, 10).Value > 0 Then
                Worksheets(k).Cells(outputrow, 10).Interior.ColorIndex = 4 'set color to green if positive change
            ElseIf Worksheets(k).Cells(outputrow, 10).Value < 0 Then
                Worksheets(k).Cells(outputrow, 10).Interior.ColorIndex = 3 'set color to red if negative change
            End If
                       
            If stockstart = 0 Then ' safety against stock starting value of 0
                Worksheets(k).Cells(outputrow, 11).Value = "N/A"
            Else
                Worksheets(k).Cells(outputrow, 11).Value = ((stockend - stockstart) / stockstart)
            End If
            
            Worksheets(k).Cells(outputrow, 12).Value = volumesum
            
            ' check if output is new greatest value
            If Worksheets(k).Cells(outputrow, 11).Value > greatestincrease And Worksheets(k).Cells(outputrow, 11).Value <> "N/A" Then
                greatestincrease = Worksheets(k).Cells(outputrow, 11).Value
                greatestincreasename = stockname
            Else
            End If
            
            If Worksheets(k).Cells(outputrow, 11).Value < greatestdecrease Then
                greatestdecrease = Worksheets(k).Cells(outputrow, 11).Value
                greatestdecreasename = stockname
            Else
            End If
            
            If volumesum > greatestvolume Then
                greatestvolume = volumesum
                greatestvolumename = stockname
            Else
            End If
            
            stockstart = Worksheets(k).Cells(i, 6).Value
            stockname = Worksheets(k).Cells(i, 1).Value
            volumesum = 0
            outputrow = outputrow + 1
            
        End If
        
    Next
    
    ' output data from final stock on each sheet
    Worksheets(k).Cells(outputrow, 9).Value = stockname
    Worksheets(k).Cells(outputrow, 10).Value = (stockend - stockstart)
    
    If Worksheets(k).Cells(outputrow, 10).Value > 0 Then
        Worksheets(k).Cells(outputrow, 10).Interior.ColorIndex = 4
    ElseIf Worksheets(k).Cells(outputrow, 10).Value < 0 Then
        Worksheets(k).Cells(outputrow, 10).Interior.ColorIndex = 3
    End If
    
    If stockstart = 0 Then ' safety against stock starting value of 0
        Worksheets(k).Cells(outputrow, 11).Value = "N/A"
    Else
        Worksheets(k).Cells(outputrow, 11).Value = ((stockend - stockstart) / stockstart)
    End If
    
    Worksheets(k).Cells(outputrow, 12).Value = volumesum
    
    ' check if final sheet output is new greatest value
    If Worksheets(k).Cells(outputrow, 11).Value > greatestincrease And Worksheets(k).Cells(outputrow, 11).Value <> "N/A" Then
        greatestincrease = Worksheets(k).Cells(outputrow, 11).Value
        greatestincreasename = stockname
    Else
    End If
            
    If Worksheets(k).Cells(outputrow, 11).Value < greatestdecrease Then
        greatestdecrease = Worksheets(k).Cells(outputrow, 11).Value
        greatestdecreasename = stockname
    Else
    End If
            
    If volumesum > greatestvolume Then
        greatestvolume = volumesum
        greatestvolumename = stockname
    Else
    End If
        
    outputrow = 2
    volumesum = 0
    
    ' output greatest changes/volume
    Worksheets(k).Cells(2, 15).Value = greatestincreasename
    Worksheets(k).Cells(2, 16).Value = greatestincrease
    Worksheets(k).Cells(3, 15).Value = greatestdecreasename
    Worksheets(k).Cells(3, 16).Value = greatestdecrease
    Worksheets(k).Cells(4, 15).Value = greatestvolumename
    Worksheets(k).Cells(4, 16).Value = greatestvolume
    
    
Next

End Sub


