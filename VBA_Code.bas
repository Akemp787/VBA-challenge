Attribute VB_Name = "Module1"
Sub StockAnalysis()

 For Each Ws In Worksheets
        
    ' Defining variables
    Dim WorksheetName As String
    Dim i As Long
    Dim j As Long
    Dim Stockticker As Long
    Dim Percentchng As Double
    Dim Grtincrease As Double
    Dim Grtdecrease As Double
    Dim Grtvolume As Double
    Dim Lastrow_A As Long
    Dim Lastrow_I As Long
    
    
    ' Creating column names
    Ws.Range("I1").Value = "Ticker"
    Ws.Range("J1").Value = "Yearly Change"
    Ws.Range("K1").Value = "Percent Change"
    Ws.Range("L1").Value = "Total Stock Volume"
    Ws.Range("P1").Value = "Ticker"
    Ws.Range("Q1").Value = "Value"
    Ws.Range("O2").Value = "Greatest % Increase"
    Ws.Range("O3").Value = "Greatest % Decrease"
    Ws.Range("O4").Value = "Greatest Total Volume"

    'Setting Last row of colum A
    Lastrow_A = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    'Setting variables
    Stockticker = 2
    j = 2
    
   'Loop through all rows
    For i = 2 To Lastrow_A
            
        'Ticker name change in column I
        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
        Ws.Cells(Stockticker, 9).Value = Ws.Cells(i, 1).Value
                
        'Yearly Change in column J
        Ws.Cells(Stockticker, 10).Value = Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value
                
        'Setting colors for yearly change: Red and Green
        If Ws.Cells(Stockticker, 10).Value < 0 Then
        Ws.Cells(Stockticker, 10).Interior.ColorIndex = 3
            
        Else: Ws.Cells(Stockticker, 10).Interior.ColorIndex = 4
                
        End If
                    
        ' Creating percent change formulas
        If Ws.Cells(j, 3).Value <> 0 Then
        Percentchng = ((Ws.Cells(i, 6).Value - Ws.Cells(j, 3).Value) / Ws.Cells(j, 3).Value)
        Ws.Cells(Stockticker, 11).Value = Format(Percentchng, "Percent")
                    
        Else: Ws.Cells(Stockticker, 11).Value = Format(0, "Percent")
                    
        End If
                               
        
        ' Creating total volume formulas
        Ws.Cells(Stockticker, 12).Value = WorksheetFunction.Sum(Range(Ws.Cells(j, 7), Ws.Cells(i, 7)))
        Stockticker = Stockticker + 1
        j = i + 1
            
        End If
        
    Next i
            
        'Setting the Last row for column I
        Lastrow_I = Ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Set values for greatest increase, decrease, & volume
        Grtincrease = Ws.Cells(2, 11).Value
        Grtdecrease = Ws.Cells(2, 11).Value
        Grtvolume = Ws.Cells(2, 12).Value
        
        'Creating a For Loop to go through the data evaluating greatest volume, increase, and decrease
        For i = 2 To Lastrow_I
            
                
            'Calculation for greatest increase
            If Ws.Cells(i, 11).Value > Grtincrease Then
            Grtincrease = Ws.Cells(i, 11).Value
            Ws.Cells(2, 16).Value = Ws.Cells(i, 9).Value
                
            Else
            Grtincrease = Grtincrease
                
            End If
                
            'Calculation for greatest decrease
            If Ws.Cells(i, 11).Value < Grtdecrease Then
            Grtdecrease = Ws.Cells(i, 11).Value
            Ws.Cells(3, 16).Value = Ws.Cells(i, 9).Value
                
            Else
            Grtdecrease = Grtdecrease
                
            End If
            
            
            'Calculation for greatest volume
            If Ws.Cells(i, 12).Value > Grtvolume Then
            Grtvolume = Ws.Cells(i, 12).Value
            Ws.Cells(4, 16).Value = Ws.Cells(i, 9).Value
                
            Else
            Grtvolume = Grtvolume
                
            End If
        
                
            'Place results in appropriate cells
            Ws.Cells(2, 17).Value = Format(Grtincrease, "Percent")
            Ws.Cells(3, 17).Value = Format(Grtdecrease, "Percent")
            Ws.Cells(4, 17).Value = Format(Grtvolume, "Scientific")
            
        Next i
    
     
Next Ws

End Sub
