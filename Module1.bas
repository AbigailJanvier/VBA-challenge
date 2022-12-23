Attribute VB_Name = "Module1"

Sub InsertData():

Dim i, lastrow As Long
Dim ws As Worksheet

Call TurnOffStuff

    For Each ws In ThisWorkbook.Worksheets


    'find the last row of data for the For loop
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'titles for the new columns
    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Yearly Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"
    
    
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    ws.Cells(2, 14).Value = " Greatest % Increase"
    ws.Cells(3, 14).Value = " Greatest % Decrease"
    ws.Cells(4, 14).Value = " Greatest Total Volume"
    
    ' populate the new columns starting from row 2
    Count = 1

    ' get the maximum percent change
    max_per = 0
    
    ' get minimum percent change
    min_per = 0
    
    'get maximum volume
    max_vol = 0
    
    'input the new data
        For i = 2 To lastrow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            Count = Count + 1
            
            'ticker symbol
            ws.Cells(Count, 8).Value = ws.Cells(i, 1).Value
            
            'yearly change
            Top = ws.Range("A:A").Find(What:=ws.Cells(Count, 8).Value).Row
            
            ws.Cells(Count, 9).Value = ws.Cells(i, 6).Value - ws.Cells(Top, 3).Value
            
                If ws.Cells(Count, 9).Value < 0 Then
                
                    ws.Cells(Count, 9).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Cells(Count, 9).Interior.ColorIndex = 4
                    
                End If
                
               
            'percent change
            
            ws.Cells(Count, 10).Value = FormatPercent((ws.Cells(Count, 9).Value / ws.Cells(Top, 3).Value))
            
                    If ws.Cells(Count, 10).Value > max_per Then
                    max_per = ws.Cells(Count, 10).Value
                    ws.Cells(2, 16).Value = max_per
                    ws.Cells(2, 15).Value = ws.Cells(Count, 8).Value
                    
                    End If
                   
                        If ws.Cells(Count, 10).Value < min_per Then
                        min_per = ws.Cells(Count, 10).Value
                        ws.Cells(3, 16).Value = min_per
                        ws.Cells(3, 15).Value = ws.Cells(Count, 8).Value
                   
                        End If
                   
            'total stock volume
            ws.Cells(Count, 11).Value = "=SUM(" & ws.Range(ws.Cells(Top, 7), ws.Cells(i, 7)).Address(False, False) & ")"
                
                             If ws.Cells(Count, 11).Value > max_vol Then
                             max_vol = ws.Cells(Count, 11).Value
                             ws.Cells(4, 16).Value = max_vol
                             ws.Cells(4, 15).Value = ws.Cells(Count, 8).Value
                             End If
                        
            End If
        Next i
    
    Next ws
    
    Call TurnOnStuff
        
End Sub



Sub TurnOffStuff():

    Application.Calculation = xlCalculationManual
    
End Sub



Sub TurnOnStuff():

    Application.Calculation = xlCalculationAutomatic
    
End Sub


