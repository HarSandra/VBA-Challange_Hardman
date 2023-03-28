Sub ABC_Testing()
        
        For Each sm In Worksheets
        
            Dim WorksheetName As String
            Dim i As Long
            Dim j As Long
            Dim TickCount As Long
            Dim LastRowA As Long
            Dim LastRowl As Long
            Dim PerChange As Double
            Dim Great_Incr As Double
            Dim Great_Decr As Double
            Dim GreatVol As Double
            
            WorksheetName = sm.Name
            
            sm.Cells(1, 9).Value = "Ticker"
            sm.Cells(1, 10).Value = "Yearly Change"
            sm.Cells(1, 11).Value = "Percent Change"
            sm.Cells(1, 12).Value = "Total Stock Volume"
            sm.Cells(2, 15).Value = "Greatest % Increase"
            sm.Cells(3, 15).Value = "Greatest % Decrease"
            sm.Cells(4, 15).Value = "Greatest Total Volume"
            sm.Cells(1, 16).Value = "Ticker"
            sm.Cells(1, 17).Value = "Value"
            
            TickCount = 2
            
            j = 2
            
            LastRowA = sm.Cells(Rows.Count, 1).End(xlUp).Row
            
                For i = 2 To LastRowA
                
                    If sm.Cells(i + 1, 1).Value <> sm.Cells(i, 1).Value Then
                    
                    sm.Cells(TickCount, 9).Value = sm.Cells(i, 1).Value
                    
                    sm.Cells(TickCount, 10).Value = sm.Cells(i, 6).Value - sm.Cells(j, 3).Value
                    
                        If sm.Cells(TickCount, 10).Value < 0 Then
                        
                        sm.Cells(TickCount, 10).Interior.Color = RGB(238, 75, 43)
                        
                        Else
                        
                        sm.Cells(TickCount, 10).Interior.Color = RGB(57, 255, 20)
                        
                        End If
                        
                        If sm.Cells(j, 3).Value <> 0 Then
                        PerChange = ((sm.Cells(i, 6).Value - sm.Cells(j, 3).Value) / sm.Cells(j, 3).Value)
                        
                        sm.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                        
                        Else
                        
                        sm.Cells(TickCount, 11).Value = Format(0, "Percent")
                        
                        End If
                        
                        sm.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(sm.Cells(j, 7), sm.Cells(i, 7)))
                        
                        TickCount = TickCount + 1
                        
                        j = i + 1
                        
                        End If
                        
                    Next i
                    
                LastRowl = sm.Cells(Rows.Count, 9).End(xlUp).Row
                
                GreatVol = sm.Cells(2, 12).Value
                Great_Incr = sm.Cells(2, 11).Value
                Great_Decr = sm.Cells(2, 11).Value
                
                For i = 2 To LastRowl
                
                    If sm.Cells(i, 12).Value > GreatVol Then
                    GreatVol = sm.Cells(i, 12).Value
                    sm.Cells(4, 16).Value = sm.Cells(i, 9).Value
                    
                    Else
                    
                    GreatVol = GreatVol
                    
                    End If
                    
                    If sm.Cells(i, 11).Value > Great_Incr Then
                    Great_Incr = sm.Cells(i, 11).Value
                    sm.Cells(2, 16).Value = sm.Cells(i, 9).Value
                    
                    Else
                    
                    Great_Incr = Great_Incr
                    
                    End If
                    
                    If sm.Cells(i, 11).Value < Great_Decr Then
                    Great_Decr = sm.Cells(i, 11).Value
                    sm.Cells(3, 16).Value = sm.Cells(i, 9).Value
                    
                    Else
                    
                    Great_Decr = Great_Decr
                    
                    End If
                    
                sm.Cells(2, 17).Value = Format(Great_Incr, "Percent")
                sm.Cells(3, 17).Value = Format(Great_Decr, "Percent")
                sm.Cells(4, 17).Value = Format(GreatVol, "Scientific")
                
                Next i
                
            Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
        Next sm
        
End Sub

