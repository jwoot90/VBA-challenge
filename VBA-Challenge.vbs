   Total_Stock_Value = Total_Stock_Value + ws.Cells(i, 7).Value
            
          
            
            Total_Close_Value = ws.Cells(i, 6).Value
            
            Yearly_Change = Total_Open_Value - Total_Close_Value
            
            Yearly_Percentage = (Total_Open_Value - Total_Close_Value) / Total_Open_Value
            ws.Range("I" & Summary_Table_Row).Value = Symbol_Name
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = Yearly_Percentage
            ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Value
            ws.Range("K:K").NumberFormat = "00.00%"
            
            Select Case Yearly_Change
                Case Is > 0
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                 Case Is < 0
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                     Case Else
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
            End Select
            'Update Variables to go to next symbol
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            Total_Stock_Value = 0
            
            Symbol_Name = ws.Cells(i + 1, 1).Value
            
            Total_Open_Value = ws.Cells(i + 1, 3).Value
            Yearly_Change = 0
            Yearly_Percentage = 0
            
        Else
             'Add to the Total Stock Value
             Total_Stock_Value = Total_Stock_Value + ws.Cells(i, 7).Value
        End If
                                                        

    Next i
    'worksheetfunction.min()
    ws.Range("Q2") = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("Q2").NumberFormat = "0.00%"
    
    ws.Range("Q3") = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
    'worksheetfunction.match()copy content of first function into this function
    'find columns in summary table
    Ticker_1_Index = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K").Value, 0)
    Ticker_2_Index = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K").Value, 0)
    Ticker_3_Index = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L").Value, 0)
    
    
    
    ws.Range("P2") = WorksheetFunction.Index(ws.Range("I:I"), Ticker_1_Index)
    ws.Range("P3") = WorksheetFunction.Index(ws.Range("I:I"), Ticker_2_Index)
    ws.Range("P4") = WorksheetFunction.Index(ws.Range("I:I"), Ticker_3_Index)
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
    Next ws
    
End Sub









