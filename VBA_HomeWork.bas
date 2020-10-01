Attribute VB_Name = "Module1"
Sub HW_Test():

Dim ws As Worksheet
For Each ws In Worksheets


    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Ticker_Name As String

    Dim Yearly_Change As Double

    Dim Opening_Price_Begining_Year As Double
    Opening_Price_Begining_Year = ws.Cells(2, 3).Value

    Dim Percent_Change As Double

    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0


        For I = 2 To LastRow
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
                Ticker_Name = ws.Cells(I, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            
                Yearly_Change = ws.Cells(I, 6).Value - Opening_Price_Begining_Year
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
                    If Opening_Price_Begining_Year = 0 Then
                        ws.Range("K" & Summary_Table_Row).Value = "n/a"
                    
                        
                    Else: Percent_Change = (Yearly_Change / Opening_Price_Begining_Year)
                        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.0%"
                            If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                            Else:
                                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                            End If
                    End If
                    
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                ws.Range("L" & Summary_Table_Row).NumberFormat = "General"
                
                Opening_Price_Begining_Year = ws.Cells(I + 1, 3).Value
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
        
            Else:
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(I, 7).Value
        
            End If
        
        Next I
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    Dim LastRow1 As Long
    LastRow1 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ws.Cells(2, 17).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 11), ws.Cells(LastRow1, 11)))
    ws.Cells(2, 17).NumberFormat = "0.0%"
    
    ws.Cells(3, 17).Value = Application.WorksheetFunction.Min(ws.Range(ws.Cells(2, 11), ws.Cells(LastRow1, 11)))
    ws.Cells(3, 17).NumberFormat = "0.0%"
    
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range(ws.Cells(2, 12), ws.Cells(LastRow1, 12)))
    ws.Cells(4, 17).NumberFormat = "General"
    
    Dim n As Integer
        
        For n = 2 To LastRow1
            If ws.Cells(n, 11).Value = ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(n, 9).Value
                
                ElseIf ws.Cells(n, 11).Value = ws.Cells(3, 17).Value Then
                    ws.Cells(3, 16).Value = ws.Cells(n, 9).Value
                
                ElseIf ws.Cells(n, 12).Value = ws.Cells(4, 17).Value Then
                    ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
            End If
         
         Next n
         
                 
Next
    
End Sub

