Attribute VB_Name = "Module1"
Sub Summarize_Stock_Data()

Dim open_price, close_price, yearly_change, percent_change As Single
Dim last_row, next_ticker, total_Volume As LongLong
Dim ws As Worksheet
Dim summ_row As Integer 'Summary iterator
Dim final_summ_row() As String

    For Each ws In ActiveWorkbook.Worksheets
        '----------------
        'Generate summary
        '----------------
            summ_row = 2
            next_ticker = 2
            
            'Get # of rows in the active worksheet
            With ActiveSheet
                last_row = ws.Cells(.Rows.Count, "A").End(xlUp).Row
                'MsgBox (last_row)
            End With
    
            'Place the headers for summary
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            'print the Totals
            For next_ticker = 2 To last_row
                open_price = ws.Cells(next_ticker, 3)
                total_Volume = ws.Cells(next_ticker, 7).Value
                ticker = ws.Cells(next_ticker, 1).Value
                While ws.Cells(next_ticker, 1).Value = ws.Cells(next_ticker + 1, 1).Value
                   total_Volume = total_Volume + ws.Cells(next_ticker + 1, 7).Value
                   next_ticker = next_ticker + 1
                Wend
                ws.Cells(summ_row, 9).Value = ticker
                close_price = ws.Cells(next_ticker, 6)
                yearly_change = close_price - open_price
                
                If yearly_change < 0 Then
                    ws.Cells(summ_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summ_row, 10).Interior.ColorIndex = 4
                End If
                
                ws.Cells(summ_row, 10).Value = yearly_change
                If open_price <> 0 Then
                    percent_change = (yearly_change / open_price)
                    ws.Cells(summ_row, 11).Value = percent_change
                    ws.Cells(summ_row, 11).NumberFormat = "0.00%"
                Else
                    ws.Cells(summ_row, 11).Value = Null
                End If
                ws.Cells(summ_row, 12).Value = total_Volume
                summ_row = summ_row + 1
            Next next_ticker
    
        '----------------------
        'Generate final summary
        '----------------------
            sum_summ_row = 2
        
            'Get # of rows in the active worksheet summary
            With ActiveSheet
               lastr = Cells(.Rows.Count, "K").End(xlUp).Row
            End With
            
            'Place the headers for final summary
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
        
            'Greatest % increase
            Index = loc_of_max_increase(ws.Range("K2:K" & lastr)) 'get location of maximum value
            final_summ_row1 = Split(Index, "$") 'get row number
            final_row1 = final_summ_row1(2)
            ws.Cells(2, 15).Value = ws.Range("I" & final_row1).Value
            ws.Cells(2, 16).Value = ws.Range("K" & final_row1).Value
            ws.Cells(2, 16).NumberFormat = "0.00%"
            
            'Greatest % Decrease
            Index = loc_of_max_decrease(ws.Range("K2:K" & lastr)) 'get location of maximum value
            final_summ_row2 = Split(Index, "$") 'get row number
            final_row2 = final_summ_row2(2)
            ws.Cells(3, 15).Value = ws.Range("I" & final_row2).Value
            ws.Cells(3, 16).Value = ws.Range("K" & final_row2).Value
            ws.Cells(3, 16).NumberFormat = "0.00%"
        
            'Greatest Total Volume
            Index = loc_of_max_vol(ws.Range("L2:L" & lastr)) 'get location of maximum value
            final_summ_row3 = Split(Index, "$") 'get row number
            final_row3 = final_summ_row3(2)
            ws.Cells(4, 15).Value = ws.Range("I" & final_row3).Value
            ws.Cells(4, 16).Value = ws.Range("L" & final_row3).Value
        
    Next ws
 

End Sub

Function loc_of_max_increase(max_increase As Range) As String
 
    loc_of_max_increase = WorksheetFunction.Index(max_increase, WorksheetFunction.Match(WorksheetFunction.Max(max_increase), max_increase, 0)).Address()

End Function

Function loc_of_max_decrease(max_decrease As Range) As String
 
    loc_of_max_decrease = WorksheetFunction.Index(max_decrease, WorksheetFunction.Match(WorksheetFunction.Min(max_decrease), max_decrease, 0)).Address()

End Function

Function loc_of_max_vol(max_vol As Range) As String
 
    loc_of_max_vol = WorksheetFunction.Index(max_vol, WorksheetFunction.Match(WorksheetFunction.Max(max_vol), max_vol, 0)).Address()

End Function


