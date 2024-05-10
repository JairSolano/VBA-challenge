Attribute VB_Name = "Module2"
Sub stock_ticker_qt()
Dim ws As Worksheet
Dim Stock_Ticker As String
Dim Stock_Total As Double
Dim Stock_Change As Double
Dim Percent_Change As Double
Dim Open_Price As Double
Dim Start As Long
Dim Summary_Table_Row As Long
Dim j As Long
Dim find_value As Long
Dim LastRow As Long
Dim LastSummaryRow As Long

' Loop across all sheets in the worksheet
For Each ws In Worksheets

' Identifing variables
Stock_Total = 0
Summary_Table_Row = 2
Start = 2
j = 0

' Taking into consideration that not all ws will have the same number of rows
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastSummaryRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

' To find total stock volume and list corresdoning ticker
For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Stock_Ticker = ws.Cells(i, 1).Value
        Stock_Total = Stock_Total + ws.Cells(i, 7).Value
        
     ' Finding Open_Price
        If Stock_Total = 0 Then
            ' Print the results
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = 0
            ws.Range("K" & 2 + j).Value = "%" & 0
            ws.Range("L" & 2 + j).Value = 0
        Else
            ' Find First non zero starting value
            If ws.Cells(Start, 3).Value = 0 Then
                For find_value = Start To i
                    If ws.Cells(find_value, 3).Value <> 0 Then
                        Start = find_value
                        Exit For
                    End If
                Next find_value
            End If
            Open_Price = ws.Cells(Start, 3).Value
            Stock_Change = ws.Cells(i, 6).Value - Open_Price
            If Open_Price <> 0 Then
                Percent_Change = Stock_Change / Open_Price
            Else
                Percent_Change = 0
            End If
            Start = i + 1
            ' Write the results to the worksheet
            ws.Range("I" & Summary_Table_Row).Value = Stock_Ticker
            ws.Range("J" & Summary_Table_Row).Value = Stock_Change
            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            ws.Range("L" & Summary_Table_Row).Value = Stock_Total
            
            ' Increment the row counter for the summary table per loop
            Summary_Table_Row = Summary_Table_Row + 1
        End If
' Reset variables for new stock ticker
        Stock_Total = 0
        j = j + 1
    Else
        Stock_Total = Stock_Total + ws.Cells(i, 7).Value
    End If
Next i
    'Headers for Summary Table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
         
         'Conditional formating fill color
         For i = 2 To LastSummaryRow
                    If ws.Cells(i, 10).Value > 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(i, 10).Value < 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(i, 10).Interior.ColorIndex = 0
            
            End If
         
             If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 11).Value < 0 Then
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                Else
                    ws.Cells(i, 11).Interior.ColorIndex = 0
            End If
                'Adding functionality (Min & Max)
                Greatest_Increase = ws.Application.WorksheetFunction.max(Range("K2:K" & LastSummaryRow))
                Greatest_Decrease = ws.Application.WorksheetFunction.Min(Range("K2:K" & LastSummaryRow))
                Greatest_Volume = ws.Application.WorksheetFunction.max(Range("L2:L" & LastSummaryRow))
                'Print the results
                ws.Range("Q2") = Greatest_Increase
                ws.Range("Q2").NumberFormat = "0.00%"
                ws.Range("Q3") = Greatest_Decrease
                ws.Range("Q3").NumberFormat = "0.00%"
                ws.Range("Q4") = Greatest_Volume
                    'Locating corresponding ticker to value
                    If ws.Cells(i, 11).Value = Greatest_Increase Then
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                    ElseIf ws.Cells(i, 11) = Greatest_Decrease Then
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                    ElseIf ws.Cells(i, 12).Value = Greatest_Volume Then
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                    End If
                'Headers for results
                ws.Range("O2") = "Greatest%Increase"
                ws.Range("O3") = "Greatest%Decrease"
                ws.Range("O4") = "Greatest Total Volume"
                ws.Range("P1") = "Ticker"
                ws.Range("Q1") = "Value"
                
                    
        Next i
        
Next ws

End Sub

