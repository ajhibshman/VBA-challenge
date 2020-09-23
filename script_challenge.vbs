Attribute VB_Name = "Module1"
Sub stocks()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim ticker As Integer
ticker = 1

Dim i As Long

'define ticker begin and ending row numbers
Dim begin As Long
Dim ending As Long

'define summary table rows
Dim sumticker As Double
sumticker = 2

'create total stock volume counter
Dim tsv As Double
tsv = 0

'create opening price placeholder
Dim priceopen As Double

'create closing price placeholder
Dim priceclose As Double

'create summary table
ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "Yearly Change"
ws.Columns("J").Style = "Currency"

ws.Cells(1, 11).Value = "Percent Change"
ws.Columns("K").Style = "Percent"
ws.Columns("K").NumberFormat = "0.00%"

ws.Cells(1, 12).Value = "Total Stock Volume"



For i = 2 To ws.Range("A" & Rows.Count).End(xlUp).Row

'populate ticker symbols into summary table
If ws.Cells(i, ticker).Value <> ws.Cells(i - 1, ticker) Then
    ws.Cells(sumticker, 9).Value = ws.Cells(i, ticker).Value
    begin = i
    For e = i To ws.Range("A" & Rows.Count).End(xlUp).Row
        If ws.Cells(e, ticker).Value <> ws.Cells(e + 1, ticker) Then
            ending = e
            Exit For
        End If
    Next e
        
    
    'populate yearly change into summary table
    priceopen = ws.Cells(i, 3).Value
    
        For k = begin To ending
        
        If ws.Cells(k, ticker) <> ws.Cells(k + 1, ticker) Then
            priceclose = ws.Cells(k, 6).Value
            ws.Cells(sumticker, 10).Value = priceclose - priceopen
            
            If priceopen <> 0 Then
                ws.Cells(sumticker, 11).Value = (priceclose - priceopen) / priceopen
            Else: ws.Cells(sumticker, 11).Value = "undefined"
            End If
            
            Exit For
                
        End If
        
        Next k
            
    
    'populate total stock volume into summary table
    For j = begin To ending
        
        If ws.Cells(j, ticker).Value = ws.Cells(sumticker, 9) Then
          tsv = tsv + ws.Cells(j, 7).Value
        End If
        
    Next j
    
    ws.Cells(sumticker, 12).Value = tsv
    tsv = 0
    
sumticker = sumticker + 1
End If

Next i

'final formatting
ws.Columns("I:L").AutoFit

'conditional formatting
For f = 2 To ws.Range("J" & Rows.Count).End(xlUp).Row

    If ws.Cells(f, 10).Value > 0 Then
        ws.Cells(f, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(f, 10).Value < 0 Then
        ws.Cells(f, 10).Interior.ColorIndex = 3
    End If
Next f

Next ws


        




End Sub




