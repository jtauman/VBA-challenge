Attribute VB_Name = "Module2"
Option Explicit

Sub ticker_challenge():
Dim ws As Worksheet
Dim first As Double
Dim symbol, x As String
Dim opn, cls, vol, percentchange As Double
Dim i, j, lastrow As Long
Dim y As Variant

For Each ws In Worksheets

i = 2 'i is iteration through data
j = 2 'using j as counter for output locations

'loading the first values before running the loop
symbol = ws.Cells(i, 1).Value
first = ws.Cells(i, 3).Value
vol = 0

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'writing column headers
ws.Range("J1").Value = "Ticker Symbol"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("J1:M1").Font.Bold = True
ws.Columns("L").NumberFormat = "0.00%"


For i = 2 To lastrow

x = ws.Cells(i + 1, 1).Value 'is the cell 1 past the current i value

'if first dates value = 0, then keep going to the next cell until the value doesn't = 0
If first = 0 Then
    first = ws.Cells(i + 1, 3).Value
End If

'checking that symbols match
If x <> symbol Then
    ws.Cells(j, 10).Value = ws.Cells(i, 1).Value 'ticker symbol
    
    If vol = 0 Then     'if sum of vol = 0(all values are 0), then set cell contents to 0
        ws.Cells(j, 11).Value = 0
        ws.Cells(j, 12).Value = 0
    
    Else: ws.Cells(j, 11).Value = ws.Cells(i, 6).Value - first 'yearly change
            ws.Cells(j, 12).Value = ws.Cells(j, 11).Value / first '% change
    
    End If
    
    If ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 11).Value < 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 3
    Else: ws.Cells(j, 11).Interior.ColorIndex = 15
    End If
        
    'Cells(j, 12).Style = "percent"

    ws.Cells(j, 13).Value = vol
    
    'resetting i to next symbol and resetting first
    symbol = ws.Cells(i + 1, 1).Value
    first = ws.Cells(i + 1, 3).Value
    vol = 0
    
    'counter for output
    j = j + 1
    
'function to sum the volumes within a ticker symbol
Else: vol = vol + ws.Cells(i + 1, 7).Value

    
End If

Next i

'CHALLENGE'


'setting row and column labels

ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"
ws.Range("P2:P4").Font.Bold = True

ws.Range("Q1").Value = "Ticker Symbol"
ws.Range("R1").Value = "Value"
ws.Range("Q1:R1").Font.Bold = True
ws.Range("R2:R3").NumberFormat = "0.00%"


'find max % increase, min % increase and max total volume
Dim maxrow, minrow, r, lastrow2 As Long
Dim maxper, minper As Double
Dim maxvol As Double

lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
maxper = Application.WorksheetFunction.max(ws.Range("L:L"))
minper = Application.WorksheetFunction.min(ws.Range("L:L"))
maxvol = Application.WorksheetFunction.max(ws.Range("M:M"))

For r = 2 To lastrow2
    If ws.Cells(r, 12).Value = maxper Then
        ws.Range("Q2").Value = ws.Cells(r, 10).Value
        ws.Range("R2").Value = ws.Cells(r, 12).Value
        
    ElseIf ws.Cells(r, 12).Value = minper Then
        ws.Range("Q3").Value = ws.Cells(r, 10).Value
        ws.Range("R3").Value = ws.Cells(r, 12).Value
    
    ElseIf ws.Cells(r, 13).Value = maxvol Then
        ws.Range("Q4").Value = ws.Cells(r, 10).Value
        ws.Range("R4").Value = ws.Cells(r, 13).Value
    End If
    
Next r

Next ws

End Sub
