Attribute VB_Name = "Module1"
Option Explicit


Sub ticker():
Dim first, lastrow As Long
Dim symbol, x As String
Dim opn, cls, vol, percentchange As Double
Dim i, j As Long
Dim y As Variant


i = 2 'i is iteration through data
j = 2 'using j as counter for output locations

'loading the first values before running the loop
symbol = Cells(i, 1).Value
first = Cells(i, 3).Value
vol = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'writing column headers
Range("J1").Value = "Ticker Symbol"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("J1:M1").Font.Bold = True
Columns("L").NumberFormat = "0.00%"


For i = 2 To lastrow

x = Cells(i + 1, 1).Value 'is the cell 1 past the current i value

If first = 0 Then
    first = Cells(i + 1, 3).Value
End If

'checking that symbols match
If x <> symbol Then
    Cells(j, 10).Value = Cells(i, 1).Value 'ticker symbol


If vol = 0 Then
    Cells(j, 11).Value = "0"
    Cells(j, 12).Value = "0"
    
Else: Cells(j, 11).Value = Cells(i, 6).Value - first 'yearly change
        Cells(j, 12).Value = Cells(j, 11).Value / first '% change
    
End If
 
 If Cells(j, 11).Value > 0 Then
        Cells(j, 11).Interior.ColorIndex = 4
    ElseIf Cells(j, 11).Value < 0 Then
        Cells(j, 11).Interior.ColorIndex = 3
    Else: Cells(j, 11).Interior.ColorIndex = 15
    End If
   
    Cells(j, 13).Value = vol
    
    'resetting i to next symbol and resetting first
    symbol = Cells(i + 1, 1).Value
    first = Cells(i + 1, 3).Value
    vol = 0
    
    'counter for output
    j = j + 1
    
'function to sum the volumes within a ticker symbol
Else: vol = vol + Cells(i + 1, 7).Value

End If



Next i

'CHALLENGE'


'setting row and column labels

Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"
Range("P2:P4").Font.Bold = True

Range("Q1").Value = "Ticker Symbol"
Range("R1").Value = "Value"
Range("Q1:R1").Font.Bold = True
Range("R2:R3").NumberFormat = "0.00%"


'find max % increase, min % increase and max total volume
Dim maxrow, minrow, r, lastrow2 As Long
Dim maxper, minper As Double
Dim maxvol As Double

lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row
maxper = Application.WorksheetFunction.max(Range("L:L"))
minper = Application.WorksheetFunction.min(Range("L:L"))
maxvol = Application.WorksheetFunction.max(Range("M:M"))

For r = 2 To lastrow2
    If Cells(r, 12).Value = maxper Then
        Range("Q2").Value = Cells(r, 10).Value
        Range("R2").Value = Cells(r, 12).Value
        
    ElseIf Cells(r, 12).Value = minper Then
        Range("Q3").Value = Cells(r, 10).Value
        Range("R3").Value = Cells(r, 12).Value
    
    ElseIf Cells(r, 13).Value = maxvol Then
        Range("Q4").Value = Cells(r, 10).Value
        Range("R4").Value = Cells(r, 13).Value
    End If
    
Next r


End Sub
