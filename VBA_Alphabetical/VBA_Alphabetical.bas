Attribute VB_Name = "Module1"
Sub alphabetical()

'Perform for each worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

'Declare Variables
Dim x As Integer
Dim lastrow As Long
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double

'Output value will be in this row number
x = 2
Yearly_Change = 0
Percent_Change = 0
Stock_Volume = 0

Cells(1, 10).Value = "Ticker"
Cells(1, 17).Value = "Ticker"
Cells(1, 18).Value = "Value"
Cells(2, 16).Value = "Greatest % Increase"
Cells(3, 16).Value = "Greatest % Decrease"
Cells(4, 16).Value = "Greatest Total Volume"
'Range("J1").Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Stock Volume"


'Find the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    Cells(x, 13).Value = Stock_Volume
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       Cells(x, 10).Value = Cells(i, 1).Value
       Open_Price = Cells(i, 3).Value
       Close_Price = Cells(i, 6).Value
       Yearly_Change = Close_Price - Open_Price
       Cells(x, 11).Value = Yearly_Change
       
            'Calculate percent_change
            If (Close_Price = 0 And Open_Price = 0) Then
            Percent_Change = 0
            ElseIf (Close_Price <> 0 And Open_Price = 0) Then
            Percent_Change = 1
            Else: Percent_Change = Yearly_Change / Open_Price
            Cells(x, 12).Value = Percent_Change
            Cells(x, 12).NumberFormat = "0.00%"
            
                If Percent_Change >= 0 Then
                Cells(x, 12).Interior.Color = vbGreen
                Else
                Cells(x, 12).Interior.Color = vbRed
                End If
       
            End If
    'Add one to the loop iteration
    x = x + 1
    'Reset Yearly Change to 0
    Yearly_Change = 0
    'Reset Percent Change to 0
    Percent_Change = 0
    'Reset Stock_Volume Change to 0
    Stock_Volume = 0
    
    
    End If
    
Next i
        
Next ws

End Sub

