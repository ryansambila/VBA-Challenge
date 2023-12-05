Attribute VB_Name = "Module1"
Sub Stock_Ticker()

For Each ws In Worksheets

'Set header names
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Add Header Color
ws.Range("I1", "L1").Interior.ColorIndex = 15
ws.Range("A1", "G1").Interior.ColorIndex = 15

'Set Values
Dim Ticker As String
Dim YRChange As Double
Dim Percent As Double
Dim Volume As Double
Volume = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Keep track of stock ticker in summery table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
j = 2

'Loop through Tickers
For i = 2 To lastrow

    'Calculating Volume of each row
    Volume = Volume + ws.Cells(i, 7).Value

    'Check for Unique stock tickers
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(Summary_Table_Row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
    'Set Ticker
    Ticker = ws.Cells(i, 1).Value
        
    'add to the Summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    ws.Range("L" & Summary_Table_Row).Value = Volume

    'If statemetn to retreave percent change
    If ws.Cells(j, 3).Value <> 0 Then
        Percent = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
        ws.Cells(Summary_Table_Row, 11).Value = Format(Percent, "Percent")
       
        'Format color for yearly change
        If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
        Else
        
        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4

    End If
        
    Else

    ws.Cells(Summary_Table_Row, 11).Value = Format(0, "Percent")
    
    End If
      
    'Reset Volume
    Volume = 0

    'Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1

    Else

    End If

Next i

Dim Increase As Double
Dim Decrease As Double
Dim greatVol As Double
Dim Value As Double

'Set header names
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Add header Color
ws.Range("O2", "O4").Interior.ColorIndex = 15
ws.Range("P1", "Q1").Interior.ColorIndex = 15

Increase = ws.Cells(2, 10).Value
Decrease = ws.Cells(2, 10).Value
greatVol = ws.Cells(2, 12).Value


'Loop for the Greatest Increase & Decrease
For i = 2 To lastrow
    Value = ws.Cells(i, 10).Value
    vol = ws.Cells(i, 12).Value

    'Find the Greatest Increase
    If Increase < Value Then
        Increase = Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    End If
    
    'Find the Greatest Decrease
    If Decrease > Value Then
        Decrease = Value
        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    End If
     
    
     'Find the Greatest Volume
        If greatVol < vol Then
        greatVol = vol
        ws.Cells(4, 17).Value = greatVol
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    End If
    
Next i

Next ws

End Sub

