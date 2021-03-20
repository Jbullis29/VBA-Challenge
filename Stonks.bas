Attribute VB_Name = "Module1"
Sub stonks()
Dim Percent_change As Double
Dim Total_Volume As Double
Dim ws As Worksheet
Dim i, ii, iii, iiii As Double
Dim lastrow, lastrow_two As Double
Dim Close_price As Double
Dim open_price As Double

    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly_Change"
        ws.Range("K1") = "Percent_Change"
        ws.Range("L1") = "Volume"
        ws.Range("M1") = "Last_Close"
        ws.Range("N1") = "First_Open"
        
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Volume = ws.Cells(i, 7).Value
        Close_price = ws.Cells(i, 6).Value
        Ticker = ws.Cells(i, 1).Value
        Total_Volume = Total_Volume + Volume
        ws.Range("I" & Summary_Table_Row + 1).Value = Ticker
        ws.Range("L" & Summary_Table_Row + 1).Value = Total_Volume
        ws.Range("M" & Summary_Table_Row + 1).Value = Close_price
        Total_Volume = 0
        Summary_Table_Row = Summary_Table_Row + 1
        Else
        Volume = ws.Cells(i, 7).Value
        Close_price = ws.Cells(i, 6).Value
        Total_Volume = Total_Volume + Volume
        ws.Range("L" & Summary_Table_Row + 1).Value = Total_Volume
    
        End If
    Next i
For ii = lastrow To 2 Step -1

    If ws.Cells(ii, 1).Value <> ws.Cells(ii - 1, 1).Value Then
    open_price = ws.Cells(ii, 3).Value
    ws.Range("N" & Summary_Table_Row).Value = open_price
    Summary_Table_Row = Summary_Table_Row - 1
Else
    open_price = ws.Cells(ii, 3).Value
 
    End If

Next ii

lastrow_two = ws.Cells(Rows.Count, 14).End(xlUp).Row
For iii = 2 To lastrow_two
    Yearly_Change = ws.Cells(iii, 13).Value - ws.Cells(iii, 14).Value
    open_price = ws.Cells(111, 14).Value
    Close_price = ws.Cells(iii, 13).Value
    Percent_change = (Close_price / open_price) - 1
    On Error Resume Next
    Summary_Table_Row = Summary_Table_Row + 1
        ws.Range("J" & Summary_Table_Row - 1).Value = Yearly_Change
    ws.Range("K" & Summary_Table_Row - 1).Value = Percent_change
    ws.Range("K" & Summary_Table_Row).NumberFormat = "%0.00"
    Percent_change = (Close_price + 1 / open_price + 1) - 1
    Next iii
For iiii = 3 To lastrow
    If ws.Cells(iiii, 11).Value <= 0 Then
    ws.Cells(iiii, 11).Interior.ColorIndex = 3
    Else: ws.Cells(iiii, 11).Interior.ColorIndex = 4
    End If
Next iiii
ws.Range("J2") = ""
ws.Range("K2") = ""

 Next ws
End Sub
