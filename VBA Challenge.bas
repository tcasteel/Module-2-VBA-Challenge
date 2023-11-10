Attribute VB_Name = "Module1"
Sub test()
    
    Dim ticker As String
    Dim Year As Double
    Dim volume As Double
    Dim sumtable As Integer
    Dim Change As Long
    Dim tick_amt As Integer
    Dim Lookup_Array1 As Range
    Dim Lookup_Array2 As Range
    Dim Return_Array As Range
    
    For Each ws In Worksheets

        Set Lookup_Array1 = ws.Range("K:K")
        Set Lookup_Array2 = ws.Range("L:L")
        Set Return_Array = ws.Range("I:I")

    tick_amt = 0
    Percent = 0
    Year = 0
    volume = 0
    totalstock = 0
    sumtable = 2

    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Vaule"
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            Year = ws.Cells(i, 6).Value - ws.Cells(i - tick_amt, 3).Value
            Percent = (ws.Cells(i, 6).Value - ws.Cells((i - tick_amt), 3).Value) / ws.Cells((i - tick_amt), 3).Value
            ws.Range("K" & sumtable).Value = Percent
            ws.Range("K" & sumtable).NumberFormat = "0.00%"
            ws.Range("I" & sumtable).Value = ticker
            ws.Range("J" & sumtable).Value = Year
            ws.Range("L" & sumtable).Value = volume
            tick_amt = 0
            sumtable = sumtable + 1
            volume = 0
        Else
        volume = volume + ws.Cells(i, 7).Value
        tick_amt = tick_amt + 1
        
        End If
            Year = 0
            Percent = 0
            
        Next i
    
    For i = 2 To lastrow
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
 
    End If
        Next i
        
    For i = 2 To lastrow
    If ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 11).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 11).Interior.ColorIndex = 3
 
    End If
        Next i
        
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Columns(11))
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Columns(11))
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Columns(12))
    
    ws.Range("P2").Value = Application.WorksheetFunction.XLookup(ws.Range("Q2").Value, Lookup_Array1, Return_Array)
    ws.Range("P3").Value = Application.WorksheetFunction.XLookup(ws.Range("Q3").Value, Lookup_Array1, Return_Array)
    ws.Range("P4").Value = Application.WorksheetFunction.XLookup(ws.Range("Q4").Value, Lookup_Array2, Return_Array)
    
    Next ws
    End Sub
    



Sub Clear()
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For Each ws In Worksheets
    For i = 2 To lastrow
    ws.Cells(i, 9).Clear
    ws.Cells(i, 10).Clear
    ws.Cells(i, 11).Clear
    ws.Cells(i, 12).Clear
    ws.Range("P2:Q4").Clear
    
    Next i
    Next ws
End Sub

