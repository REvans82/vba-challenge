Sub VBAof_WallStreet()
Dim LastRow As Long
Dim Counter As Long
Dim WB As Workbook
Dim WS As Worksheet
Dim Ticker As String
Dim TotalStock As Double
Dim Counter_Result As Long
Dim Open_Price As Double
Dim Close_Price As Double


Set WB = ActiveWorkbook
Set WS = WB.ActiveSheet
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
Counter = 2
Ticker = WS.Cells(Counter, 1)
Open_Price = WS.Cells(Counter, 3)
TotalStock = 0
Counter_Result = 2
'WS.Range("I2:L100000").ClearContents
Do While Counter <= LastRow

If WS.Cells(Counter, 1) <> Ticker Then
    WS.Cells(Counter_Result, 9) = Ticker
    WS.Cells(Counter_Result, 10) = TotalStock
    Close_Price = WS.Cells(Counter - 1, 6)
    WS.Cells(Counter_Result, 11) = Close_Price - Open_Price
    If WS.Cells(Counter_Result, 11) < 0 Then
        WS.Cells(Counter_Result, 11).Interior.Color = vbRed
    Else
        WS.Cells(Counter_Result, 11).Interior.Color = vbGreen
    End If
    If Close_Price = 0 Then
        If Open_Price = 0 Then
            WS.Cells(Counter_Result, 12) = FormatPercent(0, 2)
        Else
            WS.Cells(Counter_Result, 12) = FormatPercent(-1, 2)
        End If
    Else
        WS.Cells(Counter_Result, 12) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
    End If
    Ticker = WS.Cells(Counter, 1)
    TotalStock = WS.Cells(Counter, 7)
    Open_Price = WS.Cells(Counter, 3)
    Counter_Result = Counter_Result + 1
Else
    TotalStock = TotalStock + WS.Cells(Counter, 7)
End If

Counter = Counter + 1
Loop
    Close_Price = WS.Cells(Counter - 1, 6)
    WS.Cells(Counter_Result, 11) = Close_Price - Open_Price
    If WS.Cells(Counter_Result, 11) < 0 Then
        WS.Cells(Counter_Result, 11).Interior.Color = vbRed
    Else
        WS.Cells(Counter_Result, 11).Interior.Color = vbGreen
    End If
    If Close_Price = 0 Then
        If Open_Price = 0 Then
            WS.Cells(Counter_Result, 12) = FormatPercent(0, 2)
        Else
            WS.Cells(Counter_Result, 12) = FormatPercent(-1, 2)
        End If
    Else
        WS.Cells(Counter_Result, 12) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)
    End If
    WS.Cells(Counter_Result, 9) = WS.Cells(Counter - 1, 1)
    WS.Cells(Counter_Result, 10) = TotalStock
End Sub
