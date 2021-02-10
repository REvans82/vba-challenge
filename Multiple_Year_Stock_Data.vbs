{\rtf1\ansi\ansicpg1252\cocoartf2578
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fnil\fcharset0 Verdana;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue0;}
{\*\expandedcolortbl;;\cssrgb\c0\c0\c0;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs24 \cf2 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 Sub VBAof_WallStreet()\
\
' Declared variables:\
' LastRow = Used to identify and save the last row in each tab\
' Counter = Start at row #2 and used to count the cell based on it's Ticker value and save\
' WB (workbook) = is an object of the full document with tabs,\
' WS (worksheet) = is an object of the active worksheet\
' Ticker = Saves the ticker label into ActiveWorksheet as a String\
' TotalStock = Number representing the total of each Ticker's volume based sequence\
' Counter_Results = saves the count results through the Ticker/TotalStock Volume columns\
' Open_Price = Saved value for Open Price in ActiveWorkSheet\
' Close_Price = Saved value ofr Closing Price in ActiveWorksheet\
\
Dim LastRow As Long\
Dim Counter As Long\
Dim WB As Workbook\
Dim WS As Worksheet\
Dim Ticker As String\
Dim TotalStock As Double\
Dim Counter_Result As Long\
Dim Open_Price As Double\
Dim Close_Price As Double\
\
' Loop #1 - Sets value of Ticker and Total Stock Volume in the newly added columns while looping progresses- this occurs in columns 9 -10.\
' Special code added to clear cells and clean up area to allow for multiple runs of code. Counts the total # of rows in column 1.\
\
\
Set WB = ActiveWorkbook\
Set WS = WB.ActiveSheet\
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row\
Counter = 2\
Ticker = WS.Cells(Counter, 1)\
Open_Price = WS.Cells(Counter, 3)\
TotalStock = 0\
Counter_Result = 2\
WS.Range("I2:L100000").ClearContents\
Do While Counter <= LastRow\
\
' Loop #2 - For columns 9-12, the Counter logic is used to determine the changes in Ticker cell value, then reset to apply other Ticker sequence\
' labels and totals starting at cell A2 through very end of column. In columns 11-12, Yearly Price and Yearly Percent is dervied with an index operation\
' by calculating the result of Closing Price minus Opening Price yearly result and then divide by Closing Price to determine its percentage.\
'  Also, the Open Price is reset by saving Opening Price and performing calc on all tabs. Condition is included for when Yearly price reult is 0 (no change) or when Opening /Closing is\
' 0 (100% decrease).Also, uses conditonal formatting logic to reflect positive price changes in green and negative in red.\
\
If WS.Cells(Counter, 1) <> Ticker Then\
    WS.Cells(Counter_Result, 9) = Ticker\
    WS.Cells(Counter_Result, 10) = TotalStock\
    Close_Price = WS.Cells(Counter - 1, 6)\
    WS.Cells(Counter_Result, 11) = Close_Price - Open_Price\
    If WS.Cells(Counter_Result, 11) < 0 Then\
        WS.Cells(Counter_Result, 11).Interior.Color = vbRed\
    Else\
        WS.Cells(Counter_Result, 11).Interior.Color = vbGreen\
    End If\
    If Close_Price = 0 Then\
        If Open_Price = 0 Then\
            WS.Cells(Counter_Result, 12) = FormatPercent(0, 2)\
        Else\
            WS.Cells(Counter_Result, 12) = FormatPercent(-1, 2)\
        End If\
    Else\
        WS.Cells(Counter_Result, 12) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)\
    End If\
    Ticker = WS.Cells(Counter, 1)\
    TotalStock = WS.Cells(Counter, 7)\
    Open_Price = WS.Cells(Counter, 3)\
    Counter_Result = Counter_Result + 1\
Else\
    TotalStock = TotalStock + WS.Cells(Counter, 7)\
End If\
\
' Loop details and instructions to keeep counting and build Ticker sequence, Total Stock results, and color-coded Yearly change results in red (negative) and green (positive).\
' Also, the Open Price is reset by saving Opening Price and performing calc on all tabs.\
' Condition is included for when Yearly price reult is 0 (no change) or when Opening /Closing is 0 (100% decrease).\
\
Counter = Counter + 1\
Loop\
    Close_Price = WS.Cells(Counter - 1, 6)\
    WS.Cells(Counter_Result, 11) = Close_Price - Open_Price\
    If WS.Cells(Counter_Result, 11) < 0 Then\
        WS.Cells(Counter_Result, 11).Interior.Color = vbRed\
    Else\
        WS.Cells(Counter_Result, 11).Interior.Color = vbGreen\
    End If\
    If Close_Price = 0 Then\
        If Open_Price = 0 Then\
            WS.Cells(Counter_Result, 12) = FormatPercent(0, 2)\
        Else\
            WS.Cells(Counter_Result, 12) = FormatPercent(-1, 2)\
        End If\
    Else\
        WS.Cells(Counter_Result, 12) = FormatPercent((Close_Price - Open_Price) / Close_Price, 2)\
    End If\
    WS.Cells(Counter_Result, 9) = WS.Cells(Counter - 1, 1)\
    WS.Cells(Counter_Result, 10) = TotalStock\
End Sub\
\
}