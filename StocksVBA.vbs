Attribute VB_Name = "Module1"
'Assignment 2 - The VBA of Wall Street

Sub foreachws()


Dim WS As Worksheet

    Application.ScreenUpdating = False

    For Each WS In Worksheets

        WS.Select

        Call StocksVBA

    Next

    Application.ScreenUpdating = True

End Sub


Sub StocksVBA()

Dim WS As Worksheet

  Dim LastRow As Long
  Dim TKR As String
  Dim i As Long
  Dim tkrcnt As Long


  Set WS = ActiveSheet
  LastRow = WS.Cells.SpecialCells(xlCellTypeLastCell).Row
  i = 2
  TKR = Cells(2, 1).Value
  tkrcnt = 2

  'sort by ticker and date
  Columns("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes, Key2:=Range("B1"), Order2:=xlAscending, Header:=xlYes


  'add columns
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"



  'loop through cells until last row reached
   Do Until Cells(i, 1) = Cells(LastRow + 1, 1)
        Dim x As Long
        Dim vol As LongLong

        
        vol = 0
        x = i
    
     'loop until change in ticker and count volume for current ticker
     Do While Cells(i, 1).Value = TKR

        Cells(tkrcnt, 9) = Cells(i, 1).Value
        vol = vol + Cells(i, 7).Value
        
     
       i = i + 1
     Loop
     

    'perform calculations to fill in column values
    Cells(tkrcnt, 10).Value = Format(Cells(i - 1, 6).Value - Cells(x, 3).Value, "0.00")
    
    If Cells(tkrcnt, 10).Value >= 0 Then Cells(tkrcnt, 10).Interior.ColorIndex = 4 Else Cells(tkrcnt, 10).Interior.ColorIndex = 3
    
    If Cells(tkrcnt, 10) = 0 Then Cells(tkrcnt, 11) = 0 Else Cells(tkrcnt, 11).Value = Format((Cells(i - 1, 6) - Cells(x, 3).Value) / Cells(i - 1, 6) * 100, "0.00") + "%"
   
    Cells(tkrcnt, 12).Value = vol
    
    TKR = Cells(i, 1).Value
    tkrcnt = tkrcnt + 1


  Loop

  'variable for use in determining highest percent change, lowest percent change, and greatest total volume
  Dim max As Double
  Dim min As Double
  Dim vol2 As LongLong

  Dim tagmax As String
  Dim tagmin As String
  Dim tagvol As String

  max = 0
  min = 0
  vol2 = 0

  Cells(1, 15).Value = "Ticker"
  Cells(1, 16).Value = "Value"
  Cells(2, 14).Value = "Greatest % Increase"
  Cells(3, 14).Value = "Greatest % Decrease"
  Cells(4, 14).Value = "Greatest Total Volume"

For i = 2 To LastRow
    If Cells(i, 11).Value > max Then
       max = Cells(i, 11).Value
       tagmax = Cells(i, 9).Value
    End If
    
    If Cells(i, 11).Value < min Then
        min = Cells(i, 11).Value
        tagmin = Cells(i, 9).Value
    End If

    
    If Cells(i, 12) > vol2 Then
        vol2 = Cells(i, 12)
        tagvol = Cells(i, 9).Value
    End If
    

Next

    'insert data into table
    Cells(2, 15).Value = tagmax
    Cells(2, 16).Value = Format((max * 100), "0.00") + "%"
    Cells(3, 15).Value = tagmin
    Cells(3, 16).Value = Format((min * 100), "0.00") + "%"
    Cells(4, 15).Value = tagvol
    Cells(4, 16).Value = vol2
    
    ActiveSheet.UsedRange.Columns.AutoFit

End Sub



