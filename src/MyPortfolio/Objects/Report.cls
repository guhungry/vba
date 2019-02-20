VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Global Variable
Private info As New SheetInfo

''''''''''''''''''''''
' Last Row property
''''''''''''''''''''''
Public Property Get LastRow() As Long
    LastRow = info.LastRow
End Property

''''''''''''''''''''''
' Active Cell property
''''''''''''''''''''''
Public Property Get ActCell() As String
    ActCell = info.ActCell
End Property

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Report.Cells, 3
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportMake
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Generate stock transaction report
'---------------------------------------------------------------------------------------
'
Public Sub ReportMake()
    Dim curDate As Date
    Dim curStock As String
    Dim myComment As String
    Dim myCell As Range, endCell As Range, tran As Range, target As Range
    Dim LastRow As Long, EndRow As Long, StartRow As Long, NumStock As Long
    
    'Initialize Variables
    ReportClearData
    Stock.Initialize
    Transaction.Initialize
    Report.Activate
    Set myCell = Report.Cells(6, 1)
    LastRow = 11
    Do While LastRow <= Transaction.LastRow
        Set tran = Transaction.Range("C" & LastRow)
        curStock = tran.Cells(1, 1).value
        curDate = tran.Cells(1, 0).value
        
        'Set Header in New Row
        If curStock <> tran.Cells(0, 1).value Then
            StartRow = LastRow
            NumStock = 0
            myCell.Cells(1, 3).Formula = "=" & WCRange.ToAddress(tran.Cells(1, 1))
            'Very very super ultra slow so disabled for now
            'Report.Hyperlinks.Add Anchor:=myCell.cells(1, 3), Address:="", SubAddress:=WCRange.ToAddress(tran.cells(1, 1))
            Set myCell = myCell.Cells(2, 1)
        End If
        
        'Find EndRow
        NumStock = NumStock + 1
        EndRow = LastRow + 1
        Set endCell = Transaction.Range("C" & EndRow)
        Do While EndRow <= Transaction.LastRow
            Set endCell = Transaction.Range("C" & EndRow)
            If endCell.Cells(1, 0).value <> curDate Or endCell.Cells(1, 1).value <> curStock Then
                Exit Do
            End If
            EndRow = EndRow + 1
        Loop
        EndRow = EndRow - 1

        'Insert Data
        myCell.Cells(1, 1).value = curDate
        myCell.Cells(1, 4).value = "=" & SumTransaction("J", LastRow, EndRow)
        
        'Transaction
        SetBuySellDividendCommission myCell, LastRow, EndRow
        
        'Remain Stock
        myCell.Cells(1, 5).FormulaR1C1 = GetStockRemaining() & "+R[1]C"
        
        'Insert Header in Last Row
        If curStock <> endCell.Cells(1, 1).value Or EndRow = Transaction.LastRow Then
            Set myCell = myCell.Cells(2, 1)
            Set target = myCell.Cells(-NumStock, 1)
            target.Cells(1, 4).value = "=" & SumTransaction("J", StartRow, EndRow)
            
            'Transaction
            SetBuySellDividendCommission target, StartRow, EndRow
            
            'Remain Stock
            target.Cells(1, 5).FormulaR1C1 = GetStockRemaining()
            
            If target.Cells(1, 5).value > 0 Then
                target.Cells(1, 6).value = "=-RC[-2]/RC[-1]"
                'Get Latest Stock Price
                Dim latestPrice As Range
                Set latestPrice = Report.Range(target.Cells(1, 7).Address)
                latestPrice.Formula = Stock.GetLatestStockPrice(curStock, myComment)
                WCRange.SetComment latestPrice, myComment

                target.Cells(1, 8).value = "=RC[-1]*RC[-3]"
                target.Cells(1, 9).value = "=RC[-1]+RC[-5]"
                target.Cells(1, 10).value = "=IF(RC[-6]>0,RC[-1]*0.01,-RC[-1]/RC[-6])"
            End If
        End If
        
        Set myCell = myCell.Cells(2, 1)
        LastRow = EndRow + 1
    Loop

    'Main Header
    Set myCell = Report.Range("D5")
    LastRow = Transaction.LastRow
    If LastRow > 10 Then
        myCell.Cells(1, 1).value = "=" & SumTransaction("J", 11, LastRow)
        myCell.Cells(1, 8).value = "=" & GetStockValue("D", 11, LastRow)
        myCell.Cells(1, 11).value = "=" & GetStockValue("E", 11, LastRow)
        myCell.Cells(1, 14).value = "=" & SumTransaction("I", 11, LastRow)
        myCell.Cells(1, 15).value = "=" & GetStockValue("F", 11, LastRow)
        LastRow = WCRange.LastRow(Report.Cells)
        myCell.Cells(1, 5).value = "=SUM(H6:H" & LastRow & ")"
        myCell.Cells(1, 6).value = "=SUM(I6:I" & LastRow & ")"
        myCell.Cells(1, 7).value = "=IF(RC[-6]>0,RC[-1]*0.01,-RC[-1]/RC[-6])"
    End If
    
    'Styling
    Dim myActive As String
    myActive = info.ActCell
    Initialize
    ReportStyleAdd
    If myActive <> "" Then
        Report.Range(myActive).Select
    End If
End Sub

Private Function SumTransaction(col As String, StartRow As Long, EndRow As Long)
    SumTransaction = "SUM(Transaction!" & col & StartRow & ":" & col & EndRow & ")"
End Function

Private Sub SetBuySellDividendCommission(current As Range, StartRow As Long, EndRow As Long)
    SetBuyTransaction current, StartRow, EndRow
    SetSellTransaction current, StartRow, EndRow
    SetDividendTransaction current, StartRow, EndRow
    current.Cells(1, 17).value = GetStockCommission(StartRow, EndRow)
End Sub

Private Sub SetBuyTransaction(current As Range, StartRow As Long, EndRow As Long)
    current.Cells(1, 11).value = GetStockQuantity("D", StartRow, EndRow)
    current.Cells(1, 13).value = "=-" & GetStockValue("D", StartRow, EndRow)
    SetStockCost (current.Cells(1, 11))
End Sub

Private Sub SetSellTransaction(current As Range, StartRow As Long, EndRow As Long)
    current.Cells(1, 14).value = GetStockQuantity("E", StartRow, EndRow)
    current.Cells(1, 16).value = "=" & GetStockValue("E", StartRow, EndRow)
    SetStockCost (current.Cells(1, 14))
End Sub

Private Sub SetDividendTransaction(current As Range, StartRow As Long, EndRow As Long)
    current.Cells(1, 18).value = "=" & GetStockValue("F", StartRow, EndRow)
    ClearIfZero (current.Cells(1, 18))
End Sub

Private Function GetStockRemaining()
    GetStockRemaining = "=RC[6]-RC[9]"
End Function

Private Sub SetStockCost(quantity As Range)
    If quantity.value = 0 Then
        quantity.value = ""
        quantity.Cells(1, 3).value = ""
    Else
        quantity.Cells(1, 2).FormulaR1C1 = "=RC[1]/RC[-1]"
    End If
End Sub

Private Sub ClearIfZero(quantity As Range)
    If quantity.value = 0 Then
        quantity.value = ""
    End If
End Sub

Private Function GetStockCommission(StartRow As Long, EndRow As Long)
    GetStockCommission = "=-" & SumTransaction("I", StartRow, EndRow)
End Function

Private Function GetStockQuantity(col As String, StartRow As Long, EndRow As Long)
    GetStockQuantity = "=SUMIF(Transaction!" & col & StartRow & ":" & col & EndRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
End Function

Private Function GetStockValue(col As String, StartRow As Long, EndRow As Long)
    GetStockValue = "SUMIF(Transaction!" & col & StartRow & ":" & col & EndRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!J" & StartRow & ":J" & EndRow & ")"
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReportClearData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear report data
'---------------------------------------------------------------------------------------
Private Sub ReportClearData()
    'Clear Header
    Dim Header As Range
    Set Header = Report.Range("D5")
    Header.value = 0
    Header.Cells(1, 5).value = 0
    Header.Cells(1, 6).value = 0
    Header.Cells(1, 7).value = 0
    Header.Cells(1, 8).value = 0
    Header.Cells(1, 11).value = 0
    Header.Cells(1, 14).value = 0
    Header.Cells(1, 15).value = 0

    'Clear Data
    If LastRow > 5 Then
        Report.Range("A6:R" & LastRow).Delete Shift:=xlUp
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportStyleAdd
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Copy styling template to report data
'---------------------------------------------------------------------------------------
Private Sub ReportStyleAdd()
    Dim source As Range, target As Range
    
    If LastRow > 5 Then
        Set source = Report.Range("A2:R2")
        Set target = Report.Range("A6:R" & LastRow)
        
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
End Sub
