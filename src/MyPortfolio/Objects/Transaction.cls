VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Transaction"
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

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Initialize Global Variables - Must be call before doing anything
'---------------------------------------------------------------------------------------
'
Public Sub Initialize()
    info.Init Transaction.Cells, 10, "B"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Validate Transaction Type
' Author    : guhungry
' Date      : 2015-12-17
' Param     : val - Uppercased, Trimmed Transaction Type String
' Purpose   : Transaction Type has only B - Buy, S - Sell, D - Dividend
'---------------------------------------------------------------------------------------
Function ValidateTransactionType(val As String) As Boolean
    ValidateTransactionType = WCRegEx.IsMatch(val, "^BSD$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionAdd
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Add stock transaction (Buy, Sell, Dividend) to Transaction Sheet, Check error, add record and styling
'---------------------------------------------------------------------------------------
'
Sub btnTransactionAdd()
    'Define variables
    Dim err As String: err = ""
    Dim source As Range, target As Range
    Set source = Transaction.Range("B6")
    Set target = Transaction.Range("B11")

    'Validate Data
    If IsEmpty(source.Cells(1, 1).value) Then
        source.Cells(1, 1).value = "=Today()"
    End If
    If Not (IsDate(source.Cells(1, 1).value)) Then
        err = err & "Date is not valid." & vbNewLine
    End If
    If Trim(source.Cells(1, 2).value) = "" Then
        err = err & "Please insert Ticker." & vbNewLine
    End If
    source.Cells(1, 3).value = UCase(Trim(source.Cells(1, 3).value))
    If Not (ValidateTransactionType(source.Cells(1, 3).value)) Then
        source.Cells(1, 3).value = "B"
    End If
    If Not (WCValidate.IsNumber(source.Cells(1, 4).value)) Then
        err = err & "Please insert quantity." & vbNewLine
    End If
    If Not (WCValidate.IsNumber(source.Cells(1, 5).value)) Then
        err = err & "Please insert price." & vbNewLine
    End If
    If IsEmpty(source.Cells(1, 6).value) Then
        source.Cells(1, 6).value = Setting.ValueFromName("DEF_VAT")
    End If
    If IsEmpty(source.Cells(1, 7).value) Then
        source.Cells(1, 7).value = Setting.ValueFromName("DEF_COM")
    End If
    If IsEmpty(source.Cells(1, 8).value) Then
        source.Cells(1, 8).value = 0
    End If
    If err <> "" Then
        MsgBox err
        Exit Sub
    End If

    'Add Data
    target.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    target.Cells(0, 1).value = source.Cells(1, 1).value
    target.Cells(0, 2).value = source.Cells(1, 2).value
    If source.Cells(1, 3).value = "B" Then
        target.Cells(0, 3).value = source.Cells(1, 4).value
    ElseIf source.Cells(1, 3).value = "S" Then
        target.Cells(0, 4).value = source.Cells(1, 4).value
    Else
        target.Cells(0, 5).value = source.Cells(1, 4).value
    End If
    target.Cells(0, 6).value = source.Cells(1, 5).value
    target.Cells(0, 7).FormulaR1C1 = "=IF(RC[-4]>0, -1*RC[-4]*RC[-1], IF(RC[-3]>0, RC[-3]*RC[-1], RC[-2]*RC[-1]))"
    target.Cells(0, 8).FormulaR1C1 = "=IF(OR(RC[-5]>0, RC[-4]>0),-MAX(ABS(RC[-1])*" & source.Cells(1, 7).value & "," & source.Cells(1, 8).value & ")*(1+" & source.Cells(1, 6).value & "),0)"
    target.Cells(0, 9).FormulaR1C1 = "=RC[-1]+RC[-2]"

    'Styling
    Initialize
    TransactionStyleAdd
    info.RestoreActiveCell
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Sort stock transaction by date
'---------------------------------------------------------------------------------------
'
Sub btnTransactionSortByDate()
    TransactionSort "B", xlDescending, "C", xlAscending
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByTicker
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Sort stock transaction by stock name
'               And make report and portfolio
'---------------------------------------------------------------------------------------
'
Sub btnTransactionSortByTicker()
    Dim myActive As String
    TransactionSort "C", xlAscending, "B", xlDescending

    'Make Report & Portfolio
    Application.ScreenUpdating = False
    Report.Initialize
    myActive = Report.ActCell
    Report.ReportMake
    Report.Range(myActive).Select
    Portfolio.Initialize
    myActive = Portfolio.ActCell
    Portfolio.PortfolioMake
    Portfolio.Range(myActive).Select
    Transaction.Activate
    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TransactionStyleAdd
' Author    : guhungry
' Date      : 2010-07-06
' Purpose   : Copy styling template to report data
'---------------------------------------------------------------------------------------
'
Private Sub TransactionStyleAdd()
    Dim source As Range, target As Range

    If info.LastRow > 10 Then
        Application.ScreenUpdating = False
        Set source = Transaction.Range("B9:K9")
        Set target = Transaction.Range("B11:K" & info.LastRow)

        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByTicker
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Based stock transaction sort procedure
' @param Column1    the 1st column to sort
' @param Order1     the 1st sort order
' @param Column2    the 2nd column to sort
' @param Order2     the 2nd sort order
'---------------------------------------------------------------------------------------
'
Private Sub TransactionSort(Column1 As String, Order1 As XlSortOrder, Column2 As String, Order2 As XlSortOrder)
    Dim target As Range

    Initialize
    Application.ScreenUpdating = False
    If info.LastRow > 10 Then
        Set target = Transaction.Range("B11:K" & info.LastRow)

        'Sort
        Transaction.Sort.SortFields.Clear
        Transaction.Sort.SortFields.Add Key:=Range(Column1 & "11:" & Column1 & info.LastRow), SortOn:=xlSortOnValues, Order:=Order1, DataOption:=xlSortNormal
        Transaction.Sort.SortFields.Add Key:=Range(Column2 & "11:" & Column2 & info.LastRow), SortOn:=xlSortOnValues, Order:=Order2, DataOption:=xlSortNormal
        With Transaction.Sort
            .SetRange target
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        info.RestoreActiveCell
    End If
    Application.ScreenUpdating = True
End Sub
