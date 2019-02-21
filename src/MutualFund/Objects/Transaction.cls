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
' Procedure : btnTransactionAdd
' Author    : guhungry
' Date      : 2012-03-28
' Purpose   : Add stock transaction (Buy, Sell, Dividend) to Transaction Sheet, Check error, add record and styling
'---------------------------------------------------------------------------------------
'
Sub btnTransactionAdd()
    'Define variables
    Dim err As String
    Dim source As Range, Target As Range
    err = ""
    Set source = Transaction.Range("B6")
    Set Target = Transaction.Range("B11")
    
    NAV.Initialize
    
    'Validate Data
    If IsEmpty(source.Cells(1, 1).value) Then
        source.Cells(1, 1).value = "=Today()"
    End If
    If Not (IsDate(source.Cells(1, 1).value)) Then
        err = err & "Date is not valid." & vbNewLine
    End If
    'Validate Transaction Type (Buy, Sale, Dividend)
    If Not (WCRegEx.IsMatch("^[BSD]$", source.Cells(1, 2).value)) Then
        source.Cells(1, 2).value = "B"
    End If
    source.Cells(1, 2).value = UCase(source.Cells(1, 2).value)
    'Validate Value, Quantity, Price (NAV)
    If (source.Cells(1, 2).value = "D") Then
        'Empty the Value
        source.Cells(1, 3).value = ""
        'Validate NAV
        If Not (WCValidate.IsNumber(source.Cells(1, 5).value)) Then
            err = err & "Please insert price (dividend)." & vbNewLine
        End If
        'Validate Quantity
        If Not (WCValidate.IsNumber(source.Cells(1, 4).value)) Then
            err = err & "Please insert quantity." & vbNewLine
        End If
    Else
        ' Validate Value
        If Not (WCValidate.IsNumber(source.Cells(1, 3).value)) Then
            err = err & "Please insert value." & vbNewLine
        Else
            source.Cells(1, 3).value = Abs(source.Cells(1, 3).value)
        End If
        'Fill With Latest NAV
        source.Cells(1, 5).value = NAV.GetValueNAV(source.Cells(1, 1).value, source.Cells(1, 2).value)
        If (source.Cells(1, 5).value = -1) Then
            source.Cells(1, 5).value = ""
            err = err & "Please update data from web." & vbNewLine
        End If
        If err = "" Then
            'Calculate Quantity
            If source.Cells(1, 2).value = "S" Then
                source.Cells(1, 4).value = Application.RoundUp(source.Cells(1, 3).value / source.Cells(1, 5).value, 5)
            Else
                source.Cells(1, 4).value = Application.RoundUp(source.Cells(1, 3).value / source.Cells(1, 5).value, 6)
            End If
            source.Cells(1, 4).value = Application.RoundDown(source.Cells(1, 4).value, 4)
            'Set NAV Price
            source.Cells(1, 5).value = NAV.GetValueNAV(source.Cells(1, 1).value, "N")
        End If
    End If
    If err <> "" Then
        MsgBox err
        Exit Sub
    End If
    
    'Add Data
    Target.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    Target.Cells(0, 1).value = source.Cells(1, 1).value
    If source.Cells(1, 2).value = "B" Then
        Target.Cells(0, 3).value = source.Cells(1, 4).value
        Target.Cells(0, 2).value = -source.Cells(1, 3).value
    ElseIf source.Cells(1, 2).value = "S" Then
        Target.Cells(0, 4).value = source.Cells(1, 4).value
        Target.Cells(0, 2).value = source.Cells(1, 3).value
    Else
        Target.Cells(0, 5).value = source.Cells(1, 4).value
    End If
    Target.Cells(0, 6).value = source.Cells(1, 5).value
    Target.Cells(0, 7).FormulaR1C1 = "=ROUND(IF(RC[-4]>0, -1*RC[-4]*RC[-1], IF(RC[-3]>0, RC[-3]*RC[-1], RC[-2]*RC[-1])), 2)"
    If (source.Cells(1, 2).value <> "D") Then
        Target.Cells(0, 8).FormulaR1C1 = "=(RC[-6]-RC[-1])"
    Else
        Target.Cells(0, 2).value = Target.Cells(0, 7).value
    End If
    Target.Cells(0, 9).value = source.Cells(1, 6).value
    
    'Styling
    Initialize
    TransactionStyleAdd
    info.RestoreActiveCell
    btnTransactionSortByDate
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnTransactionSortByDate
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Sort stock transaction by date
'---------------------------------------------------------------------------------------
'
Sub btnTransactionSortByDate()
    Dim myActive As String
    TransactionSort "B", xlDescending
    
    'Make Report
    Application.ScreenUpdating = False
    Report.Initialize
    myActive = Report.ActCell
    Report.ReportClearData
    Report.ReportMake
    Report.Range(myActive).Select
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
    Dim source As Range, Target As Range
    
    If info.LastRow > 10 Then
        Application.ScreenUpdating = False
        Set source = Transaction.Range("B9:J9")
        Set Target = Transaction.Range("B11:J" & info.LastRow)
        
        source.Copy
        Target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
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
'---------------------------------------------------------------------------------------
'
Private Sub TransactionSort(Column1 As String, Order1 As XlSortOrder)
    Dim Target As Range

    Initialize
    Application.ScreenUpdating = False
    If info.LastRow > 10 Then
        Set Target = Transaction.Range("B11:J" & info.LastRow)
        
        'Sort
        Transaction.Sort.SortFields.Clear
        Transaction.Sort.SortFields.Add Key:=Range(Column1 & "11:" & Column1 & info.LastRow), SortOn:=xlSortOnValues, Order:=Order1, DataOption:=xlSortNormal
        With Transaction.Sort
            .SetRange Target
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
