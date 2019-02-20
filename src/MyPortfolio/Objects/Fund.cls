VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fund"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : Validate Transaction Type
' Author    : guhungry
' Date      : 2015-12-17
' Param     : val - Uppercased, Trimmed Transaction Type String
' Purpose   : Transaction Type has only D - Deposit, W - Withdrawn, I - Interest
'---------------------------------------------------------------------------------------
Function ValidateTransactionType(val As String) As Boolean
    ValidateTransactionType = WCRegEx.IsMatch(val, "^[DWI]$")
End Function

'---------------------------------------------------------------------------------------
' Procedure : btnFundAdd
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Add fund to Fund Sheet, Check error, add record and styling
'---------------------------------------------------------------------------------------
'2012-07-07 Add Remark Column
Sub btnFundAdd()
    'Define variables
    Dim err As String
    Dim source, target As Range
    err = ""
    Set source = Fund.Range("B5")
    Set target = Fund.Range("B11")

    'Validate Data
    If Trim(source.Cells(1, 1).value) = "" Then
        err = err & "Please insert Bank." & vbNewLine
    End If
    If Trim(source.Cells(1, 2).value) = "" Then
        source.Cells(1, 2).value = "=Today()"
    End If
    If Not (IsDate(source.Cells(1, 2).value)) Then
        err = err & "Date is not valid." & vbNewLine
    End If
    source.Cells(1, 3).value = Trim(UCase(source.Cells(1, 3).value))
    If Not (ValidateTransactionType(source.Cells(1, 3).value)) Then
        source.Cells(1, 3).value = "D"
    End If
    If Not (WCValidate.IsNumber(source.Cells(1, 4).value)) Then
        err = err & "Please insert amount." & vbNewLine
    End If
    If err <> "" Then
        MsgBox err
        Exit Sub
    End If
    
    'Add Data
    target.EntireRow.Insert xlDown, xlFormatFromRightOrBelow
    target.Cells(0, 1).value = source.Cells(1, 1).value
    target.Cells(0, 2).value = source.Cells(1, 2).value
    target.Cells(0, 3).value = source.Cells(1, 3).value
    If target.Cells(0, 3).value = "W" Then
        target.Cells(0, 4).value = -source.Cells(1, 4).value
    Else
        target.Cells(0, 4).value = source.Cells(1, 4).value
    End If
    target.Cells(0, 6).value = source.Cells(1, 5).value
    
    'Refresh Data
    FundRefreshData
    
    MsgBox "Add Fund successfully"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnFundRefreshData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Refresh fund data
'---------------------------------------------------------------------------------------
'
Public Sub btnFundRefreshData()
    FundRefreshData
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FundRefreshData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Sort fund data, add formular and styling
'---------------------------------------------------------------------------------------
'2012-07-07 Style Remark Column
Private Sub FundRefreshData()
    Dim source As Range, target  As Range
    Dim info As New SheetInfo
    
    Set source = Fund.Range("B8:G8")
    info.Init Fund.Cells, 9, "B"
    Application.ScreenUpdating = False
    If info.LastRow > 10 Then
        Set target = Fund.Range("B11:G" & info.LastRow)
        
        'Sort
        Fund.Sort.SortFields.Clear
        Fund.Sort.SortFields.Add Key:=Range("B11:B" & info.LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Fund.Sort.SortFields.Add Key:=Range("C11:C" & info.LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With Fund.Sort
            .SetRange target
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        'Styling
        source.Copy
        target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        info.RestoreActiveCell
        
        'Add Formular
        Dim i As Integer
        i = info.LastRow
        Do While i >= 11
            Fund.Cells(i, 6).FormulaR1C1 = "=IF(RC[-4]=R[1]C[-4],RC[-1]+R[1]C,RC[-1])"
            i = i - 1
        Loop
    End If
    Fund.Cells(10, 6).value = "=SUM(E11:E" & info.LastRow & ")"
    Application.ScreenUpdating = True
End Sub

