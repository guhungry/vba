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
    Dim myComment As String
    Dim myCell As Range, endCell As Range, tran As Range, Target As Range
    Dim LastRow As Long, EndRow As Long, StartRow As Long, NumStock As Long
    
    'Initialize Variables
    Report.EnableCalculation = False
    NAV.Initialize
    Transaction.Initialize
    Report.Activate
    Set myCell = Report.Cells(6, 1)
    LastRow = 11
    StartRow = LastRow
    Do While LastRow <= Transaction.LastRow
        Set tran = Transaction.Range("C" & LastRow)
        curDate = tran.Cells(1, 0).value
        
        'Set Header in New Row
        If StartRow = LastRow Then
            NumStock = 0
            myCell.Cells(1, 3).value = Trim(WCRegEx.Match(DataFromWeb.Range(Setting.ValueFromName("FUND_NAME")).value, "^[^(]+"))
            Set myCell = myCell.Cells(2, 1)
        End If
        
        'Find EndRow
        NumStock = NumStock + 1
        EndRow = LastRow + 1
        Set endCell = Transaction.Range("C" & EndRow)
        Do While EndRow <= Transaction.LastRow
            Set endCell = Transaction.Range("C" & EndRow)
            If endCell.Cells(1, 0).value <> curDate Then
                Exit Do
            End If
            EndRow = EndRow + 1
        Loop
        EndRow = EndRow - 1

        'Insert Data
        myCell.Cells(1, 1).value = curDate
        myCell.Cells(1, 4).value = "=SUM(Transaction!C" & EndRow & ":C" & LastRow & ")"
        'Buy
        myCell.Cells(1, 11).value = "=SUMIF(Transaction!D" & EndRow & ":D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
        myCell.Cells(1, 13).value = "=-SUMIF(Transaction!D" & EndRow & ":D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        myCell.Cells(1, 12).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
        'Sell
        myCell.Cells(1, 14).value = "=SUMIF(Transaction!E" & EndRow & ":E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
        myCell.Cells(1, 16).value = "=SUMIF(Transaction!E" & EndRow & ":E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        myCell.Cells(1, 15).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
        'Dividend
        myCell.Cells(1, 18).value = "=SUMIF(Transaction!F" & EndRow & ":F" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & LastRow & ")"
        'Remain Stock
        myCell.Cells(1, 5).FormulaR1C1 = "=RC[6]-RC[9]+R[1]C"
        'Commission
        myCell.Cells(1, 17).value = "=-SUM(Transaction!I" & EndRow & ":I" & LastRow & ")"
        
        'Insert Header in Last Row
        If EndRow = Transaction.LastRow Then
            Set myCell = myCell.Cells(2, 1)
            Set Target = myCell.Cells(-NumStock, 1)
            Target.Cells(1, 4).value = "=SUM(Transaction!C" & EndRow & ":C" & StartRow & ")"
            'Buy
            Target.Cells(1, 11).value = "=SUMIF(Transaction!D" & EndRow & ":D" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
            Target.Cells(1, 13).value = "=-SUMIF(Transaction!D" & EndRow & ":D" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            Target.Cells(1, 12).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
            'Sell
            Target.Cells(1, 14).value = "=SUMIF(Transaction!E" & EndRow & ":E" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ")"
            Target.Cells(1, 16).value = "=SUMIF(Transaction!E" & EndRow & ":E" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            Target.Cells(1, 15).FormulaR1C1 = "=IFERROR(RC[1]/RC[-1]," & Chr$(34) & Chr$(34) & ")"
            'Dividend
            Target.Cells(1, 18).value = "=SUMIF(Transaction!F" & EndRow & ":F" & StartRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C" & EndRow & ":C" & StartRow & ")"
            'Remain Stock
            Target.Cells(1, 5).FormulaR1C1 = "=RC[6]-RC[9]"
            'Commission
            Target.Cells(1, 17).value = "=-SUM(Transaction!I" & EndRow & ":I" & StartRow & ")"

            Report.EnableCalculation = True
            If Target.Cells(1, 5).value > 0 Then
                Target.Cells(1, 6).value = "=-RC[-2]/RC[-1]"
                'Get Latest Stock Price
                Target.Cells(1, 7).value = NAV.GetLatestNAV(myComment)
                Target.Cells(1, 7).AddComment myComment
                Target.Cells(1, 7).Comment.Shape.TextFrame.AutoSize = True
                Target.Cells(1, 8).value = "=RC[-1]*RC[-3]"
                Target.Cells(1, 9).value = "=RC[-1]+RC[-5]"
                Target.Cells(1, 10).value = "=IF(RC[-6]>0,RC[-1]*0.01,-RC[-1]/RC[-6])"
            End If
            Report.EnableCalculation = False
        End If
        
        Set myCell = myCell.Cells(2, 1)
        LastRow = EndRow + 1
    Loop

    'Main Header
    Set myCell = Report.Range("D5")
    LastRow = Transaction.LastRow
    If LastRow > 10 Then
        myCell.Cells(1, 1).value = "=SUM(Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 8).value = "=-SUMIF(Transaction!D11:D" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 11).value = "=SUMIF(Transaction!E11:E" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
        myCell.Cells(1, 14).value = "=SUM(Transaction!I11:I" & LastRow & ")"
        myCell.Cells(1, 15).value = "=SUMIF(Transaction!F11:F" & LastRow & "," & Chr$(34) & ">0" & Chr$(34) & ",Transaction!C11:C" & LastRow & ")"
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
    Report.Range(myActive).Select
    Report.EnableCalculation = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReportClearData
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Clear report data
'---------------------------------------------------------------------------------------
'
Public Sub ReportClearData()
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
'
Private Sub ReportStyleAdd()
    Dim source As Range, Target As Range
    
    If LastRow > 5 Then
        Set source = Report.Range("A2:R2")
        Set Target = Report.Range("A6:R" & LastRow)
        
        source.Copy
        Target.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If
End Sub

