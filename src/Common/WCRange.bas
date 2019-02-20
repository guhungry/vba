Attribute VB_Name = "WCRange"
Option Explicit

' |---------------------
' | Module  : Range Utilities
' | Author    : guhungry
' | Date      : 2019-02-20
' |---------------------

'---------------------------------------------------------------------------------------
' Function  : LastRow
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get last row number that have data in given column
' @param area      the search Range
' @param column the search column
' @return           the last row number that have data
' Example          WCRange.LastRow(Portfolio.Cells, "A") => 50
'---------------------------------------------------------------------------------------
Public Function LastRow(area As Range, Optional column As String = "A") As Long
    LastRow = area.Cells(area.Cells.Rows.Count, column).End(xlUp).row
End Function

'---------------------------------------------------------------------------------------
' Function  : LastColumn
' Author    : guhungry
' Date      : 2010-07-09
' Purpose   : Get last row number that have data in given column
' @param Cells      the search Range
' @param RowNum     the search row
' @return           the last column number that have data
' Example           WCRange.LastColumn(BalanceSheet.Cells, 2) => 50
'---------------------------------------------------------------------------------------
Public Function LastColumn(area As Range, Optional row As Long = 1) As Long
    LastColumn = area.Cells(row, area.Cells.Columns.Count).End(xlToLeft).column
End Function

'---------------------------------------------------------------------------------------
' Function  : ColumnName
' Author    : guhungry
' Date      : 2010-07-14
' Purpose   : Get column name of a cell.
' @param Cell       the cell
' @return           the column name
' Example           ColumnName(Portfolio.Cells(1, 5)) => 'E'
'---------------------------------------------------------------------------------------
Public Function ColumnName(area As Range) As String
    ColumnName = WCRegEx.Match(area.Address, "[^\d$]+")
End Function

'---------------------------------------------------------------------------------------
' Function  : ToAddress - Convert Range to Address
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get address of given range, use for reference or etc.
' @param Cell      the range
' @return           the absolute address or range
' Example           WCRange.ToAddress(Portfolio.Range("B30:G33")) => 'Portfolio!$B$30:$G$33'
'---------------------------------------------------------------------------------------
Public Function ToAddress(area As Range) As String
    ToAddress = "'" & area.Parent.name & "'!" & Replace(area.Address, "$", "")
End Function

'---------------------------------------------------------------------------------------
' Function  : ValueFromName
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Get value of the named variable
' @param value      the named variable
' @return           the value of named variable
' Example           WCRange.ValueFromName(Setting.Names("DATE_WEB")) => "A2"
'---------------------------------------------------------------------------------------
Public Function ValueFromName(value As name) As String
    Dim S As String
    Dim HasRef As Boolean
    Dim R As Range
    On Error Resume Next
    Set R = value.RefersToRange
    
    HasRef = (err.Number = 0)
    If HasRef = True Then
        S = R.text
    Else
        S = value.RefersTo
        If StrComp(Mid(S, 2, 1), Chr(34), vbBinaryCompare) = 0 Then
            ' text constant
            S = Mid(S, 3, Len(S) - 3)
        Else
            ' numeric contant
            S = Mid(S, 2)
        End If
    End If
    ValueFromName = S
End Function

'---------------------------------------------------------------------------------------
' Procedure : Delete Named Range
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Delete all Names in Sheet
' @param this       the Names to be deleted
' Example           WCRange.DeleteNames(names)
'---------------------------------------------------------------------------------------
Public Sub DeleteNames(value As Names)
    For Each n In value
        n.Delete
    Next n
End Sub
