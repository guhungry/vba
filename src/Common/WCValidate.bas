Attribute VB_Name = "WCValidate"
Option Explicit

'---------------------------------------------------------------------------------------
' Function  : IsNumber
' Author    : guhungry
' Date      : 2010-07-01
' Purpose   : Check if cell is not empty and is number
' @param Cell       the cell to test
' @return           Is not empty and is number
' Example           WCValidate.IsNumber(Portfolio.Range("G30"))
'---------------------------------------------------------------------------------------
'
Public Function IsNumber(value As Variant) As Boolean
    IsNumber = (Not (IsEmpty(value))) And IsNumeric(value)
End Function
