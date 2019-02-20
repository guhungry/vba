Attribute VB_Name = "Utils"
Option Explicit

'---------------------------------------------------------------------------------------
' Function  : IFF
' Author    : guhungry
' Date      : 2010-07-08
' Purpose   : IF in macro style
' @param Expression     the logic expression
' @param ValueTrue      the value if true
' @param ValueFalse     the value if false
' @return           returns ValueTrue if true else returns ValueFalse
' Example           IFF(1=1, "True", "False") => "True"
'---------------------------------------------------------------------------------------
'
Public Function IFF(Expression As Boolean, ValueTrue As Variant, ValueFalse As Variant) As Variant
    If Expression Then
        IFF = ValueTrue
    Else
        IFF = ValueFalse
    End If
End Function
