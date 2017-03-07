Attribute VB_Name = "GeneralHelper"
Option Explicit

Public Function CSafeInt(ByVal vvntExpression As Variant) As Integer
On Error GoTo CSafeIntVbErr

    Const strFunctionName As String = "CSafeInt"

    Dim intRetValue As Long

    If Len(vvntExpression) > 0 Then
        intRetValue = CInt(vvntExpression)
    End If

    CSafeInt = intRetValue

    Exit Function

CSafeIntVbErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function CSafeDbl(ByVal vvntExpression As Variant) As Double
On Error GoTo CSafeDblVbErr

    Const strFunctionName As String = "CSafeDbl"

    Dim dblRetValue As Double

    If Len(vvntExpression) > 0 Then
        dblRetValue = CDbl(vvntExpression)
    End If

    CSafeDbl = dblRetValue

    Exit Function

CSafeDblVbErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function CSafeLng(ByVal vvntExpression As Variant) As Long
On Error GoTo CSafeLngVbErr

    Const strFunctionName As String = "CSafeLng"

    Dim LngRetValue As Long

    If Len(vvntExpression) > 0 Then
        LngRetValue = CLng(vvntExpression)
    End If

    CSafeLng = LngRetValue

    Exit Function

CSafeLngVbErr:

    Err.Raise Err.Number, Err.Source, Err.Description

End Function
