Attribute VB_Name = "GeneralHelper"
Option Explicit
'TW     18/10/2005  MAR223

Public Function CSafeInt(ByVal vvntExpression As Variant) As Integer
On Error GoTo CSafeIntVbErr
    Dim strFunctionName As String
    strFunctionName = "CSafeInt"
    Dim bytRetValue As Integer
    If Len(vvntExpression) > 0 Then
        bytRetValue = CInt(vvntExpression)
    End If
    CSafeInt = bytRetValue
    Exit Function
CSafeIntVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeDbl(ByVal vvntExpression As Variant) As Double
On Error GoTo CSafeDblVbErr
    Dim strFunctionName As String
    strFunctionName = "CSafeDbl"
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
    Dim strFunctionName As String
    strFunctionName = "CSafeLng"
    Dim lngRetValue As Long
    If Len(vvntExpression) > 0 Then
        lngRetValue = CLng(vvntExpression)
    End If
    CSafeLng = lngRetValue
    Exit Function
CSafeLngVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeDate(ByVal vvntExpression As Variant) As Date
On Error GoTo CSafeDateVbErr
    Dim strFunctionName As String
    strFunctionName = "CSafeDate"
    Dim dteRetValue As Date
    If Len(vvntExpression) > 0 Then
        dteRetValue = CDate(vvntExpression)
    End If
    CSafeDate = dteRetValue
    Exit Function
CSafeDateVbErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
