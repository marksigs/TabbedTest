VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConvertAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Workfile:     ConvertAssist.cls
' Copyright:    Copyright © 1999 Marlborough Stirling
'
' Description:  Data type conversion routines
' Dependencies:
' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'PSC   04/08/99 Created
'------------------------------------------------------------------------------------------
Option Explicit
Public Sub StringToByteArray(ByRef rbytByteArray() As Byte, ByVal vstrText As String)
' header ----------------------------------------------------------------------------------
' description:
'   copies each character in vstrText into the relevant position of rbytByteArray
' pass:
'   rbytByteArray()  array of bytes into which the text characters are to be transferred
'   strText          string containing the text to be transferred into the byte array
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo StringToByteArrayVbErr
    Dim strFunctionName As String
    strFunctionName = "StringToByteArray"
    Dim lngTextLength As Long
    Dim lngIndex As Long
    Dim objErrAssist As New ErrAssist
    lngTextLength = Len(vstrText)
    ' Check string plus null terminator will fit into the byte array
    If lngTextLength > UBound(rbytByteArray) Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4InvalidParameter, _
                                "String: '" & vstrText & "'" & _
                                " can only be " & UBound(rbytByteArray) & " characters long"
    End If
       
    For lngIndex = 1 To lngTextLength
        rbytByteArray(lngIndex - 1) = CByte(Asc(Mid(vstrText, lngIndex, 1)))
    Next
    Set objErrAssist = Nothing
    Exit Sub
StringToByteArrayVbErr:
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function ByteArrayToString(ByRef rbytByteArray() As Byte) As String
' header ----------------------------------------------------------------------------------
' description:
'   creates a string representation of a byte array
' pass:
'   rbytByteArray()  array of bytes from which the string is to be formed
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ByteArrayToStringVbErr
    Dim strFunctionName As String
    strFunctionName = "ByteArrayToString"
    Dim lngArrayLimit As Long
    Dim strString As String
    Dim lngIndex As Long
    lngArrayLimit = UBound(rbytByteArray)
    lngIndex = 0
    Do While lngIndex <= lngArrayLimit And rbytByteArray(lngIndex) <> CByte(0)
        strString = strString & Chr(rbytByteArray(lngIndex))
        lngIndex = lngIndex + 1
    Loop
    ByteArrayToString = strString
    Exit Function
ByteArrayToStringVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeDate(ByVal vvntExpression As Variant) As Date
' header ----------------------------------------------------------------------------------
' description:
'   creates a date representation of an expression
' pass:
'   vvntExpression  Expression to be converted to a date
' Returns:
'   CSafeDate   Converted Expression
'------------------------------------------------------------------------------------------
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
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeLng(ByVal vvntExpression As Variant) As Long
' header ----------------------------------------------------------------------------------
' description:
'   creates a long representation of an expression
' pass:
'   vvntExpression  Expression to be converted to a long
' Returns:
'   CSafeLng   Converted Expression
'------------------------------------------------------------------------------------------
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
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeDbl(ByVal vvntExpression As Variant) As Double
' header ----------------------------------------------------------------------------------
' description:
'   creates a double representation of an expression
' pass:
'   Expression  Expression to be converted to a double
' Returns:
'   CSafeDbl   Converted Expression
'------------------------------------------------------------------------------------------
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
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeByte(ByVal vvntExpression As Variant) As Byte
' header ----------------------------------------------------------------------------------
' description:
'   creates a byte representation of an expression
' pass:
'   Expression  Expression to be converted to a byte
' Returns:
'   CSafeByte   Converted Expression
'------------------------------------------------------------------------------------------
On Error GoTo CSafeByteVbErr
    Dim strFunctionName As String
    strFunctionName = "CSafeByte"
    Dim bytRetValue As Byte
    If Len(vvntExpression) > 0 Then
        bytRetValue = CByte(Asc(vvntExpression))
    End If
    CSafeByte = bytRetValue
    Exit Function
CSafeByteVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeInt(ByVal vvntExpression As Variant) As Integer
' header ----------------------------------------------------------------------------------
' description:
'   creates an integer representation of an expression
' pass:
'   Expression  Expression to be converted to an integer
' Returns:
'   CSafeInt   Converted Expression
'------------------------------------------------------------------------------------------
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
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function CSafeBool(ByVal vvntExpression As Variant) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   creates an boolean representation of an expression
' pass:
'   Expression  Expression to be converted to a boolean
' Returns:
'   CSafeInt   Converted Expression
'------------------------------------------------------------------------------------------
On Error GoTo CSafeBoolVbErr
    Dim strFunctionName As String
    strFunctionName = "CSafeBool"
    Dim bytRetValue As Boolean
    If Len(vvntExpression) > 0 Then
        bytRetValue = CBool(vvntExpression)
    End If
    CSafeBool = bytRetValue
    Exit Function
CSafeBoolVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
