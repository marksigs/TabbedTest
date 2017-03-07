Attribute VB_Name = "ConvertAssist"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module        :   ConvertAssist
' Description   :   Conversion utilities
' Change history
' Prog      Date        Description
' RF        14/12/2005  MAR867 Complete pack handling changes -
'                              Added file based on INGDUK $\Code\1 DEV Code\Administration Interface\omAdmin\ConvertAssist.bas
'                              version 2.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


Public Function ByteArrayToString(ByRef rbytByteArray() As Byte) As String
' header ----------------------------------------------------------------------------------
' description:
'   creates a string representation of a byte array
' pass:
'   rbytByteArray()  array of bytes from which the string is to be formed
'------------------------------------------------------------------------------------------
On Error GoTo ByteArrayToStringVbErr

    Const strFunctionName As String = "ByteArrayToString"
    
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

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeDate"
    
    Dim dteRetValue As Date

    If Len(vvntExpression) > 0 Then
        dteRetValue = CDate(vvntExpression)
    End If
    
    CSafeDate = dteRetValue

    Exit Function

CSafeDateVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeLng"
    
    Dim lngRetValue As Long

    If Len(vvntExpression) > 0 Then
        lngRetValue = CLng(vvntExpression)
    End If
    
    CSafeLng = lngRetValue

    Exit Function

CSafeLngVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeDbl"
    
    Dim dblRetValue As Double

    If Len(vvntExpression) > 0 Then
        dblRetValue = CDbl(vvntExpression)
    End If
    
    CSafeDbl = dblRetValue

    Exit Function

CSafeDblVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeByte"
    
    Dim bytRetValue As Byte

    If Len(vvntExpression) > 0 Then
        bytRetValue = CByte(Asc(vvntExpression))
    End If
    
    CSafeByte = bytRetValue

    Exit Function

CSafeByteVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeInt"
    
    Dim bytRetValue As Integer

    If Len(vvntExpression) > 0 Then
        bytRetValue = CInt(vvntExpression)
    End If
    
    CSafeInt = bytRetValue

    Exit Function

CSafeIntVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

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

    Const strFunctionName As String = "CSafeBool"
    
    Dim bytRetValue As Boolean

    If Len(vvntExpression) > 0 Then
        bytRetValue = CBool(vvntExpression)
    End If
    
    CSafeBool = bytRetValue

    Exit Function

CSafeBoolVbErr:

    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

End Function

