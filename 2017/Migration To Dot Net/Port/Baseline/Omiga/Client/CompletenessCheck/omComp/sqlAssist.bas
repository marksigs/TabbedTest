Attribute VB_Name = "sqlAssist"
Option Explicit


Private Const cstrAPOSTROPHE As String = "'"
Private Const cstrWILDCARD As String = "*"
 
 '=============================================
 'Enumeration Declaration Section
 '=============================================

Public Enum DBDATATYPE
    dbdtNotStored = 0
    dbdtString = 1
    dbdtInt = 2
    dbdtDouble = 3
    dbdtDate = 4
    dbdtGuid = 5
    dbdtCurrency = 6
    dbdtComboId = 7
    dbdtBoolean = 8
    dbdtLong = 9
    dbdtDateTime = 10
End Enum

Public Enum DATETIMEFORMAT
    dtfDateTime = 0      ' date and time
    dtfDate = 1          ' date only
    dtfTime = 2          ' time only
End Enum

Public Enum FIELDDEFINDEX
    efdiElementName
    efdiType
    efdiNoNull
    efdiKey
    efdiComboGroup
    efdiFormatString
    efdiTotal
End Enum

Public Enum GUIDSTYLE
    guidBinary = 0
    guidHyphen = 1
    guidLiteral = 2
End Enum


Private Function GetStringDelimiter() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    GetStringDelimiter = "'"
End Function

Private Function CheckApostrophes(ByVal strIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Doubles each ' in a string. This is needed for names with
'   an apostrophe in them. E.g. "O'Conner" converts to "O''Conner"
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    
    Dim nPos As Integer
    Dim strOut As String
    Dim nLen As Integer
    Dim thisChar As String
    Dim i As Integer
    
    nLen = Len(strIn)
    
    For nPos = 1 To nLen Step 1
        thisChar = Mid$(strIn, nPos, 1)
        strOut = strOut + thisChar
        If thisChar = cstrAPOSTROPHE Then
            ' add an extra apostrophe
            strOut = strOut + cstrAPOSTROPHE
        End If
    Next nPos
    
    CheckApostrophes = strOut

End Function

Public Function sqlFormatString(ByVal vstrIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert a string into appropriate SQL for a string datatype
' pass:
' return:
'   FormatString            formatted string
' Raise Errors:
'------------------------------------------------------------------------------------------

    Dim strIn As String, _
        strOut As String, _
        strDelimiter As String
    
    strDelimiter = GetStringDelimiter
        
    strIn = vstrIn
    
    If Right(strIn, 1) = "*" Then
        strIn = Left(vstrIn, Len(vstrIn) - 1) & "%"
        strOut = " " & sqlGetLikeOperator() & " "
    Else
        strOut = "= "
    End If
    
    strOut = strOut & strDelimiter & CheckApostrophes(strIn) & strDelimiter
    
    sqlFormatString = strOut

End Function

Public Function sqlGetLikeOperator() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    sqlGetLikeOperator = "LIKE"
    
End Function

Public Function sqlFormatDate(ByVal vdtmIn As Date, _
    Optional ByVal vdtfDateTimeFormat As DATETIMEFORMAT = dtfDateTime) As String
' header ----------------------------------------------------------------------------------
' description:
'   Formats a date to the appropriate SQL.
' pass:
'   vdtmIn                  date to be formatted
'   vdtfDateTimeFormat      required format
' return:
'   FormatDate              formatted value
' Raise Errors:
'------------------------------------------------------------------------------------------

    Const strFunctionName As String = "sqlFormatDate"

    Dim strDate As String
    
    Dim eDbEngineType As DBPROVIDER
    eDbEngineType = GetDbProvider
    
    Select Case vdtfDateTimeFormat
    Case dtfDateTime                ' date and time format
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case omiga4DBPROVIDERSQLServer
            strDate = "CONVERT(DATETIME,'" & strDate & "',120)"
        Case omiga4DBPROVIDEROracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
            ' not implemented
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist].", strFunctionName
        End Select
    Case dtfDate                    ' date only
        strDate = Format(vdtmIn, "YYYY-MM-DD")
        Select Case eDbEngineType
        Case omiga4DBPROVIDERSQLServer
            ' fixme
            ' not implemented
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist].", strFunctionName
        Case omiga4DBPROVIDEROracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD')"
        End Select
    Case Else
        ' not implemented
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist].", strFunctionName
    End Select
    
    sqlFormatDate = "=" & strDate

End Function

Public Function sqlFormatGuid(ByVal strSrcGuid As String, Optional ByVal eGuidStyle As GUIDSTYLE = guidLiteral) As String
' header ----------------------------------------------------------------------------------
' description:  Converts a guid string from one format to another.
' pass:
'   strSrcGuid  The guid string to be converted.
'   eGuidStyle  The style of the return string.
'               guidHyphen converts from "DA6DA163412311D4B5FA00105ABB1680" to
'               "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
'               guidBinary converts from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to
'               "DA6DA163412311D4B5FA00105ABB1680" format.
'               guidLiteral converts from either guidBinary or guidHyphen format to
'               a format that can be used as a literal input parameter for the
'               database. This is the default to maintain backwards compatibility with
'               Omiga 4 Phase 2.
' return:       The converted guid string.
' AS            05/03/01    First version
'------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "sqlFormatGuid"
    Dim strTgtGuid As String
    Dim intIndex1 As Integer
    Dim intIndex2 As Integer
        
    ' By default return string unchanged, i.e., if style is guidHyphen but the source string
    ' is currently not in binary format, then it will be returned unchanged (perhaps it is
    ' already in hyphenated format?).
    strTgtGuid = strSrcGuid
    
    If eGuidStyle = guidHyphen Then
        ' Convert from "DA6DA163412311D4B5FA00105ABB1680" to "{63A16DDA-2341-D411-B5FA-00105ABB1680}"
        
        Debug.Assert Len(strSrcGuid) = 32
        
        If Len(strSrcGuid) = 32 Then
            strTgtGuid = "{"
            For intIndex1 = 0 To 15
                intIndex2 = ConvertGuidIndex(intIndex1)
                strTgtGuid = strTgtGuid & Mid$(strSrcGuid, (intIndex2 * 2) + 1, 2)
                If intIndex1 = 3 Or intIndex1 = 5 Or intIndex1 = 7 Or intIndex1 = 9 Then
                    strTgtGuid = strTgtGuid & "-"
                End If
            Next intIndex1
            strTgtGuid = strTgtGuid & "}"
        Else
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist]." & strFunctionName, "strSrcGuid length must be 32"
        End If
        
    ElseIf eGuidStyle = guidBinary Then
        ' Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
        
        Debug.Assert Len(strSrcGuid) = 38
        
        If Len(strSrcGuid) = 38 Then
            Dim intOffset As Integer
            strTgtGuid = ""
            intOffset = 2
            For intIndex1 = 0 To 15
                intIndex2 = ConvertGuidIndex(intIndex1)
                strTgtGuid = strTgtGuid & Mid$(strSrcGuid, (intIndex2 * 2) + intOffset, 2)
                If intIndex1 = 3 Or intIndex1 = 5 Or intIndex1 = 7 Or intIndex1 = 9 Then
                    intOffset = intOffset + 1
                End If
            Next intIndex1
        Else
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist]." & strFunctionName, "strSrcGuid length must be 38"
        End If
    
    ElseIf eGuidStyle = guidLiteral Then
        ' Convert guid into a format that can be used as a literal input parameter to the database.
        ' This assumes that the database type is raw for Oracle, or binary/varbinary for SQL Server.
        
        If Len(strSrcGuid) = 38 Then
            ' Guid is in hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}",
            ' so convert to binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
            strSrcGuid = sqlFormatGuid(strSrcGuid, guidBinary)
        End If
        
        Debug.Assert Len(strSrcGuid) = 32
        
        If Len(strSrcGuid) = 32 Then
            Select Case GetDbProvider
            Case omiga4DBPROVIDERSQLServer
                ' e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                strTgtGuid = "0x" & strSrcGuid
            Case omiga4DBPROVIDEROracle
                ' e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                strTgtGuid = "HEXTORAW('" & strSrcGuid & "')"
            End Select
        End If
        
    Else
        Debug.Assert 0
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist]." & strFunctionName, "strSrcGuid length must be 32"
    End If
    
    sqlFormatGuid = strTgtGuid

End Function

Public Function sqlGuidToString(ByRef rbytGuid() As Byte, Optional ByVal eDataType As DataTypeEnum = adVarBinary, Optional ByVal eGuidStyle As GUIDSTYLE = guidBinary) As String
' header ----------------------------------------------------------------------------------
' description:  Converts a guid returned (as binary) from the database into a guid string.
' pass:
'   rbytGuid   The byte array to be converted. This can be an adBinary, adVarBinary, or
'                   adGUID read from the database.
'
'   eDataType       The ADO data type of the byte array. Unless you explicitly set the
'                   data type, e.g., in ADODB.CommandCreateParameter, it defaults to:
'
'                   Database Type               ADO Type
'                   ------------------------------------
'                   Oracle raw                  adVarBinary
'                   SQL Server binary           adBinary
'                   SQL Server varbinary        adVarBinary
'                   SQL Server uniqueidentifier adGUID
'
'                   The datatype on the database can be coerced to a different ADO data type
'                   by explicitly setting the type, e.g., in ADODB.CommandCreateParameter.
'
'                   If you explicitly set the data type to adGUID for SQL Server, you MUST
'                   set eDataType to adGUID when calling sqlGuidToString, otherwise you can
'                   allow eDataType to default to adVarBinary.
'
'                   The default of adVarBinary matches the data type returned for binary fields
'                   in the Omiga 4 Phase 2 Oracle database.
'
'   eGuidStyle      The required style for the resulting guid string, e.g.,
'                       guidBinary = "DA6DA163412311D4B5FA00105ABB1680".
'                       guidHyphen = "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
'                   The default of guidBinary gives the same format as Omiga 4 Phase 2.
'
' return:           The guid string.
'
' AS                28/02/01    First version
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "sqlGuidToString"

    Dim strGuid As String
    
    If eDataType = adGUID And GetDbProvider = omiga4DBPROVIDERSQLServer Then
        ' The data type has explicitly been set to adGUID.
        ' SQL Server adGUID data type is already in hyphenated format.
        strGuid = rbytGuid
        If eGuidStyle = guidBinary Then
            ' Convert from hyphenated to binary format.
            strGuid = sqlFormatGuid(strGuid, guidBinary)
        End If
    ElseIf eDataType = adGUID Or eDataType = adBinary Or eDataType = adVarBinary Then        '
        ' For Oracle the data type defaults to adVarBinary.
        ' Guids are always read from Oracle as binary arrays, irrespective of whether the
        ' ADO data type is adGUID or adVarBinary.
        ' For SQL Server the data type defaults to adBinary.
        ' Convert binary array to guid string.
        strGuid = sqlByteArrayToGuidString(rbytGuid, eGuidStyle)
    Else
        Debug.Assert 0
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist]." & strFunctionName, "eDataType must be adGUID or adBinary or adVarBinary"
    End If
    
    sqlGuidToString = strGuid

End Function

Public Function sqlDateToString(ByVal vvarDate As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
    sqlDateToString = Right(Str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("yyyy", vvarDate)), 4)
End Function

Public Function sqlDateTimeToString(ByVal vvarDate As Variant) As String
' header ----------------------------------------------------------------------------------
' description:       Converts a datetime into a string
' pass:              a datetime
' return:            a string in dd/mm/yyy hh:mm:ss format
' Raise Errors:
'DRC  SYS4714             22/05/02
'------------------------------------------------------------------------------------------
    sqlDateTimeToString = Right(Str(DatePart("d", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("m", vvarDate) + 100), 2) & "/" & Right(Str(DatePart("yyyy", vvarDate)), 4) & _
                          " " & Right(Str(DatePart("h", vvarDate) + 100), 2) & ":" & Right(Str(DatePart("n", vvarDate) + 100), 2) & ":" & Right(Str(DatePart("s", vvarDate) + 100), 2)
End Function

#If GENERIC_SQL Then

Public Function sqlGuidStringToByteArray(ByVal strGuid As String) As Byte()
' header ----------------------------------------------------------------------------------
' description:      Converts a guid string into a byte array.
' pass:
'   strGuid         The guid string to be converted.
'                   Can be in either binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680",
'                   or hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
' return:           The byte array.
' AS                05/03/01    First version
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "sqlGuidStringToByteArray"

    Dim rbytGuid(15) As Byte
    
    If Len(strGuid) = 38 Then
        ' Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
        strGuid = sqlFormatGuid(strGuid, guidBinary)
    End If
    
    If Len(strGuid) = 32 Then
        ' Convert from "DA6DA163412311D4B5FA00105ABB1680" to byte array.
        Dim intIndex As Integer
        For intIndex = 0 To UBound(rbytGuid)
            rbytGuid(intIndex) = CByte("&H" & Mid(strGuid, (intIndex * 2) + 1, 2))
        Next intIndex
    Else
        Debug.Assert 0
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist]." & strFunctionName, "strGuid length must be 32"
    End If
    
    sqlGuidStringToByteArray = rbytGuid

End Function

Public Function sqlByteArrayToGuidString(ByRef rbytByteArray() As Byte, Optional ByVal eGuidStyle As GUIDSTYLE = guidBinary) As String
' header ----------------------------------------------------------------------------------
' description:  Converts a byte array into a guid string.
' pass:
'   rbytByteArray   The byte array to be converted.
'   eGuidStyle      The style of the return string.
'                   guidHyphen is "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
'                   guidBinary is "DA6DA163412311D4B5FA00105ABB1680" format.
' return:           The guid string.
' AS                31/01/01    First version
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "sqlByteArrayToGuidString"

    Dim strGuid As String
    
    Debug.Assert UBound(rbytByteArray) = 15
    
    If UBound(rbytByteArray) = 15 Then
        If eGuidStyle = guidHyphen Then
            ' "{63A16DDA-2341-D411-B5FA-00105ABB1680}" format.
            Dim intIndex1 As Integer
            Dim intIndex2 As Integer
            strGuid = "{"
            
            For intIndex1 = 0 To 15
                intIndex2 = ConvertGuidIndex(intIndex1)
                strGuid = strGuid & Right$(Hex$(rbytByteArray(intIndex2) + &H100), 2)
                If intIndex1 = 3 Or intIndex1 = 5 Or intIndex1 = 7 Or intIndex1 = 9 Then
                    strGuid = strGuid & "-"
                End If
            Next intIndex1
            
            strGuid = strGuid & "}"
        ElseIf eGuidStyle = guidBinary Then
            ' "DA6DA163412311D4B5FA00105ABB1680" format.
            Dim intIndex As Integer
            For intIndex = 0 To 15
                strGuid = strGuid & Right$(Hex$(rbytByteArray(intIndex) + &H100), 2)
            Next intIndex
        Else
            Debug.Assert 0
            errRaiseError _
                omiga4NotImplemented, _
                "[sqlAssist]." & strFunctionName, "GUIDSTYLE must be guidHypen or guidBinary"
        End If
    Else
        Debug.Assert 0
        errRaiseError _
            omiga4NotImplemented, _
            "[sqlAssist]." & strFunctionName, "rbytByteArray size must be 16"
    End If

    sqlByteArrayToGuidString = strGuid

End Function

Private Function ConvertGuidIndex(ByVal intIndex1 As Integer)
' header ----------------------------------------------------------------------------------
' description:  Helper function for sqlByteArrayToGuidString and sqlFormatGuid.
' pass:
' return:
' AS            31/01/01    First version
'------------------------------------------------------------------------------------------
    Dim intIndex2 As Integer
    Select Case intIndex1
    Case 0
        intIndex2 = 3
    Case 1
        intIndex2 = 2
    Case 2
        intIndex2 = 1
    Case 3
        intIndex2 = 0
    Case 4
        intIndex2 = 5
    Case 5
        intIndex2 = 4
    Case 6
        intIndex2 = 7
    Case 7
        intIndex2 = 6
    Case Else
        intIndex2 = intIndex1
    End Select
    ConvertGuidIndex = intIndex2
End Function

#End If 'GENERIC_SQL

