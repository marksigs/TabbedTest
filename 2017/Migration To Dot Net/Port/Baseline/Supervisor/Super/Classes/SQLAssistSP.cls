VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLAssistSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :
' Description   :
' Change history
' Prog      Date        Description
' DJP       09/06/01    Created. Supervisor specific Oracle/SQL Server conversion routines
' DJP       21/02/02    BM0318 - updated file to latest Core version to fix undefined variable errors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllTablesName
' Description   : Returns a string that contains the name of the table that contains a list of
'                 all tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        21/03/2007  EP2_1224 - Run-time error if combo valuename entered with apostrophe
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const cstrAPOSTROPHE As String = "'"
Private Const cstrWILDCARD As String = "*"

Public Enum GUIDSTYLE
    guidBinary = 0
    guidHyphen = 1
    guidLiteral = 2
End Enum

Public Enum DATETIMEFORMAT
    dtfDateTime = 0      ' date and time
    dtfDate = 1          ' date only
    dtfTime = 2          ' time only
End Enum

Public Function GetAllTablesName(Optional sDatabaseKey As String = "") As String
    On Error GoTo GetAllTablesNameVbErr:
    Dim nType As enumDatabaseTypes
    Dim sTable As String
    
    nType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
    
    If nType = INDEX_ORACLE Then
        sTable = "USER_TABLES"
    ElseIf nType = INDEX_SQL_SERVER Then
        sTable = "SYSOBJECTS"
    Else
        g_clsErrorHandling.RaiseError errDatabaseNotSupported
    End If
    
    GetAllTablesName = sTable
    
    Exit Function
GetAllTablesNameVbErr:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllTablesFieldName
' Description   : Returns a string that contains the name of the field that contains a list of
'                 all tables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAllTablesFieldName(Optional sDatabaseKey As String = "") As String
    On Error GoTo Failed:
    Dim nType As enumDatabaseTypes
    Dim sField As String
    
    nType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)

    If nType = INDEX_ORACLE Then
        sField = "TABLE_NAME"
    ElseIf nType = INDEX_SQL_SERVER Then
        sField = "NAME"
    Else
        g_clsErrorHandling.RaiseError errDatabaseNotSupported
    End If
    
    GetAllTablesFieldName = sField
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAllSynonymsName
' Description   : Returns a string that contains the name of the table that contains a list of
'                 all synonyms.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAllSynonymsName(Optional sDatabaseKey As String = "") As String
    On Error GoTo Failed:
    Dim nType As enumDatabaseTypes
    Dim sName As String
    
    nType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
    
    If nType = INDEX_ORACLE Then
        sName = "USER_SYNONYMS"
    ElseIf nType = INDEX_SQL_SERVER Then
        sName = "SYSOBJECTS"
    Else
        g_clsErrorHandling.RaiseError errDatabaseNotSupported
    End If
    
    GetAllSynonymsName = sName
    
    Exit Function
Failed:
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function
Public Function FormatDate(ByVal vdtmIn As Date, _
    Optional ByVal vdtfDateTimeFormat As DATETIMEFORMAT = dtfDate, Optional sDatabaseKey As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   Formats a date to the appropriate SQL.
' pass:
'   vdtmIn                  date to be formatted
'   vdtfDateTimeFormat      required format
' return:
'   FormatDate              formatted value
'------------------------------------------------------------------------------------------
On Error GoTo FormatDateVbErr

    Const strFunctionName As String = "FormatDate"
    Dim strDate As String
    
    Dim eDbEngineType As enumDatabaseTypes
    eDbEngineType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
    
    Select Case vdtfDateTimeFormat
    Case dtfDateTime
        ' date and time format
        
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case INDEX_SQL_SERVER
            strDate = "CONVERT(DATETIME,'" & strDate & "',102)"
        Case INDEX_ORACLE
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
            ' not implemented
            g_clsErrorHandling.RaiseError errDatabaseNotSupported
        End Select
    Case dtfDate
        ' date only but time defaulted to 00:00:00
        
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case INDEX_SQL_SERVER
            ' DLM SQLServer Port
            strDate = "CONVERT(DATETIME,'" & strDate & "',120)"
        Case INDEX_ORACLE
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    Case Else
            g_clsErrorHandling.RaiseError errDatabaseNotSupported
        End Select
    Case dtfTime
        ' not implemented
         g_clsErrorHandling.RaiseError errGeneralError, " Formatting as Time only has not been implemented"
    Case Else
        ' not implemented
        g_clsErrorHandling.RaiseError errDatabaseNotSupported
    End Select
    
    FormatDate = strDate
    
    Exit Function
    
FormatDateVbErr:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function TruncateDateColumn(strDateColumn As String, Optional sDatabaseKey As String = "") As String
' header ----------------------------------------------------------------------------------
' description:
'   Takes a column name and puts the SQL arround it to set the time portion to all 0's
' pass:
' return:
'------------------------------------------------------------------------------------------
    On Error GoTo TruncateDateColumnVbErr
    Const strFunctionName As String = "TruncateDateColumn"
    Dim eDbEngineType As enumDatabaseTypes
    Dim strColumn As String
    Dim strTruncatedDate As String
    
    If Len(strDateColumn) > 0 Then
        strColumn = UCase(strDateColumn)
        
        eDbEngineType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
            
        Select Case eDbEngineType
            Case INDEX_SQL_SERVER
                strTruncatedDate = "CONVERT(DATETIME, " & strColumn & " , 6 )"
            Case INDEX_ORACLE
                strTruncatedDate = "trunc(" & strColumn & ")"
            Case Else
                ' not implemented
                g_clsErrorHandling.RaiseError errDatabaseNotSupported
        End Select
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "DATE field is empty"
    End If
    
    TruncateDateColumn = strTruncatedDate
    
    Exit Function
TruncateDateColumnVbErr:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function GetSystemDate(Optional sDatabaseKey As String = "") As String
' header ----------------------------------------------------------------------------------
' description:  Returns the system date as used in as SQL search
' pass:
' return:       system date as used in a SQL search
'------------------------------------------------------------------------------------------
On Error GoTo GetSystemDateVbErr
    Const strFunctionName As String = "GetSystemDate"
    Dim eDbEngineType As enumDatabaseTypes
    
    eDbEngineType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
    Dim strDate As String
    
     Select Case eDbEngineType
        Case INDEX_SQL_SERVER
            ' DLM SQLServer Port
            strDate = "GetDate()"
         Case INDEX_ORACLE
            strDate = "SYSDATE"
        Case Else
            ' not implemented
            g_clsErrorHandling.RaiseError errDatabaseNotSupported
        End Select
    
    GetSystemDate = strDate
    
    Exit Function
GetSystemDateVbErr:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function
Public Function FormatGuid(ByVal strSrcGuid As String, Optional ByVal eGuidStyle As GUIDSTYLE = guidLiteral, Optional sDatabaseKey As String = "") As String
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
    On Error GoTo FormatGuidVbErr
    
    Const strFunctionName As String = "FormatGuid"
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
            g_clsErrorHandling.RaiseError errGeneralError, "strSrcGuid length must be 32"
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
            g_clsErrorHandling.RaiseError errGeneralError, "strSrcGuid length must be 38"
        End If
    
    ElseIf eGuidStyle = guidLiteral Then
        ' Convert guid into a format that can be used as a literal input parameter to the database.
        ' This assumes that the database type is raw for Oracle, or binary/varbinary for SQL Server.
        
        If Len(strSrcGuid) = 38 Then
            ' Guid is in hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}",
            ' so convert to binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680".
            strSrcGuid = FormatGuid(strSrcGuid, guidBinary)
        End If
        
        Debug.Assert Len(strSrcGuid) = 32
        Dim nDatabaseType As enumDatabaseTypes
        nDatabaseType = g_clsDataAccess.GetDatabaseType(sDatabaseKey)
        
        If Len(strSrcGuid) = 32 Then
            Select Case nDatabaseType
            Case INDEX_SQL_SERVER
                ' e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                strTgtGuid = "0x" & strSrcGuid
            Case INDEX_ORACLE
                ' e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                strTgtGuid = "HEXTORAW('" & strSrcGuid & "')"
            End Select
        End If
        
    Else
        Debug.Assert 0
        g_clsErrorHandling.RaiseError errGeneralError, "strGuid length must be 32"
    End If
    
    FormatGuid = strTgtGuid

    Exit Function

FormatGuidVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION

End Function
Private Function ConvertGuidIndex(ByVal intIndex1 As Integer)
' header ----------------------------------------------------------------------------------
' description:  Helper function for ByteArrayToGuidString and FormatGuid.
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
Public Function FormatSQLDate(sDate As String) As String
    sDate = GetFullYear(sDate)
    ' DJP SQL Server port
    FormatSQLDate = g_clsSQLAssistSP.FormatDate(sDate)
End Function
Public Function GuidStringToByteArray(ByVal strGuid As Variant) As Byte()
' header ----------------------------------------------------------------------------------
' description:      Converts a guid string into a byte array.
' pass:
'   strGuid         The guid string to be converted.
'                   Can be in either binary format, e.g., "DA6DA163412311D4B5FA00105ABB1680",
'                   or hyphenated format, e.g., "{63A16DDA-2341-D411-B5FA-00105ABB1680}".
' return:           The byte array.
' AS                05/03/01    First version
'------------------------------------------------------------------------------------------
    On Error GoTo GuidStringToByteArrayVbErr
    Const strFunctionName As String = "GuidStringToByteArray"

    Dim rbytGuid(15) As Byte
    
    If TypeName(strGuid) <> "Byte()" Then
        If Len(strGuid) = 38 Then
            ' Convert from "{63A16DDA-2341-D411-B5FA-00105ABB1680}" to "DA6DA163412311D4B5FA00105ABB1680"
            strGuid = FormatGuid(strGuid, guidBinary)
        End If
        
        If Len(strGuid) = 32 Then
            ' Convert from "DA6DA163412311D4B5FA00105ABB1680" to byte array.
            Dim intIndex As Integer
            For intIndex = 0 To UBound(rbytGuid)
                rbytGuid(intIndex) = CByte("&H" & Mid(strGuid, (intIndex * 2) + 1, 2))
            Next intIndex
        Else
            Debug.Assert 0
            g_clsErrorHandling.RaiseError errGeneralError, "strGuid length must be 32"
        End If
        GuidStringToByteArray = rbytGuid
    Else
        GuidStringToByteArray = strGuid
    End If


    Exit Function

GuidStringToByteArrayVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function
Public Function GuidToString(ByRef rbytGuid() As Byte, Optional ByVal eDataType As DataTypeEnum = adVarBinary, Optional ByVal eGuidStyle As GUIDSTYLE = guidBinary, Optional sDatabase As String = "") As String
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
'                   set eDataType to adGUID when calling GuidToString, otherwise you can
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
On Error GoTo GuidToStringErr
    Const strFunctionName As String = "GuidToString"

    Dim strGuid As String
    Dim nDatabaseType As enumDatabaseTypes
    nDatabaseType = g_clsDataAccess.GetDatabaseType(sDatabase)
    
    If eDataType = adGUID And nDatabaseType = INDEX_SQL_SERVER Then
        ' The data type has explicitly been set to adGUID.
        ' SQL Server adGUID data type is already in hyphenated format.
        strGuid = rbytGuid
        If eGuidStyle = guidBinary Then
            ' Convert from hyphenated to binary format.
            strGuid = FormatGuid(strGuid, guidBinary)
        End If
    ElseIf eDataType = adGUID Or eDataType = adBinary Or eDataType = adVarBinary Then        '
        ' For Oracle the data type defaults to adVarBinary.
        ' Guids are always read from Oracle as binary arrays, irrespective of whether the
        ' ADO data type is adGUID or adVarBinary.
        ' For SQL Server the data type defaults to adBinary.
        ' Convert binary array to guid string.
        strGuid = ByteArrayToGuidString(rbytGuid, eGuidStyle)
    Else
        Debug.Assert 0
        g_clsErrorHandling.RaiseError errGeneralError, "eDataType must be adGUID or adBinary or adVarBinary"
    End If
    
    GuidToString = strGuid

GuidToStringErr:

End Function
Public Function FormatString( _
            ByVal vstrIn As String, _
            Optional ByVal blnIsInputToADOCommandObject As Boolean = False) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert a string into appropriate SQL for a string datatype
' pass:
' return:
'   formatted string
'------------------------------------------------------------------------------------------
    Dim strDelimiter As String
    strDelimiter = GetStringDelimiter
    
    If Not blnIsInputToADOCommandObject Then
        FormatString = strDelimiter + CheckApostrophes(vstrIn) + strDelimiter
    Else  ' Do not check for aposterphes and do not add delimiters
        FormatString = vstrIn
    End If
    Exit Function

End Function
Public Function ByteArrayToGuidString(ByRef rbytByteArray() As Byte, Optional ByVal eGuidStyle As GUIDSTYLE = guidBinary) As String
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
    On Error GoTo ByteArrayToGuidStringVbErr
    Const strFunctionName As String = "ByteArrayToGuidString"

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
            g_clsErrorHandling.RaiseError errGeneralError, "GUIDSTYLE must be guidHypen or guidBinary"
        End If
    Else
        Debug.Assert 0
        g_clsErrorHandling.RaiseError errGeneralError, "rbytByteArray size must be 16"
    End If

    ByteArrayToGuidString = strGuid

    Exit Function

ByteArrayToGuidStringVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.DESCRIPTION
End Function
Private Function CheckApostrophes(ByVal strIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Doubles each ' in a string. This is needed for names with
'   an apostrophe in them. E.g. "O'Conner" converts to "O''Conner"
' pass:
' return:
'------------------------------------------------------------------------------------------
'TW Rewritten 21/03/2007 EP2_1224
'On Error GoTo errhandler
'
'    Dim nPos As Integer
'    Dim strOut As String
'    Dim nLen As Integer
'    Dim thisChar As String
'    Dim i As Integer
'
'    nLen = Len(strIn)
'
'    For nPos = 1 To nLen Step 1
'        thisChar = Mid$(strIn, nPos, 1)
'        strOut = strOut + thisChar
'        If thisChar = cstrAPOSTROPHE Then
'            ' add an extra apostrophe
'            strOut = strOut + cstrAPOSTROPHE
'        End If
'    Next nPos
'
'    CheckApostrophes = strOut
'
'    Exit Function
'
'errhandler:
    CheckApostrophes = Replace(strIn, "'", "''")

End Function
Public Function FormatWildcardedString(ByVal vstrIn As String, _
    ByRef rblnWildcardFound As Boolean, _
    Optional ByVal blnIsInputToADOCommandObject As Boolean = False) As String
' header ----------------------------------------------------------------------------------
' description:
'   Convert a string into appropriate SQL, including the substitution of
'   wildcard character into correct SQL syntax
'   E.g. "Smit*" is converted to "'Smit%'"
'   SR : 02/05/00 - If the output of this is to be assigned to a parameter of
'                   a command object, then do not check for aposterphes
' pass:
' return:
'   FormatWildcardedString  formatted string
' Raise Errors:             omiga4err104
'------------------------------------------------------------------------------------------
On Error GoTo FormatWildcardedStringVbErr

    Const strFunctionName As String = "FormatWildcardedString"
    
    Dim strTemp As String
    Dim lngPos As Long
    
    rblnWildcardFound = False
    
    ' check if input string contains wildcard character;
    ' this test will also cope with zero length input strings
    lngPos = InStr(1, vstrIn, cstrWILDCARD, vbTextCompare)
    If lngPos = 0 Then
        ' no wildcard
        If Not blnIsInputToADOCommandObject Then
            strTemp = FormatString(vstrIn, False)  'Check for Aposterphes
        Else
            strTemp = FormatString(vstrIn, True)  ' Do not check for Apostrphes
        End If
    Else
        Dim lngLen As Long
        lngLen = Len(vstrIn)
        If lngPos = lngLen Then
            ' wildcard is valid, i.e. at the end of the input string
            
            strTemp = Mid(vstrIn, 1, lngLen - 1) & "%"
            
            If Not blnIsInputToADOCommandObject Then
                strTemp = FormatString(strTemp, False)
            Else
                strTemp = FormatString(strTemp, True)
            End If
            rblnWildcardFound = True
        Else
            ' "Invalid search string"
            g_clsErrorHandling.RaiseError errGeneralError, "Wildcard found at a position other than the end of a string: " & vstrIn
        End If
    End If
        
    
    FormatWildcardedString = strTemp
    
    Exit Function
FormatWildcardedStringVbErr:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetLikeOperator() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    GetLikeOperator = "LIKE"

End Function
Private Function GetStringDelimiter() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------

    GetStringDelimiter = "'"

End Function

