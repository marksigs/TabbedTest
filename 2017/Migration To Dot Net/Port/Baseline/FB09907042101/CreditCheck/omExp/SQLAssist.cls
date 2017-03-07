VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SQLAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Helper object for handling database-specific
'               formatting when building SQL statements
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' RF    16/07/99 Created
' MCS   02/09/99 GetSystemDate added
' MCS   04/10/99 GetFormattedValue modified to raise error
' RF    22/10/99 Fixed DateTimeToString
' DJP   25/04/00 Added TruncateDateColumn
' SR    02/05/00 Modified method FormatWildCardedString, FormatString
'                (Added new parameter blnIsInputToADOCommandObject)
' APS   08/09/00 Corrected the SQL Server CONVERT statement to convert dates
'                into ANSI format
' DJP   23/01/01 Added GetColumnDayOfWeek
' AS    05/03/01 CC012 Added GUID support for SQL Server and Oracle
' DM	14/03/01 SYS1949 Removed raise error from SQLServer side of run time switch.
'------------------------------------------------------------------------------------------
Option Explicit

Private Const cstrAPOSTROPHE As String = "'"
Private Const cstrWILDCARD As String = "*"
 
 ' Private variables
 Private m_objErrAssist As ErrAssist
 Private m_objADOAssist As ADOAssist
 
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

#If GENERIC_SQL Then
Public Enum GUIDSTYLE
    guidBinary = 0
    guidHyphen = 1
    guidLiteral = 2
End Enum
#End If

Private Sub Class_Initialize()
    Set m_objErrAssist = New ErrAssist
    Set m_objADOAssist = New ADOAssist
End Sub

Private Sub Class_Terminate()
    Set m_objErrAssist = Nothing
    Set m_objADOAssist = Nothing
End Sub

#If GENERIC_SQL Then

Public Function FormatGuid(ByVal strSrcGuid As String, Optional ByVal eGuidStyle As GUIDSTYLE = guidLiteral) As String
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
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "strSrcGuid length must be 32"
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
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "strSrcGuid length must be 38"
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
        
        If Len(strSrcGuid) = 32 Then
            Select Case m_objADOAssist.GetDBEngine
            Case SQLServer
                ' e.g., "0xDA6DA163412311D4B5FA00105ABB1680"
                strTgtGuid = "0x" & strSrcGuid
            Case Oracle
                ' e.g., "HEXTORAW('DA6DA163412311D4B5FA00105ABB1680')"
                strTgtGuid = "HEXTORAW('" & strSrcGuid & "')"
            End Select
        End If
        
    Else
        Debug.Assert 0
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "strGuid length must be 32"
    End If
    
    FormatGuid = strTgtGuid

    Exit Function

FormatGuidVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

#Else

Public Function FormatGuid(ByVal strGuid As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'   formatted string
'------------------------------------------------------------------------------------------
On Error GoTo errhandler

    
    Select Case m_objADOAssist.GetDBEngine
    Case SQLServer
        strGuid = "0x" & strGuid
    Case Oracle
        strGuid = "hextoraw('" & strGuid & "')"
    End Select
    
    
    FormatGuid = strGuid
    
    Exit Function

errhandler:

End Function

#End If     ' GENERIC_SQL

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
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 551, _
                "Wildcard found at a position other than the end of a string: " & vstrIn
        End If
    End If
        
    
    FormatWildcardedString = strTemp
    
    Exit Function
    
FormatWildcardedStringVbErr:
    
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    
    FormatWildcardedString = ""

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetSystemDate() As String
' header ----------------------------------------------------------------------------------
' description:  Returns the system date as used in as SQL search
' pass:
' return:       system date as used in a SQL search
'MCS   02/09/99 Created
'------------------------------------------------------------------------------------------
On Error GoTo GetSystemDateVbErr

    Const strFunctionName As String = "GetSystemDate"
    
    
    Dim eDbEngineType As DBENGINETYPE
    eDbEngineType = m_objADOAssist.GetDBEngine
    Dim strDate As String
    
     Select Case eDbEngineType
        Case SQLServer
            ' DLM SQLServer Port
            strDate = "GetDate()"
         Case Oracle
            strDate = "SYSDATE"
        Case Else
            ' not implemented
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        End Select
    
    GetSystemDate = strDate
    
        
    Exit Function
    
GetSystemDateVbErr:
    
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    
    GetSystemDate = ""

    Err.Raise Err.Number, Err.Source, Err.Description
    
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

Private Function CheckApostrophes(ByVal strIn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Doubles each ' in a string. This is needed for names with
'   an apostrophe in them. E.g. "O'Conner" converts to "O''Conner"
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo errhandler
    
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
    
    Exit Function

errhandler:

End Function

Public Function TruncateDateColumn(strDateColumn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Takes a column name and puts the SQL arround it to set the time portion to all 0's
' pass:
' return:
'------------------------------------------------------------------------------------------
    On Error GoTo TruncateDateColumnVbErr
    Const strFunctionName As String = "TruncateDateColumn"
    Dim eDbEngineType As DBENGINETYPE
    Dim strColumn As String
    Dim strTruncatedDate As String
    
    If Len(strDateColumn) > 0 Then
        strColumn = UCase(strDateColumn)
        
        eDbEngineType = m_objADOAssist.GetDBEngine
            
        Select Case eDbEngineType
            Case SQLServer
                strTruncatedDate = "CONVERT(CHAR, " & strColumn & " , 6 )"
            Case Oracle
                strTruncatedDate = "trunc(" & strColumn & ")"
            
            Case Else
                ' not implemented
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        End Select
    Else
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingFieldDesc
    End If
    
    TruncateDateColumn = strTruncatedDate
    
    Exit Function
TruncateDateColumnVbErr:
    
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FormatDate(ByVal vdtmIn As Date, _
    Optional ByVal vdtfDateTimeFormat As DATETIMEFORMAT = dtfDate) As String
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
    
    Dim eDbEngineType As DBENGINETYPE
    eDbEngineType = m_objADOAssist.GetDBEngine
    
    Select Case vdtfDateTimeFormat
    Case dtfDateTime
        ' date and time format
        
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case SQLServer
            strDate = "CONVERT(DATETIME,'" & strDate & "',102)"
        Case Oracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Case Else
            ' not implemented
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
        End Select
    Case dtfDate
        ' date only but time defaulted to 00:00:00
        
        strDate = Format(vdtmIn, "YYYY-MM-DD HH:MM:SS")
        Select Case eDbEngineType
        Case SQLServer
            ' DLM SQLServer Port
            strDate = "CONVERT(DATETIME,'" & strDate & "',120)"
        Case Oracle
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        End Select
    Case dtfTime
        ' not implemented
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented, _
            " Formatting as Time only has not been implemented"
    Case Else
        ' not implemented
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    End Select
    
    
    FormatDate = strDate
    
    Exit Function
    
FormatDateVbErr:
    
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    
    FormatDate = ""

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FormatDateString(ByRef rstrDate As String, _
    Optional ByVal vdtfRequiredFormat As DATETIMEFORMAT = dtfDateTime) As String
' header ----------------------------------------------------------------------------------
' description:
'   Formats 'dd/mm/yyyy' date to appropriate SQL.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo FormatDateStringVbErr

    Const strFunctionName As String = "FormatDateString"


    
    Dim dtmDate As Date
    dtmDate = rstrDate
    
    Select Case vdtfRequiredFormat
    
        Case dtfDateTime
            If Len(rstrDate) < 19 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidDateTimeFormat, _
                "Enumerated vdtfRequiredFormat value of " & vdtfRequiredFormat & "Requested for DateTime"
            End If
        Case dtfDate
            If Len(rstrDate) <> 10 Then
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4InvalidDateTimeFormat, _
                 "Enumerated vdtfRequiredFormat value of " & vdtfRequiredFormat & " Requested for Date"
            Else
                dtmDate = rstrDate & " 00:00:00" ' default DATES to date + 00:00:00
            End If
        Case dtfTime
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented, _
                " Formatting as Time only has not been implemented"
        Case Else
                m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    
    End Select
    
    FormatDateString = FormatDate(dtmDate, vdtfRequiredFormat)
    
    
    Exit Function

FormatDateStringVbErr:

        If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    FormatDateString = ""

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

#If GENERIC_SQL Then

Public Function GuidToString(ByRef rbytGuid() As Byte, Optional ByVal eDataType As DataTypeEnum = adVarBinary, Optional ByVal eGuidStyle As GUIDSTYLE = guidBinary) As String
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
    
    If eDataType = adGUID And m_objADOAssist.GetDBEngine = SQLServer Then
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
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "eDataType must be adGUID or adBinary or adVarBinary"
    End If
    
    GuidToString = strGuid

GuidToStringErr:

End Function

#Else

Public Function GuidToString(bytArray() As Byte) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strGuid As String
    
    strGuid = ""
    
    For i = 0 To 15
        strGuid = strGuid & Right$(Hex$(bytArray(i) + &H100), 2)
    Next i
    
    GuidToString = strGuid

End Function

#End If ' GENERIC_SQL

Public Function DateToString(ByVal vvarDate As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    
    DateToString = Format$(vvarDate, "DD/MM/YYYY")
        
End Function

Public Function DateTimeToString(ByVal vvarDate As Variant) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'   Date and time string, e.g. "22/10/1999 14:12:37"
'------------------------------------------------------------------------------------------
    
    DateTimeToString = Format$(vvarDate, "DD/MM/YYYY HH:NN:SS")
        
End Function

Public Function GetFormattedValue(ByVal vstrValue As String, _
    ByVal veDbDataType As DBDATATYPE) As String
' header ----------------------------------------------------------------------------------
' description:
'   Format the value vstrValue into the SQL string appropriate to its datatype as
'   indicated by veDbDataType
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "GetFormattedValue"
    
    Dim strFormattedValue As String
    
    If vstrValue = "" Then
        ' If value is blank/null then always return the NULL value
        GetFormattedValue = "NULL"
        Exit Function
    End If


    Select Case veDbDataType
    Case dbdtGuid
        strFormattedValue = FormatGuid(vstrValue)
    Case dbdtString
        strFormattedValue = FormatString(vstrValue)
    Case dbdtInt, _
        dbdtComboId, _
        dbdtCurrency, _
        dbdtBoolean, _
        dbdtDouble, _
        dbdtLong
        ' value is unchanged
        strFormattedValue = vstrValue
    Case dbdtDate
        strFormattedValue = FormatDateString(vstrValue, dtfDate)
    Case dbdtDateTime
        strFormattedValue = FormatDateString(vstrValue, dtfDateTime)
    Case Else
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InValidDataTypeValue, _
            "Data type " & veDbDataType & " is not valid"
    End Select
        
    GetFormattedValue = strFormattedValue
    
    
End Function

Public Function GetColumnDayOfWeek(strColumn As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Sets the required SQL to allow a SQL query to search on the day portion of a date. For
'   example, on Oracle we need: SELECT DATE_FIELD FROM TABLE WHERE TO_CHAR(DATE_FIELD,'D') = 1
'   which would find only those dates who fall on a Sunday. This method returns the
'   "TO_CHAR_(DATE_FIELD,'D')" part.
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim strDayOfWeek As String
    Const strFunctionName As String = "GetColumnDayOfWeek"
    Dim eDbEngineType As DBENGINETYPE

    On Error GoTo GetColumnDayOfWeekVbErr
    
    If Len(strColumn) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Column name must be entered"
    End If
    
    eDbEngineType = m_objADOAssist.GetDBEngine
    
    Select Case eDbEngineType
    Case SQLServer
        strDayOfWeek = "DATEPART(WEEKDAY, " & strColumn & ")"
    Case Oracle
        strDayOfWeek = "TO_CHAR(" & strColumn & ",'D')"
    Case Else
        ' not implemented
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
    End Select

    GetColumnDayOfWeek = strDayOfWeek
    
    Exit Function

GetColumnDayOfWeekVbErr:

        If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

#If GENERIC_SQL Then

Public Function GuidStringToByteArray(ByVal strGuid As String) As Byte()
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
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "strGuid length must be 32"
    End If
    
    GuidStringToByteArray = rbytGuid

    Exit Function

GuidStringToByteArrayVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
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
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "GUIDSTYLE must be guidHypen or guidBinary"
        End If
    Else
        Debug.Assert 0
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "rbytByteArray size must be 16"
    End If

    ByteArrayToGuidString = strGuid

    Exit Function

ByteArrayToGuidStringVbErr:
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
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

#End If 'GENERIC_SQL
