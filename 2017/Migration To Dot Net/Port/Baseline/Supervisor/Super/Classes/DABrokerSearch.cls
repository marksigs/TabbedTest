VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DABrokerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : DABrokerSearch
' Description   : Stores all search criteria and performs the search against the
'                 INTRODUCER table
'
' Change history
' Prog      Date        Description
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Stores the search criteria specified in the gui

Private Type SearchDetails
    sFirmName As String
    sFSARef As String
    sAddress As String
' TW 16/01/2007 EP2_859
    sIntroducerId As String
' TW 16/01/2007 EP2_859 End
End Type

Private m_SearchDetails As SearchDetails

' Set
' Get
Public Sub Find(objDABrokerTable As DABrokerTable)
    On Error GoTo Failed

    Dim sSQL As String
    Dim sCompanyNameSearch As String
    Dim sForenameSearch As String
    Dim sSurnameSearch As String
    Dim sTownSearch As String
    Dim sIntTypeSearch As String
    Dim rs As ADODB.Recordset
    Dim strWhereAnd As String
    Dim strCriteria As String
    
    strWhereAnd = " AND"
    
    'Firmname
    If Len(m_SearchDetails.sFirmName) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " USERFORENAME LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFirmName)
        strCriteria = strCriteria & " OR USERSURNAME LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFirmName)
        strWhereAnd = " AND"
    End If
    
' TW 16/01/2007 EP2_859
    'Introducer Id
    If Len(m_SearchDetails.sIntroducerId) > 0 Then
        If InStr(1, m_SearchDetails.sIntroducerId, "*") = 0 Then
            strCriteria = strCriteria & strWhereAnd & " F.INTRODUCERID = '" & m_SearchDetails.sIntroducerId & "'"
        Else
            strCriteria = strCriteria & strWhereAnd & " F.INTRODUCERID LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sIntroducerId)
        End If
        strWhereAnd = " AND"
    End If
' TW 16/01/2007 EP2_859 End
    
    'FSARef
    If Len(m_SearchDetails.sFSARef) > 0 Then
' TW 16/01/2007 EP2_859
        If InStr(1, m_SearchDetails.sFSARef, "*") = 0 Then
            strCriteria = strCriteria & strWhereAnd & " FSARef = '" & m_SearchDetails.sFSARef & "'"
        Else
' TW 16/01/2007 EP2_859 End
            strCriteria = strCriteria & strWhereAnd & " FSARef LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFSARef)
' TW 16/01/2007 EP2_859
        End If
' TW 16/01/2007 EP2_859 End
        strWhereAnd = " AND"
    End If
        
    If Len(m_SearchDetails.sAddress) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " (STREET LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR DISTRICT LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR TOWN LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR COUNTY LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR I.POSTCODE LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress) & ") "
        strWhereAnd = " AND"
    End If
    
    sSQL = "select DISTINCT F.INTRODUCERID, REPLACE(REPLACE(LTRIM(ISNULL(USERFORENAME, '') + ' ' + ISNULL(USERSURNAME, '')), '  ', ' '), '  ', ' ') AS USERNAME, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(BUILDINGORHOUSENAME, '') + ' ' + ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + ISNULL(FLATNUMBER, '') + ' ' + ISNULL(STREET, '') + ' ' + ISNULL(DISTRICT, '') + ' ' + ISNULL(TOWN, '') + ' ' + ISNULL(COUNTY, '') + ' ' + " & _
                        "ISNULL(I.POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, FSARef " & _
                "from " & _
                "INTRODUCERFIRM F INNER JOIN PRINCIPALFIRM P ON F.PRINCIPALFIRMID = P.PRINCIPALFIRMID " & _
                "INNER JOIN INTRODUCER I ON F.INTRODUCERID = I.INTRODUCERID " & _
                "INNER JOIN ORGANISATIONUSER O ON I.USERID = O.USERID " & _
                "WHERE (ISNULL(P.PACKAGERINDICATOR, 0) = 0) " & _
                strCriteria

    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    TableAccess(objDABrokerTable).SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetFirmName(ByVal sName As String)
    m_SearchDetails.sFirmName = sName
End Sub
Public Sub SetIntroducerID(ByVal sName As String)
' TW 16/01/2007 EP2_859
    m_SearchDetails.sIntroducerId = sName
End Sub
Public Sub SetFSARef(ByVal sName As String)
    m_SearchDetails.sFSARef = sName
End Sub

Public Sub SetAddress(ByVal sAddress As String)
    m_SearchDetails.sAddress = sAddress
End Sub









