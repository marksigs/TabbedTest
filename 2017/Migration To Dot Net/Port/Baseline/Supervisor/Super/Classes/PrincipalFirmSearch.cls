VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrincipalFirmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : FirmPackagerSearch
' Description   : Stores all search criteria and performs the search against the
'                 PrincipalFirm table
'
' Change history
' Prog      Date        Description
' TW        18/11/2006  EP2_132 ECR20/21 Correction for finding Principal Firms
' TW        16/01/2007  EP2_859 - Principal Firms/Network display and search
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Stores the search criteria specified in the gui
Private Type SearchDetails
    sFirmName As String
    sFSARef As String
    sUnit As String
    sAddress As String
End Type

Private m_SearchDetails As SearchDetails

' Set
Public Sub SetFirmName(ByVal sName As String)
    m_SearchDetails.sFirmName = sName
End Sub
Public Sub SetUnit(ByVal sName As String)
' TW 16/01/2007 EP2_859
    m_SearchDetails.sUnit = sName
End Sub
Public Sub SetFSARef(ByVal sName As String)
    m_SearchDetails.sFSARef = sName
End Sub

Public Sub SetAddress(ByVal sAddress As String)
    m_SearchDetails.sAddress = sAddress
End Sub

' Get
Public Function GetFirmName() As String
    GetFirmName = m_SearchDetails.sFirmName
End Function
Public Function GetUnit() As String
' TW 16/01/2007 EP2_859
    GetUnit = m_SearchDetails.sUnit
End Function
Public Function GetFSARef() As String
    GetFSARef = m_SearchDetails.sFSARef
End Function

Public Function GetAddress() As String
    GetAddress = m_SearchDetails.sAddress
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Find
' Description   : Using the criteria set in SearchDetails, performs the search against the table
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Find(objPrincipalFirmTable As PrincipalFirmTable)
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
    
' TW 18/11/2006 EP2_132
'    strWhereAnd = " WHERE "
    strWhereAnd = " WHERE ISNULL(PACKAGERINDICATOR, 0) = 0 AND "
' TW 18/11/2006 EP2_132 End
    
    'Firmname
    If Len(m_SearchDetails.sFirmName) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " PRINCIPALFIRMNAME LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sFirmName)
        strWhereAnd = " AND"
    End If
    
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
    
' TW 16/01/2007 EP2_859
    'Unit
    If Len(m_SearchDetails.sUnit) > 0 Then
        If InStr(1, m_SearchDetails.sUnit, "*") = 0 Then
            strCriteria = strCriteria & strWhereAnd & " UNITID = '" & m_SearchDetails.sUnit & "'"
        Else
            strCriteria = strCriteria & strWhereAnd & " UNITID LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sUnit)
        End If
        strWhereAnd = " AND"
    End If
' TW 16/01/2007 EP2_859 End
    
    'Address
    If Len(m_SearchDetails.sAddress) > 0 Then
        strCriteria = strCriteria & strWhereAnd & " (ADDRESSLINE1 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE2 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE3 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE4 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE5 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR ADDRESSLINE6 LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress)
        strCriteria = strCriteria & " OR POSTCODE LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sAddress) & ") "
        strWhereAnd = " AND"
    End If
    
' TW 16/01/2007 EP2_859
'    sSQL = "Select * " & _
'        "From PRINCIPALFIRM" & _
'        strCriteria
    sSQL = "Select PRINCIPALFIRMID, FSARef, UNITID, PrincipalFirmName, " & _
                "REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "POSTCODE " & _
                        "From PRINCIPALFIRM" & _
        strCriteria
' TW 16/01/2007 EP2_859 End

    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    TableAccess(objPrincipalFirmTable).SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



