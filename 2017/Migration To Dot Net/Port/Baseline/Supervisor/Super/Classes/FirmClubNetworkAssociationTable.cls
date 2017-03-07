VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmClubNetAssocTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   FirmClubNetAssocTable
' Description   :   Contains all access functions for the FirmClubNetworkAssociation table.
' Change history
' Change history
' Prog     Date         Description
' TW       17/10/2006   EP2_15 - Created
' TW       06/12/2006  EP2_330 - Error linking a Packager to an a Association
' TW       08/12/2006  EP2_360 - Ar firm networks is currently showing mortgage club. Should be showing Network.
' TW       20/02/2007  EP2_1492 - Error Linking Principal Firms / Networks to Mortgage Clubs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

'Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String

' Local constants
Private Const m_sTableName = "FIRMCLUBNETWORKASSOCIATION"
Private Const DELETE_FLAG_SET As Integer = 1

Private m_sFirstMatch As String

Friend Sub SetFindAvailableBrokerAssociationRecords(strARFirmID As String)
    
'    m_sFirstMatch = "SELECT DISTINCT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, null AS AGREEDPROCFEERATE, null AS VOLPROCFEEADJUSTMENT, null AS FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID" & _
'                    " FROM MORTGAGECLUBNETWORKASSOCIATION M LEFT JOIN FIRMCLUBNETWORKASSOCIATION F ON M.CLUBNETWORKASSOCIATIONID = F.CLUBNETWORKASSOCIATIONID " & _
'                    " WHERE F.ARFIRMID <> '" & strARFirmID & "'"
    
    
'    m_sFirstMatch = "SELECT DISTINCT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, null AS AGREEDPROCFEERATE, null AS VOLPROCFEEADJUSTMENT, null AS FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID" & _
'                    " FROM MORTGAGECLUBNETWORKASSOCIATION M LEFT JOIN FIRMCLUBNETWORKASSOCIATION F ON M.CLUBNETWORKASSOCIATIONID = F.CLUBNETWORKASSOCIATIONID " & _
'                    " WHERE IsNull(PACKAGERINDICATOR, 0) = 0 " & _
'                    " AND M.CLUBNETWORKASSOCIATIONID NOT IN (SELECT CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION WHERE ARFIRMID = '" & strARFirmID & "')"
    
'' TW 06/12/2006 EP2_330
'    m_sFirstMatch = "SELECT DISTINCT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, null AS AGREEDPROCFEERATE, null AS VOLPROCFEEADJUSTMENT, null AS FIRMCLUBSEQNO, M.CLUBNETWORKASSOCIATIONID" & _
'                    " FROM MORTGAGECLUBNETWORKASSOCIATION M LEFT JOIN FIRMCLUBNETWORKASSOCIATION F ON M.CLUBNETWORKASSOCIATIONID = F.CLUBNETWORKASSOCIATIONID " & _
'                    " WHERE IsNull(PACKAGERINDICATOR, 0) = 0 " & _
'                    " AND M.CLUBNETWORKASSOCIATIONID NOT IN (SELECT CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION WHERE ARFIRMID = '" & strARFirmID & "')"
'' TW 06/12/2006 EP2_330 End
    
' TW 08/12/2006 EP2_360
    m_sFirstMatch = "SELECT DISTINCT PRINCIPALFIRMNAME AS MORTGAGECLUBNETWORKASSOCNAME, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, null AS AGREEDPROCFEERATE, null AS VOLPROCFEEADJUSTMENT, null AS FIRMCLUBSEQNO, M.PRINCIPALFIRMID" & _
                    " FROM PRINCIPALFIRM M LEFT JOIN FIRMCLUBNETWORKASSOCIATION F ON M.PRINCIPALFIRMID = F.PRINCIPALFIRMID " & _
                    " WHERE IsNull(PACKAGERINDICATOR, 0) = 0 " & _
                    " AND M.PRINCIPALFIRMID NOT IN (SELECT ISNULL(PRINCIPALFIRMID, 0) FROM FIRMCLUBNETWORKASSOCIATION WHERE ARFIRMID = '" & strARFirmID & "')"
' TW 08/12/2006 EP2_360 End
    
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub

Friend Sub SetFindAvailableFirmAssociationRecords(strPackagerID As String, intPackagerIndicator As Integer)
    
' TW 20/02/2007 EP2_1492
    m_sFirstMatch = "SELECT DISTINCT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
                    " ISNULL(STREET, '') + ' ' + " & _
                    " ISNULL(DISTRICT, '') + ' ' + " & _
                    " ISNULL(TOWN, '') + ' ' + " & _
                    " ISNULL(COUNTY, '')) " & _
                    " AS ADDRESS, null AS AGREEDPROCFEERATE, null AS VOLPROCFEEADJUSTMENT, null AS FIRMCLUBSEQNO, M.CLUBNETWORKASSOCIATIONID" & _
                    " FROM MORTGAGECLUBNETWORKASSOCIATION M LEFT JOIN FIRMCLUBNETWORKASSOCIATION F ON M.CLUBNETWORKASSOCIATIONID = F.CLUBNETWORKASSOCIATIONID " & _
                    " WHERE IsNull(PACKAGERINDICATOR, 0) = " & intPackagerIndicator & _
                    " AND M.CLUBNETWORKASSOCIATIONID NOT IN (SELECT ISNULL(CLUBNETWORKASSOCIATIONID, 0) FROM FIRMCLUBNETWORKASSOCIATION WHERE PRINCIPALFIRMID = '" & strPackagerID & "')"
' TW 20/02/2007 EP2_1492 End
    
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub


Friend Sub SetFindLinkedBrokerAssociationRecords(strARFirmID As String)
    
'    m_sFirstMatch = "SELECT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, AGREEDPROCFEERATE, VOLPROCFEEADJUSTMENT, FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION F INNER JOIN MORTGAGECLUBNETWORKASSOCIATION M ON F.CLUBNETWORKASSOCIATIONID = M.CLUBNETWORKASSOCIATIONID " & _
'                    "WHERE IsNull(PACKAGERINDICATOR, 0) = 0 AND F.ARFIRMID = '" & strARFirmID & "'"
    
' TW 08/12/2006 EP2_360
'    m_sFirstMatch = "SELECT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, AGREEDPROCFEERATE, VOLPROCFEEADJUSTMENT, FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION F INNER JOIN MORTGAGECLUBNETWORKASSOCIATION M ON F.CLUBNETWORKASSOCIATIONID = M.CLUBNETWORKASSOCIATIONID " & _
'                    " WHERE F.ARFIRMID = '" & strARFirmID & "'"
    
    m_sFirstMatch = "SELECT PRINCIPALFIRMNAME AS MORTGAGECLUBNETWORKASSOCNAME, REPLACE(REPLACE(LTRIM(" & _
                        "ISNULL(ADDRESSLINE1, '') + ' ' + ISNULL(ADDRESSLINE2, '') + ' ' + ISNULL(ADDRESSLINE3, '') + ' ' + ISNULL(ADDRESSLINE4, '') + ' ' + ISNULL(ADDRESSLINE5, '') + ' ' + ISNULL(ADDRESSLINE6, '') + ' ' + " & _
                        "ISNULL(POSTCODE, '')) , '  ', ' '), '  ', ' ') AS ADDRESS, " & _
                        "AGREEDPROCFEERATE, VOLPROCFEEADJUSTMENT, FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID, F.PRINCIPALFIRMID FROM FIRMCLUBNETWORKASSOCIATION F INNER JOIN PRINCIPALFIRM M ON F.PRINCIPALFIRMID = M.PRINCIPALFIRMID " & _
                    " WHERE F.ARFIRMID = '" & strARFirmID & "'"
' TW 08/12/2006 EP2_360 End
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub

Friend Sub SetFindLinkedFirmAssociationRecords(strPackagerID As String)
    
'    m_sFirstMatch = "SELECT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
'                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
'                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
'                    " ISNULL(STREET, '') + ' ' + " & _
'                    " ISNULL(DISTRICT, '') + ' ' + " & _
'                    " ISNULL(TOWN, '') + ' ' + " & _
'                    " ISNULL(COUNTY, '')) " & _
'                    " AS ADDRESS, AGREEDPROCFEERATE, VOLPROCFEEADJUSTMENT, FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION F INNER JOIN MORTGAGECLUBNETWORKASSOCIATION M ON F.CLUBNETWORKASSOCIATIONID = M.CLUBNETWORKASSOCIATIONID " & _
'                    " WHERE IsNull(PACKAGERINDICATOR, 0) = 1 AND F.PRINCIPALFIRMID = '" & strPackagerID & "'"
    
    m_sFirstMatch = "SELECT MORTGAGECLUBNETWORKASSOCNAME, LTRIM(ISNULL(BUILDINGORHOUSENAME, '') + ' ' + " & _
                    " ISNULL(BUILDINGORHOUSENUMBER, '') + ' ' + " & _
                    " ISNULL(FLATNUMBER, '') + ' ' + " & _
                    " ISNULL(STREET, '') + ' ' + " & _
                    " ISNULL(DISTRICT, '') + ' ' + " & _
                    " ISNULL(TOWN, '') + ' ' + " & _
                    " ISNULL(COUNTY, '')) " & _
                    " AS ADDRESS, AGREEDPROCFEERATE, VOLPROCFEEADJUSTMENT, FIRMCLUBSEQNO, F.CLUBNETWORKASSOCIATIONID FROM FIRMCLUBNETWORKASSOCIATION F INNER JOIN MORTGAGECLUBNETWORKASSOCIATION M ON F.CLUBNETWORKASSOCIATIONID = M.CLUBNETWORKASSOCIATIONID " & _
                    " WHERE F.PRINCIPALFIRMID = '" & strPackagerID & "'"
    
    m_clsTableAccess.Create m_sTableName, m_sFirstMatch
End Sub





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Class_Initialize
' Description   :   VB Calls this method when the class is created. Do any initialisation here
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    
    m_clsTableAccess.Create m_sTableName
    
    ' Key fields
    colFields.Add "FirmClubSeqNo"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' Order criteria used when executing queries
    m_clsTableAccess.SetOrderCriteria "CLUBNETWORKASSOCIATIONID"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetFirmClubSeqNo() As String
    GetFirmClubSeqNo = CStr(m_clsTableAccess.GetVal("FirmClubSeqNo"))
End Function
Friend Function GetPrincipalFirmID() As String
    GetPrincipalFirmID = CStr(m_clsTableAccess.GetVal("PrincipalFirmID"))
End Function
Friend Function GetPrincipalFirmName() As String
    GetPrincipalFirmName = CStr(m_clsTableAccess.GetVal("PrincipalFirmName"))
End Function
Friend Function GetMortgageClubNetworkAssocName() As String
    GetMortgageClubNetworkAssocName = CStr(m_clsTableAccess.GetVal("MORTGAGECLUBNETWORKASSOCNAME"))
End Function

Friend Function GetARFirmID() As String
    GetARFirmID = CStr(m_clsTableAccess.GetVal("ARFirmID"))
End Function
Friend Function GetAgreedProcFeeRate() As String
    GetAgreedProcFeeRate = CStr(m_clsTableAccess.GetVal("AgreedProcFeeRate"))
End Function
Friend Function GetVolProcFeeAdjustment() As String
    GetVolProcFeeAdjustment = CStr(m_clsTableAccess.GetVal("VolProcFeeAdjustment"))
End Function
Friend Function GetClubNetworkAssociationID() As String
    GetClubNetworkAssociationID = CStr(m_clsTableAccess.GetVal("ClubNetworkAssociationId"))
End Function

Friend Function GetAddress() As String
    GetAddress = CStr(m_clsTableAccess.GetVal("Address"))
End Function




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetFirmClubSeqNo(sValue As Variant) As Boolean
    SetFirmClubSeqNo = m_clsTableAccess.SetVal("FirmClubSeqNo", CVar(sValue))
End Function
Friend Function SetPrincipalFirmID(sValue As Variant) As Boolean
    SetPrincipalFirmID = m_clsTableAccess.SetVal("PrincipalFirmID", CVar(sValue))
End Function
Friend Function SetARFirmID(sValue As Variant) As Boolean
    SetARFirmID = m_clsTableAccess.SetVal("ARFirmID", CVar(sValue))
End Function
Friend Function SetClubNetworkAssociationId(sValue As Variant) As Boolean
    SetClubNetworkAssociationId = m_clsTableAccess.SetVal("ClubNetworkAssociationId", CVar(sValue))
End Function
Friend Function SetDeleteFlag(Optional bDelete As Boolean = False) As Boolean
    Dim nDelete As Integer
    
    If bDelete Then
        nDelete = 1
    Else
        nDelete = 0
    End If
    
    SetDeleteFlag = m_clsTableAccess.SetVal("DELETEFLAG", CVar(nDelete))

End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implements Table Access Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_clsTableAccess.GetIsDeleted()
End Function

Private Sub TableAccess_SetDeleteFields(colFields As Collection)
    m_clsTableAccess.SetDeleteFields colFields
End Sub

Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_clsTableAccess.GetTableDataFromSearch sSearch
End Sub

Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function
Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub
Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub
Private Sub TableAccess_DeleteAllRows()
    m_clsTableAccess.DeleteAllRows
End Sub

Private Sub TableAccess_DeleteRow(colMatchData As Collection)
    m_clsTableAccess.DeleteRow colMatchData
End Sub

Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function

Private Function TableAccess_GetKeyMatchFields()
    Set TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function

Private Function TableAccess_GetKeyMatchValues() As Variant
    Set TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function

Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function
Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function
Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function
Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
        m_clsTableAccess.ValidateData sDesc
End Sub
Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function
Private Sub TableAccess_MoveFirst()
    m_clsTableAccess.MoveFirst
End Sub
Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
End Function
Private Function TableAccess_RecordCount() As Integer
    TableAccess_RecordCount = m_clsTableAccess.RecordCount()
End Function

Private Function TableAccess_SetDate(sField As String, vval As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vval)
End Function

Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub

Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)
    m_clsTableAccess.SetKeyMatchValues colValues
End Sub

Private Function TableAccess_SetNextRow() As Boolean
    TableAccess_SetNextRow = m_clsTableAccess.SetNextRow()
End Function
Private Sub TableAccess_SetPopulateType(enumType As PopulateType)
    
End Sub
Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub

Private Sub TableAccess_SetUpdated()
    m_clsTableAccess.SetUpdated
End Sub
Private Function TableAccess_SetVal(sField As Variant, vval As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vval)
End Function
Private Sub TableAccess_Update()
    m_clsTableAccess.Update
End Sub
Private Sub TableAccess_SetPopulateNull()
    m_clsTableAccess.SetPopulateNull
End Sub

Private Function TableAccess_GetSearch() As String
    TableAccess_GetSearch = m_sSearch
End Function
Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function
' DJP This can be added later if required.
Public Function TableAccess_GetUpdateValues() As Collection
End Function
' DJP This can be added later if required.
'Public Function TableAccess_GetUpdateFields() As Collection
'End Function
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetFirmClubSeqNo()
    colListLine.Add GetMortgageClubNetworkAssocName()
    colListLine.Add GetAddress()
    colListLine.Add GetAgreedProcFeeRate()
    colListLine.Add GetVolProcFeeAdjustment()
    colListLine.Add GetClubNetworkAssociationID()
End Sub
Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub

Public Function TableAccess_GetDatabase() As String
    TableAccess_GetDatabase = m_clsTableAccess.GetDatabase()
End Function

Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_clsTableAccess.SetOrderCriteria sCriteria
End Sub

Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function


Private Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub
