VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntermediaryTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : IntermediaryTable
' Description   : Contains all data access functions for the Intermediary Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DJP       28/06/01    SQL Server port
' DJP       04/08/01    Add access routines for Intermediary table names
' STB       15/11/2001  SYS2550: SQL-Server enabled.
' AW        08/01/02    SYS3560: Added GetOtherSystemNumber()
' STB       11/01/02    SYS3581: GetRowOfData() returns the literal version of
'                       the intermediary type (rather than numeric).
' STB       25/03/02    SYS4312: Added additional methods to tag Optimus numbers
'                       onto records.
' DB        06/01/03    SYS5457 Modify the promotion of Intermediaries
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const m_sTableName = "INTERMEDIARY"
Private Const m_sProgramTypeComboGroup = "BatchProgramType"
Private Const m_sOrganisationTable As String = "INTERMEDIARYORGANISATION"
Private Const m_sIndividualTable As String = "INTERMEDIARYINDIVIDUAL"

' Private data
Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_bIsIndividual As Boolean

Implements TableAccess


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
        
    Dim colFields As Collection
    Dim colDupKeys As Collection
    
    On Error GoTo Failed
    
    Set colFields = New Collection
    Set colDupKeys = New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "INTERMEDIARYGUID"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    'Duplicate keys
    colDupKeys.Add "INTERMEDIARYPANELID"
    m_clsTableAccess.SetDuplicateKeys colDupKeys
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_AddRow
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetIsDeleted
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetIsDeleted() As Boolean
    TableAccess_GetIsDeleted = m_clsTableAccess.GetIsDeleted()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetDeleteFields
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetDeleteFields(colFields As Collection)
    m_clsTableAccess.SetDeleteFields colFields
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetTableDataFromSearch
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_GetTableDataFromSearch(sSearch As String, Optional bUseExistingSearch As Boolean = False)
    m_clsTableAccess.GetTableDataFromSearch sSearch
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetLength
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetLength(sField As String) As Integer
    TableAccess_GetLength = m_clsTableAccess.GetLength(sField)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_CloseRecordSet
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_CloseRecordSet()
    m_clsTableAccess.CloseRecordSet
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_DeleteAllRows
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_DeleteAllRows()
    m_clsTableAccess.DeleteAllRows
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_DeleteRow
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_DeleteRow(colMatchData As Collection)
    m_clsTableAccess.DeleteRow colMatchData
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_DoesRecordExist
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetDate
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetDate(sField As String, Optional ctrl As Object = Nothing, Optional ctrlLabel As Object = Nothing, Optional bCheck As Boolean = True, Optional bGetTime As Boolean = False) As Variant
    TableAccess_GetDate = m_clsTableAccess.GetDate(sField)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetKeyMatchFields
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetKeyMatchFields()
    Set TableAccess_GetKeyMatchFields = m_clsTableAccess.GetKeyMatchFields()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetKeyMatchValues
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetKeyMatchValues() As Variant
    Set TableAccess_GetKeyMatchValues = m_clsTableAccess.GetKeyMatchValues()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetRecordSet
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetRecordSet() As ADODB.Recordset
    Set TableAccess_GetRecordSet = m_clsTableAccess.GetRecordSet()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetTableData
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetTableData(Optional enumPopulateType As PopulateType = POPULATE_KEYS, Optional enumClassOption As ClassOption) As ADODB.Recordset
    Set TableAccess_GetTableData = m_clsTableAccess.GetTableData(enumPopulateType)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetUpdated
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetUpdated() As Boolean
    TableAccess_GetUpdated = m_clsTableAccess.GetUpdated()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetVal
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetVal(sField As Variant) As Variant
    TableAccess_GetVal = m_clsTableAccess.GetVal(sField)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_ValidateData
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_ValidateData(Optional sDesc As String = "")
    m_clsTableAccess.ValidateData sDesc
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_MatchRecords
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_MatchRecords(colFields As Collection, colMatchData As Collection) As ADODB.Recordset
    Set TableAccess_MatchRecords = m_clsTableAccess.MatchRecords(colFields, colMatchData)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_MoveFirst
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_MoveFirst()
    m_clsTableAccess.MoveFirst
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_MoveNext
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_MoveNext() As Boolean
    TableAccess_MoveNext = m_clsTableAccess.MoveNext()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_RecordCount
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_RecordCount() As Integer
    TableAccess_RecordCount = m_clsTableAccess.RecordCount()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetDate
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetKeyMatchFields
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetKeyMatchFields(colFields As Collection)
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetKeyMatchValues
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetKeyMatchValues(colValues As Collection)
    m_clsTableAccess.SetKeyMatchValues colValues
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetNextRow
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_SetNextRow() As Boolean
    TableAccess_SetNextRow = m_clsTableAccess.SetNextRow()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetPopulateType
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetPopulateType(enumType As PopulateType)
    'Stub.
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetRecordSet
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetRecordSet(rs As ADODB.Recordset)
    m_clsTableAccess.SetRecordSet rs
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetUpdated
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetUpdated()
    m_clsTableAccess.SetUpdated
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetVal
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_Update
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_Update()
    m_clsTableAccess.Update
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetPopulateNull
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetPopulateNull()
    m_clsTableAccess.SetPopulateNull
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetSearch
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetSearch() As String
    'Stub.
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetTable
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetTable() As String
    TableAccess_GetTable = m_sTableName
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetUpdateValues
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TableAccess_GetUpdateValues() As Collection
    'Stub.
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_Create
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetRowOfData
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)

    Dim clsIntermediary As Intermediary
    
    'We need the help class to return the string version of the type.
    Set clsIntermediary = New Intermediary
    
    'DB SYS5457 06/01/03 - Remove need to display the type of intermediary
    If m_bIsIndividual Then
        colListLine.Add GetIndividualTitleText
        colListLine.Add GetIndividualForeName
        colListLine.Add GetIndividualSurName
        'colListLine.Add GetIndividualForeName() & " " & GetIndividualSurName() & ", " & clsIntermediary.GetIntermediaryTypeAsText(GetIntermediaryType()), OBJECT_DESCRIPTION_KEY
        colListLine.Add GetIndividualForeName() & " " & GetIndividualSurName(), OBJECT_DESCRIPTION_KEY
    Else
        colListLine.Add GetOrgName
        colListLine.Add GetOrgCommisionNo
        
        'colListLine.Add GetOrgName & ", " & clsIntermediary.GetIntermediaryTypeAsText(GetIntermediaryType()), OBJECT_DESCRIPTION_KEY
        colListLine.Add GetOrgName, OBJECT_DESCRIPTION_KEY
    End If
    
    colListLine.Add GetActiveFrom
    colListLine.Add GetActiveTo
    colListLine.Add GetPanelID
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_DeleteRecords
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    
    On Error GoTo Failed
    
    m_clsTableAccess.DeleteRecords sDatabase
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetDatabase
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetDatabase(sDatabaseKey As String)
    m_clsTableAccess.SetDatabase sDatabaseKey
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetDatabase
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TableAccess_GetDatabase() As String
    TableAccess_GetDatabase = m_clsTableAccess.GetDatabase()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetOrderCriteria
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetOrderCriteria(sCriteria As String)
    m_clsTableAccess.SetOrderCriteria sCriteria
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetDuplicateKeys
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetDuplicateKeys(colDupKeys As Collection)
    m_clsTableAccess.SetDuplicateKeys colDupKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetDuplicateKeys
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetDuplicateKeys() As Collection
    Set TableAccess_GetDuplicateKeys = m_clsTableAccess.GetDuplicateKeys()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_SetCollectionField
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_SetCollectionField(sField As String)
    m_clsTableAccess.SetCollectionField sField
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_GetCollectionField
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TableAccess_GetCollectionField() As String
    TableAccess_GetCollectionField = m_clsTableAccess.GetCollectionField()
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_ApplyFilter
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TableAccess_CancelFilter
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaries
' Description   : Use to retrieve information to construct the top-level treeview nodes.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetIntermediaries(Optional sType As String = "", Optional vGUID As Variant = "")

    Dim sSQL As String
    Dim sIntPrefix As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    'SB: Format SQL appropriate to specific DB server version.
    If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
        sIntPrefix = "Int"
    
        sSQL = "select " & _
                    "intermediaryorganisation.name NAME," & _
                    "intermediaryorganisation.COMPANYCOMMISIONNUMBER, " & _
                    "intermediaryindividual.forename," & _
                    "intermediaryindividual.surname," & _
                    "intermediaryindividual.title," & _
                    sIntPrefix & ".IntermediaryGuid," & _
                    "INT.IsAttachedToIntermediaryguid," & _
                    "IntView.IntermediaryType," & _
                    "IntView.ChildCount " & _
               "From " & _
                    m_sTableName & " INT, " & _
                    "Intermediaryorganisation, " & _
                    "intermediaryindividual, " & _
                    "(SELECT  " & _
                        "count(IsAttachedToIntermediaryGuid) ChildCount, " & _
                        "IsAttachedToIntermediaryGuid, IntermediaryType,IntermediaryGUID " & _
                    "From  " & _
                        "Intermediary  " & _
                    "Where " & _
                        "IsAttachedToIntermediaryGuid is not null " & _
                    "group by " & _
                        "IntermediaryType,IsAttachedToIntermediaryGuid, IntermediaryGUID) IntView " & _
               "Where " & _
                    "INT.IntermediaryGuid = IntView.IsAttachedToIntermediaryGuid(+)"

        If Len(sType) > 0 Then
            sSQL = sSQL & " and " & sIntPrefix & ".IntermediaryType = " & sType
        End If
        
        sSQL = sSQL & " and intermediaryorganisation.INTERMEDIARYGUID(+) = " & sIntPrefix & ".INTERMEDIARYGUID AND intermediaryindividual.INTERMEDIARYGUID(+) = " & sIntPrefix & ".INTERMEDIARYGUID "
        
        If Len(vGUID) > 0 Then
            sSQL = sSQL & " and INT.intermediaryguid = " & g_clsSQLAssistSP.FormatString(vGUID)
        End If
        
        sSQL = sSQL & " order by IntermediaryGuid"
    Else
        sIntPrefix = "Interm"
                       
        sSQL = "SELECT IntermediaryOrganisation.Name NAME, IntermediaryOrganisation.CompanyCommisionNumber, IntermediaryIndividual.Forename, IntermediaryIndividual.Surname, IntermediaryIndividual.Title, " & sIntPrefix & ".IntermediaryGuid, Interm.IsAttachedToIntermediaryguid, IntView.IntermediaryType, IntView.ChildCount " & _
               "FROM " & m_sTableName & " Interm LEFT OUTER JOIN (SELECT count(IsAttachedToIntermediaryGuid) ChildCount, IsAttachedToIntermediaryGuid, IntermediaryType, IntermediaryGUID FROM Intermediary WHERE IsAttachedToIntermediaryGuid is not null group by IntermediaryType, IsAttachedToIntermediaryGuid, IntermediaryGUID) IntView ON Interm.IntermediaryGuid = Intview.IsAttachedToIntermediaryGuid " & _
                "LEFT OUTER JOIN IntermediaryOrganisation ON IntermediaryOrganisation.IntermediaryGuid = Interm.IntermediaryGuid LEFT OUTER JOIN IntermediaryIndividual ON IntermediaryIndividual.IntermediaryGuid = Interm.IntermediaryGuid "
        
        If Len(sType) > 0 Then
            sSQL = sSQL & " WHERE " & sIntPrefix & ".IntermediaryType = " & sType
        End If
        
        If Len(vGUID) > 0 Then
            sSQL = sSQL & " WHERE Interm.IntermediaryGuid = " & g_clsSQLAssistSP.FormatGuid(vGUID)
        End If
    
        sSQL = sSQL & " ORDER BY Interm.IntermediaryGuid"
    End If
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.numer, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetParentIntermediaryGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetParentIntermediaryGUID() As Variant
    
    Dim vVal As Variant
    
    ' DJP SQL Server port, return variant
    vVal = m_clsTableAccess.GetVal("ISATTACHEDTOINTERMEDIARYGUID")
    
    GetParentIntermediaryGUID = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryType
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIntermediaryType() As Integer
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("INTERMEDIARYTYPE")
    GetIntermediaryType = CInt(Val(vVal))

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetCommunicationMethod
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetCommunicationMethod() As Integer
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("COMMUNICATIONMETHOD")
    GetCommunicationMethod = CInt(Val(vVal))

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetProcFeePaymentMethod
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetProcFeePaymentMethod() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("PROCURATIONFEEPAYMENTMETHOD")
    GetProcFeePaymentMethod = vVal
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOtherSystemNumber
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOtherSystemNumber() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("OTHERSYSTEMNUMBER")
    GetOtherSystemNumber = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOtherSystemSuffix
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOtherSystemSuffix() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("OTHERSYSTEMSUFFIX")
    GetOtherSystemSuffix = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActiveFrom
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetActiveFrom() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("INTERMEDIARYACTIVEFROM")
    GetActiveFrom = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetProcurationFeeLTVLimit
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetProcurationFeeLTVLimit() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("PROCURATIONFEELTVLIMIT")
    GetProcurationFeeLTVLimit = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetContactDetailsGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetContactDetailsGUID() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("CONTACTDETAILSGUID")
    GetContactDetailsGUID = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetAddressGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAddressGUID() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("ADDRESSGUID")
    GetAddressGUID = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetTown
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetTown() As Variant
    GetTown = m_clsTableAccess.GetVal("TOWN")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActiveTo
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetActiveTo() As String
    
    Dim vTmp As Variant
    
    vTmp = m_clsTableAccess.GetDate("INTERMEDIARYACTIVETO")
    GetActiveTo = vTmp
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReportFrequency
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetReportFrequency(vVal As Variant) As Boolean
    SetReportFrequency = m_clsTableAccess.SetVal("REPORTFREQUENCY", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIntermediaryType
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIntermediaryType(ByVal sType As String) As Boolean
    SetIntermediaryType = m_clsTableAccess.SetVal("INTERMEDIARYTYPE", sType)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetIntermediaryGuid
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetIntermediaryGuid(vGUID As Variant) As Boolean
    SetIntermediaryGuid = m_clsTableAccess.SetVal("INTERMEDIARYGUID", vGUID)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetActiveTo
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetActiveTo(vVal As Variant) As Boolean
    SetActiveTo = m_clsTableAccess.SetVal("INTERMEDIARYACTIVETO", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetActiveFrom
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetActiveFrom(dDate As Date) As Boolean
    SetActiveFrom = m_clsTableAccess.SetVal("INTERMEDIARYACTIVEFROM", dDate)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetPanelID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetPanelID(sPanelID As String) As Boolean
    SetPanelID = m_clsTableAccess.SetVal("INTERMEDIARYPANELID", sPanelID)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddressGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetAddressGUID(vGUID As Variant) As Boolean
    SetAddressGUID = m_clsTableAccess.SetVal("ADDRESSGUID", vGUID)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetContactGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetContactGUID(vGUID As Variant) As Boolean
    SetContactGUID = m_clsTableAccess.SetVal("CONTACTDETAILSGUID", vGUID)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetParentGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetParentGUID(vGUID As Variant) As Boolean
    SetParentGUID = m_clsTableAccess.SetVal("ISATTACHEDTOINTERMEDIARYGUID", vGUID)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetCommunicationMethod
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetCommunicationMethod(vVal As Variant) As Boolean
    SetCommunicationMethod = m_clsTableAccess.SetVal("COMMUNICATIONMETHOD", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcurationFee
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetProcurationFee(vVal As Variant) As Boolean
    SetProcurationFee = m_clsTableAccess.SetVal("PROCURATIONFEELTVLIMIT", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetProcurationFeeMethod
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetProcurationFeeMethod(vVal As Variant) As Boolean
    SetProcurationFeeMethod = m_clsTableAccess.SetVal("PROCURATIONFEEPAYMENTMETHOD", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetOtherSystemNumber
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetOtherSystemNumber(vVal As Variant) As Boolean
    SetOtherSystemNumber = m_clsTableAccess.SetVal("OTHERSYSTEMNUMBER", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetOtherSystemSuffix
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetOtherSystemSuffix(vVal As Variant) As Boolean
    SetOtherSystemSuffix = m_clsTableAccess.SetVal("OTHERSYSTEMSUFFIX", vVal)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryGUID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIntermediaryGUID() As Variant
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("INTERMEDIARYGUID")
    GetIntermediaryGUID = vVal
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetNoChildren
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetNoChildren() As Long
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("CHILDCOUNT")
    GetNoChildren = CLng(Val(vVal))
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetPanelID
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetPanelID() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("INTERMEDIARYPANELID")
    GetPanelID = CStr(vVal)
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReportFrequency
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetReportFrequency() As Variant
    GetReportFrequency = m_clsTableAccess.GetVal("REPORTFREQUENCY")
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOrgName
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOrgName() As String
    
    Dim vVal As Variant
    
    On Error GoTo Failed
    
    vVal = m_clsTableAccess.GetVal("NAME")
    GetOrgName = CStr(vVal)
    
    Exit Function
    
Failed:
    'There is no Name assosicated with this intermediary
    GetOrgName = "<Company Name>"
    Err.Clear
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOrgCommisionNo
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOrgCommisionNo() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("COMPANYCOMMISIONNUMBER")
    GetOrgCommisionNo = CStr(vVal)
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualDevArea
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualDevArea() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("KEYDEVELOPMENTAREA")
    GetIndividualDevArea = CStr(vVal)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualSurName
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualSurName() As String
    
    Dim vVal As Variant
        
    On Error GoTo Failed
    
    vVal = m_clsTableAccess.GetVal("SURNAME")
    GetIndividualSurName = CStr(vVal)
    
    Exit Function
    
Failed:
    Err.Clear
    GetIndividualSurName = "<Surname>"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualFullName
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualFullName() As String
    GetIndividualFullName = GetIndividualForeName & " " & GetIndividualSurName
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualDevAreaText
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualDevAreaText() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("DEVELOPMENTAREATEXT")
    GetIndividualDevAreaText = CStr(vVal)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualTitleText
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualTitleText() As String
    
    Dim vVal As Variant
    
    vVal = m_clsTableAccess.GetVal("TITLETEXT")
    GetIndividualTitleText = CStr(vVal)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualTitle
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualTitle() As String
        
    Dim vVal As Variant
    
    On Error GoTo Failed
    
    vVal = m_clsTableAccess.GetVal("TITLE")
    GetIndividualTitle = CStr(vVal)
    
    Exit Function
    
Failed:
    Err.Clear
    GetIndividualTitle = "<Title>"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualForeName
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetIndividualForeName() As String
    
    Dim vVal As Variant
    
    On Error GoTo Failed
    
    vVal = m_clsTableAccess.GetVal("FORENAME")
    GetIndividualForeName = CStr(vVal)
    
    Exit Function
    
Failed:
    Err.Clear
    GetIndividualForeName = "<Forename>"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetParentIntermediaryListField
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetParentIntermediaryListField() As String
    GetParentIntermediaryListField = "INTERMEDIARYPANELID"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryGUIDField
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIntermediaryGUIDField() As String
    GetIntermediaryGUIDField = "INTERMEDIARYGUID"
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryDetails
' Description   : Retrieves all relevant Intermediary details given the GUID used to populate
'                 the listview on frmMain.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetIntermediaryDetails(vGUID As Variant, Optional sType As String = "")

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim sComboGroup As String
    Dim sTitleCombo As String
    
    On Error GoTo Failed
    
    sComboGroup = "COMBOVALUE"
    sTitleCombo = "Title"
        
    If sType = INDIVIDUAL Then
        m_bIsIndividual = True
        
        'SB: Format SQL appropriate to specific DB server version.
        If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
            sSQL = "SELECT " & _
                        "INT.*, IND.*, " & _
                        "ComboTitle.ValueName TitleText " & _
                   "FROM " & _
                        m_sTableName & " INT, " & _
                        m_sIndividualTable & " IND, " & _
                        sComboGroup & " ComboTitle " & _
                   "WHERE " & _
                        "IND.INTERMEDIARYGUID (+) = INT.INTERMEDIARYGUID AND " & _
                        "IND.INTERMEDIARYGUID  = " & g_clsSQLAssistSP.FormatString(vGUID) & " AND " & _
                        "ComboTitle.GroupName = " & g_clsSQLAssistSP.FormatString(sTitleCombo) & "AND  " & _
                        "ComboTitle.ValueID = IND.Title "
        Else
            sSQL = "SELECT INT.*, IND.*, ComboTitle.ValueName TitleText FROM " & m_sTableName & " INT LEFT OUTER JOIN " & m_sIndividualTable & " IND ON  IND.INTERMEDIARYGUID = INT.INTERMEDIARYGUID, " & sComboGroup & " ComboTitle WHERE IND.INTERMEDIARYGUID  = "
            
            'SB: vGUID could be an empty string if a non-record node is clicked.
            If CStr(vGUID) = "" Then
                sSQL = sSQL & "''"
            Else
                sSQL = sSQL & g_clsSQLAssistSP.FormatGuid(vGUID)
            End If
            
            sSQL = sSQL & " AND ComboTitle.GroupName = " & g_clsSQLAssistSP.FormatString(sTitleCombo) & " AND ComboTitle.ValueID = IND.Title "
        End If
    Else
        m_bIsIndividual = False
        
        'SB: Format SQL appropriate to specific DB server version.
        If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
            sSQL = "Select INT.*, ORG.Name, ORG.COMPANYCOMMISIONNUMBER from " & m_sTableName & " INT, " & m_sOrganisationTable & _
            " ORG Where ORG.INTERMEDIARYGUID (+) = INT.INTERMEDIARYGUID AND " & _
            " ORG.INTERMEDIARYGUID  = " & g_clsSQLAssistSP.FormatString(vGUID)
        Else
            sSQL = "SELECT INT.* , ORG.Name, ORG.COMPANYCOMMISIONNUMBER FROM " & m_sTableName & " INT LEFT OUTER JOIN " & m_sOrganisationTable & " ORG ON ORG.INTERMEDIARYGUID = INT.INTERMEDIARYGUID WHERE ORG.INTERMEDIARYGUID = "
            
            'SB: vGUID could be an empty string if a non-record node is clicked.
            If CStr(vGUID) = "" Then
                sSQL = sSQL & "''"
            Else
                sSQL = sSQL & g_clsSQLAssistSP.FormatGuid(vGUID)
            End If
        End If
    End If
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryName
' Description   : Returns the name of the intermediary
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIntermediaryName(sType As String) As String
    
    On Error GoTo Failed
    
    If sType = INDIVIDUAL Then
        GetIntermediaryName = GetIndividualFullName
    Else
        GetIntermediaryName = GetOrgName
    End If
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetChildIntsWithChildren
' Description   : Used to retrieve any child node information BELOW the top-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetChildIntsWithChildren(vGUID As Variant, sType As String)

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    'SB: Format SQL appropriate to specific DB server version.
    If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
        sSQL = "select " & _
                    "intermediaryorganisation.name NAME," & _
                    "intermediaryorganisation.COMPANYCOMMISIONNUMBER, " & _
                    "intermediaryindividual.forename," & _
                    "intermediaryindividual.surname," & _
                    "intermediaryindividual.title," & _
                    "INT.IntermediaryGuid," & _
                    "INT.IntermediaryGuid," & _
                    "INTView.IsAttachedToIntermediaryguid," & _
                    "IntView.IntermediaryType," & _
                    "IntView.ChildCount "
        sSQL = sSQL & _
               "From " & _
                    m_sTableName & " INT, " & _
                    "Intermediaryorganisation, " & _
                    "intermediaryindividual, " & _
                    "(SELECT  " & _
                        "count(IsAttachedToIntermediaryGuid) ChildCount, " & _
                        "IsAttachedToIntermediaryGuid," & _
                        "IntermediaryType," & _
                        "IntermediaryGUID " & _
                    "From  " & _
                        "Intermediary  " & _
                    "Where " & _
                        "IsAttachedToIntermediaryGuid is not null " & _
                    "group by " & _
                        "IntermediaryType,IsAttachedToIntermediaryGuid, IntermediaryGUID) IntView "
        sSQL = sSQL & "Where " & _
                            "INT.IntermediaryGuid = IntView.IsAttachedToIntermediaryGuid(+) and " & _
                            "intermediaryorganisation.INTERMEDIARYGUID(+) = Int.INTERMEDIARYGUID AND " & _
                            "intermediaryindividual.INTERMEDIARYGUID(+) = Int.INTERMEDIARYGUID  and " & _
                            "int.IntermediaryType = " & sType & " and " & _
                            "INT.IsAttachedToIntermediaryGuid = " & g_clsSQLAssistSP.FormatString(vGUID)
    Else
        sSQL = "SELECT INTERMEDIARYORGANISATION.NAME NAME, INTERMEDIARYORGANISATION.COMPANYCOMMISIONNUMBER, INTERMEDIARYINDIVIDUAL.FORENAME, INTERMEDIARYINDIVIDUAL.SURNAME, INTERMEDIARYINDIVIDUAL.TITLE, INT.INTERMEDIARYGUID, INT.INTERMEDIARYGUID, IntView.IsAttachedToIntermediaryguid, IntView.IntermediaryType, IntView.ChildCount " & _
            "FROM INTERMEDIARY INT LEFT OUTER JOIN (SELECT COUNT(IsAttachedToIntermediaryGuid) ChildCount, IsAttachedToIntermediaryGuid, IntermediaryType, IntermediaryGUID FROM " & _
            "Intermediary WHERE IsAttachedToIntermediaryGuid IS NOT NULL GROUP BY IntermediaryType, IsAttachedToIntermediaryGuid, IntermediaryGUID) IntView ON IntView.IsAttachedToIntermediaryGuid = INT.INTERMEDIARYGUID LEFT OUTER JOIN Intermediaryorganisation ON intermediaryorganisation.INTERMEDIARYGUID = INT.INTERMEDIARYGUID LEFT OUTER JOIN intermediaryindividual ON intermediaryindividual.INTERMEDIARYGUID = INT.INTERMEDIARYGUID " & _
            "WHERE INT.IntermediaryType = " & sType & " AND INT.IsAttachedToIntermediaryGuid = " & g_clsSQLAssistSP.FormatGuid(vGUID)
    End If
    
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.numer, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetValidParentIntermediaries
' Description   : Returns all intermeidiaries whose type is contained within the specified
'                 collection of Validation Types.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetValidParentIntermediaries(ByRef colTypesList As Collection)

    Dim sSQL As String
    Dim iIndex As Integer
    Dim sTypesList As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    'Only attempt to get any records if an intermediary type was specified.
    If colTypesList.Count > 0 Then
        'Build the IN (x,y,z) SQL clause of intermediary types.
        For iIndex = 1 To colTypesList.Count
            sTypesList = sTypesList & colTypesList(iIndex) & ", "
        Next iIndex
        
        'Strip any trailing ", "
        If Right$(sTypesList, 2) = ", " Then
            sTypesList = Left$(sTypesList, Len(sTypesList) - 2)
        End If
        
        'Build a SQL query to obtain any possible intermediaries of the specified type(s).
        sSQL = "SELECT INTER.* FROM " & m_sTableName & " INTER WHERE INTERMEDIARYTYPE IN (" & sTypesList & ")"
    
        Set rs = g_clsDataAccess.GetTableData(, sSQL)
        m_clsTableAccess.SetRecordSet rs
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetPanelIDForGUID
' Description   : Returns the PanelID for the given interemdiaryguid. Uses the filter method
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetPanelIDForGUID(vGUID As Variant) As String
    
    Dim sFilter As String
    Dim sPanel As String
    
    On Error GoTo Failed
    
    If Len(vGUID) > 0 Then
        sFilter = "INTERMEDIARYGUID = " & g_clsSQLAssistSP.FormatString(vGUID)
        m_clsTableAccess.ApplyFilter sFilter
        sPanel = GetPanelID
        m_clsTableAccess.CancelFilter
        GetPanelIDForGUID = sPanel
    End If
    
    Exit Function
    
Failed:
    m_clsTableAccess.CancelFilter
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIndividualTable
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetIndividualTable() As String
    GetIndividualTable = m_sIndividualTable
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetOrganisationTable
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetOrganisationTable() As String
    GetOrganisationTable = m_sOrganisationTable
End Function


