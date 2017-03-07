VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : ApplicationTable
' Description   : Contains functions to read/write/search the application table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "APPLICATION"
Private m_sSearch As String
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    colFields.Add "APPLICATIONNUMBER"
    
    m_clsTableAccess.SetKeyMatchFields colFields
End Sub
Public Function GetApplicationNumber() As Variant
    GetApplicationNumber = m_clsTableAccess.GetVal("APPLICATIONNUMBER")
End Function
Public Function GetAppFactFindNumber() As Variant
    GetAppFactFindNumber = m_clsTableAccess.GetVal("MAXFACTFIND")
End Function
Public Function GetStageName() As Variant
    GetStageName = m_clsTableAccess.GetVal("STAGENAME")
End Function
Public Function GetForeName() As Variant
    GetForeName = m_clsTableAccess.GetVal("FIRSTFORENAME")
End Function
Public Function GetSurName() As Variant
    GetSurName = m_clsTableAccess.GetVal("SURNAME")
End Function
Public Function GetTitle() As Variant
    GetTitle = m_clsTableAccess.GetVal("TITLE")
End Function
Public Function GetPostcode() As Variant
    GetPostcode = m_clsTableAccess.GetVal("POSTCODE")
End Function
Public Function GetBuildingOrHouseNumber() As Variant
    GetBuildingOrHouseNumber = m_clsTableAccess.GetVal("BUILDINGORHOUSENUMBER")
End Function
Public Function GetFlatNumber() As Variant
    GetFlatNumber = m_clsTableAccess.GetVal("FLATNUMBER")
End Function
Public Function GetStreet() As Variant
    GetStreet = m_clsTableAccess.GetVal("STREET")
End Function
Public Function GetTown() As Variant
    GetTown = m_clsTableAccess.GetVal("TOWN")
End Function
Public Function GetAddress() As String
    On Error GoTo Failed
    Dim sAddress As String
    Dim sFlat As String
    Dim sBuilding As String
    Dim sPostcode As String
    Dim sStreet As String
    Dim sTown As String
    
    sPostcode = GetPostcode()
    
    If Len(sPostcode) > 0 Then
        sAddress = sPostcode + ", "
    End If
    
    sFlat = GetFlatNumber()
    
    If Len(sFlat) > 0 Then
        sAddress = sAddress + "Flat " + sFlat + ", "
    End If

    sBuilding = GetBuildingOrHouseNumber()

    If Len(sBuilding) > 0 Then
        sAddress = sAddress + sBuilding + ", "
    End If

    sStreet = GetStreet()

    If Len(sStreet) > 0 Then
        sAddress = sAddress + sStreet + ", "
    End If

    sTown = GetTown()

    If Len(sTown) > 0 Then
        sAddress = sAddress + sTown
    End If

    GetAddress = sAddress

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
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
Private Function TableAccess_DoesRecordExist(colMatchData As Collection, Optional colMatchFields As Collection) As Boolean
    TableAccess_DoesRecordExist = m_clsTableAccess.DoesRecordExist(colMatchData, colMatchFields)
End Function
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

Private Function TableAccess_SetDate(sField As String, vVal As Variant) As Boolean
    TableAccess_SetDate = m_clsTableAccess.SetDate(sField, vVal)
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
Private Function TableAccess_SetVal(sField As Variant, vVal As Variant) As Boolean
    TableAccess_SetVal = m_clsTableAccess.SetVal(sField, vVal)
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
    colListLine.Add GetApplicationNumber()
    colListLine.Add GetSurName
    colListLine.Add GetForeName
    colListLine.Add GetAddress()
    'colListLine.Add GetStageName()
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActiveApplications
' Description   : Returns all active applications. An active application is one that doesn't have
'                 an entry in the applicationstage table for Cancelled or Complete. This function
'                 will return the lead application name and address, too.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetActiveApplications(Optional sAppNo As String, Optional sSurName As String, Optional sForeName As String, Optional bUseExistingSearch As Boolean = False)
    On Error GoTo Failed
    Dim sSearch As String
    Dim rs As ADODB.Recordset
    
'    sSearch = "SELECT " + _
'        "DISTINCT A.APPLICATIONNUMBER, CV.TITLE,CV.SURNAME, CV.FIRSTFORENAME, " + _
'        "AD.POSTCODE , AD.BUILDINGORHOUSENUMBER, AD.FLATNUMBER, AD.STREET, AD.TOWN " + _
'    "FROM " + _
'        "APPLICATIONSTAGE A, CUSTOMERROLE CR, CUSTOMERVERSION CV, CUSTOMERADDRESS CA, ADDRESS AD " + _
'    "WHERE " + _
'        "A.APPLICATIONNUMBER NOT IN  ( " + _
'        "SELECT " + _
'            "B.APPLICATIONNUMBER " + _
'        "FROM " + _
'            "APPLICATIONSTAGE B " + _
'        "Where " + _
'            "B.STAGENUMBER IN ( "
'
'    sSearch = sSearch + "SELECT " + _
'                "VALUEID " + _
'            "FROM " + _
'                "COMBOVALUE " + _
'            "WHERE " + _
'                "UPPER(GROUPNAME) = 'APPLICATIONSTAGE' AND " + _
'                "VALUEID IN  ( " + _
'                "SELECT " + _
'                    "VALUEID " + _
'                    "FROM " + _
'                    "COMBOVALIDATION " + _
'                "WHERE " + _
'                    "(UPPER(VALIDATIONTYPE) = 'C' OR UPPER(VALIDATIONTYPE) = 'MC') AND " + _
'                    "UPPER(GROUPNAME) = 'APPLICATIONSTAGE'))) " + _
'        "AND CR.APPLICATIONNUMBER = A.APPLICATIONNUMBER AND " + _
'        "CV.CUSTOMERNUMBER = CR.CUSTOMERNUMBER AND " + _
'        "CV.CUSTOMERVERSIONNUMBER = CR.CUSTOMERVERSIONNUMBER AND " + _
'        "CV.CUSTOMERVERSIONNUMBER = CA.CUSTOMERVERSIONNUMBER AND " + _
'        "CR.APPLICATIONNUMBER = A.APPLICATIONNUMBER AND " + _
'        "CR.CUSTOMERORDER = 1  AND " + _
'        "CV.CUSTOMERNUMBER = CA.CUSTOMERNUMBER(+) AND " + _
'        "CA.ADDRESSGUID = AD.ADDRESSGUID(+) "
'
'
'        If Len(sAppNo) > 0 Then
'            sSearch = sSearch + " AND A.APPLICATIONNUMBER LIKE " + g_clsFormProcessing.FormatWildcardString(sAppNo)
'        End If
'
'        If Len(sSurname) > 0 Then
'            sSearch = sSearch + " AND CV.SURNAME LIKE " + g_clsFormProcessing.FormatWildcardString(sSurname)
'        End If
'
'        If Len(sForename) > 0 Then
'            sSearch = sSearch + " AND CV.FIRSTFORENAME LIKE " + g_clsFormProcessing.FormatWildcardString(sForename)
'        End If
'
'
'    sSearch = sSearch + "AND CA.addresstype = (" + _
'        "SELECT " + _
'            "VALUEID " + _
'        "FROM " + _
'            "COMBOVALUE " + _
'        "WHERE " + _
'            "UPPER(GROUPNAME) = 'CUSTOMERADDRESSTYPE' AND " + _
'            "VALUEID= ( " + _
'            "SELECT " + _
'                "VALUEID " + _
'            "FROM " + _
'                "COMBOVALIDATION " + _
'            "WHERE " + _
'                "VALIDATIONTYPE = 'H' and UPPER(GROUPNAME) = 'CUSTOMERADDRESSTYPE')) " + _
'            "ORDER BY A.APPLICATIONNUMBER"
        
    sSearch = "SELECT Distinct a.APPLICATIONNUMBER, CV.Title, CV.SURNAME, CV.FIRSTFORENAME, " + _
    "ADDRESSDETAILS.POSTCODE, ADDRESSDETAILS.BUILDINGORHOUSENUMBER, ADDRESSDETAILS.FLATNUMBER, ADDRESSDETAILS.STREET, ADDRESSDETAILS.TOWN " + _
    "From " + _
    "APPLICATIONSTAGE A, CUSTOMERROLE CR, CUSTOMERVERSION CV, " + _
        "( SELECT CUSTOMERNUMBER, CUSTOMERVERSIONNUMBER,  AD.POSTCODE , AD.BUILDINGORHOUSENUMBER, AD.FLATNUMBER, AD.STREET, AD.TOWN " + _
          "From " + _
            "COMBOVALIDATION, CUSTOMERADDRESS, ADDRESS AD " + _
          "Where " + _
            "CUSTOMERADDRESS.ADDRESSTYPE = COMBOVALIDATION.VALUEID AND " + _
            "CUSTOMERADDRESS.ADDRESSGUID = AD.ADDRESSGUID AND " + _
            "VALIDATIONTYPE = 'H' and GROUPNAME = 'CustomerAddressType') ADDRESSDETAILS "
    
    sSearch = sSearch + "Where " + _
    "A.APPLICATIONNUMBER NOT IN (SELECT " + _
            "APPLICATIONNUMBER " + _
         "From " + _
            "COMBOVALIDATION , APPLICATIONSTAGE " + _
         "Where " + _
            "(VALIDATIONTYPE = 'C' OR VALIDATIONTYPE = 'MC')  AND " + _
            "COMBOVALIDATION.GROUPNAME = 'ApplicationStage' AND " + _
            "APPLICATIONSTAGE.STAGENUMBER = COMBOVALIDATION.VALUEID) AND " + _
    "CR.APPLICATIONNUMBER = A.APPLICATIONNUMBER AND " + _
    "CV.CUSTOMERNUMBER = CR.CUSTOMERNUMBER AND " + _
    "CV.CUSTOMERVERSIONNUMBER = CR.CUSTOMERVERSIONNUMBER AND " + _
    "CV.CUSTOMERVERSIONNUMBER = ADDRESSDETAILS.CUSTOMERVERSIONNUMBER(+) AND " + _
    "CR.APPLICATIONNUMBER = A.APPLICATIONNUMBER AND " + _
    "CR.CUSTOMERORDER = 1  AND " + _
    "CV.CUSTOMERNUMBER = ADDRESSDETAILS.CUSTOMERNUMBER(+) "

    If Len(sAppNo) > 0 Then
        sSearch = sSearch + " AND A.APPLICATIONNUMBER LIKE " + g_clsFormProcessing.FormatWildcardString(sAppNo)
    End If
    
    If Len(sSurName) > 0 Then
        sSearch = sSearch + " AND CV.SURNAME LIKE " + g_clsFormProcessing.FormatWildcardString(sSurName)
    End If
    
    If Len(sForeName) > 0 Then
        sSearch = sSearch + " AND CV.FIRSTFORENAME LIKE " + g_clsFormProcessing.FormatWildcardString(sForeName)
    End If
    
    sSearch = sSearch + " ORDER BY A.APPLICATIONNUMBER "

    m_clsTableAccess.GetTableDataFromSearch sSearch, bUseExistingSearch
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


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
