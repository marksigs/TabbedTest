VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdIncentiveTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdIncentiveTable
' Description   : Manages all data access to the Incentives table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' DJP       03/12/01    SYS2912 SQL Server locking problem, update GetIncentives
' SDS       29/01/02    SYS3320 - Necessary changes in code to incorporate replacement of "Benefit Type" textbox with Combobox
' PE        13/01/07    EP2_836 - Error adding an incentive to a product
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements TableAccess

' Private data
Private m_sSearch As String
Private m_clsTableAccess As TableAccess
Private m_colMatchValues As Collection

'Constants
Private Const m_sTableName = "INCENTIVE"
Private Const m_sVersionField As String = "INCENTIVEVERSIONNUMBER"

' BMIDS Change history
' Prog      Date        Description
' GD        05/06/02    BMIDS0016 Changes to GetIncentives
Private Sub Class_Initialize()
    Dim colFields As New Collection

    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "IncentiveGUID"
    
    If g_clsVersion.DoesVersioningExist() Then
        colFields.Add m_sVersionField
    End If
    
    m_clsTableAccess.SetKeyMatchFields colFields

End Sub

Friend Function GetIncentiveGUID() As Variant
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("IncentiveGUID")

    ' DJP SQL Server port
    GetIncentiveGUID = vVal
End Function

Friend Function GetType() As String
    GetType = m_clsTableAccess.GetVal("Type")
End Function

Friend Function GetDescription() As String
    GetDescription = m_clsTableAccess.GetVal("Description")
End Function

Friend Function GetIncentiveBenefitType() As String
    GetIncentiveBenefitType = m_clsTableAccess.GetVal("IncentiveBenefitType")
End Function

Friend Function GetIncentiveBenefitTypeDescription() As String
     GetIncentiveBenefitTypeDescription = m_clsTableAccess.GetVal("TypeDescription")
End Function

Friend Function GetAmount() As String
    GetAmount = CStr(m_clsTableAccess.GetVal("Amount"))
End Function

Friend Function GetPercentage() As String
    GetPercentage = CStr(m_clsTableAccess.GetVal("Percentage"))
End Function

Friend Function GetPercentageMax() As String
    GetPercentageMax = CStr(m_clsTableAccess.GetVal("PercentageMaximum"))
End Function

Friend Function SetIncentiveGUID(vGUID As Variant) As Boolean
    SetIncentiveGUID = m_clsTableAccess.SetVal("IncentiveGUID", vGUID)

    ' And versioning if it exists
    g_clsVersion.HandleVersioning Me

End Function

Friend Function SetDescription(sDescription As String) As Boolean
    SetDescription = m_clsTableAccess.SetVal("Description", CVar(sDescription))
End Function

Friend Function SetAmount(sAmount As String) As Boolean
    SetAmount = m_clsTableAccess.SetVal("Amount", CVar(sAmount))
End Function

Friend Function SetPercentage(sPercentage As String) As Boolean
    SetPercentage = m_clsTableAccess.SetVal("Percentage", CVar(sPercentage))
End Function

Friend Function SetPercentageMax(sPercentageMax As String) As Boolean
    SetPercentageMax = m_clsTableAccess.SetVal("PercentageMaximum", CVar(sPercentageMax))
End Function

Friend Function SetIncentiveBenefitType(sIncentiveBenefitType As String) As Boolean
    SetIncentiveBenefitType = m_clsTableAccess.SetVal("IncentiveBenefitType", CVar(sIncentiveBenefitType))
End Function

Friend Function SetProductCode(sProductCode As String) As Boolean
    SetProductCode = m_clsTableAccess.SetVal("MortgageProductCode", sProductCode)
End Function

Friend Function SetProductStartDate(sStartDate As String) As Boolean
    SetProductStartDate = m_clsTableAccess.SetVal("StartDate", sStartDate)
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

Public Function TableAccess_GetUpdateValues() As Collection
End Function

Public Sub TableAccess_Create(sTable As String, Optional sFirstMatch As String = "")
    m_clsTableAccess.Create sTable, sFirstMatch
End Sub

Public Sub TableAccess_GetRowOfData(colListLine As Collection, Optional bIncludeHeader As Boolean = True)
    colListLine.Add GetType()
    colListLine.Add GetIncentiveBenefitTypeDescription()
    colListLine.Add GetAmount()
    colListLine.Add GetPercentage()
    colListLine.Add GetPercentageMax()
    colListLine.Add GetDescription()
End Sub

Public Sub TableAccess_SetTableClass(clsTable As TableAccess)
    Set m_clsTableAccess = clsTable
End Sub

Public Sub GetIncentives(incType As IncentiveType, colProductValues As Collection, Optional bAllFields As Boolean = False, Optional colIncentiveGUIDs As Collection)
    On Error GoTo Failed
    Dim sCode As String
    Dim sType As String
    Dim sSearch As String
    Dim sVersion As String
    Dim sDatabase As String
    Dim sStartDate As String
    Dim sSubSearch As String
    Dim sFields As String
    Dim sIncentiveTable As String
    Dim colMatchFields As Collection
    Dim colMatchValues As Collection
    Dim bDoProductSearch As Boolean
    Dim clsMortgageProduct As New MortgageProductTable
    Dim clsTableAccess As TableAccess
    Dim bDoSearch As Boolean
    
    bDoSearch = True
    If incType = Exclusive Then
        sType = "Exclusive"
        sIncentiveTable = "ExclusiveIncentive"
    ElseIf incType = Inclusive Then
        sType = "Inclusive"
        sIncentiveTable = "InclusiveIncentive"
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Invalid Incentive Type: " & incType
    End If
        
    sDatabase = m_clsTableAccess.GetDatabase
    
    bDoProductSearch = True
    Set clsTableAccess = clsMortgageProduct
    Set colMatchFields = clsTableAccess.GetKeyMatchFields()
    
    If bAllFields Then
        sFields = "SELECT INCENTIVE.* "
    Else
        sFields = "SELECT " & g_clsSQLAssistSP.FormatString(sType) & """TYPE"", INCENTIVE.* "
    End If
    
    'Added by Sunil ,  29-01-2002
    If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
        sFields = sFields & " , NVL(ComboValue.VALUENAME , '')  TYPEDESCRIPTION "
    Else
        sFields = sFields & " , isnull(ComboValue.VALUENAME , '')  TYPEDESCRIPTION "
    End If

    
    If Not colIncentiveGUIDs Is Nothing Then
        If colIncentiveGUIDs.Count > 0 Then
            Dim vIncentiveGUID As Variant
            Dim sIncentiveGUID As String
            Dim nThisGUID As Long
            Dim nCount As Long
            
            sSearch = sFields & "FROM " & _
                        "INCENTIVE, " & sIncentiveTable & " " & _
                        " , ComboValue "  'Added by Sunil , 29-01-2002
            sSearch = sSearch & "WHERE INCENTIVE.INCENTIVEGUID IN "

            sSubSearch = "("
            
            nCount = colIncentiveGUIDs.Count
            nThisGUID = 1
            
            For Each vIncentiveGUID In colIncentiveGUIDs
                sIncentiveGUID = g_clsSQLAssistSP.GuidToString(CStr(vIncentiveGUID))
                sIncentiveGUID = g_clsSQLAssistSP.FormatGuid(sIncentiveGUID)
                
                sSubSearch = sSubSearch & sIncentiveGUID
                If nThisGUID < nCount Then
                    sSubSearch = sSubSearch & ", "
                End If
                nThisGUID = nThisGUID + 1
            Next
            
            sSubSearch = sSubSearch & ") AND " & sIncentiveTable & ".INCENTIVEGUID = INCENTIVE.INCENTIVEGUID "
            sSubSearch = sSubSearch & " AND ((INCENTIVE.INCENTIVEBENEFITTYPE = ComboValue.VALUEID)  and ComboValue.GROUPNAME = 'IncentiveBenefitType') "
            sSearch = sSearch & sSubSearch
            
            bDoProductSearch = False
        End If
    End If
    
    If bDoProductSearch Then
        If Not colProductValues Is Nothing Then
            ' DJP SQL Server port - moved to DataAccess
            If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
                         sSearch = sFields & "FROM " & _
                                    "INCENTIVE , ComboValue " & _
                                    " where (INCENTIVE.INCENTIVEBENEFITTYPE (+)= ComboValue.VALUEID) " & _
                                    " and ComboValue.GROUPNAME = 'IncentiveBenefitType' and INCENTIVE.INCENTIVEGUID IN "
                                                     
                        sSubSearch = g_clsDataAccess.GenerateSearchCriteria(colMatchFields, colProductValues, sDatabase)
                        
                        sSearch = sSearch & "(SELECT " & _
                                                "INCENTIVEGUID " & _
                                            "FROM " & _
                                                sIncentiveTable & " " & _
                                            "WHERE " & _
                                                sSubSearch & ")"
                        
                        
            Else
                        sSearch = sFields & "FROM " & _
                                    "INCENTIVE left outer Join  ComboValue " & _
                                    " on (INCENTIVE.INCENTIVEBENEFITTYPE = ComboValue.VALUEID) " & _
                                    " and ComboValue.GROUPNAME = 'IncentiveBenefitType' WHERE INCENTIVE.INCENTIVEGUID IN "
                                                     ' "INCENTIVE WHERE INCENTIVE.INCENTIVEGUID IN "
                        sSubSearch = g_clsDataAccess.GenerateSearchCriteria(colMatchFields, colProductValues, sDatabase)
                        
                        sSearch = sSearch & "(SELECT " & _
                                                "INCENTIVEGUID " & _
                                            "FROM " & _
                                                sIncentiveTable & " " & _
                                            "WHERE " & _
                                                sSubSearch & ")"
            End If
        Else
            bDoSearch = False
        End If
    End If
        
    If bDoSearch Then
        m_clsTableAccess.GetTableDataFromSearch sSearch
    Else
        m_clsTableAccess.GetTableData POPULATE_EMPTY
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub TableAccess_DeleteRecords(Optional sDatabase As String = "")
    On Error GoTo Failed
    m_clsTableAccess.DeleteRecords sDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DeleteIncentives(Optional sDatabaseName As String)
    On Error GoTo Failed
    ' Need to get the mortgage product code and start date
    Dim colMatchValues As Collection
    Dim sProductCode As String
    Dim sProductStartDate As String
    Dim sCommand As String
    
        
    Set colMatchValues = m_clsTableAccess.GetKeyMatchValues
    
    ' Two keys
    If colMatchValues.Count = 2 Then
        sProductCode = colMatchValues(1)
        sProductStartDate = colMatchValues(2)
    
        ' Inclusive
        ' DJP SQL Server port
        sCommand = "DELETE " & _
                    "FROM " & _
                        "INCENTIVE " & _
                    "WHERE " & _
                        "INCENTIVEGUID IN (" & _
                            "SELECT " & _
                                "INCENTIVEGUID " & _
                            "FROM " & _
                                "INCLUSIVEINCENTIVE " & _
                            "WHERE " & _
                                "MORTGAGEPRODUCTCODE = " & g_clsSQLAssistSP.FormatString(sProductCode) & " AND " & _
                                g_clsSQLAssistSP.TruncateDateColumn("STARTDATE", sDatabaseName) & " = " & g_clsSQLAssistSP.FormatDate(sProductStartDate, , sDatabaseName) & ")"
        
        g_clsDataAccess.ExecuteCommand sCommand, sDatabaseName
    
        ' Exclusive
        ' DJP SQL Server port
        sCommand = "DELETE " & _
                    "FROM " & _
                        "INCENTIVE " & _
                    "WHERE " & _
                        "INCENTIVEGUID IN (" & _
                            "SELECT " & _
                                "INCENTIVEGUID " & _
                            "FROM " & _
                                "EXCLUSIVEINCENTIVE " & _
                            "WHERE " & _
                                "MORTGAGEPRODUCTCODE = " & g_clsSQLAssistSP.FormatString(sProductCode) & " AND " & _
                                g_clsSQLAssistSP.TruncateDateColumn("STARTDATE", sDatabaseName) & " = " & g_clsSQLAssistSP.FormatDate(sProductStartDate, , sDatabaseName) & ")"
    
        g_clsDataAccess.ExecuteCommand sCommand, sDatabaseName
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, " DeleteIncentives - no keys present"
    End If
    
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
Public Sub GetIncentivesForPromotions(incType As IncentiveType, colProductValues As Collection, Optional bAllFields As Boolean = False, Optional colIncentiveGUIDs As Collection)

'GD BMIDS BMIDS00016 - Extra method reqd to allow exclusion of TYPEDESCRIPTION, since the recordset conflicts with the
'database table structure when trying to promote the table to the target database

    On Error GoTo Failed
    Dim sCode As String
    Dim sType As String
    Dim sSearch As String
    Dim sVersion As String
    Dim sDatabase As String
    Dim sStartDate As String
    Dim sSubSearch As String
    Dim sFields As String
    Dim sIncentiveTable As String
    Dim colMatchFields As Collection
    Dim colMatchValues As Collection
    Dim bDoProductSearch As Boolean
    Dim clsMortgageProduct As New MortgageProductTable
    Dim clsTableAccess As TableAccess
    Dim bDoSearch As Boolean
    
    bDoSearch = True
    If incType = Exclusive Then
        sType = "Exclusive"
        sIncentiveTable = "ExclusiveIncentive"
    ElseIf incType = Inclusive Then
        sType = "Inclusive"
        sIncentiveTable = "InclusiveIncentive"
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Invalid Incentive Type: " & incType
    End If
        
    sDatabase = m_clsTableAccess.GetDatabase
    
    bDoProductSearch = True
    Set clsTableAccess = clsMortgageProduct
    Set colMatchFields = clsTableAccess.GetKeyMatchFields()
    
    If bAllFields Then
        sFields = "SELECT INCENTIVE.* "
    Else
        sFields = "SELECT " & g_clsSQLAssistSP.FormatString(sType) & """TYPE"", INCENTIVE.* "
    End If
    'GD BMIDS00016  start
    'Added by Sunil ,  29-01-2002
    'If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
    '    sFields = sFields & " , NVL(ComboValue.VALUENAME , '')  TYPEDESCRIPTION "
    'Else
    '    sFields = sFields & " , isnull(ComboValue.VALUENAME , '')  TYPEDESCRIPTION "
    'End If
    'GD BMIDS end

    
    If Not colIncentiveGUIDs Is Nothing Then
        If colIncentiveGUIDs.Count > 0 Then
            Dim vIncentiveGUID As Variant
            Dim sIncentiveGUID As String
            Dim nThisGUID As Long
            Dim nCount As Long
            

            sSearch = sFields & "FROM " & _
                        "INCENTIVE, " & sIncentiveTable & " " & _
                        " , ComboValue "  'Added by Sunil , 29-01-2002
            sSearch = sSearch & "WHERE INCENTIVE.INCENTIVEGUID IN "

            sSubSearch = "("
            
            
            nCount = colIncentiveGUIDs.Count
            nThisGUID = 1
            
            For Each vIncentiveGUID In colIncentiveGUIDs
                sIncentiveGUID = g_clsSQLAssistSP.GuidToString(CStr(vIncentiveGUID))
                sIncentiveGUID = g_clsSQLAssistSP.FormatGuid(sIncentiveGUID)
                
                sSubSearch = sSubSearch & sIncentiveGUID
                If nThisGUID < nCount Then
                    sSubSearch = sSubSearch & ", "
                End If
                nThisGUID = nThisGUID + 1
            Next
            
            sSubSearch = sSubSearch & ") AND " & sIncentiveTable & ".INCENTIVEGUID = INCENTIVE.INCENTIVEGUID "
            sSubSearch = sSubSearch & " AND ((INCENTIVE.INCENTIVEBENEFITTYPE *= ComboValue.VALUEID)  and ComboValue.GROUPNAME = 'IncentiveBenefitType') "
            sSearch = sSearch & sSubSearch
            
            bDoProductSearch = False
        End If
    End If
    
    If bDoProductSearch Then
        If Not colProductValues Is Nothing Then
            ' DJP SQL Server port - moved to DataAccess
            If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
                         sSearch = sFields & "FROM " & _
                                    "INCENTIVE , ComboValue " & _
                                    " where (INCENTIVE.INCENTIVEBENEFITTYPE (+)= ComboValue.VALUEID) " & _
                                    " and ComboValue.GROUPNAME = 'IncentiveBenefitType' and INCENTIVE.INCENTIVEGUID IN "
                                                     
                        sSubSearch = g_clsDataAccess.GenerateSearchCriteria(colMatchFields, colProductValues, sDatabase)
                        
                        sSearch = sSearch & "(SELECT " & _
                                                "INCENTIVEGUID " & _
                                            "FROM " & _
                                                sIncentiveTable & " " & _
                                            "WHERE " & _
                                                sSubSearch & ")"
                        
                        
            Else
                        sSearch = sFields & "FROM " & _
                                    "INCENTIVE left outer Join  ComboValue " & _
                                    " on (INCENTIVE.INCENTIVEBENEFITTYPE = ComboValue.VALUEID) " & _
                                    " and ComboValue.GROUPNAME = 'IncentiveBenefitType' WHERE INCENTIVE.INCENTIVEGUID IN "
                                                     ' "INCENTIVE WHERE INCENTIVE.INCENTIVEGUID IN "
                        sSubSearch = g_clsDataAccess.GenerateSearchCriteria(colMatchFields, colProductValues, sDatabase)
                        
                        sSearch = sSearch & "(SELECT " & _
                                                "INCENTIVEGUID " & _
                                            "FROM " & _
                                                sIncentiveTable & " " & _
                                            "WHERE " & _
                                                sSubSearch & ")"
            End If
        Else
            bDoSearch = False
        End If
    End If
        
    If bDoSearch Then
        m_clsTableAccess.GetTableDataFromSearch sSearch
    Else
        m_clsTableAccess.GetTableData POPULATE_EMPTY
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
