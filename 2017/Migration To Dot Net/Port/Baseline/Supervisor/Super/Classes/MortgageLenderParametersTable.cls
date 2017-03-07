VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortLenderParamsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements TableAccess
Private m_clsTableAccess As TableAccess
Private Const m_sTableName = "MORTGAGELENDERPARAMETERS"
Private m_sSearch As String
Private Const m_sVersionField As String = "MORTGAGELENDERVERSIONNUMBER"
Private Sub Class_Initialize()
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "OrganisationID"
    m_clsTableAccess.SetKeyMatchFields colFields

    ' If versioning exists, we'll need to add that too...
    If g_clsVersion.DoesVersioningExist() Then
        colFields.Add m_sVersionField
    End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' General Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Data Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function GetOrganisationID() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("OrganisationID")
    GetOrganisationID = CStr(vVal)
End Function
Friend Function GetAccInterestInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccInterestInd")
    GetAccInterestInd = CStr(vVal)
End Function
Friend Function GetAccInterestPayableInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccInterestPayableInd")
    GetAccInterestPayableInd = CStr(vVal)
End Function
Friend Function GetFirstPaymentDueDateInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("FirstPaymentDueDateInd")
    GetFirstPaymentDueDateInd = CStr(vVal)
End Function
Friend Function GetAccountingStartMonth() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccountingStartMonth")
    GetAccountingStartMonth = CStr(vVal)
End Function
Friend Function GetInterestChargedInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("InterestChargedInd")
    GetInterestChargedInd = CStr(vVal)
End Function
Friend Function GetAccInterestRoundingFactor() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccInterestRoundingFactor")
    GetAccInterestRoundingFactor = CStr(vVal)
End Function
Friend Function GetAccInterestRoundingDirection() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccInterestRoundingDirection")
    GetAccInterestRoundingDirection = CStr(vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function GetDevonRulingInd() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("DevonRulingInd")
'    GetDevonRulingInd = CStr(vVal)
'End Function
'Friend Function GetInterestChangeInd() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("InterestChangeInd")
'    GetInterestChangeInd = CStr(vVal)
'End Function
'Friend Function GetPartAndPartRounding() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("PartAndPartRounding")
'    GetPartAndPartRounding = CStr(vVal)
'End Function
'BMIDS736 End
Friend Function GetMonthlyPaymentInd() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("MonthlyPaymentInd")
    GetMonthlyPaymentInd = CStr(vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function GetNetGrossProfileInd() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("NetGrossProfileInd")
'    GetNetGrossProfileInd = CStr(vVal)
'End Function
'BMIDS736 End
Friend Function GetPaymentRoundingFactor() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PaymentRoundingFactor")
    GetPaymentRoundingFactor = CStr(vVal)
End Function
Friend Function GetPaymentRoundingDirection() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PaymentRoundingDirection")
    GetPaymentRoundingDirection = CStr(vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function GetInterestCalcRoundingFactor() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("InterestCalcRoundingFactor")
'    GetInterestCalcRoundingFactor = CStr(vVal)
'End Function
'Friend Function GetInterestCalcRoundingDirection() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("InterestCalcRoundingDirection")
'    GetInterestCalcRoundingDirection = CStr(vVal)
'End Function
'BMIDS736 End
Friend Function GetShiftInMonths() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ShiftInMonths")
    GetShiftInMonths = CStr(vVal)
End Function
Friend Function GetPositionInMonth() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("PositionInMonth")
    GetPositionInMonth = CStr(vVal)
End Function
Friend Function GetShiftInDays() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ShiftInDays")
    GetShiftInDays = CStr(vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function GetIncludeACCInterest() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("IncludeACCInterest")
'    GetIncludeACCInterest = CStr(vVal)
'End Function
'Friend Function GetIncludeInACCInterest() As String
'    Dim vVal As Variant
'    vVal = m_clsTableAccess.GetVal("IncludeInACCInterest")
'    GetIncludeInACCInterest = CStr(vVal)
'End Function
'BMIDS736

'BMIDS736 GHun
Friend Function GetCoolingOffPeriod() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("CoolingOffPeriod")
    GetCoolingOffPeriod = CStr(vVal)
End Function

Friend Function GetSpecifiedPaymentDay() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("SpecifiedPaymentDay")
    GetSpecifiedPaymentDay = CStr(vVal)
End Function

Friend Function GetDaysInYear() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("DaysInYear")
    GetDaysInYear = CStr(vVal)
End Function

Friend Function GetAccruedDaysIncluded() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccruedDaysIncluded")
    GetAccruedDaysIncluded = CStr(vVal)
End Function

Friend Function GetAccruedDaysAdded() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("AccruedDaysAdded")
    GetAccruedDaysAdded = CStr(vVal)
End Function

Friend Function GetAPRMonth() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("APRMonth")
    GetAPRMonth = CStr(vVal)
End Function

Friend Function GetWhenAccIntRounded() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("WhenAccIntRounded")
    GetWhenAccIntRounded = CStr(vVal)
End Function

Friend Function GetWhenPaymentRounded() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("WhenPaymentRounded")
    GetWhenPaymentRounded = CStr(vVal)
End Function

Friend Function GetWhenBalanceRounded() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("WhenBalanceRounded")
    GetWhenBalanceRounded = CStr(vVal)
End Function

Friend Function GetBalanceRoundingFactor() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BalanceRoundingFactor")
    GetBalanceRoundingFactor = CStr(vVal)
End Function

Friend Function GetBalanceRoundingDirection() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("BalanceRoundingDirection")
    GetBalanceRoundingDirection = CStr(vVal)
End Function

Friend Function GetChargeRoundingFactor() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ChargeRoundingFactor")
    GetChargeRoundingFactor = CStr(vVal)
End Function

Friend Function GetChargeRoundingDirection() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("ChargeRoundingDirection")
    GetChargeRoundingDirection = CStr(vVal)
End Function

Friend Function GetOutputPaymentSchedule() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("OutputPaymentSchedule")
    GetOutputPaymentSchedule = CStr(vVal)
End Function

Friend Function GetOutputBalanceSchedule() As String
    Dim vVal As Variant
    vVal = m_clsTableAccess.GetVal("OutputBalanceSchedule")
    GetOutputBalanceSchedule = CStr(vVal)
End Function
'BMIDS736 End

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Set Access Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function SetOrganisationID(vVal As Variant) As Boolean
   SetOrganisationID = m_clsTableAccess.SetVal("OrganisationID", vVal)
   
   g_clsVersion.HandleVersioning Me
End Function
Friend Function SetAccInterestInd(vVal As Variant) As Boolean
   SetAccInterestInd = m_clsTableAccess.SetVal("AccInterestInd", vVal)
End Function
Friend Function SetAccInterestPayableInd(vVal As Variant) As Boolean
   SetAccInterestPayableInd = m_clsTableAccess.SetVal("AccInterestPayableInd", vVal)
End Function
Friend Function SetFirstPaymentDueDateInd(vVal As Variant) As Boolean
   SetFirstPaymentDueDateInd = m_clsTableAccess.SetVal("FirstPaymentDueDateInd", vVal)
End Function
Friend Function SetAccountingStartMonth(vVal As Variant) As Boolean
   SetAccountingStartMonth = m_clsTableAccess.SetVal("AccountingStartMonth", vVal)
End Function
Friend Function SetInterestChargedInd(vVal As Variant) As Boolean
   SetInterestChargedInd = m_clsTableAccess.SetVal("InterestChargedInd", vVal)
End Function
Friend Function SetACCInterestRoundingFactor(vVal As Variant) As Boolean
   SetACCInterestRoundingFactor = m_clsTableAccess.SetVal("ACCInterestRoundingFactor", vVal)
End Function
Friend Function SetAccInterestRoundingDirection(vVal As Variant) As Boolean
   SetAccInterestRoundingDirection = m_clsTableAccess.SetVal("AccInterestRoundingDirection", vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function SetDevonRulingInd(vVal As Variant) As Boolean
'   SetDevonRulingInd = m_clsTableAccess.SetVal("DevonRulingInd", vVal)
'End Function
'Friend Function SetInterestChangeInd(vVal As Variant) As Boolean
'   SetInterestChangeInd = m_clsTableAccess.SetVal("InterestChangeInd", vVal)
'End Function
'Friend Function SetPartAndPartRounding(vVal As Variant) As Boolean
'   SetPartAndPartRounding = m_clsTableAccess.SetVal("PartAndPartRounding", vVal)
'End Function
'BMIDS736 End
Friend Function SetMonthlyPaymentInd(vVal As Variant) As Boolean
   SetMonthlyPaymentInd = m_clsTableAccess.SetVal("MonthlyPaymentInd", vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function SetNetGrossProfileInd(vVal As Variant) As Boolean
'   SetNetGrossProfileInd = m_clsTableAccess.SetVal("NetGrossProfileInd", vVal)
'End Function
'BMIDS736 End
Friend Function SetPaymentRoundingFactor(vVal As Variant) As Boolean
   SetPaymentRoundingFactor = m_clsTableAccess.SetVal("PaymentRoundingFactor", vVal)
End Function
Friend Function SetPaymentRoundingDirection(vVal As Variant) As Boolean
   SetPaymentRoundingDirection = m_clsTableAccess.SetVal("PaymentRoundingDirection", vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function SetInterestCalcRoundingFactor(vVal As Variant) As Boolean
'   SetInterestCalcRoundingFactor = m_clsTableAccess.SetVal("InterestCalcRoundingFactor", vVal)
'End Function
'Friend Function SetInterestCalcRoundingDirection(vVal As Variant) As Boolean
'   SetInterestCalcRoundingDirection = m_clsTableAccess.SetVal("InterestCalcRoundingDirection", vVal)
'End Function
'BMIDS736 End
Friend Function SetShiftInMonths(vVal As Variant) As Boolean
   SetShiftInMonths = m_clsTableAccess.SetVal("ShiftInMonths", vVal)
End Function
Friend Function SetPositionInMonth(vVal As Variant) As Boolean
   SetPositionInMonth = m_clsTableAccess.SetVal("PositionInMonth", vVal)
End Function
Friend Function SetShiftInDays(vVal As Variant) As Boolean
   SetShiftInDays = m_clsTableAccess.SetVal("ShiftInDays", vVal)
End Function
'BMIDS736 GHun No longer used
'Friend Function SetIncludeACCInterest(vVal As Variant) As Boolean
'   SetIncludeACCInterest = m_clsTableAccess.SetVal("IncludeACCInterest", vVal)
'End Function
'Friend Function SetIncludeInACCInterest(vVal As Variant) As Boolean
'   SetIncludeInACCInterest = m_clsTableAccess.SetVal("IncludeInACCInterest", vVal)
'End Function
'BMIDS736 End

'BMIDS736 GHun
Friend Function SetCoolingOffPeriod(vVal As Variant) As Boolean
   SetCoolingOffPeriod = m_clsTableAccess.SetVal("CoolingOffPeriod", vVal)
End Function
Friend Function SetSpecifiedPaymentDay(vVal As Variant) As Boolean
   SetSpecifiedPaymentDay = m_clsTableAccess.SetVal("SpecifiedPaymentDay", vVal)
End Function
Friend Function SetDaysInYear(vVal As Variant) As Boolean
   SetDaysInYear = m_clsTableAccess.SetVal("DaysInYear", vVal)
End Function
Friend Function SetAccruedDaysIncluded(vVal As Variant) As Boolean
   SetAccruedDaysIncluded = m_clsTableAccess.SetVal("AccruedDaysIncluded", vVal)
End Function
Friend Function SetAccruedDaysAdded(vVal As Variant) As Boolean
   SetAccruedDaysAdded = m_clsTableAccess.SetVal("AccruedDaysAdded", vVal)
End Function
Friend Function SetAPRMonth(vVal As Variant) As Boolean
   SetAPRMonth = m_clsTableAccess.SetVal("APRMonth", vVal)
End Function
Friend Function SetWhenAccIntRounded(vVal As Variant) As Boolean
   SetWhenAccIntRounded = m_clsTableAccess.SetVal("WhenAccIntRounded", vVal)
End Function
Friend Function SetWhenPaymentRounded(vVal As Variant) As Boolean
   SetWhenPaymentRounded = m_clsTableAccess.SetVal("WhenPaymentRounded", vVal)
End Function
Friend Function SetWhenBalanceRounded(vVal As Variant) As Boolean
   SetWhenBalanceRounded = m_clsTableAccess.SetVal("WhenBalanceRounded", vVal)
End Function
Friend Function SetBalanceRoundingFactor(vVal As Variant) As Boolean
   SetBalanceRoundingFactor = m_clsTableAccess.SetVal("BalanceRoundingFactor", vVal)
End Function
Friend Function SetBalanceRoundingDirection(vVal As Variant) As Boolean
   SetBalanceRoundingDirection = m_clsTableAccess.SetVal("BalanceRoundingDirection", vVal)
End Function
Friend Function SetChargeRoundingFactor(vVal As Variant) As Boolean
   SetChargeRoundingFactor = m_clsTableAccess.SetVal("ChargeRoundingFactor", vVal)
End Function
Friend Function SetChargeRoundingDirection(vVal As Variant) As Boolean
   SetChargeRoundingDirection = m_clsTableAccess.SetVal("ChargeRoundingDirection", vVal)
End Function
Friend Function SetOutputPaymentSchedule(vVal As Variant) As Boolean
   SetOutputPaymentSchedule = m_clsTableAccess.SetVal("OutputPaymentSchedule", vVal)
End Function
Friend Function SetOutputBalanceSchedule(vVal As Variant) As Boolean
   SetOutputBalanceSchedule = m_clsTableAccess.SetVal("OutputBalanceSchedule", vVal)
End Function
'BMIDS736 End


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

'Private Function TableAccess_DoUpdateSets() As Boolean
'    TableAccess_DoUpdateSets = m_clsTableAccess.DoUpdateSets()
'End Function
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
