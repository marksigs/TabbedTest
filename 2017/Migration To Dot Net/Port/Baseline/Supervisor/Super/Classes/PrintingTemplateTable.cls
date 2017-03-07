VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintingTemplateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : PaymentProcessingTable
' Description   : Contains all data access functions for the PaymentProcessingBatchParameters Table
' Change history
' Prog      Date        Description
' AA        13/02/01    Added class
' DJP       18/06/01    Removed GetNextTemplateID as it wasn't being used anymore.
' DJP       22/06/01    SQL Server port, use global SQLAssist not local one
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS HISTORY
'
' Prog      Date        Description
' DB        04/11/02    BMIDS00840 - Added Get & Set methods for edit/view before print.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MARS HISTORY
'
' Prog      Date        Description
' GHun      13/07/2005  MAR7    Integrate local printing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'EPSOM HISTORY
'
' Prog      Date        Description
' TW        25/01/2007  EP2_990 - Apply EP1257 to Epsom 2 Supervisor (CC56 -  Gemini printing  Supervisor changes)
' GHun      13/02/2007  EP2_1363 Added get and set for EmailRecipientType
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_clsTableAccess As TableAccess
Private m_sSearch As String
Private m_bUpdated As Boolean
Private Const m_sTableName As String = "HostTemplate"
Private Const m_sTemplateTable As String = "AvailableTemplates"
Private Const m_sComboDocumentGroup As String = "PrintDocumentType"
Private Const m_sComboUserRole As String = "UserRole"

Implements TableAccess

Private Sub Class_Initialize()
    On Error GoTo Failed
    
    Dim colFields As New Collection
    
    Set m_clsTableAccess = New TableAccess
           
    m_clsTableAccess.Create m_sTableName
    
    colFields.Add "HOSTTEMPLATEID"
    m_clsTableAccess.SetKeyMatchFields colFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

'MAR7 GHun
Friend Function GetPrintToWeb() As Variant
    GetPrintToWeb = m_clsTableAccess.GetVal("URLPostInd")
End Function

Friend Function SetPrintToWeb(vVal As Variant) As Boolean
    SetPrintToWeb = m_clsTableAccess.SetVal("URLPostInd", vVal)
End Function
'MAR7 End

Private Sub TableAccess_AddRow()
    m_clsTableAccess.AddRow
End Sub

Private Sub TableAccess_ApplyFilter(Optional sSearch As String)
    m_clsTableAccess.ApplyFilter sSearch
End Sub

Private Sub TableAccess_CancelFilter()
    m_clsTableAccess.CancelFilter
End Sub

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
    colListLine.Add GetHostTemplateID
    colListLine.Add GetDPSTemplateID
    colListLine.Add GetTemplateGroupText
    colListLine.Add GetHostTemplateName
    colListLine.Add GetHostTemplateDescription
    colListLine.Add GetMinRoleLevelText
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

Public Sub GetPrintTemplates()

    Dim sSQL As String
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    'SQL Server port, use global SQLAssist not local one
    sSQL = "Select " & m_sTableName & ".*, ComboDocument.ValueName GroupTypeText, ComboUserRole.ValueName RoleTypeText " & _
           "From " & m_sTableName & ", ComboValue ComboDocument, ComboValue ComboUserRole " & _
           "Where ComboDocument.GroupName = " & g_clsSQLAssistSP.FormatString(m_sComboDocumentGroup) & " And ComboDocument.ValueID = TemplateGroupId And " & _
           "ComboUserRole.GroupName = " & g_clsSQLAssistSP.FormatString(m_sComboUserRole) & " AND ComboUserRole.ValueID = MinRoleLevel"

    Set rs = g_clsDataAccess.GetTableData(m_sTableName, sSQL)
    
    ValidateRecordset rs, "HostTemplates"
    m_clsTableAccess.SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Function GetHostTemplateID() As Variant
    GetHostTemplateID = m_clsTableAccess.GetVal("HOSTTEMPLATEID")
End Function

Friend Function GetPrintManagerMethod() As Variant
    GetPrintManagerMethod = m_clsTableAccess.GetVal("PDMMETHOD")
End Function

Friend Function GetDPSTemplateID() As Variant
    GetDPSTemplateID = m_clsTableAccess.GetVal("DPSTEMPLATEID")
End Function

Friend Function GetTemplateGroupID() As Variant
    GetTemplateGroupID = m_clsTableAccess.GetVal("TEMPLATEGROUPID")
End Function

Friend Function GetTemplateGroupText() As Variant
    GetTemplateGroupText = m_clsTableAccess.GetVal("GroupTypeText")
End Function

Friend Function GetHostTemplateName() As Variant
    GetHostTemplateName = m_clsTableAccess.GetVal("HOSTTEMPLATENAME")
End Function

Friend Function GetHostTemplateDescription() As Variant
    GetHostTemplateDescription = m_clsTableAccess.GetVal("HOSTTEMPLATEDESCRIPTION")
End Function

Friend Function GetMinRoleLevel() As Variant
    GetMinRoleLevel = m_clsTableAccess.GetVal("MINROLELEVEL")
End Function

Friend Function GetMinRoleLevelText() As Variant
    GetMinRoleLevelText = m_clsTableAccess.GetVal("ROLETYPETEXT")
End Function

Friend Function GetMaxCopies() As Variant
    GetMaxCopies = m_clsTableAccess.GetVal("MAXCOPIES")
End Function

Friend Function GetDefaultCopies() As Variant
    GetDefaultCopies = m_clsTableAccess.GetVal("DEFAULTCOPIES")
End Function

Friend Function GetCustSpecificInd() As Variant
    GetCustSpecificInd = m_clsTableAccess.GetVal("CUSTOMERSPECIFICIND")
End Function

Friend Function GetEditBeforePrintInd() As Variant
    GetEditBeforePrintInd = m_clsTableAccess.GetVal("EDITBEFOREPRINT")
End Function
Friend Function GetViewBeforePrintInd() As Variant
    GetViewBeforePrintInd = m_clsTableAccess.GetVal("VIEWBEFOREPRINT")
End Function

'*=[MC]BMIDS770 - CC072 GET PREVENTEDITINDMS VALUE
Friend Function GetPreventEditInDMS() As Variant
    GetPreventEditInDMS = m_clsTableAccess.GetVal("PREVENTEDITINDMS")
End Function


Friend Function GetRecipientTypeID() As Variant
    GetRecipientTypeID = m_clsTableAccess.GetVal("RECIPIENTTYPE")
End Function

Friend Function GetPrintMenuAccessInd() As Variant
    GetPrintMenuAccessInd = m_clsTableAccess.GetVal("PRINTMENUACCESSIND")
End Function

Friend Function GetPrinterDestinationType() As Variant
    GetPrinterDestinationType = m_clsTableAccess.GetVal("PRINTERDESTINATIONTYPE")
End Function

Friend Function GetRemotePrinterLocation() As Variant
    GetRemotePrinterLocation = m_clsTableAccess.GetVal("REMOTEPRINTERLOCATION")
End Function

Friend Function GetInactiveIndicator() As Variant
    GetInactiveIndicator = m_clsTableAccess.GetVal("INACTIVEINDICATOR")
End Function

Friend Function SetTemplateID(vVal As Variant) As Boolean
    SetTemplateID = m_clsTableAccess.SetVal("HOSTTEMPLATEID", vVal)
End Function

Friend Function SetDPSTemplateID(vVal As Variant) As Boolean
    SetDPSTemplateID = m_clsTableAccess.SetVal("DPSTEMPLATEID", vVal)
End Function

Friend Function SetTemplateGroupID(vVal As Variant) As Boolean
    SetTemplateGroupID = m_clsTableAccess.SetVal("TEMPLATEGROUPID", vVal)
End Function

Friend Function SetHostTemplateName(vVal As Variant) As Boolean
    SetHostTemplateName = m_clsTableAccess.SetVal("HOSTTEMPLATENAME", vVal)
End Function

Friend Function SetHostTemplateDescription(vVal As Variant) As Boolean
    SetHostTemplateDescription = m_clsTableAccess.SetVal("HOSTTEMPLATEDESCRIPTION", vVal)
End Function

Friend Function SetMinRoleLevel(vVal As Variant) As Boolean
    SetMinRoleLevel = m_clsTableAccess.SetVal("MINROLELEVEL", vVal)
End Function

Friend Function SetMaxCopies(vVal As Variant) As Boolean
    SetMaxCopies = m_clsTableAccess.SetVal("MAXCOPIES", vVal)
End Function

Friend Function SetDefaultCopies(vVal As Variant) As Boolean
    SetDefaultCopies = m_clsTableAccess.SetVal("DEFAULTCOPIES", vVal)
End Function

Friend Function SetCustomerSpecificInd(vVal As Variant) As Boolean
    SetCustomerSpecificInd = m_clsTableAccess.SetVal("CUSTOMERSPECIFICIND", vVal)
End Function

Friend Function SetEditBeforePrintInd(vVal As Variant) As Boolean
    SetEditBeforePrintInd = m_clsTableAccess.SetVal("EDITBEFOREPRINT", vVal)
End Function
Friend Function SetViewBeforePrintInd(vVal As Variant) As Boolean
    SetViewBeforePrintInd = m_clsTableAccess.SetVal("VIEWBEFOREPRINT", vVal)
End Function

'*=[MC]BMIDS770 - CC072 - PREVENT EDIT ONCE DOCUMENTS STORED IN DMS
Friend Function SetPreventEditInDMS(vVal As Variant) As Boolean
    SetPreventEditInDMS = m_clsTableAccess.SetVal("PREVENTEDITINDMS", vVal)
End Function
'*=[MC]EDIT SECTION END CC072



Friend Function SetRecipientType(vVal As Variant) As Boolean
    SetRecipientType = m_clsTableAccess.SetVal("RECIPIENTTYPE", vVal)
End Function

Friend Function SetPrintMenuAccessInd(vVal As Variant) As Boolean
    SetPrintMenuAccessInd = m_clsTableAccess.SetVal("PRINTMENUACCESSIND", vVal)
End Function

Friend Function SetPrintManagerMethodName(vVal As Variant) As Boolean
    SetPrintManagerMethodName = m_clsTableAccess.SetVal("PDMMETHOD", vVal)
End Function

Friend Function SetPrintDestinationType(vVal As Variant) As Boolean
    SetPrintDestinationType = m_clsTableAccess.SetVal("PRINTERDESTINATIONTYPE", vVal)
End Function

Friend Function SetRemotePrinterLocation(vVal As Variant) As Boolean
    SetRemotePrinterLocation = m_clsTableAccess.SetVal("REMOTEPRINTERLOCATION", vVal)
End Function

Friend Function SetInactiveIndicator(vVal As Variant) As Boolean
    SetInactiveIndicator = m_clsTableAccess.SetVal("INACTIVEINDICATOR", vVal)
End Function

Public Function GetListField() As String
    GetListField = "ValueName"
End Function

'MAR7 GHun Add functions for Document Delivery Type
Friend Function GetDocumentDeliveryType() As Variant
    GetDocumentDeliveryType = m_clsTableAccess.GetVal("DELIVERYTYPE")
End Function
'
Friend Function SetDocumentDeliveryType(vVal As Variant) As Boolean
    SetDocumentDeliveryType = m_clsTableAccess.SetVal("DELIVERYTYPE", vVal)
End Function

Friend Function GetFirstPagePrinterTray() As Variant
    GetFirstPagePrinterTray = m_clsTableAccess.GetVal("FIRSTPAGEPRINTERTRAY")
End Function

Friend Function SetFirstPagePrinterTray(vVal As Variant) As Boolean
    SetFirstPagePrinterTray = m_clsTableAccess.SetVal("FIRSTPAGEPRINTERTRAY", vVal)
End Function

Friend Function GetOtherPagesPrinterTray() As Variant
    GetOtherPagesPrinterTray = m_clsTableAccess.GetVal("OTHERPAGESPRINTERTRAY")
End Function

Friend Function SetOtherPagesPrinterTray(vVal As Variant) As Boolean
    SetOtherPagesPrinterTray = m_clsTableAccess.SetVal("OTHERPAGESPRINTERTRAY", vVal)
End Function
'MAR7 - End
' TW 25/01/2007 EP2_990
Friend Function GetGeminiPrintMode() As Variant
    GetGeminiPrintMode = m_clsTableAccess.GetVal("GEMINIPRINTMODE")
End Function
'
Friend Function SetGeminiPrintMode(vVal As Variant) As Boolean
    SetGeminiPrintMode = m_clsTableAccess.SetVal("GEMINIPRINTMODE", vVal)
End Function

Friend Function GetPrecreateForPackInd() As Variant
    GetPrecreateForPackInd = m_clsTableAccess.GetVal("PRECREATEFORPACK")
End Function

Friend Function SetPrecreateForPackInd(vVal As Variant) As Boolean
    SetPrecreateForPackInd = m_clsTableAccess.SetVal("PRECREATEFORPACK", vVal)
End Function
' TW 25/01/2007 EP2_990 End

'EP2_1363 GHun
Friend Function GetEmailRecipientType() As Variant
    GetEmailRecipientType = m_clsTableAccess.GetVal("EMAILRECIPIENTTYPE")
End Function

Friend Function SetEmailRecipientType(vVal As Variant) As Boolean
    SetEmailRecipientType = m_clsTableAccess.SetVal("EMAILRECIPIENTTYPE", vVal)
End Function
'EP2_1363 End
