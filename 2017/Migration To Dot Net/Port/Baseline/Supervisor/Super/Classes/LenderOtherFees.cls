VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OtherFees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : OtherFees
' Description   : Manages all form processing for the Other Fees tab on the Lender Details
'                 form.
' Change history
' Prog      Date        Description
' AA        20/02/01    Changed function ValidateScreenData from a Function to a sub
' DJP       16/06/01    SQL Server port - use variants for guid keys
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Constants
Private Const ORGANISATION_ID_FIELD = 1

' Private data
Private m_clsOtherFeesTable As LenderOtherFeesTable

' DJP SQL Server port
Private m_vLenderCode As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Initialise
' Description   : Called to initialise the Legal Fees tab. Will load all data required by this tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(Optional bIsEdit As Boolean)
    On Error GoTo Failed
    
    
    ' Get the OtherFees, if there are any
    TableAccess(m_clsOtherFeesTable).GetTableData
    
    ' Set the datasource of the datagrid to the data we just retrieved.
    Set frmLenderDetails.dgOtherFees.DataSource = TableAccess(m_clsOtherFeesTable).GetRecordSet
    frmLenderDetails.dgOtherFees.Enabled = True
    
    ' Set the field names of the datagrid
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetGridFields
' Description   : Sets the fields and states of each column in the datagrid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetGridFields()
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim sOrgID As String
    
    ' First, OrganisationID. Not visible, but needs the lender code copied in
    sOrgID = g_clsSQLAssistSP.GuidToString(CStr(m_vLenderCode))
    fields.sField = "OrganisationID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = sOrgID
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Next, LenderOtherFeeName has to be entered
    fields.sField = "LenderOtherFeeName"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Fee Name must be entered"
    fields.sTitle = "Name"
    colFields.Add fields
    
    ' StartDate has to be entered
    fields.sField = "LenderOtherFeeStartDate"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Start Date must be entered"
    fields.sTitle = "Start Date"
    fields.bDateField = True
    colFields.Add fields
    
    ' Fee Amount
    fields.sField = "LenderOtherFeeAmount"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Fee Amount must be entered"
    fields.sTitle = "Amount"
    fields.bDateField = False
    
    colFields.Add fields
    
    frmLenderDetails.dgOtherFees.SetColumns colFields, "EditOtherFees", "Other Fees"
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenFields
' Description   : Populates all fields on this tab, but there aren't any as it is a listview.
'                 This method is called externally so still has to be supported even though it
'                 doesn't do anything.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetScreenFields()
    On Error GoTo Failed
        
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoUpdates
' Description   : Commits all data on this tab to the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DoUpdates()
    On Error GoTo Failed
    
    TableAccess(m_clsOtherFeesTable).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitTab
' Description   : Called externally to this class and performs any tab initialisation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitTab()

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeyValues
' Description   : Sets the key values to identify the list of Other Fees to be displayed on this
'                 tab.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeyValues(colValues As Collection)
    On Error GoTo Failed
    ' DJP SQL Server port
    Dim vLenderCode As Variant
    
    vLenderCode = colValues.Item(ORGANISATION_ID_FIELD)
    
    If Len(vLenderCode) > 0 Then
        m_vLenderCode = vLenderCode

        TableAccess(m_clsOtherFeesTable).SetKeyMatchValues colValues
    Else
        g_clsErrorHandling.RaiseError errEmptyLenderCode
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError errCanNotUpdateBatchJob
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called externally to this class and performs any tab validation, if required.
'                 Name and start date must be unique.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ValidateScreenData()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = frmLenderDetails.dgOtherFees.ValidateRows()
    
    ' If the above failed it will have displayed its own error, so just break out.
    If bRet = False Then
        g_clsErrorHandling.RaiseError errDontDisplayError
    End If

    bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsOtherFeesTable)
    
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Other Fees must have unique Name and Start Date"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Class_Initialize()
    On Error GoTo Failed
    ' Create table classes
    
    If m_clsOtherFeesTable Is Nothing Then
        Set m_clsOtherFeesTable = New LenderOtherFeesTable
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

