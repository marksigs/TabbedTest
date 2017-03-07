VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdSpecialGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : MortProdSpecialGroup
' Description   : Provides functionality for the Special Group tab on Mortgage Product.
' Change history
' Prog      Date        Description
' DJP       20/11/01    SYS2912 SQL Server locking problem.
' DJP       10/12/01    SYS2831 Client variants.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


' Private data
Private m_bIsEdit As Boolean
Private m_colProductKeys As Collection
Private m_sProductCode As String
Private m_sProductStartDate As String
Private m_clsOtherSpecialGroup As MortProdSpecialGroupTable
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Called when this class is created. Create all module data as required.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'-------------------------------------------------------------------------------------------------
' BMIDS Change history
' Prog      Date        Description
' GD        23/05/02    BMIDS00014
'-------------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    If m_clsOtherSpecialGroup Is Nothing Then
        Set m_clsOtherSpecialGroup = New MortProdSpecialGroupTable
    End If
    Set clsTableAccess = m_clsOtherSpecialGroup
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub Initialise(Optional bIsEdit As Boolean)
    On Error GoTo Failed
    Dim rsGroup As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim nPopulateType As PopulateType
        
    m_bIsEdit = bIsEdit
    
    Set clsTableAccess = m_clsOtherSpecialGroup
    nPopulateType = POPULATE_EMPTY
    
    If m_bIsEdit Then
        nPopulateType = POPULATE_KEYS
        clsTableAccess.SetKeyMatchValues m_colProductKeys
    End If
    
    Set rsGroup = clsTableAccess.GetTableData(nPopulateType)
    clsTableAccess.ValidateData
    
    Set frmProductDetails.dgSpecialGroup.DataSource = rsGroup
    frmProductDetails.dgSpecialGroup.Enabled = True
    
    SetGridFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub SetGridFields()
'GD BMIDS 0014
    On Error GoTo Failed
    Dim fields As FieldData
    Dim colFields As New Collection
    
    Dim colComboValues As New Collection
    Dim colComboIDS As New Collection
    
    ' Mortgage Product Code
    fields.sField = "MortgageProductCode"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sProductCode
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Mortgage Product Start Date
    fields.sField = "StartDate"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = m_sProductStartDate
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Special Group

    GetSpecialGroupsAsCollection colComboValues, colComboIDS
    
    'GD BMIDS 0014
    'Group Type Sequence Number
    fields.sField = "GroupTypeSequenceNumber"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = ""
    fields.sError = ""
    fields.sTitle = "" ' No need to display, since this column is not displayed

    Set fields.colComboValues = Nothing
    colFields.Add fields

    fields.sField = "GroupType"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Special Group must be entered"
    fields.sTitle = "Special Group"
    fields.sOtherField = "GroupTypeSequenceNumber"
    'GD BMIDS0014

    Set fields.colComboValues = colComboValues
    Set fields.colComboIDS = colComboIDS

    colFields.Add fields
    
    frmProductDetails.dgSpecialGroup.SetColumns colFields, "SpecialGroup", "Special Groups"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetScreenFields()

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub DoUpdates()
    On Error GoTo Failed
    Dim clsBandedTable As BandedTable
    
    Set clsBandedTable = m_clsOtherSpecialGroup
    
    clsBandedTable.SetUpdateValues m_colProductKeys
    clsBandedTable.SetUpdateSets
    clsBandedTable.DoUpdateSets

    TableAccess(m_clsOtherSpecialGroup).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveScreenData()
    On Error GoTo Failed
    
    DoUpdates
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub InitTab()
    On Error GoTo Failed
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function ValidateScreenData()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = frmProductDetails.dgSpecialGroup.ValidateRows()

    If bRet Then
        bRet = Not g_clsFormProcessing.CheckForDuplicates(m_clsOtherSpecialGroup)
    
        If bRet = False Then
            frmProductDetails.dgSpecialGroup.SetGridFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Special Group must be unique"
        End If
    End If
    
    ValidateScreenData = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetProductKeys(colValues As Collection)
    On Error GoTo Failed
    
    If colValues Is Nothing Then
        g_clsErrorHandling.RaiseError errKeysEmpty, "MortProdSpecialGroup.SetProductKeys"
    End If
    
    m_sProductCode = colValues(PRODUCT_CODE_KEY)
    m_sProductStartDate = colValues(PRODUCT_START_DATE_KEY)
    
    Set m_colProductKeys = colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetIsEdit(Optional bEdit As Boolean = True)
    m_bIsEdit = bEdit
End Sub

Public Function IsEdit()
    IsEdit = m_bIsEdit
End Function
Private Sub GetSpecialGroupsAsCollection(colValues As Collection, colIDS As Collection)
    'GD BMIDS0014
    On Error GoTo Failed
    Dim sType As String
    Dim sSearch As String
    Dim sTypeText As String

    Dim rs As ADODB.Recordset
    


    sSearch = "select valuename GROUPTYPE,valueid GROUPTYPESEQUENCENUMBER from combovalue where groupname ='SpecialGroup'"

    Set rs = g_clsDataAccess.GetTableData("COMBOVALUE", sSearch)
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        While Not rs.EOF
            sTypeText = rs.fields("GROUPTYPE").Value
            colValues.Add sTypeText
            
            sType = rs.fields("GROUPTYPESEQUENCENUMBER").Value
            colIDS.Add sType
            rs.MoveNext
        Wend
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No Special Groups exist"
    End If
    rs.Close
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
