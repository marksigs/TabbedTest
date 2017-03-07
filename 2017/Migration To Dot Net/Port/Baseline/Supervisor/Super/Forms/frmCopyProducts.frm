VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "*\A..\..\MSGOCX\MSGOCX.vbp"
Begin VB.Form frmCopyProducts 
   Caption         =   "Copy Product"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frmCopyProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4500
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   5835
      TabIndex        =   5
      Top             =   1260
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   0
      Left            =   2460
      TabIndex        =   0
      Top             =   60
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   6
   End
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   2
      Left            =   2460
      TabIndex        =   2
      Top             =   780
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Mask            =   "##/##/####"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtProductDetails 
      Height          =   315
      Index           =   1
      Left            =   2460
      TabIndex        =   1
      Top             =   420
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   50
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   780
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "Start Time"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   915
   End
   Begin VB.Label lblProductDetails 
      Caption         =   "Mortgage Product Name"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "New Start Date"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblMainProdDetails 
      Caption         =   "New Mortgage Product Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmCopyProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :
' Description   :
'
' Change history
' Prog      Date        Description
' STB
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Change history
' Prog      Date        Description
' GD        06/06/2002  BMIDS00016
' JD        12/07/2004  BMIDS775  added copying of mortgage product AdditionalFeatures
' HMA       09/12/2004  BMIDS959  Remove MortgageProductBands table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const constMortgageProduct As Integer = 1
Private Const constAdditionalParameter As Integer = 2
Private Const constPurposeOfLoan As Integer = 3
Private Const constTypeOfBuyer As Integer = 4
Private Const constChannelEligibility As Integer = 5
Private Const constOtherFees As Integer = 6
Private Const constBaseRates As Integer = 7

Private m_colMatchValues As Collection

' Constants
Private Const constTableCount As Integer = 2
Private Const PRODUCT_CODE = 0
Private Const PRODUCT_NAME = 1
Private Const START_DATE = 2
Private m_RowValues(constTableCount) As UpdateValues
Private m_ReturnCode As MSGReturnCode
Private Sub cmdCancel_Click()
    Me.Hide
End Sub
Public Sub SetKeys(colMatchValues As Collection)
    On Error GoTo Failed
    Set m_colMatchValues = colMatchValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    If g_clsFormProcessing.DoMandatoryProcessing(Me) Then
        ValidateScreenData
        CopyProduct
        
        SetReturnCode
        Me.Hide
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SaveChangeRequest(clsTableAccess As TableAccess)
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sProductCode As String
    Dim sDate As String
    
    sProductCode = txtProductDetails(PRODUCT_CODE).Text
    g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), sDate, GET_CONTROL_VALUE
    
    Set colMatchValues = New Collection
    
    colMatchValues.Add sProductCode
    colMatchValues.Add sDate
    
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ValidateScreenData()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsMortProd As TableAccess
    Dim bTimeValid As Boolean
     
    Set colMatchValues = New Collection
    Set clsMortProd = New MortgageProductTable
    
    GetNewKeyMatchValues colMatchValues
    
    If clsMortProd.DoesRecordExist(colMatchValues) Then
        txtProductDetails(PRODUCT_CODE).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Product already exists - please enter a unique Product Code and Start Date"
    End If
    
    ' Validate the time
       
    bTimeValid = g_clsValidation.ValidateTime(txtStartTime.ClipText)
    
    If Not bTimeValid Then
        txtStartTime.SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Time must be between 00:00:00 and 23:59:59"
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetNewKeyMatchValues(colMatchValues As Collection)
    On Error GoTo Failed
    Dim sProductCode As String
    Dim sStartDate As String
    
    sProductCode = txtProductDetails(PRODUCT_CODE).Text
    g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), sStartDate, GET_CONTROL_VALUE
    
    colMatchValues.Add sProductCode
    colMatchValues.Add sStartDate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CopyProduct()
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim vProductCode As Variant
    Dim vDate As Variant
    Dim colIntRates As Collection
    Dim vIntRateKey As Variant
    Dim sTime As String
    
    vProductCode = txtProductDetails(PRODUCT_CODE).Text
    g_clsFormProcessing.HandleDate txtProductDetails(START_DATE), vDate, GET_CONTROL_VALUE
    
    sTime = txtStartTime.ClipText
    
    If Len(sTime) > 0 Then
        vDate = vDate + " " + txtStartTime.Text
    End If
    
    Set colIntRates = New Collection
    g_clsDataAccess.BeginTrans
    
    On Error GoTo Failed
    Set clsTableAccess = New MortgageProductTable

    CopyProductTable clsTableAccess, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdParamsTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdChannelEligTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdTypeOfBuyerTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdPurpOfLoanTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdLanguageTable, m_colMatchValues, vProductCode, vDate
    ' BMIDS959  Remove MortgageProductBands table
    CopyProductTable New OtherFeeTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdTypicalAPRTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdTypeofAppEligTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdEmpEligTable, m_colMatchValues, vProductCode, vDate
    CopyProductTable New MortProdPropLocTable, m_colMatchValues, vProductCode, vDate
'GD BMIDS00016
    CopyProductTable New MortProdProdCondTable, m_colMatchValues, vProductCode, vDate


    CopyProductTable New MortProdIntRatesTable, m_colMatchValues, vProductCode, vDate
    
    CopyProductTable New MortProdSpecialGroupTable, m_colMatchValues, vProductCode, vDate
    
    'JD BMIDS775
    CopyProductTable New MortProdAddtFeaTable, m_colMatchValues, vProductCode, vDate

    
    CopyIncentives Inclusive, vProductCode, vDate
    CopyIncentives Exclusive, vProductCode, vDate
    g_clsDataAccess.CommitTrans
    
    SaveChangeRequest clsTableAccess
    
    Exit Sub
Failed:
    g_clsDataAccess.RollbackTrans
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CopyIncentives(nIncentive As IncentiveType, vProductCode As Variant, vDate As Variant)
    On Error GoTo Failed
    ' Need to copy Incentive, InclusiveIncentive and ExclusiveIncentive
    ' First, copy the Incentive table and create a new GUID
    Dim clsTableAccess As TableAccess
    Dim clsIncentive As New MortProdIncentiveTable
    Dim clsInclusiveIncentive As New MortProdIncIncentiveTable
    Dim nThisField As Long
    Dim rsInclusiveIncentive As ADODB.Recordset
    Dim rsIncentive As ADODB.Recordset
    Dim rsClone As ADODB.Recordset
    Dim vOldGUID As Variant
    Dim vNewGUID As Variant
    Dim colMatchValues As New Collection
    Dim rowInclusiveIncentive As Variant
    Dim rowIncentive As Variant
    Dim nFields As Long
    Dim colIncentiveKey As Collection
    Dim vBookMark As Variant
    Dim nNumRows As Long
    Dim vLastBookmark As Variant
    Dim nRow As Long
    
    ' First, the Inclusive ones...
    Set clsTableAccess = clsInclusiveIncentive
    clsInclusiveIncentive.SetType nIncentive

    clsTableAccess.SetKeyMatchValues m_colMatchValues
    Set rsInclusiveIncentive = clsTableAccess.GetTableData()

    If clsTableAccess.RecordCount > 0 Then
        rowInclusiveIncentive = rsInclusiveIncentive.GetRows()

        ' Need to copy everything for this row in the incentive table
        'rsInclusiveIncentive.MoveFirst
        Set rsClone = rsInclusiveIncentive.Clone

        'rowInclusiveIncentive = rsInclusiveIncentive.GetRows()

        rsInclusiveIncentive.MoveLast
        vLastBookmark = rsInclusiveIncentive.Bookmark

        rsInclusiveIncentive.MoveFirst
        vBookMark = rsInclusiveIncentive.Bookmark
        nRow = 0
        
        While (Not rsInclusiveIncentive.EOF) And rsInclusiveIncentive.Bookmark <= vLastBookmark
            Set clsTableAccess = clsInclusiveIncentive
            vBookMark = rsInclusiveIncentive.Bookmark

            nFields = UBound(rowInclusiveIncentive, 1) + 1
            ' Copy the existing fields
            rsClone.AddNew

            ' All fielnds in the Incentive table
            For nThisField = 0 To nFields - 1
                rsClone(nThisField).Value = rowInclusiveIncentive(nThisField, nRow)
            Next

            ' Get the current GUID for this record
            clsTableAccess.SetRecordSet rsInclusiveIncentive
            vOldGUID = clsInclusiveIncentive.GetIncentiveGUID()
            
            clsTableAccess.SetRecordSet rsClone
            vNewGUID = clsInclusiveIncentive.SetIncentiveGUID()
            
            clsInclusiveIncentive.SetMortgageProductCode CStr(vProductCode)
            clsInclusiveIncentive.SetStartDate vDate

            ' Find the record in the Inclusive table
            Set clsTableAccess = clsIncentive

            Set colIncentiveKey = New Collection
            colIncentiveKey.Add vOldGUID
            clsTableAccess.SetKeyMatchValues colIncentiveKey

            Set rsIncentive = clsTableAccess.GetTableData()

            If rsIncentive.RecordCount = 1 Then
                rowIncentive = rsIncentive.GetRows()
                ' Should only be one

                nFields = UBound(rowIncentive, 1) + 1

                ' Copy the existing fields
                rsIncentive.AddNew

                For nThisField = 0 To nFields - 1
                    rsIncentive(nThisField).Value = rowIncentive(nThisField, 0)
                Next

                ' Set the new GUID
                clsIncentive.SetIncentiveGUID vNewGUID
                clsTableAccess.Update
                clsTableAccess.CloseRecordSet

            End If
            nRow = nRow + 1
            rsInclusiveIncentive.Bookmark = vBookMark
            rsInclusiveIncentive.MoveNext
        Wend

        Set clsTableAccess = clsInclusiveIncentive
        clsTableAccess.Update
        clsTableAccess.CloseRecordSet
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CopyProductTable(clsTableAccess As TableAccess, colMatchValues As Collection, Optional vKey1 As Variant, Optional vKey2 As Variant)
    Dim sProductCode As String
    Dim sStartDate As String
    Dim vDate As Variant
    Dim rowData As Variant
    Dim rs As ADODB.Recordset
    Dim nFields As Long
    Dim nRows As Long
    Dim nThisRow As Long
    Dim nThisField As Long
    On Error GoTo Failed

    If g_clsDataAccess.DoesTableExist(clsTableAccess.GetTable()) Then
        If clsTableAccess.GetTable = "INTERESTRATETYPE" Then
           Dim clsInterestRateTable As MortProdIntRatesTable
           
           Set clsInterestRateTable = clsTableAccess
           clsInterestRateTable.GetInterestRates colMatchValues
           Set rs = clsTableAccess.GetRecordSet()
        Else
            clsTableAccess.SetKeyMatchValues colMatchValues
            Set rs = clsTableAccess.GetTableData()
        End If
    
        If clsTableAccess.RecordCount > 0 Then
            rowData = rs.GetRows()
            nRows = UBound(rowData, 2)
        
            For nThisRow = 0 To nRows
                nFields = UBound(rowData, 1) + 1
                ' Copy the existing fields
                
                rs.AddNew
                For nThisField = 0 To nFields - 1
                    rs(nThisField).Value = rowData(nThisField, nThisRow)
                Next
                
                SaveKeys clsTableAccess, vKey1, vKey2
            Next
            
            clsTableAccess.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveKeys(clsTableAccess As TableAccess, _
                     vKey1 As Variant, _
                     Optional vKey2 As Variant)
    ' Save the keys on the current row
    On Error GoTo Failed
    
    If TypeOf clsTableAccess Is MortgageProductTable Then
        Dim clsMortProd As MortgageProductTable
        Set clsMortProd = clsTableAccess
        clsMortProd.SetMortgageProductCode vKey1
        clsMortProd.SetStartDate vKey2
    
    ElseIf TypeOf clsTableAccess Is MortProdParamsTable Then
        Dim clsMortProdParams As MortProdParamsTable
        
        Set clsMortProdParams = clsTableAccess
        clsMortProdParams.SetMortgageProductCode vKey1
        clsMortProdParams.SetStartDate vKey2
    
    ElseIf TypeOf clsTableAccess Is MortProdChannelEligTable Then
        Dim clsMortProdChan As MortProdChannelEligTable
        
        Set clsMortProdChan = clsTableAccess
        clsMortProdChan.SetProductCode CStr(vKey1)
        clsMortProdChan.SetStartDate vKey2
    
    ElseIf TypeOf clsTableAccess Is MortProdTypeOfBuyerTable Then
        Dim clsMortProdBuyer As MortProdTypeOfBuyerTable
        
        Set clsMortProdBuyer = clsTableAccess
        clsMortProdBuyer.SetProductCode CStr(vKey1)
        clsMortProdBuyer.SetStartDate vKey2
    
    ElseIf TypeOf clsTableAccess Is MortProdPurpOfLoanTable Then
        Dim clsMortPurpLoan As MortProdPurpOfLoanTable
        
        Set clsMortPurpLoan = clsTableAccess
        clsMortPurpLoan.SetProductCode CStr(vKey1)
        clsMortPurpLoan.SetStartDate vKey2
    
    ElseIf TypeOf clsTableAccess Is MortProdIntRatesTable Then
        Dim clsMortIntRates As MortProdIntRatesTable
        
        Set clsMortIntRates = clsTableAccess
        clsMortIntRates.SetMortgageProductCode CStr(vKey1)
        clsMortIntRates.SetStartDate vKey2

    ElseIf TypeOf clsTableAccess Is MortProdSpecialGroupTable Then
        Dim clsMortSpecialGroup As MortProdSpecialGroupTable
        
        Set clsMortSpecialGroup = clsTableAccess
        clsMortSpecialGroup.SetProductCode CStr(vKey1)
        clsMortSpecialGroup.SetStartDate vKey2

    ElseIf TypeOf clsTableAccess Is MortProdLanguageTable Then
        Dim clsMortLanguage As MortProdLanguageTable
        Dim sProductName As String
        
        sProductName = txtProductDetails(PRODUCT_NAME).Text
        
        Set clsMortLanguage = clsTableAccess
        clsMortLanguage.SetProductCode CStr(vKey1)
        clsMortLanguage.SetStartDate vKey2
        
        clsMortLanguage.SetProductName sProductName
        
        ' BMIDS959  Remove MortgageProductBands table
        
    ElseIf TypeOf clsTableAccess Is OtherFeeTable Then
        Dim clsOtherFee As OtherFeeTable
        
        Set clsOtherFee = clsTableAccess
        clsOtherFee.SetProductCode CStr(vKey1)
        clsOtherFee.SetStartDate vKey2

    ElseIf TypeOf clsTableAccess Is MortProdTypicalAPRTable Then
        Dim clsTypicalAPR As MortProdTypicalAPRTable
        
        Set clsTypicalAPR = clsTableAccess
        clsTypicalAPR.SetProductCode CStr(vKey1)
        clsTypicalAPR.SetStartDate vKey2

    ElseIf TypeOf clsTableAccess Is MortProdTypeofAppEligTable Then
        Dim clsTypeOfApp As MortProdTypeofAppEligTable
        
        Set clsTypeOfApp = clsTableAccess
        clsTypeOfApp.SetProductCode CStr(vKey1)
        clsTypeOfApp.SetStartDate vKey2

    ElseIf TypeOf clsTableAccess Is MortProdEmpEligTable Then
        Dim clsTypeOfEmpElig As MortProdEmpEligTable
        
        Set clsTypeOfEmpElig = clsTableAccess
        clsTypeOfEmpElig.SetProductCode CStr(vKey1)
        clsTypeOfEmpElig.SetStartDate vKey2
    ElseIf TypeOf clsTableAccess Is MortProdPropLocTable Then
        Dim clsPropLocation As MortProdPropLocTable
        
        Set clsPropLocation = clsTableAccess
        clsPropLocation.SetProductCode CStr(vKey1)
        clsPropLocation.SetStartDate vKey2
'GD BMIDS00016
    ElseIf TypeOf clsTableAccess Is MortProdProdCondTable Then
        Dim clsMortProdCond As MortProdProdCondTable

        Set clsMortProdCond = clsTableAccess
        clsMortProdCond.SetProductCode CStr(vKey1)
        clsMortProdCond.SetStartDate vKey2
'JD BMIDS775
    ElseIf TypeOf clsTableAccess Is MortProdAddtFeaTable Then
        Dim clsMortProdAddtFea As MortProdAddtFeaTable

        Set clsMortProdAddtFea = clsTableAccess
        clsMortProdAddtFea.SetProductCode CStr(vKey1)
        clsMortProdAddtFea.SetStartDate vKey2
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Class " & TypeName(clsTableAccess) & " not recognised"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function GetCurrentKeys(clsTableAccess As TableAccess) As Collection
    On Error GoTo Failed
    Dim colMatchValues As New Collection
    
    If TypeOf clsTableAccess Is MortProdIncentiveTable Then
        Dim clsInclusive As MortProdIncentiveTable
        Set clsInclusive = clsTableAccess
        colMatchValues.Add clsInclusive.GetIncentiveGUID()
    
'    ElseIf TypeOf clsTableAccess Is MortProdExIncentiveTable Then
'        Dim clsExlusive As MortProdExIncentiveTable
'        Set clsExlusive = clsTableAccess
'        colMatchValues.Add clsExlusive.GetIncentiveGUID()
    End If
    Set GetCurrentKeys = colMatchValues
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub Form_Load()
    SetReturnCode MSGFailure
    EndWaitCursor
End Sub
Private Sub txtProductDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtProductDetails(Index).ValidateData
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
Private Sub txtStartTime_GotFocus()
    On Error GoTo Failed
    
    ' Make sure the cursor is at the start
    txtStartTime.Text = txtStartTime.Text

    ' Highlight the text, if there is any.
    If Len(txtStartTime.ClipText) > 0 Then
        txtStartTime.SelStart = 0
        txtStartTime.SelLength = Len(txtStartTime.Text)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

