VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortProdAddtFeaCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************************
' Class         : MortProdAddtFeaCS.cls
' Description   : This class implements base additionalfeatures class and executes methods
'                 on base class
'
' CHANGE HISTORY
' Prog      Date        Description
' MC        26/05/2004  [MC]BMIDS775/REG016/018 Regulation change Additional Features list
'*************************************************************************************************

Option Explicit

'Base Class Implementation
Implements MortProdAddtFea

Private m_clsMortProdAddtFea As MortProdAddtFea

Private Sub Class_Initialize()
    Set m_clsMortProdAddtFea = New MortProdAddtFea
End Sub


'*=[MC]The following are base call implementation methods

Private Sub MortProdAddtFea_Initialise(bIsEdit As Boolean)
    '*=[MC]Initialize AdditionalFeatures Object "MortprodAddtFea"
    m_clsMortProdAddtFea.Initialise bIsEdit
End Sub


Private Sub MortProdAddtFea_SaveScreenData()
    '*=[MC]Invoke SaveScreenData() Method
    m_clsMortProdAddtFea.SaveScreenData
End Sub

Private Sub MortProdAddtFea_SetProductKeys(colValues As Collection)
    '*=Set Product key columns to Mortgage Product AdditonalFeatures object
    m_clsMortProdAddtFea.SetProductKeys colValues
End Sub

Private Sub MortProdAddtFea_SetScreenFields()
    '*=Call SetScreenFields() to populate screen
    m_clsMortProdAddtFea.SetScreenFields
End Sub

Private Function MortProdAddtFea_ValidateScreenData() As Boolean
    '*=[MC]Validate screen data method implementation
    MortProdAddtFea_ValidateScreenData = m_clsMortProdAddtFea.ValidateScreenData
End Function


