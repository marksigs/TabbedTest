VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LenderPaymentDetailsCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LenderPaymentDetailsCS
' Description   : Client specific version of the LenderPaymentDetails class.
'                 This class will be called by a form or controlling class and
'                 any client specific code can be put in here for each client.
'                 Each method can a) override the base code, b) delegate to the
'                 base code or c) overried and delegate.
'
' Change history
' Prog  Date        Description
' STB   08-May-2002 MSMS0069 Created class.
' CL    28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Implements LenderPaymentDetails

'The underlying base class.
Private m_clsLenderPayment As LenderPaymentDetails


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Create underlying base class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_clsLenderPayment = New LenderPaymentDetails
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_DoUpdates
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_DoUpdates()
    m_clsLenderPayment.DoUpdates
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_Initialise
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_Initialise(ByVal bIsEdit As Boolean)
    m_clsLenderPayment.Initialise bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_InitTab
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_InitTab()
    m_clsLenderPayment.InitTab
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_IsEdit
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LenderPaymentDetails_IsEdit() As Boolean
    LenderPaymentDetails_IsEdit = m_clsLenderPayment.IsEdit
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_SaveScreenData
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_SaveScreenData()
    m_clsLenderPayment.SaveScreenData
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_SetScreenFields
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_SetScreenFields()
    m_clsLenderPayment.SetScreenFields
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_SetTableClass
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LenderPaymentDetails_SetTableClass(clsTableAccess As TableAccess)
    m_clsLenderPayment.SetTableClass clsTableAccess
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LenderPaymentDetails_ValidateScreenData
' Description   : Delegate method to base case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LenderPaymentDetails_ValidateScreenData() As Boolean
    LenderPaymentDetails_ValidateScreenData = m_clsLenderPayment.ValidateScreenData
End Function
