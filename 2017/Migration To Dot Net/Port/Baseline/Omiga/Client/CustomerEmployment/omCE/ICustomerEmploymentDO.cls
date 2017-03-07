VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerEmploymentDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub CreateAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                   ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub CreateContractDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                   ByRef rxmlKeysNode As IXMLDOMNode)
End Sub

Public Sub CreateNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub CreateTaxDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub DeleteOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function FindEarnedIncomeList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindEmploymentAndIncomeDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindEmploymentList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindOtherIncomeList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAccountantGUID(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetContractDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                   Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
End Function

Public Function GetEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                   Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
End Function

Public Function GetEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement, _
                                       Optional ByVal vblnRetrieveEmploymentFields As Boolean = True) As IXMLDOMNode
End Function

Public Function IsEmploymentStatus(ByVal venumEmploymentStatus As Integer, ByVal vstrEmploymentStatusId As String) As Boolean
End Function

Public Function IsMainSelfEmployed(ByVal vxmlElemRequest As IXMLDOMElement, _
                                Optional ByVal vxmlNode As IXMLDOMNode) As Boolean
End Function

Public Sub UpdateAccountantDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateContractDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateEarnedIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateEmploymentDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateNetProfitDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateOtherIncome(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Sub UpdateTaxDetails(ByVal vxmlTableElement As IXMLDOMElement)
End Sub

Public Function GetEmploymentAndSelfEmployedDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetEmploymentData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetTaxDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
End Function
'++ BMIDS00246 New function
Public Function GetEmploymentStatusList(ByVal vxmlElemRequest As IXMLDOMElement) As String
End Function

