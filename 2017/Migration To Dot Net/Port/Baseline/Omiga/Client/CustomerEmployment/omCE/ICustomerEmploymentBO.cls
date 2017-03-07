VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerEmploymentBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function CreateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindEarnedIncomeList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindEmploymentList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindEmploymentAndIncomeDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function FindOtherIncomeList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetAccountantGUID(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetCurrentEmploymentAndIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveAccountantDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                      Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveContractDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                    Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveEarnedIncome(ByVal vxmlRequest As IXMLDOMElement, _
                                 Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveEmployedDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                    Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                      Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveCurrentEmploymentAndIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                     Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveOtherIncome(ByVal vxmlRequest As IXMLDOMElement, _
                                Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement, _
                                        Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function SaveTaxDetails(ByVal vxmlRequest As IXMLDOMElement, _
                               Optional ByVal vblnDoValidation As Boolean = False) As IXMLDOMNode
End Function

Public Function UpdateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateAccountantDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateContractDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateEarnedIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateGeneralEmployment(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateNetProfitDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateOtherIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function ValidateTaxDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function GetEmploymentAndSelfEmployedDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CopyEmploymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
'New method to call GetTaxDetails
Public Function GetIncomeSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

