VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICustomerFinancialTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'-------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'PSC    17/07/02    BMIDS00006 Added CreateAccountRelationships
'                                    DeleteAccountRelationships
'                                    UpdateAccountRelationships
'                                    CreateSpecialFeature
'                                    UpdateSpecialFeature
'                                    DeleteSpecialFeature
'MV     14/08/2002  BMIDS00311  Core Ref: AWR: SYS1705 Modified SaveMortgageAccountAndLoan
'-------------------------------------------------------------------------


Public Function CreateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute CreateRegularOutgoings.VB_Description = "Business Process to return relevant data from the ""regularOutgoings"" table (within a transaction)"
End Function

Public Function CreateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlKeysNode As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateArrearsHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateBankCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateBankruptcyHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateCCJHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateDeclinedMortgage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateGroupConnection(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLiability(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateLifeProduct(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateMortgageAccount(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateRegularOutgoings(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
Attribute UpdateRegularOutgoings.VB_Description = "Updates a record within the ""RegularOutgoings"" table (Within a transaction)"
End Function

Public Function UpdateMortgageLoan(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function SaveMortgageAccountAndLoan(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function SaveMortgageAccount(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateAccountRelationships(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function CreateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function UpdateSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function

Public Function DeleteSpecialFeature(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
End Function
