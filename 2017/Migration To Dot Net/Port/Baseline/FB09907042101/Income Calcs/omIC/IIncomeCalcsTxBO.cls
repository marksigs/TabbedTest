VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IIncomeCalcsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RunIncomeCalculation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CalculateAllowableIncome(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function CalculateMaximumBorrowing(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function

Public Function RunToleranceCheck(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

End Function