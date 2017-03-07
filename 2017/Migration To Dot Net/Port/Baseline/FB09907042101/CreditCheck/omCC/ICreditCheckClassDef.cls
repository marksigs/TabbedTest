VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICreditCheckClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CreditCheckRetrieveEnum
    ccAll
    ccCurrent
    ccCAIS
End Enum
Public Function LoadApplicationCreditCheckData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40
End Function

Public Function LoadCreditCheckCustomerMiscData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40
End Function

Public Function LoadCreditCheckDetailsData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40
End Function

Public Function LoadCreditCheckMiscellaneousData(Optional ByVal intWhatToGet As CreditCheckRetrieveEnum) As FreeThreadedDOMDocument40
End Function

Public Function LoadDemoOptionsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauCAISData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauVotersRollData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauPublicInfoData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauCIFASData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauCAPSData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauAssociationsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauCorrectionLinesData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauDataHeaderData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFullBureauStandardHeaderData() As FreeThreadedDOMDocument40
End Function

Public Function LoadFBResultsData() As FreeThreadedDOMDocument40
End Function

