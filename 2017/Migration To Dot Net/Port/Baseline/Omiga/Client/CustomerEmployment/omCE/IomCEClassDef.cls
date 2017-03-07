VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IomCEClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum omCEDataDefType
    omCEDefault
    omCEEmployment
End Enum

Public Function LoadAccountantData() As FreeThreadedDOMDocument40
End Function

Public Function LoadContractDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadContractDetailsProfitData() As FreeThreadedDOMDocument40
End Function

Public Function LoadDataForEmployment() As FreeThreadedDOMDocument40
End Function

Public Function LoadEarnedIncomeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadEmployedDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadEmployedDetailsIncomeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadEmploymentAccountantGUID() As FreeThreadedDOMDocument40
End Function

Public Function LoadEmploymentData(Optional ByVal intCEDataDefType As _
                            omCEDataDefType = omCEDefault) As FreeThreadedDOMDocument40
End Function

Public Function LoadEmploymentDirectoryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadEmploymentThirdPartyData() As FreeThreadedDOMDocument40
End Function

Public Function LoadIncomeSummaryData() As FreeThreadedDOMDocument40
End Function

Public Function LoadNetProfitData() As FreeThreadedDOMDocument40
End Function

Public Function LoadSelfEmployedDetailsData() As FreeThreadedDOMDocument40
End Function

Public Function LoadSelfEmployedDetailsProfitData() As FreeThreadedDOMDocument40
End Function

Public Function LoadUnearnedIncomeData() As FreeThreadedDOMDocument40
End Function

Public Function LoadIncomeSummaryDetailsData() As FreeThreadedDOMDocument40
End Function


