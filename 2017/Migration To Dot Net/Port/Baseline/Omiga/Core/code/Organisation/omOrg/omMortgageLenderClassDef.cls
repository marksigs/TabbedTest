VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omMortgageLenderClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IomMortgageLenderClassDef
Private Function IomMortgageLenderClassDef_LoadLenderNamesData() As FreeThreadedDOMDocument40
     
    Dim strFunctionName As String
    Dim objXmlAssist As New XMLAssist
    Dim strXML As String
    strFunctionName = "IomMortgageLenderClassDef_LoadLenderNamesData"
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGELENDER"
    strXML = strXML & "<OTHERS>LENDERCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LENDERNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ORGANISATIONID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME> "
    Set IomMortgageLenderClassDef_LoadLenderNamesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
