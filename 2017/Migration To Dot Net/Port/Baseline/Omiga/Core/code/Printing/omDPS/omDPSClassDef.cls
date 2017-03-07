VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omDPSClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omDPSClassDef.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   DPS Class Definition
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     12/05/00    Created
'------------------------------------------------------------------------------------------
'Baseline History
'
'Prog   Date        Description
'HMA    29/04/2004  CORE46  Add DELIVERYENGINETYPE, ACEXML and ACEXSL to TEMPLATE table
'------------------------------------------------------------------------------------------
Option Explicit
Implements IomDPSClassDef

Private Function IomDPSClassDef_LoadTemplateData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IClassDef_LoadTemplateData"
    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    strXML = "<TABLENAME>TEMPLATE"
    strXML = strXML & "<PRIMARYKEY>TEMPLATEID<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>TEMPLATENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FILENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECURITYLEVEL<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STAGENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DEFAULTPRINTER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LANGUAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DELIVERYENGINETYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACEXML<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACEXSL<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Set IomDPSClassDef_LoadTemplateData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
