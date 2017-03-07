VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omSubClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omSubClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Class Definition
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     29/03/00    Created.
'------------------------------------------------------------------------------------------

Option Explicit

Implements IomSubClassDef

Private Function IomSubClassDef_LoadApplicationSubmissionStatusData() _
    As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomSubClassDef_LoadApplicationSubmissionStatusData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "APPLICATIONSUBMISSIONSTATUS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>SUBMISSIONSTATUS"
    strXML = strXML & "<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>SubmissionStatus</COMBO>"
    strXML = strXML & "</OTHERS>"
    strXML = strXML & "<OTHERS>SUBMISSIONDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomSubClassDef_LoadApplicationSubmissionStatusData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


