VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExperianClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for omiga4 Class Definition
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/12/99    Created
'PSC    28/02/00    Remove Dim As New
'------------------------------------------------------------------------------------------

Option Explicit

Implements IExperianClassDef

Private Function IExperianClassDef_LoadDemoOptionsData() As FreeThreadedDOMDocument40


    Const strFunctionName As String = "IExperianClassDef_LoadDemoOptionsData"

    Dim strXML As String
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CREDITCHECKDEMO"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>SEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DEMOMODEIND<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IExperianClassDef_LoadDemoOptionsData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

