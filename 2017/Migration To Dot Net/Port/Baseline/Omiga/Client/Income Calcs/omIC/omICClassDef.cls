VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omICClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      omICClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   XML definitions for Class/Tables for omIC
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'-------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
'
' MDC   03/07/02    BMIDS00654 Maximum Borrowing Calculations
'-------------------------------------------------------------------------------------------

Implements IomICClassDef

Private Function IomICClassDef_LoadIncomeMultipleSetData() As FreeThreadedDOMDocument40

Const strFunctionName As String = "IomICClassDef_LoadIncomeMultipleSetData"

Dim strXML As String
Dim objXmlAssist As New XMLAssist
    
    strXML = "<TABLENAME>INCOMEMULTIPLESET"
    strXML = strXML & "<PRIMARYKEY>INCOMEMULTIPLIERCODE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INCOMEMULTIPLIERDESCRIPTION<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SINGLEINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>JOINTINCOMEMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>HIGHESTEARNERMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>LOWESTEARNERMULTIPLE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IomICClassDef_LoadIncomeMultipleSetData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function
