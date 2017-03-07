VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omLCClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omLCClassDef.cls
'Copyright:     Copyright © 2000 Marlborough Stirling

'Description:   XML definitions for Class/Tables for omLC
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog   Date           Description
' MCS    18/07/2000     Created for performance rework.
'
'------------------------------------------------------------------------------------------


Option Explicit

Implements IomLCClassDef
Private Function IomLCClassDef_LoadLifeCoverRatesData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IomLCClassDef_LoadLifeCoverRatesData"

    Dim strXML As String

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "LIFECOVERRATES"
    strXML = strXML & "<PRIMARYKEY>LIFECOVERRATESNUMBER<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>LIFECOVERRATESSTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>COVERTYPE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICANTSGENDER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>HIGHAPPLICANTSAGE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>HIGHTERM<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>COVERTYPETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>APPLICANTSGENDERTEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ADDITIONALSMOKERRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ANNUALRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>NOTGOODHEALTHRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Dim objXmlAssist As New XMLAssist
    
    Set IomLCClassDef_LoadLifeCoverRatesData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function


