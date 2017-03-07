VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omBCClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omBCClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   XML definitions for Class/Tables for omOrg
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObjext
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' MCS    14/12/99 Created for performance rework.
'------------------------------------------------------------------------------------------

Option Explicit

Implements IomBCClassDef

Private Function IomBCClassDef_LoadBCProductData() As FreeThreadedDOMDocument40
' header ----------------------------------------------------------------------------------
' description:
'   XML class definition for the OMIGAUSER table
' pass:
' return:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IomBCClassDef_LoadBCProductData"
    
    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    strXML = strXML & "<TABLENAME>BUILDINGANDCONTENTSPRODUCT"
    strXML = strXML & "<PRIMARYKEY>BCPRODUCTNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>VALUABLESLIMIT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BCSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INCLUDESBUILDINGSCOVER<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INCLUDESCONTENTSCOVER<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>VALUABLESITEMLIMIT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set IomBCClassDef_LoadBCProductData = xmlClassDefDoc
    
    Set objXmlAssist = Nothing
    

End Function

