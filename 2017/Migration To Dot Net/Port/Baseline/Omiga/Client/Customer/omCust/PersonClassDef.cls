VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PersonClassdef"
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
'------------------------------------------------------------------------------------------

Option Explicit

Implements IPersonClassDef

Private Function IPersonClassDef_LoadPersonData() As FreeThreadedDOMDocument40

'Header------------------------------------------------------------------------
'Description : Creates the class template required for getting the XML from
'              recordset
'Pass        : None
'Return      : XML string
' *****************************************************************************
 Const strFunctionName As String = "IPersonClassDef_LoadPersonData"
 
    Dim strXML As String
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "PERSON"
    strXML = strXML & "<PRIMARYKEY>PERSONGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DATEOFBIRTH<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FIRSTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>GENDER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Sex</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>MARITALSTATUS<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>MaritalStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>NATIONALITY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Nationality</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>OTHERFORENAMES<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SECONDFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>TITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IPersonClassDef_LoadPersonData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
    
End Function

