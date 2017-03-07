VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omAUClassDef"
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
Implements IomAUClassDef
Private Function IomAUClassDef_LoadAccessAuditData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "LoadAccessAuditData"
    Const strXml As String = "<TABLENAME>ACCESSAUDIT" & _
            "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
            "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" & _
            "<OTHERS>ACCESSDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" & _
            "<OTHERS>ATTEMPTNUMBER<TYPE>dbdtInt</TYPE></OTHERS>" & _
            "<OTHERS>AUDITRECORDTYPE<TYPE>dbdtComboId</TYPE>" & _
            "<COMBO>AccessAuditType</COMBO></OTHERS>" & _
            "<OTHERS>MACHINEID<TYPE>dbdtString</TYPE></OTHERS>" & _
            "<OTHERS>SUCCESSINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
            "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IomAUClassDef_LoadAccessAuditData = objXmlAssist.load(strXml, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IomAUClassDef_LoadApplicationAccessData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "LoadApplicationAccessData"
    Const strXml As String = "<TABLENAME>APPLICATIONACCESS" & _
            "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
            "<OTHERS>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>" & _
            "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IomAUClassDef_LoadApplicationAccessData = objXmlAssist.load(strXml, TypeName(Me), strFunctionName)
     
    Set objXmlAssist = Nothing
     
End Function
Private Function IomAUClassDef_LoadChangePasswordData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "LoadChangePasswordData"
    Const strXml As String = "<TABLENAME>CHANGEPASSWORD" & _
            "<PRIMARYKEY>ACCESSAUDITGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>" & _
            "<OTHERS>ONBEHALFOFUSERID<TYPE>dbdtString</TYPE></OTHERS>" & _
            "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IomAUClassDef_LoadChangePasswordData = objXmlAssist.load(strXml, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
