VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDOAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      IDOAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:
'
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    25/01/00    Added GetNextSequenceNumberEx.
'------------------------------------------------------------------------------------------
Option Explicit
Public Enum SQL_FORMAT_TYPE
    ' fields and values in 2 separate strings,
    ' e.g. "FieldA,FieldB" and "Value1,Value2"
    sftFieldValueSeparated
    ' field=value pairs separated by commas,
    ' e.g. "FieldA = Value1, FieldB = Value2"
    sftCommaSeparated
    ' field=value pairs separated by " AND ",
    ' e.g. "FieldA = Value1 AND FieldB = Value2"
    sftAndSeparated
End Enum
Public Enum CLASS_DEF_KEY
    cdkPRIMARYKEY
    cdkOTHERS
End Enum
Public Enum CLASS_DEF_KEY_AMOUNT
    cdkaALLKEYS
    cdkaANYKEYS
End Enum
Private Const cstrNotSpecified = "Not specified"
Public Sub BuildSQLString( _
    ByVal vxmlRequest As IXMLDOMElement, _
    ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
    ByRef rstrTable As String, _
    Optional ByVal vsftSQLFormatType As SQL_FORMAT_TYPE = sftAndSeparated, _
    Optional ByVal vcdkKey As CLASS_DEF_KEY = cdkPRIMARYKEY, _
    Optional ByVal vcdkKeyValue As CLASS_DEF_KEY_AMOUNT = cdkaALLKEYS, _
    Optional ByRef rstrFieldValuePair As String = cstrNotSpecified, _
    Optional ByRef rstrFields As String = cstrNotSpecified, _
    Optional ByRef rstrValues As String = cstrNotSpecified, _
    Optional ByRef rstrSelect As String, _
    Optional ByVal vstrItemName As String = "")
End Sub
    
Public Sub Create(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
End Sub
Public Sub DeleteAll(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
End Sub
Public Sub Delete(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
End Sub
Public Function FindList(ByVal vxmlRequest As IXMLDOMElement, _
                         ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                         Optional ByVal vstrOrderByField As String = cstrNotSpecified) As IXMLDOMNode
End Function
Public Function FindListMultiple(ByVal vxmlRequest As IXMLDOMElement, _
                                 ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                 Optional ByVal vstrOrderByField As Variant, _
                                 Optional ByVal vstrItemName As String = "", _
                                 Optional ByVal vstrAdditionalCodition As String = "", _
                                 Optional ByVal vstrAndOrToCodition As String = "AND") As IXMLDOMNode
                                   
End Function
Public Function GenerateSequenceNumber( _
                    ByVal rxmlElement As IXMLDOMElement, _
                    ByVal vxmlClassDefinition As FreeThreadedDOMDocument40, _
                    ByVal vstrSequenceField As String) As Long
End Function
Public Function GetData(ByVal vxmlRequest As IXMLDOMElement, _
                        ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                        Optional ByVal vstrTableName As String = "", _
                        Optional ByVal vstrItemName As String = "") As IXMLDOMNode
End Function
Public Function GetXMLFromRecordSet(ByVal vrstRecordSet As ADODB.Recordset, _
                                    ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                    Optional ByVal vxmlInNode As IXMLDOMNode = Nothing) As IXMLDOMNode
End Function
Public Sub Update(ByVal vxmlRequest As IXMLDOMElement, ByVal vxmlClassDef As FreeThreadedDOMDocument40)
End Sub
Public Function GetNextSequenceNumber( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    ByVal vxmlClassDefinition As FreeThreadedDOMDocument40, _
                    ByVal vstrTableName As String, _
                    ByVal vstrSequenceField As String) As Long
End Function
Public Function GetComponentData(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal xmlClassDef As FreeThreadedDOMDocument40) As IXMLDOMNode
End Function
Public Sub GetXMLFromWholeRecordset(ByVal vrstRecordSet As ADODB.Recordset, _
                                    ByVal vxmlClassDef As FreeThreadedDOMDocument40, _
                                    ByVal vxmlInNode As IXMLDOMNode)
End Sub
