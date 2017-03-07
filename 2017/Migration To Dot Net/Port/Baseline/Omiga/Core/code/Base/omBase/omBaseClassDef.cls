VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omBaseClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IomBaseClassDef
Private Function IomBaseClassDef_LoadMessageData() As FreeThreadedDOMDocument40
    
    Dim strFunctionName As String
    strFunctionName = "IomBaseClassDef_LoadMessageData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MESSAGE"
    strXML = strXML & "<PRIMARYKEY>MESSAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>MESSAGETEXT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MESSAGETYPE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IomBaseClassDef_LoadMessageData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
Private Function IomBaseClassDef_LoadCurrencyData() As FreeThreadedDOMDocument40
    
    Dim strFunctionName As String
    strFunctionName = "IomBaseClassDef_LoadCurrencyData"
    Dim strXML As String
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CURRENCY"
    strXML = strXML & "<PRIMARYKEY>CURRENCYCODE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BASECURRENCYIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CURRENCYNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONVERSIONRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ROUNDINGPRECISION<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ROUNDINGDIRECTION<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ROUNDINGDIRECTION</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Set IomBaseClassDef_LoadCurrencyData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlAssist = Nothing
End Function
