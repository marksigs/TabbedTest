VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContactClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IContactClassDef

Private Function IContactClassDef_LoadContactData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IContactClassDef_LoadContactData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CONTACTDETAILS"

    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    'AW 31/07/31    EP1059
    strXML = strXML & "<OTHERS>CONTACTTITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IContactClassDef_LoadContactData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

'JR added - Omiplus24
Private Function IContactClassDef_LoadContactDetailsData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IContactClassDef_LoadContactDetailsData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CONTACTDETAILS"

    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"

    strXML = strXML & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
    'AW 31/07/31    EP1059
    strXML = strXML & "<OTHERS>CONTACTTITLEOTHER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & IContactClassDef_LoadContactTelephoneData.xml

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IContactClassDef_LoadContactDetailsData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

'JR added - Omiplus24
Private Function IContactClassDef_LoadContactTelephoneData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IContactClassDef_LoadContactTelephoneData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>CONTACTTELEPHONEDETAILS"

    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TELEPHONESEQNUM<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactTelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TELENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IContactClassDef_LoadContactTelephoneData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function
