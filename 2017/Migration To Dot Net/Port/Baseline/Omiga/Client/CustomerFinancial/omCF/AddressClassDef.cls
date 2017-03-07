VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddressClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AddressClassDef.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   ClassDef that maps XML to Database fields
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DJP    05/10/01 SYS2564/SYS2777 (child)

Option Explicit

Implements IAddressClassDef


'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    27/08/2002  BMIDS00336 Credit Check/Full Bureau Download
'------------------------------------------------------------------------------------------

Private Function IAddressClassDef_LoadAddressData() As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IAddressClassDef_LoadAddressData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "ADDRESS"
    strXML = strXML & "<PRIMARYKEY>ADDRESSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BUILDINGORHOUSENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>FLATNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>STREET<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISTRICT<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOWN<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTY<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Country</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>POSTCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DELIVERYPOINTSUFFIX<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAILSORTCODE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAFINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'DJP SYS2564/SYS2777 (child) - make call to client defined method to do any client specific
    ' field changes
    AmendClassDef strFunctionName, strXML
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set IAddressClassDef_LoadAddressData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing
End Function

Private Sub AmendClassDef(ByVal strFunctionName As String, ByRef strXML As String)
' header ----------------------------------------------------------------------------------
' description:
'   AmendClassDef. Called by every ClassDef in this class that needs to do for client
'   variantions on the Core ClassDef. Put a switch statement in (the client derivation) and
'   each method that needs to be changed can do its own variation. Does nothing in Core.
' pass:
' return:                n/a
'------------------------------------------------------------------------------------------

    'BMIDS00336 MDC 27/08/2002
    If strFunctionName = "IAddressClassDef_LoadAddressData" Then
        strXML = strXML & "<OTHERS>BFPO<TYPE>dbdtBoolean</TYPE></OTHERS>"
    End If
    'BMIDS00336 MDC 27/08/2002 - End
    
End Sub

