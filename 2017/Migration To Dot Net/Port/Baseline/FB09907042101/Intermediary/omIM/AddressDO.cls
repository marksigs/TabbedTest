VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AddressDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      AddressDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:
'               All other that use address (Ex: CustomerAddress)
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'SR    02/11/99  Created
'LD    27/06/01  SYS2386 All projects to use guidassist.bas rather than guidassist.cls
'------------------------------------------------------------------------------------------

Option Explicit

Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request (ADDRESS)
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
''------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"
    Dim strTagName As String
    'stores the XML string to passed to DOAssist (after making the modification, if any)
    Dim strModifiedXmlRequest As String
       
    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlAddressElem As IXMLDOMElement
    Dim xmlAddressGuidElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
       
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    'Append elements with table name and PrimaryKey to XMLout
    Set xmlTableElem = xmlOut.createElement("ADDRESS")
    xmlOut.appendChild xmlTableElem
    
    strTagName = "ADDRESSGUID"
    Set xmlAddressGuidElem = xmlOut.createElement(strTagName)
    xmlTableElem.appendChild xmlAddressGuidElem
    
    'Load the Input XML string and check an element exists with TABLE NAME
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlAddressElem = xmlIn.getElementsByTagName("ADDRESS").Item(0)
      
    If xmlAddressElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ADDRESS" & " tag not found"
    End If
    
    '*Check whether AddressGuid element exists. If not add one with value of
    '*newly generated Guid. If exists, check whether vlaue is mentioned, and
    '*assign generated Guid if value was not mentioned
    strTagName = "ADDRESSGUID"
    Set xmlElement = xmlAddressElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
         Set xmlElement = xmlIn.createElement(strTagName)
        xmlAddressElem.appendChild xmlElement
    End If

    If Len(xmlElement.Text) = 0 Then
        Dim strGUID As String
        strGUID = CreateGUID
        xmlElement.Text = strGUID
        'Also assign this to PrimaryKey tag in xmlOut
        xmlAddressGuidElem.Text = strGUID
    Else 'Assign the Guid passed to appropriate tag in XML out
        xmlAddressGuidElem.Text = xmlElement.Text
    End If
    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    strModifiedXmlRequest = xmlIn.xml
    objDOAssist.Create strModifiedXmlRequest, LoadAddressData
    
    Create = xmlOut.xml
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    Set xmlAddressGuidElem = Nothing
    Set xmlAddressElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
CreateVbErr:

    Set objDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    Set xmlAddressGuidElem = Nothing
    Set xmlAddressElem = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing

        
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub Delete(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "Delete"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                           
    Dim objDOAssist As DOAssist
    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    objDOAssist.Delete vstrXMLRequest, LoadAddressData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

DeleteVbErr:

    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    GetData = objDOAssist.GetData(vstrXMLRequest, LoadAddressData)

    ' TODO ************************************************************************
    '   Add Any derived data here - NOW NOTHING
    ' *****************************************************************************
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
        
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objDOAssist As DOAssist
    
    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    objDOAssist.Update vstrXMLRequest, LoadAddressData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

UpdateVbErr:
    
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vstrXMLData         base XML data stream
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'   AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "AddDerivedData"

    Dim objErrAssist As New ErrAssist

    AddDerivedData = vstrXMLData
    
    Set objErrAssist = Nothing
    
    Exit Function
            
    Set objErrAssist = Nothing
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function
    
AddDerivedDataVBErr:

    Set objErrAssist = Nothing
    
    AddDerivedData = ""

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function LoadAddressData()
    Dim strXML As String
        
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
    strXML = strXML & "</TABLENAME>"
    
    LoadAddressData = strXML
End Function

Public Function IsEmptyAddress(ByVal vstrXMLRequest As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:  Check if a customer address is empty
' pass:         vstrXMLRequest  address to check
'
' return:       IsEmptyAddress True if address is empty
'                              False if address has data
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsEmptyAddressVbErr
    
    Const strFunctionName As String = "IsEmptyAddress"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objElement As IXMLDOMElement
    
    Dim blnIsEmpty As Boolean
    blnIsEmpty = False
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set objElement = objXmlIn.getElementsByTagName("ADDRESS").Item(0)
    
    If objElement Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "ADDRESS" & "tag is missing"
    End If
    
    If Len(objXmlAssist.GetTagValue(objElement, "POSTCODE")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "FLATNUMBER")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "BUILDINGORHOUSENAME")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "BUILDINGORHOUSENUMBER")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "STREET")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "DISTRICT")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "TOWN")) = 0 And _
       Len(objXmlAssist.GetTagValue(objElement, "COUNTY")) = 0 Then
       'APS 06/09/99 - UNIT TEST REF 19 - Removed the check for the Country
       'as this maybe populated by default
       'And Len(objXMLAssist.GetTagValue(objElement, "COUNTRY")) = 0 Then
        blnIsEmpty = True
    End If
    
    IsEmptyAddress = blnIsEmpty
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objElement = Nothing
        
    Exit Function
    
IsEmptyAddressVbErr:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objElement = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function Save(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Determines the action to be performed on the associated table based on the input
'   and performs the respective action.
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Return :
'    xml String containing the Action Taken. If new record is created, the Primary Key data
'    is added to returned string
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo SaveVbErr

    Const strFunctionName As String = "Save"
        
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    Dim strAddressDef As String
    Dim strXmlOut As String
    Dim strTemp As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objDOAssist As DOAssist
    
    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    strAddressDef = LoadAddressData
    Select Case objXmlAssist.DetermineOperation(vstrXMLRequest, strAddressDef)
        Case booCreate
            strTemp = Create(vstrXMLRequest)
            strXmlOut = "<ACTION>" & "CREATE" & strTemp & "</ACTION>"
        Case booUpdate
            Call Update(vstrXMLRequest)
            strXmlOut = "<ACTION>UPDATE</ACTION>"
        Case booDelete
            Call Delete(vstrXMLRequest)
            strXmlOut = "<ACTION>DELETE</ACTION>"
        Case Else
            ' do nothig
    End Select
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Save = strXmlOut
    
    Exit Function

SaveVbErr:
    
    Set objDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

