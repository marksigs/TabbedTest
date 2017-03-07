VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "NewPropertyTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      NewPropertyTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' BG    26/05/00    Added Method SaveNewPropertyAndAddress
' SR    02/06/00    Added method 'SaveNewPropertyAndDeposit'
' JLD   06/06/00    Just send newproperty node to update newproperty
' SR    13/06/00    New method 'SaveNewProperty'
' MC    14/07/00    SYS1176 - Added SaveNewPropertyAddress method
'------------------------------------------------------------------------------------------
'BMids History:
' Prog  Date        Description
' GHun  23/09/2004  BMIDS887 Added SaveNewPropertyAndUpdateFactFind
'------------------------------------------------------------------------------------------

Option Explicit

Implements INewPropertyTxBO
Implements ObjectControl
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required

Private Const cstrROOT_NODE_NAME As String = "NEWPROPERTY"

Private Function INewPropertyTxBO_CreateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create multiple instances of the persistant data associated with this
'   business object. Invokes appropriate Data Object(s) to perform persistance.
' pass:
'   vxmlRequest  xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyTxBO_CreateNewPropertyAddress"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement

    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    strTableName = "NEWPROPERTYADDRESS"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.CreateNewPropertyAddress xmlElement
    Else
        ' For each relevant node call the data object CreateNewPropertyAddress
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.CreateNewPropertyAddress xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_CreateNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyAddressExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateNewPropertyAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_CreateNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume CreateNewPropertyAddressExit

End Function

Private Function INewPropertyTxBO_DeleteNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Delete multiple instances of the persistant data associated with this
'   business object. Invokes appropriate Data Object(s) to perform persistance
' pass:
'   vxmlRequest  xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyTxBO_DeleteNewPropertyAddress"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    Dim strTableName As String
    strTableName = "NEWPROPERTYADDRESS"
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.DeleteNewPropertyAddress xmlElement
    Else
        ' For each relevant node call the data object DeleteNewPropertyAddress
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.DeleteNewPropertyAddress xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_DeleteNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteNewPropertyAddressExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteNewPropertyAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_DeleteNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume DeleteNewPropertyAddressExit

End Function

Private Function INewPropertyTxBO_DeleteVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyTxBO_DeleteVendorDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    strTableName = "NEWPROPERTYVENDOR"

    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.DeleteVendorDetails xmlElement
    Else
        ' For each relevant node call the data object DeleteVendorDetails
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.DeleteVendorDetails xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_DeleteVendorDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeleteVendorDetailsExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteVendorDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_DeleteVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteVendorDetailsExit
End Function

Private Function INewPropertyTxBO_SaveNewProperty( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' Description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyTxBO_SaveNewProperty"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As SQLAssist
    Dim xmlTempNode As IXMLDOMNode
    
    Set objSqlAssist = New SQLAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                  oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim strApplicationNo As String
    Dim strApplicationFFNO As String
    Dim strKey As String
    Dim blnRecordExists As Boolean
    
    
    strApplicationNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONNUMBER")
    strApplicationFFNO = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
    strKey = "APPLICATIONNUMBER=" & objSqlAssist.FormatString(strApplicationNo)
    strKey = strKey & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNO
    blnRecordExists = objIADOAssist.CheckRecordExists("NEWPROPERTY", strKey)
    
    If blnRecordExists Then
        Set xmlTempNode = INewPropertyTxBO_UpdateNewProperty(vxmlRequest)
    Else
        Set xmlTempNode = INewPropertyTxBO_CreateNewProperty(vxmlRequest)
    End If
        
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
    
    Set INewPropertyTxBO_SaveNewProperty = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyExit:
        
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set xmlTempNode = Nothing
    
    Exit Function

SaveNewPropertyVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_SaveNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyExit

End Function

Private Function INewPropertyTxBO_SaveNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveNewPropertyAddress multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyTxBO_SaveNewPropertyAddress"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objIApplicationClassDef As IApplicationClassDef
    
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As SQLAssist
    Dim strApplicationNo As String
    Dim strApplicationFFNO As String
    Dim strKey As String
    Dim blnRecordExists As Boolean
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Initialise
    '
    If vxmlRequest.nodeName = "NEWPROPERTYADDRESS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("NEWPROPERTYADDRESS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "NEWPROPERTYADDRESS" & " tag not found"
        End If
    End If
    
    Set objSqlAssist = New SQLAssist
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
  
    '
    ' Main
    '
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadNewPropertyAddressData

    If Not (xmlElement Is Nothing) Then
        ' Call the appropriate DO method for this element
        strApplicationNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONNUMBER")
        strApplicationFFNO = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
        strKey = "APPLICATIONNUMBER=" & objSqlAssist.FormatString(strApplicationNo)
        strKey = strKey & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNO
        blnRecordExists = objIADOAssist.CheckRecordExists("NEWPROPERTYADDRESS", strKey)
        If blnRecordExists Then
            objINewPropertyDO.UpdateNewPropertyAddress xmlElement
        Else
            objINewPropertyDO.CreateNewPropertyAddress xmlElement
        End If
    Else
        ' For each relevant node call the appropriate data object method
        For Each xmlElement In xmlNodeList
            strApplicationNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONNUMBER")
            strApplicationFFNO = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
            strKey = "APPLICATIONNUMBER=" & objSqlAssist.FormatString(strApplicationNo)
            strKey = strKey & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNO
            blnRecordExists = objIADOAssist.CheckRecordExists("NEWPROPERTYADDRESS", strKey)
            If blnRecordExists Then
                objINewPropertyDO.UpdateNewPropertyAddress xmlElement
            Else
                objINewPropertyDO.CreateNewPropertyAddress xmlElement
            End If
        Next
    End If
    
    Set INewPropertyTxBO_SaveNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAddressExit:
    
    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    
    Exit Function

SaveNewPropertyAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_SaveNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAddressExit

End Function

Private Function INewPropertyTxBO_SaveNewPropertyAndAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndAddressVbErr

    Const strFunctionName As String = "INewPropertyTxBO_SaveNewPropertyAndAddress"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As SQLAssist
    Dim xmlTempNode As IXMLDOMNode
    
    Set objSqlAssist = New SQLAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim strApplicationNo As String
    Dim strApplicationFFNO As String
    Dim strKey As String
    Dim blnRecordExists As Boolean
    
    
    strApplicationNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONNUMBER")
    strApplicationFFNO = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
    strKey = "APPLICATIONNUMBER=" & objSqlAssist.FormatString(strApplicationNo)
    strKey = strKey & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNO
    blnRecordExists = objIADOAssist.CheckRecordExists("NEWPROPERTY", strKey)
    
    If blnRecordExists Then
        Set xmlTempNode = INewPropertyTxBO_UpdateNewPropertyDetails(vxmlRequest)
    Else
        Set xmlTempNode = INewPropertyTxBO_CreateNewPropertyDetails(vxmlRequest)
    End If
        
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
    
    blnRecordExists = objIADOAssist.CheckRecordExists("NEWPROPERTYADDRESS", strKey)
    
    If blnRecordExists Then
        Set xmlTempNode = INewPropertyTxBO_UpdateNewPropertyAddress(vxmlRequest)
    Else
        Set xmlTempNode = INewPropertyTxBO_CreateNewPropertyAddress(vxmlRequest)
    End If
    
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
        
    Set INewPropertyTxBO_SaveNewPropertyAndAddress = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAndAddressExit:
        
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set xmlTempNode = Nothing
    
    Exit Function

SaveNewPropertyAndAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_SaveNewPropertyAndAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndAddressExit

End Function

Private Function INewPropertyTxBO_SaveNewPropertyAndDeposit( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndDepositVbErr

    Const strFunctionName As String = "INewPropertyTxBO_SaveNewPropertyAndDeposit"

    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40, xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlPropertyElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objSqlAssist As SQLAssist
    Dim objIClassDef As IApplicationClassDef
    
    Set objSqlAssist = New SQLAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                  cstrROOT_NODE_NAME & " tag not found"
    End If
   
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    '--------------------------------------------------------------------------------------
    ' Save 'NEWPROPERTY'. Check whether a corresponding record exists in 'NEWPROPERTY',
    ' if so update it else create one
    '--------------------------------------------------------------------------------------
    
    Dim strApplicationNo As String, strApplicationFFNO As String, strKey As String
    Set xmlPropertyElem = vxmlRequest.getElementsByTagName("NEWPROPERTY").Item(0)
    
    strApplicationNo = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONNUMBER")
    strApplicationFFNO = m_objXmlAssist.GetNodeValue(xmlElement, ".//APPLICATIONFACTFINDNUMBER")
    strKey = "APPLICATIONNUMBER=" & objSqlAssist.FormatString(strApplicationNo)
    strKey = strKey & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNO
    
    If objIADOAssist.CheckRecordExists("NEWPROPERTY", strKey) Then
         'objINewPropertyDO.UpdateNewProperty vxmlRequest
         objINewPropertyDO.UpdateNewProperty xmlPropertyElem
    Else
        'objINewPropertyDO.CreateNewProperty vxmlRequest
        objINewPropertyDO.CreateNewProperty xmlPropertyElem
    End If
              
    '--------------------------------------------------------------------------------------
    ' Check the Operation to be performed on the table 'NewPropertyDeposit' and call the
    ' the respective method.
    '--------------------------------------------------------------------------------------
    Dim botOperationType As BOOPERATIONTYPE
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData
    
    botOperationType = m_objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
    
    Select Case botOperationType
        Case booCreate
            objIDOAssist.GenerateSequenceNumber xmlElement, xmlClassDefDoc, "NPDEPOSITSEQUENCENUMBER"
            objINewPropertyDO.CreateNewPropertyDeposit xmlElement
        Case booUpdate
            objINewPropertyDO.UpdateNewPropertyDeposit xmlElement
        Case Else
            ' N/A
    End Select
        
    Set INewPropertyTxBO_SaveNewPropertyAndDeposit = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAndDepositExit:
        
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlResponseElem = Nothing
    Set objINewPropertyDO = Nothing
    Set objIClassDef = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objSqlAssist = Nothing

    Exit Function

SaveNewPropertyAndDepositVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_SaveNewPropertyAndDeposit = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndDepositExit
End Function

Private Function INewPropertyTxBO_UpdateVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyTxBO_UpdateVendorDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    
    strTableName = "NEWPROPERTYVENDOR"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.UpdateVendorDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateVendorDetails
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.UpdateVendorDetails xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_UpdateVendorDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateVendorDetailsExit:
    
    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVendorDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_UpdateVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateVendorDetailsExit
End Function

Private Function INewPropertyTxBO_CreateVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyTxBO_CreateVendorDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    Dim strTableName As String
    strTableName = "NEWPROPERTYVENDOR"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.CreateVendorDetails xmlElement
    Else
        ' For each relevant node call the data object CreateVendorDetails
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.CreateVendorDetails xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_CreateVendorDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateVendorDetailsExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVendorDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_CreateVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateVendorDetailsExit
End Function


Private Function INewPropertyTxBO_CreateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyTxBO_CreateNewPropertyDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    Dim strTableName As String
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.CreateNewPropertyDetails xmlElement
    Else
        ' For each relevant node call the data object CreateNewPropertyDetails
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.CreateNewPropertyDetails xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_CreateNewPropertyDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyDetailsExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateNewPropertyDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_CreateNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyDetailsExit
End Function


Private Function INewPropertyTxBO_CreateNewProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyTxBO_CreateNewProperty"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    Dim strTableName As String
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.CreateNewProperty xmlElement
    Else
        ' For each relevant node call the data object CreateNewProperty
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.CreateNewProperty xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_CreateNewProperty = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyExit:

    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateNewPropertyVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_CreateNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyExit
End Function

Private Function INewPropertyTxBO_UpdateNewProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyTxBO_UpdateNewProperty"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.UpdateNewProperty xmlElement
    Else
        ' For each relevant node call the data object UpdateNewProperty
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.UpdateNewProperty xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_UpdateNewProperty = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyExit:
    
    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateNewPropertyVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_UpdateNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyExit

End Function

Private Function INewPropertyTxBO_UpdateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyTxBO_UpdateNewPropertyDetails"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.UpdateNewPropertyDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateNewPropertyDetails
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.UpdateNewPropertyDetails xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_UpdateNewPropertyDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyDetailsExit:
    
    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateNewPropertyDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_UpdateNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyDetailsExit

End Function

Private Function INewPropertyTxBO_UpdateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyTxBO_UpdateNewPropertyAddress"

    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    
    Dim strTableName As String
    strTableName = "NEWPROPERTYADDRESS"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objINewPropertyDO.UpdateNewPropertyAddress xmlElement
    Else
        ' For each relevant node call the data object UpdateNewPropertyAddress
        For Each xmlElement In xmlNodeList
            objINewPropertyDO.UpdateNewPropertyAddress xmlElement
        Next
    End If
    
    Set INewPropertyTxBO_UpdateNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyAddressExit:
    
    Set objINewPropertyDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateNewPropertyAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyTxBO_UpdateNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyAddressExit

End Function


Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing

End Sub

'BMIDS887 GHun
Private Function INewPropertyTxBO_SaveNewPropertyAndUpdateFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
  
    Const cstrFunctionName As String = "SaveNewPropertyAndUpdateFactFind"
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlAFF As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIApplicatinDO As IApplicationDO
    
On Error GoTo SaveNewPropertyAndUpdateFactFindVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlTempNode = xmlGetMandatoryNode(vxmlRequest, "NEWPROPERTY")
    If xmlGetAttributeAsBoolean(xmlTempNode, "NEWPROPERTYEXISTS") Then
        Set xmlTempNode = INewPropertyTxBO_UpdateNewProperty(vxmlRequest)
    Else
        Set xmlTempNode = INewPropertyTxBO_CreateNewProperty(vxmlRequest)
    End If
    
    m_objErrAssist.CheckResponse xmlTempNode, xmlResponseElem, True
    
    Set xmlAFF = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFIND")
    If Not xmlAFF Is Nothing Then
        If m_objContext Is Nothing Then
            Set objIApplicatinDO = New ApplicationDO
        Else
            Set objIApplicatinDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
        End If
    
        objIApplicatinDO.UpdateApplicationFactFind xmlAFF
    End If
    
    Set INewPropertyTxBO_SaveNewPropertyAndUpdateFactFind = xmlResponseElem
    
SaveNewPropertyAndUpdateFactFindExit:
    Set xmlTempNode = Nothing
    Set xmlAFF = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicatinDO = Nothing
    
    Exit Function

SaveNewPropertyAndUpdateFactFindVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    Set INewPropertyTxBO_SaveNewPropertyAndUpdateFactFind = m_objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' go to clean-up section
    Resume SaveNewPropertyAndUpdateFactFindExit
    
End Function
'BMIDS887 End

