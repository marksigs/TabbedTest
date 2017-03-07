VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "PackageManagerTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PackageManagerTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Dependencies:  PackageManagerDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     27-01-00    Created.
'RF     25/02/00    Enhanced error messaging to help fix AQR SYS0321.
'RF     29/02/00    AQR SYS0378: Incorrect "REQUEST" blocks being used.
'MH     14/03/00    AQR SYS0321: GetNextPackageNumber may fail. Changed to not use a SP
'LD     07/11/00    Explicity close database connections
'DM             14/03/01  SYS1946 - Changed name of stored procedure being called.
'DM     06/06/01    SYS1946 Sequences.
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'DM     20/07/01    SYS2511 changed parameter type to adVarChar
'------------------------------------------------------------------------------------------

Option Explicit

Implements IPackageManagerTxBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================
 
Private Const cstrTABLE_CUSTOMERPACKAGE = "CUSTOMERPACKAGE"
Private Const cstrTABLE_APPLICATIONPACKAGE = "APPLICATIONPACKAGE"

Private Function IPackageManagerTxBO_AddCustomerToPackage( _
        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   This method links a customer to a package and to the application.
'   In all instances both the customer and the package records will already exist
'   on the database
' pass:
'   vrXMLRequest
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS" and xml representation of data
'   or: TYPE="SYSERR" and <ERROR> element
' Raise Errors:
'   Handles duplicate records on customerpackage by returning omiga4err133
'------------------------------------------------------------------------------------------
On Error GoTo AddCustomerToPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_AddCustomerToPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlInputElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    
    Dim objIPackageManagerBO As IPackageManagerBO
    Dim objApplManagerBO As ApplicationManagerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIPackageManagerBO = New PackageManagerBO
        Set objApplManagerBO = New ApplicationManagerBO
    Else
        Set objIPackageManagerBO = objContext.CreateInstance(App.Title & ".PackageManagerBO")
        Set objApplManagerBO = objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
    End If

    ' create 1 to n customer packages
    
    Dim lngErrorNumber As Long
    
    Set xmlTempResponseElem = objIPackageManagerBO.CreateCustomerPackage(vxmlRequest)
    lngErrorNumber = objErrAssist.CheckResponse(xmlTempResponseElem, xmlResponseElem, False)
    
    If lngErrorNumber <> 0 Then
        Dim lngOmigaErrNo As Long
        lngOmigaErrNo = objErrAssist.GetOmigaErrorNumber(lngErrorNumber)
        If lngOmigaErrNo = oeDuplicateKey Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 133
        Else
            objErrAssist.ThrowError TypeName(Me), strFunctionName, lngOmigaErrNo
        End If
    End If
    
    ' link customer to the existing application -
    ' NB. DOM Interface methods are not availble on ApplicationManagerBO
    
    'RF 29/02/00 AQR SYS0378: Start - Incorrect "REQUEST" blocks being used.
    'Set xmlInputElem = xmlOut.createElement("REQUEST")
    Set xmlInputElem = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    'RF 29/02/00 AQR SYS0378: End
    
    objXmlAssist.CopyNode "CUSTOMERPACKAGE", "APPLICATION", vxmlRequest, xmlInputElem, True
    
    Dim strXmlResponse As String
    strXmlResponse = objApplManagerBO.AddCustomerToApplication(xmlInputElem.xml)
    
     ' if an error then re-raise it
    objErrAssist.CheckXMLResponse strXmlResponse, True
    
    Set IPackageManagerTxBO_AddCustomerToPackage = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If

AddCustomerToPackageExit:

    Set objIPackageManagerBO = Nothing
    Set objApplManagerBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlInputElem = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

AddCustomerToPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_AddCustomerToPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AddCustomerToPackageExit

End Function

 
Private Function IPackageManagerTxBO_CreateApplicationPackage( _
                    ByVal vxmlRequest As IXMLDOMElement, _
                    ByVal vxmlGeneratedKeys As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create multiple instances of the persistant data associated with this business object.
'   Invokes appropriate Data Object(s) to perform persistance.
' pass:
'   vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_CreateApplicationPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlPackageNumElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIPackageManagerDO As IPackageManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_APPLICATIONPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_APPLICATIONPACKAGE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, cstrTABLE_APPLICATIONPACKAGE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIPackageManagerDO = New PackageManagerDO
    Else
        Set objIPackageManagerDO = objContext.CreateInstance(App.Title & ".PackageManagerDO")
    End If
    
    Dim strNewPackageNumber As String, strTag As String, rblnExists As Boolean
    
    strTag = "PACKAGENUMBER"
    
    If Not xmlElement Is Nothing Then
    
        '----------------------------------------------------------------------------------
        ' Assign the new package number. If a node does not exist for this, create one
        '----------------------------------------------------------------------------------
        
        strNewPackageNumber = objXmlAssist.GetTagValue(xmlElement, strTag, rblnExists, False)
        If Not rblnExists Then
            'Create an element with this name and assign value
            Set xmlPackageNumElem = vxmlRequest.ownerDocument.createElement(strTag)
            xmlElement.appendChild xmlPackageNumElem
        Else
            Set xmlPackageNumElem = xmlElement.getElementsByTagName(strTag).Item(0)
        End If
        xmlPackageNumElem.Text = GetNextPackageNumber
        
        'Add the New package number to Generated Keys
        vxmlGeneratedKeys.appendChild xmlPackageNumElem.cloneNode(True)
               
        objIPackageManagerDO.CreateApplicationPackage xmlElement
    Else
        For Each xmlElement In xmlNodeList
        
            '----------------------------------------------------------------------------------
            ' Assign the new package number. If a node does not exist for this, create one
            '----------------------------------------------------------------------------------
            
            strNewPackageNumber = objXmlAssist.GetTagValue(xmlElement, strTag, rblnExists, False)
            If Not rblnExists Then
                'Create an element with this name and assign value
                Set xmlPackageNumElem = vxmlRequest.ownerDocument.createElement(strTag)
                xmlElement.appendChild xmlPackageNumElem
            Else
                Set xmlPackageNumElem = xmlElement.getElementsByTagName(strTag).Item(0)
            End If
            xmlPackageNumElem.Text = GetNextPackageNumber
            
            'Add the New package number to Generated Keys
            vxmlGeneratedKeys.appendChild xmlPackageNumElem.cloneNode(True)
            
            objIPackageManagerDO.CreateApplicationPackage xmlElement
        Next xmlElement
    End If
        
    Set IPackageManagerTxBO_CreateApplicationPackage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateApplicationPackageExit:

    Set objIPackageManagerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlPackageNumElem = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateApplicationPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_CreateApplicationPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationPackageExit

End Function

Private Function IPackageManagerTxBO_CreateCustomerPackage( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_CreateCustomerPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIPackageManagerDO As IPackageManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
         ' Get a list of nodes that need to be passed down to the data object
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERPACKAGE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_CUSTOMERPACKAGE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIPackageManagerDO = New PackageManagerDO
    Else
        Set objIPackageManagerDO = objContext.CreateInstance(App.Title & ".PackageManagerDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIPackageManagerDO.CreateCustomerPackage xmlElement
    Else
        ' For each relevant node call the data object Create
        For Each xmlElement In xmlNodeList
            objIPackageManagerDO.CreateCustomerPackage xmlElement
        Next
    End If
      
    Set IPackageManagerTxBO_CreateCustomerPackage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCustomerPackageExit:

    Set objIPackageManagerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateCustomerPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_CreateCustomerPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateCustomerPackageExit

End Function

Private Function IPackageManagerTxBO_CreatePackageAndApplication( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   create multiple instances of the persistant data associated with this business object
'   invokes appropriate Data Object(s) to perform persistance
' pass:
'   vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreatePackageAndApplicationVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_CreatePackageAndApplication"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlInputElem As IXMLDOMElement ' Input node to other methods being called
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlGeneratedKeys As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Dim objIPackageManagerBO As IPackageManagerBO
    Dim objIApplicationBO As IApplicationBO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERPACKAGE).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_CUSTOMERPACKAGE & " tag not found"
    End If
    
    If objContext Is Nothing Then
        Set objIPackageManagerBO = New PackageManagerBO
        Set objIApplicationBO = New ApplicationBO
    Else
        Set objIPackageManagerBO = objContext.CreateInstance(App.Title & ".PackageManagerBO")
        Set objIApplicationBO = objContext.CreateInstance(App.Title & ".ApplicationBO")
    End If
    
    '-------------------------------------------------------------------------------
    ' create a new application package
    '--------------------------------------------------------------------------------
    Set xmlInputElem = xmlOut.createElement("APPLICATIONPACKAGE")
    Set xmlElement = xmlOut.createElement("COMPANYID")
    xmlElement.Text = objXmlAssist.GetTagValue(vxmlRequest, "COMPANYID")

    Set xmlGeneratedKeys = xmlOut.createElement("PACKAGENUMBERLIST")
    Set xmlTempResponseNode = objIPackageManagerBO.ValidateApplicationPackage(xmlInputElem)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIPackageManagerBO.CreateApplicationPackage( _
                                                xmlInputElem, xmlGeneratedKeys)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                        
    ' Add the Package Key data to the request node
    AddPackageDetails vxmlRequest, xmlGeneratedKeys

    '---------------------------------------------------------------------------------
    ' create a new application and add the key data to Request
    '---------------------------------------------------------------------------------
    
    ' copy the customerpackage tag and call it application
    
    'RF 29/02/00 AQR SYS0378: Start - Incorrect "REQUEST" blocks being used.
    'Set xmlInputElem = xmlOut.createElement("REQUEST")
    Set xmlInputElem = objXmlAssist.GetRequestNodeEx(vxmlRequest)
    'RF 29/02/00 AQR SYS0378: End
    
    objXmlAssist.CopyNode "CUSTOMERPACKAGE", "APPLICATION", vxmlRequest, xmlInputElem, True
    Set xmlTempResponseNode = objIApplicationBO.ValidateApplication(xmlInputElem)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIApplicationBO.CreateApplication(xmlInputElem)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    ' Add the Application Key data to the request node
    AddApplicationDetails vxmlRequest, xmlTempResponseNode
    
    Set xmlTempResponseNode = IPackageManagerTxBO_AddCustomerToPackage(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the Package and Application Keys generated to Out document
    Dim objXmlKeys As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    
    Set objXmlNode = xmlOut.createElement("CUSTOMERPACKAGEKEYS")
    Set objXmlKeys = xmlResponseElem.appendChild(objXmlNode)
                
    AddTagsToResponse vxmlRequest, objXmlKeys, "PACKAGENUMBER"
    AddTagsToResponse vxmlRequest, objXmlKeys, "APPLICATIONNUMBER"
    AddTagsToResponse vxmlRequest, objXmlKeys, "APPLICATIONFACTFINDNUMBER"
    
    Set IPackageManagerTxBO_CreatePackageAndApplication = xmlResponseElem
    
    Exit Function

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreatePackageAndApplicationExit:

    Set objIPackageManagerBO = Nothing
    Set objIApplicationBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set objXmlKeys = Nothing
    Set objXmlNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlInputElem = Nothing
    
    Exit Function

CreatePackageAndApplicationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_CreatePackageAndApplication = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePackageAndApplicationExit

End Function

Private Sub AddApplicationDetails(ByVal robjXMLElement As IXMLDOMElement, _
                                  ByVal xmlResponse As IXMLDOMElement) _
' header ----------------------------------------------------------------------------------
' description:  Add the Application Key data from Response Elem to other element passed in.
'               Called from 'CreatePackageAndApplication' metod
' pass:         robjXMLElement - the element to append the application data to
'               xmlResponse - Element containing the Application data
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo AddApplicationDetailsVbErr
    
    Const strFunctionName As String = "AddApplicationDetails"
                                   
    Dim objXmlAssist As New XMLAssist
        
    Dim objXmlTagList As IXMLDOMNodeList
    Set objXmlTagList = robjXMLElement.getElementsByTagName("CUSTOMERPACKAGE")
    
    Dim objXmlApplicationList As IXMLDOMNodeList
    Set objXmlApplicationList = xmlResponse.getElementsByTagName("APPLICATIONNUMBER")
    
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlApplicationNode As IXMLDOMNode
    
    Dim i0 As Integer
    
    i0 = 0
    For Each objXmlNode In objXmlTagList
        Set objXmlApplicationNode = objXmlApplicationList.Item(i0)
        If Not objXmlApplicationNode Is Nothing Then
            objXmlNode.appendChild objXmlApplicationNode
            i0 = i0 + 1
        End If
    Next
        
    Set objXmlApplicationList = xmlResponse.getElementsByTagName("APPLICATIONFACTFINDNUMBER")
        
    i0 = 0
    For Each objXmlNode In objXmlTagList
        Set objXmlApplicationNode = objXmlApplicationList.Item(i0)
        If Not objXmlApplicationNode Is Nothing Then
            objXmlNode.appendChild objXmlApplicationNode
            i0 = i0 + 1
        End If
    Next
            
    Set xmlResponse = Nothing
    Set objXmlAssist = Nothing
    
    Exit Sub
    
AddApplicationDetailsVbErr:

    Set xmlResponse = Nothing
    Set objXmlAssist = Nothing

    Dim objErrAssist As New ErrAssist
    objErrAssist.AddToErrSource strFunctionName
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Sub

Private Sub AddPackageDetails(ByRef robjXMLElement As IXMLDOMElement, _
                              ByVal xmlResponse As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Add the Package Key data from Response Elem to other element passed in.
'               Called from 'CreatePackageAndApplication' metod
' pass:         robjXMLElement - the element to append the package data to
'               xmlResponse - Element containing the package data
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo AddPackageDetailsVbErr

    Const strFunctionName As String = "AddPackageDetails"
                                   
    Dim objXmlAssist As New XMLAssist
                        
    Dim objXmlTagList As IXMLDOMNodeList
    Set objXmlTagList = robjXMLElement.getElementsByTagName("CUSTOMERPACKAGE")
    
    Dim objXmlPackageList As IXMLDOMNodeList
    Set objXmlPackageList = xmlResponse.getElementsByTagName("PACKAGENUMBER")
    
    Dim objXmlNode As IXMLDOMNode
    Dim objPackageNode As IXMLDOMNode
    Dim i0 As Integer
    
    i0 = 0
    For Each objXmlNode In objXmlTagList
        Set objPackageNode = objXmlPackageList.Item(i0)
        If Not objPackageNode Is Nothing Then
            objXmlNode.appendChild objPackageNode
            i0 = i0 + 1
        End If
    Next
        
    Set xmlResponse = Nothing
    Set objXmlAssist = Nothing
    
    Exit Sub

AddPackageDetailsVbErr:
    
    Set xmlResponse = Nothing
    Set objXmlAssist = Nothing
    
    Dim objErrAssist As New ErrAssist

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub AddTagsToResponse(ByVal vobjXmlIn As IXMLDOMElement, _
                            ByRef robjXmlout As IXMLDOMElement, _
                            ByVal vstrTagName As String)

    Dim objXmlList As IXMLDOMNodeList
    
    Set objXmlList = vobjXmlIn.getElementsByTagName(vstrTagName)
    
    Dim objXmlNode As IXMLDOMNode
    For Each objXmlNode In objXmlList
        robjXmlout.appendChild objXmlNode
    Next
    
End Sub

Private Function IPackageManagerTxBO_DeleteCustomerFromPackage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerFromPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_DeleteCustomerFromPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMElement

    Dim objIPackageManagerDO As IPackageManagerDO
    Dim objApplManagerBO As ApplicationManagerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERPACKAGE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                                    oeMissingPrimaryTag, cstrTABLE_CUSTOMERPACKAGE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIPackageManagerDO = New PackageManagerDO
        Set objApplManagerBO = New ApplicationManagerBO
    Else
        Set objIPackageManagerDO = objContext.CreateInstance(App.Title & ".PackageManagerDO")
        Set objApplManagerBO = objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
    End If
       
    Dim strXmlResponse As String
    Dim strXmlInput As String
    
    If Not xmlElement Is Nothing Then
        objIPackageManagerDO.DeleteCustomerPackage xmlElement
        ' Build the request string and Delete customer from application
        Set xmlRequestNode = objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
        xmlRequestNode.appendChild xmlElement
        strXmlInput = xmlRequestNode.xml
        strXmlInput = Replace(strXmlInput, "CUSTOMERPACKAGE", "CUSTOMERROLE")
        strXmlInput = Replace(strXmlInput, "/CUSTOMERPACKAGE", "/CUSTOMERROLE")
        strXmlResponse = objApplManagerBO.DeleteCustomerFromApplication(strXmlInput)
        objErrAssist.CheckXMLResponse strXmlResponse, True
    Else
        ' For each relevant node call the data object Create
        For Each xmlElement In xmlNodeList
            objIPackageManagerDO.DeleteCustomerPackage xmlElement
             ' Delete customer from application
            Set xmlRequestNode = objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
            xmlRequestNode.appendChild xmlElement
            strXmlInput = xmlRequestNode.xml
            strXmlInput = Replace(strXmlInput, "CUSTOMERPACKAGE", "APPLICATION")
            strXmlInput = Replace(strXmlInput, "/CUSTOMERPACKAGE", "/APPLICATION")
            strXmlResponse = objApplManagerBO.DeleteCustomerFromApplication(strXmlInput)
            objErrAssist.CheckXMLResponse strXmlResponse, True
        Next
    End If
        
    Set IPackageManagerTxBO_DeleteCustomerFromPackage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
DeleteCustomerFromPackageExit:

    Set objIPackageManagerDO = Nothing
    Set objApplManagerBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlRequestNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteCustomerFromPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_DeleteCustomerFromPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteCustomerFromPackageExit
End Function

Private Function IPackageManagerTxBO_DeleteCustomerPackage( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteCustomerPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_DeleteCustomerPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIPackageManagerDO As IPackageManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_CUSTOMERPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_CUSTOMERPACKAGE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_CUSTOMERPACKAGE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIPackageManagerDO = New PackageManagerDO
    Else
        Set objIPackageManagerDO = objContext.CreateInstance(App.Title & ".PackageManagerDO")
    End If
    
    Dim strXmlResponse As String
    
    If Not xmlElement Is Nothing Then
        objIPackageManagerDO.DeleteCustomerPackage xmlElement
    Else
        ' For each relevant node call the data object Create
        For Each xmlElement In xmlNodeList
            objIPackageManagerDO.DeleteCustomerPackage xmlElement
        Next
    End If
        
    Set IPackageManagerTxBO_DeleteCustomerPackage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteCustomerPackageExit:

    Set objIPackageManagerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteCustomerPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_DeleteCustomerPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
        Set objContext = Nothing
    End If
    
    ' go to clean-up section
    Resume DeleteCustomerPackageExit

End Function

Private Function IPackageManagerTxBO_UpdateApplicationPackage( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationPackageVbErr

    Const strFunctionName As String = "IPackageManagerTxBO_UpdateApplicationPackage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIPackageManagerDO As IPackageManagerDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_APPLICATIONPACKAGE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_APPLICATIONPACKAGE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_APPLICATIONPACKAGE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIPackageManagerDO = New PackageManagerDO
    Else
        Set objIPackageManagerDO = objContext.CreateInstance(App.Title & ".PackageManagerDO")
    End If
    
    
    If Not xmlElement Is Nothing Then
        objIPackageManagerDO.UpdateApplicationPackage xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objIPackageManagerDO.UpdateApplicationPackage xmlElement
        Next xmlElement
    End If
        
    Set IPackageManagerTxBO_UpdateApplicationPackage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationPackageExit:

    Set objIPackageManagerDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationPackageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IPackageManagerTxBO_UpdateApplicationPackage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationPackageExit

End Function

Private Function GetNextPackageNumber() As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a new package number, including any prefix
' pass:     n/a
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNextPackageNumberVbErr

    Const strFunctionName As String = "GetNextPackageNumber"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objErrAssist As New ErrAssist
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlTableNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objSqlAssist As New SQLAssist
    
    Dim objADOAssist As IADOAssist
    
    Dim strNextPackageNumber As String
    Dim strCheckDigit As String
    
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    'DM     06/06/01    SYS1946 Sequences.
    #If GENERIC_SQL Then
        Dim adoParameterPackageNumber As ADODB.Parameter
        Const ProcName = "usp_GetNextPackageNo"
    #Else
        Dim adoRecSet As ADODB.Recordset
    #End If
    Dim strResponse As String
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    '------------------------------------------------------------------------------------------
    ' get next number
    '------------------------------------------------------------------------------------------
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        #If GENERIC_SQL Then
        #Else
        .CursorLocation = adUseClient ' DLM Changed not returning a record set.
        #End If
        .open
    End With
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        Set .ActiveConnection = adoConnection
        'DM     06/06/01    SYS1946 Sequences.
        #If GENERIC_SQL Then
            .CommandType = adCmdStoredProc
            .CommandText = ProcName
            'DM SYS2511 changed parameter type to adVarChar
            Set adoParameterPackageNumber = adoCommand.CreateParameter("p_NextNumber", adVarChar, adParamOutput, 12)
            adoCommand.Parameters.Append adoParameterPackageNumber
        #Else
            .CommandType = adCmdText                                        ' DLM Changed
            .CommandText = "SELECT SEQ_PACKAGENO.NextVal FROM DUAL"         ' DLM Changed
        #End If
        .Prepared = True
    End With
    
    ' Keep getting numbers and calculating the check digit until the
    ' number returned produces a valid Check Digit
    Do
        #If GENERIC_SQL Then
            adoCommand.Execute
            strNextPackageNumber = Format(adoParameterPackageNumber.Value, "0000000")
        #Else
            Set adoRecSet = adoCommand.Execute                              ' DLM Change
            If adoRecSet.EOF Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
            End If
        
            ' APS UNIT TEST REF 84 - Prefix the Next Number with leading zeros
            strNextPackageNumber = Format(adoRecSet.Fields(0).Value, "0000000") ' DLM Changed
        #End If
        strCheckDigit = GenerateCheckDigit(strNextPackageNumber)
    Loop Until Not strCheckDigit = vbNullString
  
    Set adoCommand.ActiveConnection = Nothing
        adoConnection.Close
    Set adoConnection = Nothing
    #If GENERIC_SQL Then
        Set adoParameterPackageNumber = Nothing
    #Else
        If Not adoRecSet Is Nothing Then
                If adoRecSet.State = adStateOpen Then
                        adoRecSet.Close
                End If
        End If
        Set adoRecSet = Nothing
    #End If
    GetNextPackageNumber = strNextPackageNumber & strCheckDigit
 
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
'MH This isn't a  good idea in a private function
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'        Set objContext = Nothing
'    End If
    
    Exit Function

GetNextPackageNumberVbErr:
   
    #If GENERIC_SQL Then ' DLM Changed
        Set adoParameterPackageNumber = Nothing
    #Else
        ' Close the recordset if still open
        If Not adoRecSet Is Nothing Then
                If adoRecSet.State = adStateOpen Then
                        adoRecSet.Close
                End If
        End If
            
        Set adoRecSet = Nothing
    #End If
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
                If adoConnection.State = adStateOpen Then
                        adoConnection.Close
                End If
        End If

    Set adoConnection = Nothing
        Set objXmlAssist = Nothing
    
    If objErrAssist.IsSystemError Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    Set objErrAssist = Nothing
    
'MH This isn't a  good idea in a private function
'    If Not objContext Is Nothing Then
'        objContext.SetAbort
'    End If
    
    GetNextPackageNumber = ""
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function




