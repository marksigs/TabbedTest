VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BuildingsAndCOntentsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BuildingsAndContentsBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions'
'               only
'               i.e. any code that does not require transaction support will reside in
'               this first level module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level)
'               module (Omiga4TxBO in this example)
'
'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'JLD    21/12/1999  Created
'JLD    01/02/2000  Added error to RESPONSE block if polaris failed to produce a calc.
'APS    04/02/2000  Changed strErrorText to strErrorString
'JLD    09/02/2000  Created generic interface to return from Calculate
'APS    16/03/2000  Changed Calculate function based on new Polaris output xml structure
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit

Implements IBuildingsAndContentsBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "BUILDINGSANDCONTENTS"


'BMIDS History:
'
'Prog   Date        Description
'GD     08/08/02    BMIDS00045 Changed Calculate so it doesn't call Polaris, just returns a dummy XML.

Public Function Calculate(ByVal vstrXmlInputInterface) As String
' header ----------------------------------------------------------------------------------
' description:  Calls the calculation
' pass:         vstrXmlInputInterface  xml Request data stream containing the
'                                      Input Interface for the calculation
'
' return:       RunCalcsEngine  xml Request data stream containing the
'                               Output Interface from the calculation
'
' Raise Errors: Error 235 - if the Polaris calculation returns an error
'
'------------------------------------------------------------------------------------------
On Error GoTo CalculateVbErr
    
    Const strFunctionName As String = "Calculate"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlPolarisOutDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlReturnTextNode As IXMLDOMNode
    Dim xmlStatusElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlBCCalcOutputNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    ''BMIDS00045 Dim objPolaris As PolarisInterface

    Dim strResponse As String
    
    'BMIDS00045 Set objPolaris = New PolarisInterface

    'Set xmlPolarisOutDoc = objXmlAssist.load(vstrXmlInputInterface, TypeName(Me), strFunctionName)
    'objXmlAssist.ChangeNodeName xmlPolarisOutDoc.documentElement, "BCCALCINPUT", "POLARISINPUT"
    'strResponse = objPolaris.ExecutePolaris(xmlPolarisOutDoc.documentElement.xml)
    'BMIDS00045 strResponse = objPolaris.ExecutePolaris(vstrXmlInputInterface)

    strResponse = "<BCCALCOUTPUT><STATUS>A</STATUS><BCRETURNTEXT>text</BCRETURNTEXT></BCCALCOUTPUT>"
    'obtain the response
    Set xmlPolarisOutDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        
    Set xmlStatusElem = xmlPolarisOutDoc.documentElement.selectSingleNode("//BCCALCOUTPUT/STATUS")
    Set xmlReturnTextNode = xmlPolarisOutDoc.documentElement.selectSingleNode("//BCCALCOUTPUT/BCRETURNTEXT")
    
    ' Polaris not returning xml
    If xmlStatusElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 235, "Status Tag not found in Polaris Response"
    End If
    
    If xmlReturnTextNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 235, "Returned Text Tag not found in Polaris Response"
    End If

    ' Polaris returning an error type "E"
    If xmlStatusElem.Text = "E" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 235, , xmlReturnTextNode.Text
    End If
    
    ' Interpretation of the Polaris return code
    If Not xmlStatusElem Is Nothing Then
    
        Dim objComboDO As ComboDO
        Dim strStatusValue As String
        Dim strStatusText As String
        
        If objContext Is Nothing Then
            Set objComboDO = New ComboDO
        Else
            Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        End If
        
        strStatusValue = objComboDO.GetFirstComboValueId("BCSubQuoteStatus", xmlStatusElem.Text)
        
        strStatusText = objComboDO.GetComboText("BCSubQuoteStatus", strStatusValue)
        
        xmlStatusElem.setAttribute "TEXT", strStatusText
        
        xmlStatusElem.Text = strStatusValue
                
        Set objComboDO = Nothing
        
    End If
    
    xmlResponseElem.appendChild xmlPolarisOutDoc.documentElement
    
    Calculate = xmlOut.documentElement.xml
    
CalculateExit:

    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlNode = Nothing
    Set xmlStatusElem = Nothing
    Set xmlPolarisOutDoc = Nothing
    'BMIDS00045Set objPolaris = Nothing

    Set objComboDO = Nothing

    Exit Function
    
CalculateVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Calculate = objErrAssist.CreateErrorResponse
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume CalculateExit

End Function


Private Function IBuildingsAndContentsBO_FindProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr
    
    Const strFunctionName As String = "IBuildingsAndContentsBO_FindProductList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIBCDO As IBuildingsAndContentsDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIBCDO = New BuildingsAndContentsDO
    Else
        Set objIBCDO = objContext.CreateInstance(App.Title & ".BuildingsAndContentsDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIBCDO.FindProductList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IBuildingsAndContentsBO_FindProductList = xmlResponseElem.cloneNode(True)

FindProductListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIBCDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBuildingsAndContentsBO_FindProductList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductListExit

End Function

Public Function FindProductList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindProductListVbErr

    Const strFunctionName As String = "FindProductList"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IBuildingsAndContentsBO_FindProductList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindProductList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindProductListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindProductList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductListExit

End Function



