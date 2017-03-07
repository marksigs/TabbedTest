VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IntermediaryBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      IntermediaryBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'               contains (implied) standard interface:
'               Public Function GetIndividualIntermediary(vstrXMLRequest As String) As String
'               Public Function FindIndividualIntermediaryList(vstrXMLRequest As String) As String
'
'Dependencies:  IntermediaryDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     29/10/99    Created
'DRC    04/10/01      SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
Option Explicit

Private Const cstrTABLE_NAME = "INTERMEDIARY"

Public Function FindIndividualIntermediaryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the details (full list) of Intermediary that match the criteria
' pass       :  XML string with Intermediary GUID
' return     :  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    
On Error GoTo FindIndividualIntermediaryListVbErr
    Const strFunctionName As String = "FindIndividualIntermediaryList"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    Dim objIntermediaryDO As IntermediaryDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Call ValidateFindListIndividualInput(xmlIn)  'Validate the inputs and get the updataed DOM document
    
    If objContext Is Nothing Then
        Set objIntermediaryDO = New IntermediaryDO
    Else
        Set objIntermediaryDO = objContext.CreateInstance(App.Title & ".IntermediaryDO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    'call Data Object's corresponding function
    Dim strXMLData As String
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
        
    strXMLData = objIntermediaryDO.FindIndividualIntermediaryList(xmlRequestNode.xml)
    Set xmlData = objXMLAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
    
    FindIndividualIntermediaryList = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindIndividualIntermediaryListExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXMLAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing

    Exit Function

FindIndividualIntermediaryListVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindIndividualIntermediaryList = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume FindIndividualIntermediaryListExit

End Function

Public Function GetIndividualIntermediary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a one Individual Intermediary
' pass       :  XML string with Intermediary GUID
' return     :  xml Response data stream containing results of operation
'               either: TYPE="SUCCESS" and xml representation of data
'               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetIndividualIntermediaryVbErr
    
    Const strFunctionName As String = "GetIndividualIntermediary"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Dim objIntermediaryDO As IntermediaryDO
       
    ValidateGetIndIntermediaryInput vstrXMLRequest
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIntermediaryDO = New IntermediaryDO
    Else
        Set objIntermediaryDO = objContext.CreateInstance(App.Title & ".IntermediaryDO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    'call Data Object's corresponding function
    Dim strXMLData As String
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
        
    strXMLData = objIntermediaryDO.GetIndividualIntermediary(xmlRequestNode.xml)
    Set xmlData = objXMLAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
    
    GetIndividualIntermediary = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetIndividualIntermediaryExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set objXMLAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetIndividualIntermediaryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetIndividualIntermediary = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetIndividualIntermediaryExit

End Function

Private Sub ValidateFindListIndividualInput(ByVal xmlIn As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:  Validate the input data of for the method FindIndividualIntermediaryList
'               and update the values, if required
' pass       :  XML document containg the XML Input
' return     :  XML document (updated, if any required)
'------------------------------------------------------------------------------------------
   
    On Error GoTo ValidateFindListIndividualInputVbErr

    Const strFunctionName As String = "ValidateFindListIndividualInput"
    
    Dim objErrAssist As New ErrAssist
    Dim objPanelIdElement As IXMLDOMElement
    Dim objForeNameElement As IXMLDOMElement
    Dim objSurNameElement As IXMLDOMElement
    Dim objTownElement As IXMLDOMElement

    'Check that the elements exist for PANELID, FORENAME, SURNAME, TOWN
    'and if any of these has value '*' replace it with empty string (equal to NULL)
    Set objPanelIdElement = xmlIn.getElementsByTagName("PANELID").Item(0)
    If objPanelIdElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "PANELID tag not found"
    End If
    If objPanelIdElement.Text = "*" Then
        objPanelIdElement.Text = ""
    End If
    
    Set objForeNameElement = xmlIn.getElementsByTagName("FORENAME").Item(0)
    If objForeNameElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "FORENAME tag not found"
    End If
    If objForeNameElement.Text = "*" Then
        objForeNameElement.Text = ""
    End If
    
    Set objSurNameElement = xmlIn.getElementsByTagName("SURNAME").Item(0)
    If objSurNameElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "SURNAME tag not found"
    End If
    If objSurNameElement.Text = "*" Then
        objSurNameElement.Text = ""
    End If
    
    Set objTownElement = xmlIn.getElementsByTagName("TOWN").Item(0)
    If objTownElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "TOWN tag not found"
    End If
    If objTownElement.Text = "*" Then
        objTownElement.Text = ""
    End If
    
    If objPanelIdElement.Text = "" And objForeNameElement.Text = "" And _
       objSurNameElement.Text = "" And objTownElement.Text = "" Then
        
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 256
    End If
    
    If objForeNameElement.Text <> "" And objSurNameElement.Text = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 257
    End If

    Set objErrAssist = Nothing
    Set objPanelIdElement = Nothing
    Set objForeNameElement = Nothing
    Set objSurNameElement = Nothing
    Set objTownElement = Nothing
    
    Exit Sub
ValidateFindListIndividualInputVbErr:
    
    Set objErrAssist = Nothing
    Set objPanelIdElement = Nothing
    Set objForeNameElement = Nothing
    Set objSurNameElement = Nothing
    Set objTownElement = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    ' re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ValidateGetIndIntermediaryInput(ByVal vstrXMLRequest As String)
    
    On Error GoTo ValidateGetIndIntermediaryInputVbErr
    
    Const strFunctionName As String = "ValidateGetIndIntermediaryInput"
    
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlRequestElem As IXMLDOMElement
    Dim objXMLAssist As New XMLAssist
    
   Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
   Set xmlRequestElem = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
   Dim strXMLData As String
    
    If xmlRequestElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim strTagName As String, strIntermediaryGuid As String
    
    strTagName = "INTERMEDIARYGUID"
    strIntermediaryGuid = objXMLAssist.GetTagValue(xmlRequestElem, strTagName)
    
    If strIntermediaryGuid = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 507
    End If
    
    Set xmlIn = Nothing
    Set xmlRequestElem = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
        
    Exit Sub
ValidateGetIndIntermediaryInputVbErr:

    Set xmlIn = Nothing
    Set xmlRequestElem = Nothing
    Set objXMLAssist = Nothing
      
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
        
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
   
End Sub
