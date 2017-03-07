Attribute VB_Name = "errFunctions"
Option Explicit

'------------------------------------------------------------------------
' Procedure: errGetOmigaErrorNumber
' Author: MO
' Date: 14/08/2002
' Purpose: Returns the 'omiga' error number, gather from the err.number
' Input parameters: strResponse
' Output parameters: None
'------------------------------------------------------------------------
Public Function errGetOmigaErrorNumber(ByVal vlngErrorNo As Long) As Long
    errGetOmigaErrorNumber = vlngErrorNo - vbObjectError - 512
End Function

'------------------------------------------------------------------------
' Procedure: errCheckXMLResponse
' Author: MO
' Date: 27/07/2002
' Purpose: Checks the xml response that is returned from a Business object,
'           and raises an error if applicable
' Input parameters: strResponse
' Output parameters: None
'------------------------------------------------------------------------
Public Sub errCheckXMLResponse(strResponse As String)

On Error GoTo errCheckXMLResponseErrorOccured

    Dim xmlDOMDocument As FreeThreadedDOMDocument40
    Dim objResponseNode As IXMLDOMNode
    Dim objErrorNode As IXMLDOMNode
    Dim strResponseType As String
    
    Set xmlDOMDocument = New FreeThreadedDOMDocument40
    xmlDOMDocument.validateOnParse = False
    xmlDOMDocument.setProperty "NewParser", True
    
    If xmlDOMDocument.loadXML(strResponse) = True Then
        Set objResponseNode = xmlGetMandatoryNode(xmlDOMDocument.documentElement, "//RESPONSE")
        
        strResponseType = xmlGetMandatoryAttribute(objResponseNode, "TYPE").Text
        If strResponseType <> "SUCCESS" Then
            Set objErrorNode = xmlGetMandatoryNode(objResponseNode, "ERROR")
            Err.Raise xmlGetMandatoryNode(objErrorNode, "NUMBER").Text, xmlGetMandatoryNode(objErrorNode, "SOURCE").Text, xmlGetMandatoryNode(objErrorNode, "DESCRIPTION").Text
        End If
    Else
        Err.Raise direFailedToLoadXML, "errCheckXMLResponse", "Failed to load response XML : Reason - " & xmlDOMDocument.parseError.reason
    End If
    
errCheckXMLResponseExit:
    
    Set xmlDOMDocument = Nothing
    
Exit Sub

errCheckXMLResponseErrorOccured:
    
    Err.Raise Err.Number, "errCheckXMLResponse" & " : " & Err.Source, Err.Description
    Resume errCheckXMLResponseExit

End Sub
