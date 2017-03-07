VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AdditionalQuestionsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AdditionalQuestionsBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Additional Questions Business Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'JR     26/03/01    SYS2048 Created
'                   Added FindAdditionalQuestionList
'                   Added FindAppAdditionalQuestionList
'                   Added CreateAppAdditionalQuestions
'                   Added UpdateAppAdditionalQuestions
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IAdditionalQuestionsBO

Private gobjContext As ObjectContext
Private gobjDO As IAdditionalQuestionsDO
Private gobjTxBO As IAdditionalQuestionsTxBO

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'MV     14/06/2002  BMIDS00013 - Created IAdditionalQuestionsBO_GetAdditionalQuestionsList
'MV     19/06/2002  BMIDS00013 - Modified IAdditionalQuestionsBO_GetAdditionalQuestionsList
'------------------------------------------------------------------------------------------

Private Sub IAdditionalQuestionsBO_CreateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified ApplnAdditionalQuestions entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateAppAdditionalQuestionsExit
    
    Const strFunctionName As String = "IAdditionalQuestionsBO_CreateAppAdditionalQuestions"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "APPLNADDITIONALQUESTIONS"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsTxBO")
    End If
    
    gobjTxBO.CreateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
     
CreateAppAdditionalQuestionsExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsBO_FindAdditionalQuestionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve a list of ADDITIONALQUESTIONS entities
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       ADDITIONALQUESTIONS node appended to vxmlResponseNode on exit
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
On Error GoTo FindAdditionalQuestionListExit

    Const cstrFunctionName As String = "IAdditionalQuestionsBO_FindAdditionalQuestionList"

    Dim xmlQuestionNode As IXMLDOMNode

    ' Get mandatory Question node
    Set xmlQuestionNode = xmlGetMandatoryNode(vxmlRequestNode, "QUESTION")
    xmlGetMandatoryAttribute xmlQuestionNode, "TYPE"
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsDO")
    End If
    
    gobjDO.FindAdditionalQuestionList xmlQuestionNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindAdditionalQuestionListExit:

    Set xmlQuestionNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsBO_FindAppAdditionalQuestionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve a list of APPLNADDITIONALQUESTIONS entities
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       QUESTIONS node appended to vxmlResponseNode on exit
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
On Error GoTo FindAppAdditionalQuestionListExit

    Const cstrFunctionName As String = "IAdditionalQuestionsBO_FindAppAdditionalQuestionList"

    Dim xmlAppQuestionNode As IXMLDOMNode
    
    ' Get Mandatory nodes
    Set xmlAppQuestionNode = xmlGetMandatoryNode(vxmlRequestNode, "ADDITIONALQUESTIONS")
    xmlGetMandatoryAttribute xmlAppQuestionNode, "APPLICATIONNUMBER"
    xmlGetMandatoryAttribute xmlAppQuestionNode, "APPLICATIONFACTFINDNUMBER"
           
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsDO")
    End If
    
    gobjDO.FindAppAdditionalQuestionList xmlAppQuestionNode, vxmlResponseNode
        
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindAppAdditionalQuestionListExit:

    Set xmlAppQuestionNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsBO_GetAdditionalQuestionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Private Sub IAdditionalQuestionsBO_GetAdditionalQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetAdditionalQuestionsListExit

    Const cstrFunctionName As String = "IAdditionalQuestionsBO_GetAdditionalQuestionsList"

    Dim xmlAppQuestionNode As IXMLDOMNode
    Dim xmlResponseElem  As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlAdditionalQuestionList As IXMLDOMNodeList
    Dim xmlAdditionalQuestionsNode As IXMLDOMElement
    Dim xmlPatternMatch  As IXMLDOMNode
    Dim xmlRequestElement As IXMLDOMElement
    Dim xmlTempRequest As IXMLDOMElement
    
    Dim strAppNumber  As String
    Dim strAppFFNumber  As String
    Dim strQuestionRef As String
    Dim strPattern As String
    Dim strType  As Integer
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
'Stage1:

    ' Get Mandatory nodes
    Set xmlAppQuestionNode = xmlGetMandatoryNode(vxmlRequestNode, "ADDITIONALQUESTIONSLIST")
    strType = xmlGetMandatoryAttributeText(xmlAppQuestionNode, "TYPE")
    strAppNumber = xmlGetMandatoryAttributeText(xmlAppQuestionNode, "APPLICATIONNUMBER")
    strAppFFNumber = xmlGetMandatoryAttributeText(xmlAppQuestionNode, "APPLICATIONFACTFINDNUMBER")
    
    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsDO")
    End If
    
    gobjDO.GetAdditionalQuestionsList xmlAppQuestionNode, xmlResponseElem
        
'Stage2:
    
    'Extract AdditionalQuestions Node List
    Set xmlAdditionalQuestionList = xmlResponseElem.selectNodes(".//ADDITIONALQUESTIONS")
    
    If Not xmlAdditionalQuestionList Is Nothing Then
        
        'Loop thru the Node List
        For Each xmlAdditionalQuestionsNode In xmlAdditionalQuestionList
        
            strQuestionRef = xmlGetMandatoryAttributeText(xmlAdditionalQuestionsNode, "QUESTIONREFERENCE")
            strPattern = ".//ADDITIONALQUESTIONS[@APPLICATIONNUMBER='" & strAppNumber & "' and @QUESTIONREFERENCE='" & strQuestionRef & "']"
            
            'Find the Node for the strPattern
            Set xmlPatternMatch = xmlGetNode(xmlResponseElem, strPattern)
            
                'If Not found then call CreateAppAdditionalQuestions to create a record
                'else append it to the Response
                If xmlPatternMatch Is Nothing Then
                
                    Dim xmlOut As New FreeThreadedDOMDocument40
                    xmlOut.validateOnParse = False
                    xmlOut.setProperty "NewParser", True
                    Dim xmlTempResponseElement As IXMLDOMElement
                    
                    'Creat Response element
                    Set xmlTempResponseElement = xmlOut.createElement("RESPONSE")
                    xmlOut.appendChild xmlTempResponseElement
                    xmlTempResponseElement.setAttribute "TYPE", "SUCCESS"
            
                    'Create Request Element
                    Set xmlTempRequest = xmlGetRequestNode(vxmlRequestNode)
                    Set xmlRequestElement = xmlDoc.createElement("APPLNADDITIONALQUESTIONS")
                    
                    xmlTempRequest.appendChild xmlRequestElement
                    
                    xmlRequestElement.setAttribute "APPLICATIONNUMBER", strAppNumber
                    xmlRequestElement.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNumber
                    xmlRequestElement.setAttribute "QUESTIONREFERENCE", strQuestionRef
                    xmlRequestElement.setAttribute "TYPE", strType
                    
                    'Call IAdditionalQuestionsBO_CreateAppAdditionalQuestions
                    IAdditionalQuestionsBO_CreateAppAdditionalQuestions xmlTempRequest, xmlTempResponseElement
                    errCheckXMLResponseNode xmlTempResponseElement, vxmlResponseNode, True
                    
                    'If success then append to the response node
                    xmlAdditionalQuestionsNode.setAttribute "APPLICATIONNUMBER", strAppNumber
                    xmlAdditionalQuestionsNode.setAttribute "APPLICATIONFACTFINDNUMBER", strAppFFNumber
                    
                    Set xmlTempResponseElement = Nothing
                    Set xmlOut = Nothing
                    
                End If
            
            'Append to the Response Element
            vxmlResponseNode.appendChild xmlAdditionalQuestionsNode
            
        Next
    
    End If
    
GetAdditionalQuestionsListExit:
    
    Set xmlAppQuestionNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlDoc = Nothing
    Set xmlAdditionalQuestionList = Nothing
    Set xmlAdditionalQuestionsNode = Nothing
    Set xmlPatternMatch = Nothing
    Set xmlRequestElement = Nothing
    Set xmlTempRequest = Nothing
    Set xmlAppQuestionNode = Nothing

    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub


Private Sub IAdditionalQuestionsBO_UpdateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified APPLNADDITIONALQUESTIONS entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateAppAdditionalQuestionsExit
    
    Const strFunctionName As String = "IAdditionalQuestionsBO_UpdateAppAdditionalQuestions"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsTxBO")
    End If
    
    gobjTxBO.UpdateAppAdditionalQuestions vxmlRequestNode, vxmlResponseNode
        
UpdateAppAdditionalQuestionsExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
   Set gobjContext = Nothing
End Sub

