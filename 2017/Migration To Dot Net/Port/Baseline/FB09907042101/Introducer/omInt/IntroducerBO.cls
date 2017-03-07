VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IntroducerBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------
' Class Module  <IntroducerBO>
' File:         <IntroducerBO.cls>
' Author:       Martin O'Hanlon
' Date:         20/09/2002
' Reference:    Business Object Specification--IntroducerBO.doc
' Purpose:      To interact with BM's Introducer system to search and pull back
'               introducer data
' History:
'   AQR                 AUTHOR                      DESCRIPTION           DATE
'   ---                 ------                      -----------          ------
' INWP1, BM00502        Martin O'Hanlon             Created              20/09/2002
'-----------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private gobjContext As ObjectContext

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'------------------------------------------------------------------------
' Procedure omRequest
' Author: MO
' Date: 20/09/2002
' Purpose: Main request to interface with the Introducer system
' Input parameters: vstrXMLIn
' Output parameters: String (XML)
'------------------------------------------------------------------------
Public Function omRequest(ByVal vstrXmlIn As String) As String

    On Error GoTo omRequestVbErr

    Const strFunctionName As String = "omRequest"

    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlRequestNode As IXMLDOMNode
    
  ' PSC 12/02/02 SYS4039 - Start
    If gobjContext Is Nothing Then
        errThrowError strFunctionName, oeMTSNotFound
    End If
    ' PSC 12/02/02 SYS4039 - End
    
    Set xmlInDoc = New FreeThreadedDOMDocument40
    xmlInDoc.validateOnParse = False
    xmlInDoc.setProperty "NewParser", True
    
    xmlInDoc.async = False
    
    xmlInDoc.loadXML vstrXmlIn

    Set xmlRequestNode = xmlGetMandatoryNode(xmlInDoc, "REQUEST")
    
    omRequest = DoRequest(xmlRequestNode)

omRequestExit:

    'destroy objects
    Set xmlInDoc = Nothing
    
    Exit Function

omRequestVbErr:

    omRequest = errCreateErrorResponse()

    Resume omRequestExit

End Function

'------------------------------------------------------------------------
' Procedure DoRequest
' Author: MO
' Date: 20/09/2002
' Purpose: Performs the request asked by the calling application
' Input parameters: vxmlRequestNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function DoRequest(ByVal vxmlRequestNode As IXMLDOMNode) As String

    On Error GoTo DoRequestExit
    
    Const strFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text
    Else
        strOperation = vxmlRequestNode.Attributes.getNamedItem("NAME").Text
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
        
        Case "GETINTRODUCERDATA"
            DoRequest = TESTGetIntroducerData(vxmlRequestNode)
        Case "FINDINTRODUCER"
            DoRequest = TESTFindIntroducer(vxmlRequestNode)
        Case Else
            'raise error
            errThrowError strFunctionName, oeNotImplemented, "OPERATION = " & strOperation
    End Select
    
DoRequestExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

'------------------------------------------------------------------------
' Procedure FindIntroducer
' Author: MO
' Date: 20/09/2002
' Purpose:  Calls the translation bridge and intiates the FindIntroducer
'           service
' Input parameters: vxmlRequestNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function FindIntroducer(ByVal vxmlRequestNode As IXMLDOMNode) As String

    On Error GoTo FindIntroducerExit
    
    Const strFunctionName As String = "FindIntroducer"
    
    FindIntroducer = CallBMTranslationBridge(vxmlRequestNode.xml)
    
FindIntroducerExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

'------------------------------------------------------------------------
' Procedure GetIntroducerData
' Author: MO
' Date: 20/09/2002
' Purpose:  Calls the translation bridge and intiates the GetIntroducerData
'           service, which returns all the information for one or more
'           introducers
' Input parameters: vxmlRequestNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function GetIntroducerData(ByVal vxmlRequestNode As IXMLDOMNode) As String

    On Error GoTo GetIntroducerDataExit
    
    Const strFunctionName As String = "GetIntroducerData"
    
    GetIntroducerData = CallBMTranslationBridge(vxmlRequestNode.xml)
   
GetIntroducerDataExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Function

'------------------------------------------------------------------------
' Procedure TESTGetIntroducerData
' Author: MO
' Date: 20/09/2002
' Purpose:  SHOULD NOT BE USED IN PRODUCTION, used while the BM service
'           is not available
' Input parameters: vxmlRequestNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function TESTGetIntroducerData(ByVal vxmlRequestNode As IXMLDOMNode) As String

    On Error GoTo TESTGetIntroducerDataExit
    
    Const strFunctionName As String = "TESTGetIntroducerData"
    
    Dim strConnectionString As String
    Dim strSQL As String
    Dim objConnection As ADODB.Connection
    Dim objRecordset As ADODB.Recordset
    Dim objIntroducerToSearchNode As IXMLDOMNode
    Dim objIntroducerNode As IXMLDOMElement
    
    Dim xmlOutDoc As New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    xmlOutDoc.async = False

    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
        
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    strConnectionString = adoGetDbConnectString
    
    For Each objIntroducerToSearchNode In vxmlRequestNode.selectNodes("INTRODUCER")
        
        'build SQL
        strSQL = "SELECT * " & _
                    "FROM intermediary " & _
                    "INNER JOIN address on Address.ADDRESSGUID = Intermediary.ADDRESSGUID " & _
                    "LEFT JOIN intermediaryindividual on Intermediary.intermediaryguid = Intermediaryindividual.intermediaryguid " & _
                    "LEFT JOIN intermediaryorganisation on Intermediary.intermediaryguid = intermediaryorganisation.intermediaryguid " & _
                    "LEFT JOIN contactdetails on contactdetails.CONTACTDETAILSGUID = intermediary.CONTACTDETAILSGUID " & _
                    "LEFT JOIN contacttelephonedetails on contacttelephonedetails.CONTACTDETAILSGUID = contactdetails.CONTACTDETAILSGUID and Usage = 10 " & _
                    "WHERE intermediary.INTERMEDIARYPANELID = '" & xmlGetMandatoryNode(objIntroducerToSearchNode, "@BMREFERENCE").Text & "'"
        
        Set objConnection = New ADODB.Connection
        
        objConnection.Open strConnectionString
        
        Set objRecordset = objConnection.Execute(strSQL)
        
        If objRecordset.BOF = False Or objRecordset.EOF = False Then
                
            Set objIntroducerNode = xmlOutDoc.createElement("INTRODUCER")
            
            objIntroducerNode.setAttribute "BMREFERENCE", "" & objRecordset.Fields("INTERMEDIARYPANELID")
            objIntroducerNode.setAttribute "TITLE", "Mr"
            objIntroducerNode.setAttribute "FORENAME", "" & objRecordset.Fields("FORENAME")
            objIntroducerNode.setAttribute "SURNAME", "" & objRecordset.Fields("SURNAME")
            objIntroducerNode.setAttribute "COMPANYNAME", "" & objRecordset.Fields("NAME")
            objIntroducerNode.setAttribute "BUILDINGNAME", "" & objRecordset.Fields("BUILDINGORHOUSENAME").Value
            objIntroducerNode.setAttribute "BUILDINGNO", "" & objRecordset.Fields("BUILDINGORHOUSENUMBER").Value
            objIntroducerNode.setAttribute "STREET", "" & objRecordset.Fields("STREET").Value
            objIntroducerNode.setAttribute "DISTRICT", "" & objRecordset.Fields("DISTRICT").Value
            objIntroducerNode.setAttribute "TOWN", "" & objRecordset.Fields("TOWN").Value
            objIntroducerNode.setAttribute "COUNTY", "" & objRecordset.Fields("COUNTY").Value
            objIntroducerNode.setAttribute "POSTCODE", "" & objRecordset.Fields("POSTCODE").Value
            objIntroducerNode.setAttribute "TELEPHONENO", "" & objRecordset.Fields("AREACODE").Value & " " & objRecordset.Fields("TELENUMBER").Value
            objIntroducerNode.setAttribute "PREFERREDCONTACTMETHOD", "T"
            objIntroducerNode.setAttribute "TYPEOFINTRODUCER", "Mortgage Broker"
            objIntroducerNode.setAttribute "DIRECTINDIRECT", "D"
            
            xmlResponseNode.appendChild objIntroducerNode
        
        Else
            Err.Raise vbObjectError, "TESTGETIntroducerData", "Cannot find introducer : " & xmlGetMandatoryNode(objIntroducerToSearchNode, "@BMRERENCE").Text
        End If
        
    Next
    
    TESTGetIntroducerData = xmlOutDoc.xml
    
TESTGetIntroducerDataExit:
    
    Set objConnection = Nothing
    Set objRecordset = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

'------------------------------------------------------------------------
' Procedure TESTFindIntroducer
' Author: MO
' Date: 20/09/2002
' Purpose:  SHOULD NOT BE USED IN PRODUCTION, used while the BM service
'           is not available
' Input parameters: vxmlRequestNode, vxmlResponseNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function TESTFindIntroducer(ByVal vxmlRequestNode As IXMLDOMNode) As String

    On Error GoTo TESTFindIntroducerExit
    
    Const strFunctionName As String = "TESTFindIntroducer"
    
    Dim strConnectionString As String
    Dim objConnection As ADODB.Connection
    Dim objRecordset As ADODB.Recordset
    Dim strSQL As String
    Dim strWhere As String
    Dim objIntroducerNode As IXMLDOMElement
    
    Dim xmlOutDoc As New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    
    xmlOutDoc.async = False

    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlOutDoc.appendChild(xmlResponseElem)
        
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim objAttribute As IXMLDOMAttribute
    
    strConnectionString = adoGetDbConnectString
    
    'build SQL
    strSQL = "SELECT * " & _
                    "FROM intermediary " & _
                    "INNER JOIN address on Address.ADDRESSGUID = Intermediary.ADDRESSGUID " & _
                    "LEFT JOIN intermediaryindividual on Intermediary.intermediaryguid = Intermediaryindividual.intermediaryguid " & _
                    "LEFT JOIN intermediaryorganisation on Intermediary.intermediaryguid = intermediaryorganisation.intermediaryguid " & _
                    "LEFT JOIN contactdetails on contactdetails.CONTACTDETAILSGUID = intermediary.CONTACTDETAILSGUID " & _
                    "LEFT JOIN contacttelephonedetails on contacttelephonedetails.CONTACTDETAILSGUID = contactdetails.CONTACTDETAILSGUID and Usage = 10 "
    
    For Each objAttribute In xmlGetMandatoryNode(vxmlRequestNode, "INTRODUCER").Attributes
        
        If Len(objAttribute.Text) > 0 Then
        
            Select Case objAttribute.Name
                Case "BMREFERENCE"
                    strWhere = strWhere & "INTERMEDIARYPANELID = '" & objAttribute.Text & "' AND "
                Case "DIRECTINDIRECT"
                    
                Case "NAME1"
                    strWhere = strWhere & "CONTACTFORENAME = '" & objAttribute.Text & "' AND "
                Case "NAME2"
                    strWhere = strWhere & "CONTACTSURNAME = '" & objAttribute.Text & "' AND "
                Case "TOWNPOSTCODE"
                    strWhere = strWhere & "(TOWN = '" & objAttribute.Text & "' OR POSTCODE = '" & objAttribute.Text & "') AND "
            End Select
        
        End If
        
    Next objAttribute
    
    If Len(strWhere) > 0 Then
        strWhere = "WHERE " & strWhere
        strWhere = Left(strWhere, Len(strWhere) - 4)
        strSQL = strSQL & strWhere
    End If
    
    Set objConnection = New ADODB.Connection
    
    objConnection.Open strConnectionString
    
    Set objRecordset = objConnection.Execute(strSQL)
    
    If objRecordset.BOF = False Or objRecordset.EOF = False Then
    
        While objRecordset.EOF = False
            
            Set objIntroducerNode = xmlOutDoc.createElement("INTRODUCER")
            
            objIntroducerNode.setAttribute "BMREFERENCE", "" & objRecordset.Fields("INTERMEDIARYPANELID").Value
            objIntroducerNode.setAttribute "NAME", "" & objRecordset.Fields("NAME").Value & objRecordset.Fields("FORENAME").Value & " " & objRecordset.Fields("SURNAME").Value
            objIntroducerNode.setAttribute "BUILDINGNAME", "" & objRecordset.Fields("BUILDINGORHOUSENAME").Value
            objIntroducerNode.setAttribute "BUILDINGNO", "" & objRecordset.Fields("BUILDINGORHOUSENUMBER").Value
            objIntroducerNode.setAttribute "STREET", "" & objRecordset.Fields("STREET").Value
            objIntroducerNode.setAttribute "DISTRICT", "" & objRecordset.Fields("DISTRICT").Value
            objIntroducerNode.setAttribute "TOWN", "" & objRecordset.Fields("TOWN").Value
            objIntroducerNode.setAttribute "COUNTY", "" & objRecordset.Fields("COUNTY").Value
            objIntroducerNode.setAttribute "POSTCODE", "" & objRecordset.Fields("POSTCODE").Value
            
            xmlResponseNode.appendChild objIntroducerNode
            
            Set objIntroducerNode = Nothing
            
            objRecordset.MoveNext
        Wend
        
    End If
    
    TESTFindIntroducer = xmlOutDoc.xml
    
TESTFindIntroducerExit:
    
    Set objConnection = Nothing
    Set objRecordset = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function

'------------------------------------------------------------------------
' Procedure CallBMTranslationBridge
' Author: MO
' Date: 31/09/2002
' Purpose:  This method makes the call to the BMTB
' Input parameters: vxmlRequestNode, vxmlResponseNode
' Output parameters: String (XML)
'------------------------------------------------------------------------
Private Function CallBMTranslationBridge(strRequest As String) As String

On Error GoTo CallBMTranslationBridgeExit
    
    Const strFunctionName As String = "CallBMTranslationBridge"
    
    Dim objBMTB As Object
    
    Set objBMTB = CreateObject("BMTranslationBridge.ExecuteBMService")
    
    CallBMTranslationBridge = objBMTB.ExecuteBusinessService(strRequest)

CallBMTranslationBridgeExit:
    
    Set objBMTB = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Function
