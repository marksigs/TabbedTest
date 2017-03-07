VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PrepareTemplateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PrepareTemplateData.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:
'       Handles new print template data extraction.
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'TW     26/05/2005      Created
'AS     22/06/2005      CORE152 Added tracing.
'TW     08/07/2005      CORE163 Removed references to CUSTOMERNUMBER,
'                               CUSTOMERVERSIONNUMBER
'                               and GLOBALPARAMETER name
'                               from PrepareTemplateCRUD
'TW     11/07/2005      CORE164 Pass USERID and UNITID, if supplied, to omCRUD
'TW     18/10/2005      CORE204, MAR223, BBG2445, BBGR1657
'                               Extensions to PopulateTemplateCRUD to pick up 'Primary' xml
'                               data supplied in the TEMPLATEDATA node of an omPrint request by omPack to
'                               support Pack processing in the MARS project.
'TW     19/12/2005      CORE225 Reinstated references to CUSTOMERNUMBER
'                               and CUSTOMERVERSIONNUMBER
'TW     10/01/2006      CORE228 Ensure that CUSTOMERVERSIONNUMBER is only used if CUSTOMERNUMBER is supplied
'TW     30/01/2006      CORE232 Improvements for Pack Processing
'PSC    31/01/2006      CORE233 Correct logic to test for packs
'BC     16/02/2006      CORE246 Additional logging in PopulateTemplateCRUD
'BC     20/03/2006      CORE255 Additional, optional parameter (MultiApplicant) passed in
'BC     20/03/2006      CORE255 Remove postprocref from omCRUD call for local printing and perform XSL Transform programatically
'IK     09/08/2006      CORE291 direct load of .xslt file to allow use of xsl:import
'                               (code from P. Edney)
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Private gobjContext As ObjectContext

Private Type PDMReplacementStruct
    PrimaryXML As String
    ApplicationFactFindNumber As Integer
    ApplicationNumber As String
    CustomerNumber As String
    CustomerVersionNumber As Integer
'TW 11/07/2005 CORE164
'    GlobalParamName As String
'TW 11/07/2005 CORE164 End
    PDMMethod As String
    UserId As String
    UnitId As String
'TW 11/07/2005 CORE164 End
End Type

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
    gobjTrace.StartTrace
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub


Private Function omPDMReplacement(PDM As PDMReplacementStruct) As String
Const strFunctionName As String = "omPDMReplacement"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim strDataSource As String
Dim strDataSourceKey As String
Dim strDataSourceType As String
Dim strPath As String
Dim strWork As String
Dim x As Integer
Dim xmlAttribute As IXMLDOMAttribute
Dim xmlEngineData As FreeThreadedDOMDocument40
Dim xmlDataSourceKeyNode As IXMLDOMNode
Dim xmlDataSourceKeys As IXMLDOMNodeList
Dim xmlDataSourceNode As IXMLDOMNode
Dim xmlDataSourceList As IXMLDOMNodeList
Dim xmlMethodNode As IXMLDOMNode

    On Error GoTo omPDMReplacementError:
    
    
    strPath = App.Path
    x = InStrRev(strPath, "\")
    strPath = Left$(strPath, x)
    
    Set xmlEngineData = New FreeThreadedDOMDocument40
    xmlEngineData.setProperty "NewParser", True
    
    xmlEngineData.Load strPath & "XML\NewEngine.xml"
        
    Set xmlMethodNode = xmlEngineData.selectSingleNode("//TEMPLATE[@PDMMethod='" & UCase$(PDM.PDMMethod) & "']")
    If xmlMethodNode Is Nothing Then
        Err.Raise 40003, "PrepareTemplate.omPDMReplacement", "Method '" & PDM.PDMMethod & "' not found"
    End If
    Set xmlDataSourceList = xmlMethodNode.selectNodes("DATASOURCE")
    If xmlDataSourceList.length = 0 Then
        Err.Raise 40004, "PrepareTemplate.omPDMReplacement", "No data source(s) declared for method '" & PDM.PDMMethod & "'"
    End If
    For Each xmlDataSourceNode In xmlDataSourceList
        Set xmlAttribute = xmlDataSourceNode.Attributes.getNamedItem("Name")
        If xmlAttribute Is Nothing Then
            Err.Raise 40005, "PrepareTemplate.omPDMReplacement", "No name for data source declared for method '" & PDM.PDMMethod & "'"
        End If
        strDataSource = xmlAttribute.Text
        Set xmlAttribute = xmlDataSourceNode.Attributes.getNamedItem("Type")
        If xmlAttribute Is Nothing Then
            Err.Raise 40006, "PrepareTemplate.omPDMReplacement", "No type declared for data source '" & strDataSource & "' for method '" & PDM.PDMMethod & "'"
        End If
        strDataSourceType = xmlAttribute.Text
        
        Select Case strDataSourceType
            Case "View"
                strWork = strWork & vbCrLf & "Select * From " & strDataSource & " Where"
            Case "Stored Procedure"
                strWork = strWork & vbCrLf & "Execute " & strDataSource
        End Select
        Set xmlDataSourceKeys = xmlDataSourceNode.selectNodes("KEY")
        If xmlDataSourceKeys.length = 0 Then
            Err.Raise 40007, "PrepareTemplate.omPDMReplacement", "No key(s) declared for data source '" & strDataSource & "' for method '" & PDM.PDMMethod & "'"
        End If
        x = 0
        For Each xmlDataSourceKeyNode In xmlDataSourceKeys
            Set xmlAttribute = xmlDataSourceKeyNode.Attributes.getNamedItem("Name")
            If xmlAttribute Is Nothing Then
                Err.Raise 40008, "PrepareTemplate.omPDMReplacement", "No name for key declared for data source '" & strDataSource & "' for method '" & PDM.PDMMethod & "'"
            End If
            strDataSourceKey = xmlAttribute.Text
            Select Case strDataSourceType
                Case "View"
                    Select Case strDataSourceKey
                        Case "APPLICATIONNUMBER"
                            strDataSourceKey = strDataSourceKey & " = '" & PDM.ApplicationNumber & "'"
                        Case "APPLICATIONFACTFINDNUMBER"
                            strDataSourceKey = strDataSourceKey & " = " & PDM.ApplicationFactFindNumber
                        Case "CUSTOMERNUMBER"
                            strDataSourceKey = strDataSourceKey & " = '" & PDM.CustomerNumber & "'"
                        Case "CUSTOMERVERSIONNUMBER"
                            strDataSourceKey = strDataSourceKey & " = " & PDM.CustomerVersionNumber
                        Case "PRIMARYXML"
                            strDataSourceKey = strDataSourceKey & " = " & PDM.PrimaryXML
                        Case "USERID"
                            strDataSourceKey = strDataSourceKey & " = '" & PDM.UserId & "'"
                        Case "UNITID"
                            strDataSourceKey = strDataSourceKey & " = '" & PDM.UnitId & "'"
                    End Select
                    If x > 0 Then
                        strWork = strWork & " And"
                    End If
                Case "Stored Procedure"
                    Select Case strDataSourceKey
                        Case "APPLICATIONNUMBER"
                            strDataSourceKey = "'" & PDM.ApplicationNumber & "'"
                        Case "APPLICATIONFACTFINDNUMBER"
                            strDataSourceKey = PDM.ApplicationFactFindNumber
                        Case "CUSTOMERNUMBER"
                            strDataSourceKey = "'" & PDM.CustomerNumber & "'"
                        Case "CUSTOMERVERSIONNUMBER"
                            strDataSourceKey = PDM.CustomerVersionNumber
                        Case "PRIMARYXML"
                            strDataSourceKey = PDM.PrimaryXML
                        Case "USERID"
                            strDataSourceKey = "'" & PDM.UserId & "'"
                        Case "UNITID"
                            strDataSourceKey = "'" & PDM.UnitId & "'"
                    End Select
                    If x > 0 Then
                        strWork = strWork & " ,"
                    End If
            End Select
            strWork = strWork & " " & strDataSourceKey
            x = x + 1
        Next
        Select Case strDataSourceType
            Case "View"
                strWork = strWork & " FOR XML AUTO, BINARY Base64"
            Case "Stored Procedure"
        End Select
    Next
    omPDMReplacement = strWork
TidyUp:
    Set xmlEngineData = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function
omPDMReplacementError:
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    Resume TidyUp:
End Function

Private Function PopulateTemplateCRUD(PDM As PDMReplacementStruct, ByVal strXSLPath As String, ByVal strMultiApplicant As String) As String 'BC CORE255 20/03/2006
Const strFunctionName As String = "PopulateTemplateCRUD"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim strRequest As String
Dim strResponse As String
Dim strXSL As String
Dim strCustomerNumber As String  'BC CORE255

Dim CRUDObject As omCRUDBO

Dim xmlAttribute As IXMLDOMAttribute
Dim xmlRequestNode As IXMLDOMNode
Dim xmlRequestDoc As FreeThreadedDOMDocument40
Dim xmlSchemaNode As IXMLDOMNode

Dim xmlPrimaryXMLDoc As FreeThreadedDOMDocument40
Dim xmlPrimaryNode As IXMLDOMNode
Dim xmlResponseDoc As FreeThreadedDOMDocument40
Dim xmlResponseNode As IXMLDOMNode
Dim xmlChildNode As IXMLDOMNode
Dim xmlTemplateDataNode As IXMLDOMNode   ' PSC 31/01/2006 CORE233


Dim xmlDataDoc As FreeThreadedDOMDocument40
Dim xslProcessor As IXSLProcessor
Dim xslSheet As FreeThreadedDOMDocument40
Dim xslTemplate As XSLTemplate40
Dim xmlFMAChecklistElement As IXMLDOMElement 'BC CORE255
Dim xmlCustomerRoleElement As IXMLDOMElement 'BC CORE255

Dim x As Long

    On Error GoTo PopulateTemplateCRUDError:
    
' Create a REQUEST block for omCRUD
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    Set xmlRequestNode = xmlRequestDoc.createNode(NODE_ELEMENT, "REQUEST", "")
    xmlRequestDoc.appendChild xmlRequestNode
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("USERID")
    xmlAttribute.Text = PDM.UserId
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("CRUD_OP")
    xmlAttribute.Text = "READ"
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("ENTITY_REF")
    xmlAttribute.Text = PDM.PDMMethod
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("SCHEMA_NAME")
    xmlAttribute.Text = "PrepareTemplate"
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    ' PSC 31/01/2006 CORE233 - Start
    Set xmlPrimaryXMLDoc = New FreeThreadedDOMDocument40
    xmlPrimaryXMLDoc.async = False
    xmlPrimaryXMLDoc.setProperty "NewParser", True
    xmlPrimaryXMLDoc.loadXML PDM.PrimaryXML
    
    Set xmlTemplateDataNode = xmlPrimaryXMLDoc.selectSingleNode("//TEMPLATEDATA[*]")
    ' PSC 31/01/2006 CORE233 - End
    
    
'BC 20/03/2006 CORE255 Begin
'TW 30/01/2006 CORE232
    ' PSC 31/01/2006 CORE233
'    If xmlTemplateDataNode Is Nothing Then
'' Only post processs with the xsl when not in a pack
''TW 30/01/2006 CORE232 End
'        Set xmlAttribute = xmlRequestDoc.createAttribute("postProcRef")
'        xmlAttribute.Text = PDM.PDMMethod & ".xsl"
'        xmlRequestNode.Attributes.setNamedItem xmlAttribute
''TW 30/01/2006 CORE232
'    End If
'TW 30/01/2006 CORE232 End
'BC 20/03/2006 CORE255 End
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttribute.Text = "y"
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlAttribute = xmlRequestDoc.createAttribute("COMBOTYPELOOKUP")
    xmlAttribute.Text = "Y"
    xmlRequestNode.Attributes.setNamedItem xmlAttribute
    
    Set xmlSchemaNode = xmlRequestDoc.createNode(NODE_ELEMENT, PDM.PDMMethod, "")
    xmlRequestNode.appendChild xmlSchemaNode
    
    If Len(PDM.ApplicationNumber) > 0 Then
        Set xmlAttribute = xmlRequestDoc.createAttribute("APPLICATIONNUMBER")
        xmlAttribute.Text = PDM.ApplicationNumber
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
    End If
    
    If Len(PDM.ApplicationFactFindNumber) > 0 Then
        Set xmlAttribute = xmlRequestDoc.createAttribute("APPLICATIONFACTFINDNUMBER")
        xmlAttribute.Text = PDM.ApplicationFactFindNumber
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
    End If
    
'TW 8/7/2005 CORE163
'
'TW 19/12/2005 CORE225
    'BC 20/03/20065 CORE255 Begin
    If Not (strMultiApplicant = "Y") Then
        If Len(PDM.CustomerNumber) > 0 Then
            Set xmlAttribute = xmlRequestDoc.createAttribute("CUSTOMERNUMBER")
            xmlAttribute.Text = PDM.CustomerNumber
            xmlSchemaNode.Attributes.setNamedItem xmlAttribute
    'TW 10/01/2006 CORE228
            If Len(PDM.CustomerVersionNumber) > 0 Then
                Set xmlAttribute = xmlRequestDoc.createAttribute("CUSTOMERVERSIONNUMBER")
                xmlAttribute.Text = PDM.CustomerVersionNumber
                xmlSchemaNode.Attributes.setNamedItem xmlAttribute
            End If
    'TW 10/01/2006 CORE228 End
        End If
    End If
    'BC 20/03/20065 CORE255 End

'TW 10/01/2006 CORE228
'    If Len(PDM.CustomerVersionNumber) > 0 Then
'        Set xmlAttribute = xmlRequestDoc.createAttribute("CUSTOMERVERSIONNUMBER")
'        xmlAttribute.Text = PDM.CustomerVersionNumber
'        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
'    End If
'TW 10/01/2006 CORE228 End
'TW 19/12/2005 CORE225 End
'
'    If Len(PDM.GlobalParamName) > 0 Then
'        Set xmlAttribute = xmlRequestDoc.createAttribute("NAME")
'        xmlAttribute.Text = PDM.GlobalParamName
'        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
'    End If
'TW 8/7/2005 CORE163 End
'TW 11/07/2005 CORE164
    If Len(PDM.UserId) > 0 Then
        Set xmlAttribute = xmlRequestDoc.createAttribute("USERID")
        xmlAttribute.Text = PDM.UserId
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
    End If
    If Len(PDM.UnitId) > 0 Then
        Set xmlAttribute = xmlRequestDoc.createAttribute("UNITID")
        xmlAttribute.Text = PDM.UnitId
        xmlSchemaNode.Attributes.setNamedItem xmlAttribute
    End If
'TW 11/07/2005 CORE164 End
        
    gobjTrace.TraceXML xmlRequestDoc.xml, "PopulateTemplateCRUD_request"
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.0]", "CRUDObject.OmRequest(xmlRequestDoc.xml),entry"
    Set CRUDObject = New omCRUDBO
    strResponse = CRUDObject.OmRequest(xmlRequestDoc.xml)
    
    'BC 20/03/20065 CORE255 Begin
    Set xmlResponseDoc = New FreeThreadedDOMDocument40
    
    xmlResponseDoc.async = False
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.loadXML strResponse
    
    gobjTrace.TraceXML xmlResponseDoc.xml, "PopulateTemplateCRUD_response"
    'BC 20/03/20065 CORE255 End
    
'TW 18/10/2005 CORE204
    ' PSC 31/01/2006 CORE233
    If Not xmlTemplateDataNode Is Nothing Then
        ' PSC 31/01/2006 CORE233
        Set xmlPrimaryNode = xmlTemplateDataNode
        If xmlPrimaryNode.childNodes.length > 0 Then

            
'TW 30/01/2006 CORE232
'            Set xmlResponseNode = xmlResponseDoc.selectSingleNode("//TEMPLATEDATA")
            Set xmlResponseNode = xmlResponseDoc.selectSingleNode("//" & PDM.PDMMethod)
'TW 30/01/2006 CORE232 End
            If Not xmlResponseNode Is Nothing Then
                For Each xmlChildNode In xmlPrimaryXMLDoc.childNodes
                    xmlResponseNode.appendChild xmlChildNode.cloneNode(True)
                Next
            End If
        End If
    Else
        
        'BC CORE255 21/03/2006 Begin
        If strMultiApplicant = "Y" Then
            Set xmlFMAChecklistElement = xmlResponseDoc.selectSingleNode("//" & PDM.PDMMethod)
            If Len(PDM.CustomerNumber) > 0 Then
                xmlFMAChecklistElement.setAttribute "CUSTOMERNUMBER", PDM.CustomerNumber
            Else
                Set xmlCustomerRoleElement = xmlResponseDoc.selectSingleNode("//CUSTOMERROLE")
                strCustomerNumber = xmlCustomerRoleElement.getAttribute("CUSTOMERNUMBER")
                xmlFMAChecklistElement.setAttribute "CUSTOMERNUMBER", strCustomerNumber
            End If
        End If
        
        
        'BC CORE255 21/03/2006 End
        
    End If
        
    'CORE291 Peter Edney - 04/08/2006
    'strXSL = GetFileData(strXSLPath)

    Set xslSheet = New FreeThreadedDOMDocument40
    xslSheet.async = False
    xslSheet.setProperty "NewParser", True

    'CORE291 Peter Edney - 04/08/2006
    'xslSheet.loadXML strXSL
    xslSheet.Load strXSLPath
    
    Set xslTemplate = New XSLTemplate40
    Set xslTemplate.stylesheet = xslSheet
    Set xslProcessor = xslTemplate.createProcessor()

    xslProcessor.input = xmlResponseDoc
    
    gobjTrace.TraceXML xmlResponseDoc.xml, "PopulateTemplateCRUD_preTransform" 'BC CORE246

    x = xslProcessor.Transform()
    
    strResponse = xslProcessor.output
                
'TW 18/10/2005 CORE204 End
  
    gobjTrace.TraceXML strResponse, "PopulateTemplateCRUD_postTransform" 'BC CORE246
    
    PopulateTemplateCRUD = strResponse
        
    gobjTrace.TraceMessage TypeName(Me), strFunctionName, "[1.1]", "CRUDObject.OmRequest(xmlRequestDoc.xml),exit"
    gobjTrace.TraceXML PopulateTemplateCRUD, "PopulateTemplateCRUD_response"
    
TidyUp:
    Set xmlAttribute = Nothing
    Set xmlRequestNode = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlSchemaNode = Nothing
    
    Set xmlPrimaryXMLDoc = Nothing
    Set xmlPrimaryNode = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlResponseNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlTemplateDataNode = Nothing  ' PSC 31/01/2006 CORE233
    
    Set CRUDObject = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function

PopulateTemplateCRUDError:
    PopulateTemplateCRUD = vbNullString
    App.LogEvent "Failed in PrepareTemplate.PopulateTemplateCRUD " & Err.Description, vbLogEventTypeError
    Debug.Print Err.Description
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    On Error GoTo 0
    Resume TidyUp:
End Function

Public Function PopulateTemplate(ByVal strPDMMethod As String, _
                                 ByVal strTemplatePath As String, _
                                 ByVal strXSLPath As String, _
                                 ByVal strApplicationNumber As String, _
                                 ByVal intApplicationFactFindNumber As Integer, _
                                 ByVal strCustomerNumber As String, _
                                 ByVal intCustomerVersionNumber As Integer, _
                                 ByVal strPrimaryXML As String, _
                                 ByVal strUserId As String, _
                                 ByVal strUnitId As String, _
                                 Optional ByVal blnUseCRUD As Boolean = False)

Const strFunctionName As String = "PopulateTemplate"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim PDM As PDMReplacementStruct

Dim fs As FileSystemObject 'BC CORE255

Dim strQuery As String
Dim strXML As String
Dim strXSL As String
Dim strPath As String 'BC CORE255
Dim strXMLPath As String 'BC CORE255
Dim strEnginePath As String 'BC CORE255
Dim strMultiApplicant As String 'BC CORE255

Dim xmlDataDoc As FreeThreadedDOMDocument40
Dim xslProcessor As IXSLProcessor
Dim xslSheet As FreeThreadedDOMDocument40
Dim xslTemplate As XSLTemplate40
Dim xmlEngineData As FreeThreadedDOMDocument40 'BC CORE255
Dim xmlMethodNode As IXMLDOMNode 'BC CORE255
Dim xmlMultiApplicantElement As IXMLDOMElement 'BC CORE255

Dim x As Long

    On Error GoTo PopulateTemplateError:
    
    PDM.PDMMethod = strPDMMethod
    PDM.ApplicationNumber = strApplicationNumber
    PDM.ApplicationFactFindNumber = intApplicationFactFindNumber
    PDM.CustomerNumber = strCustomerNumber
    PDM.CustomerVersionNumber = intCustomerVersionNumber
    PDM.PrimaryXML = strPrimaryXML
    PDM.UserId = strUserId
'TW 11/07/2005 CORE164
'    PDM.GlobalParamName = strGlobalParamName
    PDM.UnitId = strUnitId

    'BC CORE255 20/03/2006 Begin - Retrieve optional attribute - MultiApplicant
    Set fs = New FileSystemObject
    
    strPath = App.Path
    x = InStrRev(strPath, "\")
    strPath = Left$(strPath, x)
    strXMLPath = strPath & "XML\"
    strEnginePath = strXMLPath & "NewEngine.xml"

' Check for the existence of the New Engine xml
    If fs.FileExists(strEnginePath) Then
        Set xmlEngineData = New FreeThreadedDOMDocument40
        xmlEngineData.Load strEnginePath
    
' Check if there is an entry in the New Engine xml for the PDM Method
        Set xmlMethodNode = xmlEngineData.selectSingleNode("//TEMPLATE[@PDMMethod='" & UCase$(strPDMMethod) & "']")
        Set xmlMultiApplicantElement = xmlMethodNode 'BC CORE255 20/03/2006
        
        'BC CORE255 20/03/2006 Begin
        If Not xmlMultiApplicantElement Is Nothing Then
            If Not (xmlMultiApplicantElement.getAttribute("MultiApplicant") = "") Then
                strMultiApplicant = UCase$(xmlMultiApplicantElement.getAttribute("MultiApplicant"))
            Else
                strMultiApplicant = "N"
            End If
        End If
        
    End If
    'BC CORE255 20/03/2006 End
    
    If blnUseCRUD Then
        PopulateTemplate = PopulateTemplateCRUD(PDM, strXSLPath, strMultiApplicant) 'BC CORE255 20/03/2006
    Else
        strQuery = omPDMReplacement(PDM)
        
        strXML = GetTableDataAsXML(strQuery)
        If strXML = "" Then
            Err.Raise 40001, "PrepareTemplate.PopulateTemplate", "No xml returned from GetTableDataAsXML(" & strQuery & ")"
        End If
        
        'CORE291 Peter Edney - 04/08/2006
        'strXSL = GetFileData(strXSLPath)
        'If strXSL = "" Then
        '    Err.Raise 40002, "PrepareTemplate.PopulateTemplate", "No xml returned from GetTableDataAsXML(" & strXSLPath & ")"
        'End If
        
        Set xmlDataDoc = New FreeThreadedDOMDocument40
        xmlDataDoc.setProperty "NewParser", True
        xmlDataDoc.loadXML strXML
    
        Set xslSheet = New FreeThreadedDOMDocument40
        
        'CORE291 Peter Edney - 04/08/2006
        If Not xslSheet.Load(strXSLPath) Then
            Err.Raise 40002, "PrepareTemplate.PopulateTemplate", "No xml returned from GetTableDataAsXML(" & strXSLPath & ")"
        End If
        
        Set xslTemplate = New XSLTemplate40
        Set xslTemplate.stylesheet = xslSheet
        Set xslProcessor = xslTemplate.createProcessor()
    
        xslProcessor.input = xmlDataDoc
        x = xslProcessor.Transform()
        
        PopulateTemplate = xslProcessor.output
    End If
TidyUp:
    Set xmlDataDoc = Nothing
    Set xslProcessor = Nothing
    Set xslSheet = Nothing
    Set xslTemplate = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function

PopulateTemplateError:
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    On Error GoTo 0
    PopulateTemplate = ""
    App.LogEvent "Failed in PrepareTemplate.PopulateTemplate " & Err.Description, vbLogEventTypeError
    Resume TidyUp:
End Function

Private Function GetTableDataAsXML(ByVal strQuery As String) As String
Const strFunctionName As String = "GetTableDataAsXML"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim adoStream As ADODB.Stream
Dim connOmiga As Connection
Dim recXML As ADODB.Command
    
    On Error GoTo GetTableDataAsXMLError:
    
    Set connOmiga = New Connection
    connOmiga.ConnectionString = GetOmiga4ConnectionString()
    connOmiga.Open
    
    Set adoStream = New ADODB.Stream
    adoStream.Open
    
    Set recXML = New ADODB.Command
    recXML.ActiveConnection = connOmiga
    recXML.CommandText = strQuery
    recXML.CommandType = adCmdText
    
    recXML.Properties("Output Stream") = adoStream
    recXML.Properties("XML Root") = "TEMPLATEDATA"
    recXML.Execute , , adExecuteStream
    
    GetTableDataAsXML = Replace(adoStream.ReadText(), ">", ">" & vbCrLf)
    
    adoStream.Close
    connOmiga.Close
    
TidyUp:
    Set adoStream = Nothing
    Set recXML = Nothing
    Set connOmiga = Nothing
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function
    
GetTableDataAsXMLError:
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    On Error GoTo 0
    App.LogEvent "Failed in PrepareTemplate.GetTableDataAsXML " & Err.Description, vbLogEventTypeError
    GetTableDataAsXML = "<TEMPLATEDATA/>"
    Resume TidyUp:
End Function

Private Function GetOmiga4ConnectionString() As String
Const strFunctionName As String = "GetOmiga4ConnectionString"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim objWshShell As Object

Dim strConnection As String
Dim strRegSection As String
    
Dim strUserId As String
Dim strPassword As String

Dim lngErrNo As Long
Dim strSource As String
Dim strDescription As String
    
    On Error GoTo BuildDbConnectionStringVbErr
    Set objWshShell = CreateObject("WScript.Shell")
    
    strRegSection = "HKLM\SOFTWARE\OMIGA4\DATABASE CONNECTION\"
    
    On Error Resume Next
            
    strUserId = objWshShell.RegRead(strRegSection & "User ID")
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
            
    On Error GoTo BuildDbConnectionStringVbErr

    If Err.Number <> -2147024894 And Err.Number <> 0 Then
        Err.Raise lngErrNo, strSource, strDescription
    End If

    On Error Resume Next
            
    strPassword = objWshShell.RegRead(strRegSection & "Password")
    
    lngErrNo = Err.Number
    strSource = Err.Source
    strDescription = Err.Description
            
    On Error GoTo BuildDbConnectionStringVbErr

    If Err.Number <> -2147024894 And Err.Number <> 0 Then
        Err.Raise lngErrNo, strSource, strDescription
    End If
    
    strConnection = "Provider=SQLOLEDB;Server=" & objWshShell.RegRead(strRegSection & "Server") & ";" & "database=" & objWshShell.RegRead(strRegSection & "Database Name") & ";"

    ' If User Id is present use SQL Server Authentication else
    ' use integrated security
    If Len(strUserId) > 0 Then
        strConnection = strConnection & "UID=" & strUserId & ";" & "pwd=" & strPassword & ";"
    Else
        strConnection = strConnection & "Integrated Security=SSPI;Persist Security Info=False"
    End If
           
    GetOmiga4ConnectionString = strConnection
    
    Set objWshShell = Nothing
    
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function

BuildDbConnectionStringVbErr:
    
    Set objWshShell = Nothing
    
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function




Private Function GetFileData(ByVal strPath As String) As String
Const strFunctionName As String = "GetFileData"
gobjTrace.TraceMethodEntry TypeName(Me), strFunctionName

Dim F As Integer
Dim strWork As String

    On Error GoTo GetFileDataError:
    strWork = Space$(FileLen(strPath))
    F = FreeFile()
    Open strPath For Binary Access Read As F
    Get F, , strWork
    Close F
    GetFileData = strWork
    
TidyUp:
    gobjTrace.TraceMethodExit TypeName(Me), strFunctionName
    Exit Function
    
GetFileDataError:
    gobjTrace.TraceMethodError TypeName(Me), strFunctionName
    On Error GoTo 0
    App.LogEvent "Failed in PrepareTemplate.GetFileData " & Err.Description, vbLogEventTypeError
    Resume TidyUp:
End Function

Public Sub InitialiseTraceInterface( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStartElapsed As Single)
    
    If Not vblnTraceOn Then
        Exit Sub
    End If
    
    gobjTrace.TraceInitialiseFromParent vblnTraceOn, vstrFileName, vintStartElapsed
End Sub

