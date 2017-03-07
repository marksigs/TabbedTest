VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ImportBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ImportBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Dependencies:  List any other dependent components
'               e.g. ImportTxBO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'SR     23/03/00    Creation
'SR     03/05/00    New Method 'CreateApplication'
'SR     08/05/00    New Method 'Update Application'
'------------------------------------------------------------------------------------------

Option Explicit

Implements IImportBO
Implements ObjectControl

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Public Function ImportApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ImportVbErr

    Const strFunctionName As String = "ImportApplication"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                               
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IImportBO_ImportApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ImportApplication = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ImportExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ImportVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ImportApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ImportExit

End Function

Private Function IImportBO_ImportApplication( _
        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ImportVbErr
    
    Const strFunctionName As String = "IImportBO_ImportApplication"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim intApplStage As Integer
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    intApplStage = vxmlRequest.selectSingleNode(".//STAGENUMBER").Text
    
    Select Case intApplStage
        Case 30  'AIP
            'Create Application
            Set xmlTempResponseNode = IImportBO_CreateApplication(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case 50  'Mortgage Application
            Set xmlTempResponseNode = IImportBO_UpdateApplication(vxmlRequest)
            m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        Case Else 'Invalid
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, oeInvalidParameter
    End Select
    
        
    Set IImportBO_ImportApplication = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ImportExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ImportVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportBO_ImportApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ImportExit

End Function

Public Function CreateApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'Input Format
'<REQUEST USERID=""USER0001"" UNITID=""UNIT1"" MACHINEID=""MSGETC"">
'   <CUSTOMERLIST>
'      <CUSTOMER>
'        <CUSTOMERNUMBER/>
'        <OTHERSYSTEMCUSTOMERNUMBER>YYYY</OTHERSYSTEMCUSTOMERNUMBER>
'        <CUSTOMERVERSION>
'            <CUSTOMERNUMBER/>
'            <CUSTOMERVERSIONNUMBER/>
'            <CORRESPONDENCESALUTATION/>
'            .....
'            .....
'
'            <CUSTOMERROLETYPE>1</CUSTOMERROLETYPE>
'            <CUSTOMERORDER>5</CUSTOMERORDER>
'
'            <CUSTOMERADDRESSLIST>
'               <CUSTOMERADDRESS>
'                   <CUSTOMERNUMBER/>
'                   ......
'                   <ADDRESS>
'                       <ADDRESSGUID/>
'                       <BUILDINGORHOUSENUMBER>45</BUILDINGORHOUSENUMBER>
'                       ...........
'                   </ADDRESS>
'                </CUSTOMERADDRESS>
'            </CUSTOMERADDRESSLIST>
'
'            <CUSTOMERTELEPHONENUMBERLIST>
'                  <CUSTOMERTELEPHONENUMBER>
'                       ...........
'                   </CUSTOMERTELEPHONENUMBER>
'            </CUSTOMERTELEPHONENUMBERLIST>
'
'            <AREASOFINTERESTLIST>
'               <AREASOFINTEREST>
'                   <CUSTOMERNUMBER/>
'                   ......
'                   <INTERESTAREA>2</INTERESTAREA>"
'               </AREASOFINTEREST>
'            </AREASOFINTERESTLIST>
'
'            <ALIASPERSONLIST>
'               <ALIAS>
'                   <CUSTOMERNUMBER></CUSTOMERNUMBER>
'                   <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'                   <ALIASSEQUENCENUMBER></ALIASSEQUENCENUMBER>
'                   .............
'                   <PERSON>
'                       <PERSONGUID></PERSONGUID>
'                       <DATEOFBIRTH></DATEOFBIRTH>
'                       ........
'                   </PERSON>
'               </ALIAS>
'           </ALIASPERSONLIST>
'
'            <EMPLOYMENTLIST>
'               <EMPLOYMENT>
'                   ........
'               </EMPLOYMENT>
'            </EMPLOYMENTLIST>
'
'            <UNEARNEDINCOMELIST>
'               <UNEARNEDINCOME>
'                   ........
'               </UNEARNEDINCOME>
'            </UNEARNEDINCOMELIST>
'
'            <LOANSLIABILITIESLIST>
'               <LOANSLIABILITIES>
'                   .................
'               </LOANSLIABILITIES>
'            </LOANSLIABILITIESLIST>
'
'            <BANKCREDITCARDLIST>
'               <BANKCREDITCARD>
'                   .............
'               </BANKCREDITCARD>
'            </BANKCREDITCARDLIST>
'        </CUSTOMERVERSION>
'      </CUSTOMER>
'    </CUSTOMERLIST>
'
'    <APPLICATION>
''       <APPLICATIONNUMBER></APPLICATIONNUMBER>
'        <PACKAGENUMBER>PACK0002</PACKAGENUMBER>
'        ..............
'
'        <TYPEOFAPPLICATION>1</TYPEOFAPPLICATION>
'        <APPLICATIONVERIFICATION>
'              ...............
'         </APPLICATIONVERIFICATION>
'        <FINANCIALSUMMARY>
'           ...........
'        </FINANCIALSUMMARY>
'        <NEWPROPERTY>
'           .........
'        </NEWPROPERTY>
'        <NEWPROPERTYADDRESS>
'           ............
'            <ADDRESS>
'                <ADDRESSGUID></ADDRESSGUID>
'                ............
'            </ADDRESS>
'        </NEWPROPERTYADDRESS>
'    </APPLICATION>
' </REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CreateApplication"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                               
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IImportBO_CreateApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateApplication = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Public Function UpdateApplication(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'-----------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "UpdateApplication"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                               
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IImportBO_UpdateApplication(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateApplication = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateApplication = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IImportBO_CreateApplication( _
        ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    
    Const strFunctionName As String = "IImportBO_CreateApplication"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIImportTxBO As IImportTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIImportTxBO = New ImportTxBO
    Else
        Set objIImportTxBO = m_objContext.CreateInstance(gstrIMPORT_COMPONENT & ".ImportTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIImportTxBO.CreateApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IImportBO_CreateApplication = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIImportTxBO = Nothing
    
    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportBO_CreateApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IImportBO_UpdateApplication( _
            ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    
    Const strFunctionName As String = "IImportBO_UpdateApplication"

    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIImportTxBO As IImportTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIImportTxBO = New ImportTxBO
    Else
        Set objIImportTxBO = m_objContext.CreateInstance(gstrIMPORT_COMPONENT & ".ImportTxBO")
    End If
        
    '   call Business Transaction Object Create function
    Set xmlTempResponseNode = objIImportTxBO.UpdateApplication(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IImportBO_UpdateApplication = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateExit:
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objIImportTxBO = Nothing
    
    Exit Function

UpdateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IImportBO_UpdateApplication = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

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

