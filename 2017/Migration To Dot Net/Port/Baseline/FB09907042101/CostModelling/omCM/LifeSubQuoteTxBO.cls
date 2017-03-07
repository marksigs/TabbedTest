VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "LifeSubQuoteTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LifeSubQuoteTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for LifeSubQuote Business Object which requires transactions
'               contains (implied) standard interface:
'Dependencies:  LifeSubQuoteDO,  LifeSubQuoteBO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     29/11/1999  Created
'PSC    20/11/1999  Amend to return the created sub quote number
'PSC    01/02/2000  Added SaveRecostedBenefits()
'APS    30/05/2000  SYS0797 - Removed reference to MortgageSubQuoteNumber when creating a LifeSubQuote
'                   see DB Change 242
'MS     19/07/00    Modified for performance enhancements
'------------------------------------------------------------------------------------------

Option Explicit
Implements ILifeSubQuoteTxBO
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "LIFESUBQUOTE"
Private Const cstrTABLE_LIFEBENEFIT_NAME = "LIFEBENEFIT"
 
Public Function Create(ByVal vstrXMLRequest As String, _
                  Optional ByVal strSubQuoteCreateType As String = "NEW") As String
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object - Table (LifeSubQuote)
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Dim objLifeSubQuoteDO As LifeSubQuoteDO

    Dim strXmlOut As String
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    ' If we have more than one in the list then we need a list of keys
    If xmlNodeList.length > 1 Then
        Set xmlListElem = xmlOut.createElement("LIFESUBQUOTELIST")
        xmlOut.appendChild xmlListElem
    End If
   
    If objContext Is Nothing Then
        Set objLifeSubQuoteDO = New LifeSubQuoteDO
    Else
        Set objLifeSubQuoteDO = objContext.CreateInstance(App.Title & ".LifeSubQuoteDO")
    End If
    
    Dim intNodeIndex As Integer
    
    ' For each relevant node add key and call the data object create
    Dim strXML As String
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        strXML = BuildInXmlForCreate(objElement, xmlSubQuoteElem, strSubQuoteCreateType)
        objLifeSubQuoteDO.Create strXML
        
        If xmlListElem Is Nothing Then
            xmlOut.appendChild xmlSubQuoteElem
        Else
            xmlListElem.appendChild xmlSubQuoteElem
        End If
    Next
    
    Create = xmlOut.xml
 
    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlOut = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

CreateVbErr:
   
    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlListElem = Nothing
    Set xmlSubQuoteElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Sub UpdateLifeBenefit(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing update data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeBenefitVbErr

    Const strFunctionName As String = "UpdateLifeBenefit"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim xmlUpdateElement As IXMLDOMElement

    Dim objLifeSubQuoteDO As LifeSubQuoteDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrTABLE_LIFEBENEFIT_NAME)

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_LIFEBENEFIT_NAME & " tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objLifeSubQuoteDO = New LifeSubQuoteDO
    Else
        Set objLifeSubQuoteDO = objContext.CreateInstance(App.Title & ".LifeSubQuoteDO")
    End If

    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objLifeSubQuoteDO.UpdateLifeBenefit objElement.xml
    Next

    Set xmlUpdateElement = Nothing
    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

UpdateLifeBenefitVbErr:

    Set xmlUpdateElement = Nothing
    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
'
Public Sub DeleteLifeBenefit(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with the
'               table - 'LifeBenefit'
'               invokes appropriate Data Object(s) to perform delete
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeBenefitVbErr

    Const strFunctionName As String = "DeleteLifeBenefit"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    
    Dim objLifeSubQuoteDO As LifeSubQuoteDO

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName(cstrTABLE_LIFEBENEFIT_NAME)

    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_LIFEBENEFIT_NAME & " tag not found"
    End If

    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objLifeSubQuoteDO = New LifeSubQuoteDO
    Else
        Set objLifeSubQuoteDO = objContext.CreateInstance(App.Title & ".LifeSubQuoteDO")
    End If

    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objLifeSubQuoteDO.DeleteLifeBenefit objElement.xml
    Next

    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteLifeBenefitVbErr:

    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function BuildInXmlForCreate(ByVal xmlInElement As IXMLDOMElement, _
                                 ByRef xmlSubQuoteNo As IXMLDOMElement, _
                                 ByVal strCreateSubQuoteType As String) As String
' header ----------------------------------------------------------------------------------
' description:  Modifes the Xml string received from the BO before invoking method in DO
'               Also adds the new LifeSubQuoteNumber to strXmlOut
' pass:         objElement : the node contatining the data required to build the Input to DO
'               strCreateSubQuoteType : 'New' subquote or 'First' SubQuote
' return:       XML string required as Input to 'Create' method in DO
'------------------------------------------------------------------------------------------
On Error GoTo BuildInXmlForCreateVbErr

    Const strFunctionName As String = "BuildInXmlForCreate"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
    Dim objDOAssist As DOAssist
    Dim objErrAssist As New ErrAssist
    
    Dim xmlCreateDoc As New FreeThreadedDOMDocument40
    xmlCreateDoc.validateOnParse = False
    xmlCreateDoc.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim strSystemDate As String, strApplicationNumber As String
    Dim strAppliactionFactFindNumber As String, strQuotationType As String
    
    If objContext Is Nothing Then
            Set objDOAssist = New DOAssist
    Else
            Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'NEW FUNCTION TO BE WRITTEN IN ADOASSIST THAT RETURNS DATE/TIME/DATETIME
    'strSystemDate = objSqlAssist.GetSystemDate
    strSystemDate = Format(Now, "DD/MM/YYYY")
    
    'Fetch values from the xmlElement to local variables
    strApplicationNumber = objXmlAssist.GetTagValue(xmlInElement, "APPLICATIONNUMBER")
    strAppliactionFactFindNumber = _
                objXmlAssist.GetTagValue(xmlInElement, "APPLICATIONFACTFINDNUMBER")
    strQuotationType = objXmlAssist.GetTagValue(xmlInElement, "QUOTATIONTYPE")
    
    Set xmlTableElem = xmlCreateDoc.createElement("LIFESUBQUOTE")
    xmlCreateDoc.appendChild xmlTableElem
    
    Set xmlElement = xmlCreateDoc.createElement("APPLICATIONNUMBER")
    xmlTableElem.appendChild xmlElement
    xmlElement.Text = strApplicationNumber
    
    Set xmlElement = xmlCreateDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTableElem.appendChild xmlElement
    xmlElement.Text = strAppliactionFactFindNumber
    
'    Set xmlElement = xmlCreateDoc.createElement("MSQAPPLICATIONNUMBER")
'    xmlTableElem.appendChild xmlElement
'    xmlElement.Text = strApplicationNumber
    
'    Set xmlElement = xmlCreateDoc.createElement("MSQAPPLICATIONFACTFINDNUMBER")
'    xmlTableElem.appendChild xmlElement
'    xmlElement.Text = strAppliactionFactFindNumber
      
    Set xmlElement = xmlCreateDoc.createElement("DATEANDTIMEGENERATED")
    xmlTableElem.appendChild xmlElement
    xmlElement.Text = strSystemDate
    
    Set xmlElement = xmlCreateDoc.createElement("QUOTATIONTYPE")
    xmlTableElem.appendChild xmlElement
    xmlElement.Text = strQuotationType
        
    'Find values of LifeSubQuoteNumber
    Dim intLifeSubQuoteNumber As Integer ', intMortgageSubQuoteNumber As Integer
    Dim strLifeSubQuoteDef As String
    
    ' APS SYS0767 - Duplicate key error following new sub-quote when previous one was invalid.
    'If strCreateSubQuoteType = "FIRST" Then
    '    intLifeSubQuoteNumber = 1
'        intMortgageSubQuoteNumber = 1
    'Else
        
        Dim objIomCMClassDef As omCMClassDef
        
        Set objIomCMClassDef = New omCMClassDef
                
        strLifeSubQuoteDef = objIomCMClassDef.LoadLifeSubQuoteData
        
        intLifeSubQuoteNumber = CInt(objDOAssist.GetNextSequenceNumber(xmlInElement.xml, _
                                strLifeSubQuoteDef, "LIFESUBQUOTENUMBER"))
' APS SYS0797 - Remove reference to MortgageSubQuoteNumber field when creating a life quote
'        intMortgageSubQuoteNumber = _
'                    CInt(objXmlAssist.GetTagValue(xmlInElement, "MORTGAGESUBQUOTENUMBER"))
    'End If
        
    Set xmlElement = xmlCreateDoc.createElement("LIFESUBQUOTENUMBER")
    xmlTableElem.appendChild xmlElement
    xmlElement.Text = CStr(intLifeSubQuoteNumber)
    
    Set xmlSubQuoteNo = xmlTableElem.cloneNode(False)
    xmlSubQuoteNo.appendChild xmlElement.cloneNode(True)
    
'    Set xmlElement = xmlCreateDoc.createElement("MORTGAGESUBQUOTENUMBER")
'    xmlTableElem.appendChild xmlElement
'    xmlElement.Text = CStr(intMortgageSubQuoteNumber)
    
    BuildInXmlForCreate = xmlCreateDoc.xml
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objErrAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set xmlCreateDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function
    
BuildInXmlForCreateVbErr:
    
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set objSQLAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set xmlCreateDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If

    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function
'MS 19/07/00 Modified for performance enhancements
Private Function ILifeSubQuoteTxBO_SaveRecostedBenefits(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Deletes existing life benefits then creates the new ones
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveRecostedBenefitsVbErr

    Const strFunctionName As String = "ILifeSubQuoteTxBO_SaveRecostedBenefits"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objILifeSubQuoteDO As ILifeSubQuoteDO
'MS 19/07/00 Modified for performance enhancements
'    Dim objLifeSubQuoteDO As LifeSubQuoteDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vXMLRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objILifeSubQuoteDO = New LifeSubQuoteDO
    Else
        Set objILifeSubQuoteDO = objContext.CreateInstance(App.Title & ".LifeSubQuoteDO")
    End If
    
    objILifeSubQuoteDO.DeleteAllLifeBenefits xmlElement
'MS 19/07/00 Modified for performance enhancements
'    Set objLifeSubQuoteDO = objILifeSubQuoteDO
'    objLifeSubQuoteDO.SaveDetails xmlElement.xml
    objILifeSubQuoteDO.SaveDetails xmlElement
    
    Set ILifeSubQuoteTxBO_SaveRecostedBenefits = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveRecostedBenefitsExit:
    
    Set objILifeSubQuoteDO = Nothing
'MS 19/07/00 Modified for performance enhancements
'    Set objLifeSubQuoteDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveRecostedBenefitsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set ILifeSubQuoteTxBO_SaveRecostedBenefits = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveRecostedBenefitsExit

End Function

