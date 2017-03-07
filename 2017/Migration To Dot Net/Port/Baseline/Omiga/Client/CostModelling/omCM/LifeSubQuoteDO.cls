VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LifeSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'Workfile:      LifeSubQuote.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Dependencies:
'               LifeSubQuoteBO
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date      Description
'SR     26/11/99  Created
'PSC    20/11/1999  Amend to correct setting of classdef document
'PSC    01/02/2000  Added DeleteAllLifeBenefits()
'SR     25/04/2000 New method 'GetData'
'JLD    18/05/00    SYS0725 fixes
'MS     19/07/00    Modified for performance enhancements
'MS     01/08/00    SYS1291
'GD     22/04/02    SYS3092 Only primary keys were being populated in LifeBenefit table. Altered to that all fields now pop'd
'------------------------------------------------------------------------------------------

Option Explicit
Implements ILifeSubQuoteDO

Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request - Table (LifeSubQuote)
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------

On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    Dim strXmlOut As String
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
       
    'FIX ME when change to DOM DOC
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
       
    objDOAssist.Create vstrXMLRequest, objIomCMClassDef.LoadLifeSubQuoteData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:

    Set objDOAssist = Nothing
     
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub DeleteLifeBenefit(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with the table - 'LifeBenefit'
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLifeBenefitVbErr

    Const strFunctionName As String = "DeleteLifeBenefit"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objDOAssist As DOAssist
    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
            
    objDOAssist.Delete vstrXMLRequest, objIomCMClassDef.LoadLifeBenefitData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

DeleteLifeBenefitVbErr:

    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub UpdateLifeBenefit(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with the table 'LifeBenefit'
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLifeBenefitVbErr

    Const strFunctionName As String = "UpdateLifeBenefit"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objDOAssist As DOAssist
       
    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
    
    objDOAssist.Update vstrXMLRequest, objIomCMClassDef.LoadLifeBenefitData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateLifeBenefitVbErr:
    
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub BuildXmlForUpdateLifeBenefit(ByVal xmlLifeBenefitElem As IXMLDOMElement, _
                                    ByVal xmlKeyElem As IXMLDOMElement, _
                                    ByVal strCustomerNumber1 As String, _
                                    ByVal strCustomerVersionNumber1 As String, _
                                    ByVal strCustomerNumber2 As String, _
                                    ByVal strCustomerVersionNumber2 As String)

' header ----------------------------------------------------------------------------------
' description:
'   Builds the XML strings required for Updation of table 'LifeBenefit' ;
'   used in the method 'SaveDetails'
' pass:
'
'Return : Null
' Error :
'-------------------------------------------------------------------------------------------
On Error GoTo BuildXmlForUpdateLifeBenefitVbErr
    
    Const strFunctionName As String = "BuildXmlForUpdateLifeBenefit"
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim xmlCustomerElem As IXMLDOMElement
    Dim xmlCustomerVersionElem As IXMLDOMElement
        
    Dim strApplicants As String
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
    'Assign values to CustomerNumber1, CustomerVersionNumber1. If element are not available
    ' in the input node, create before assigning values.
    Set xmlCustomerElem = xmlLifeBenefitElem.getElementsByTagName("CUSTOMERNUMBER1").Item(0)
    If xmlCustomerElem Is Nothing Then
        Set xmlCustomerElem = xmlIn.createElement("CUSTOMERNUMBER1")
        xmlLifeBenefitElem.appendChild xmlCustomerElem
    End If
    xmlCustomerElem.Text = strCustomerNumber1
    
    Set xmlCustomerVersionElem = _
                xmlLifeBenefitElem.getElementsByTagName("CUSTOMERVERSIONNUMBER1").Item(0)
     If xmlCustomerVersionElem Is Nothing Then
        Set xmlCustomerVersionElem = xmlIn.createElement("CUSTOMERVERSIONNUMBER1")
        xmlLifeBenefitElem.appendChild xmlCustomerVersionElem
    End If
    xmlCustomerVersionElem.Text = strCustomerVersionNumber1
    
    'Assign values to CustomerNumber2, CustomerVersionNumber2. Assign null if value of
    'Applicants tag is not 'Joint'
    strApplicants = xmlLifeBenefitElem.getElementsByTagName("APPLICANTS").Item(0).Text
    If strApplicants = "" Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, oeInvalidParameter, _
                                " For APPLICANTS in table LIFEBENEFIT"
    End If
    
    Set xmlCustomerElem = xmlLifeBenefitElem.getElementsByTagName("CUSTOMERNUMBER2").Item(0)
    If xmlCustomerElem Is Nothing Then
        Set xmlCustomerElem = xmlIn.createElement("CUSTOMERNUMBER2")
        xmlLifeBenefitElem.appendChild xmlCustomerElem
    End If
    
    Set xmlCustomerVersionElem = _
                xmlLifeBenefitElem.getElementsByTagName("CUSTOMERVERSIONNUMBER2").Item(0)
     If xmlCustomerVersionElem Is Nothing Then
        Set xmlCustomerVersionElem = xmlIn.createElement("CUSTOMERVERSIONNUMBER2")
        xmlLifeBenefitElem.appendChild xmlCustomerVersionElem
    End If
    
    If UCase(strApplicants) = "APPLICANT1" Then
        xmlCustomerElem.Text = ""
        xmlCustomerVersionElem.Text = ""
    Else
        xmlCustomerElem.Text = strCustomerNumber2
        xmlCustomerVersionElem.Text = strCustomerVersionNumber2
    End If
    
    Set xmlIn = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlCustomerVersionElem = Nothing
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Exit Sub
BuildXmlForUpdateLifeBenefitVbErr:
    
    Set xmlIn = Nothing
    Set xmlCustomerElem = Nothing
    Set xmlCustomerVersionElem = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
'MS 19/07/00 Modified for performance enhancements
Private Sub GetValuesFromXmlInForSaveDetails(ByVal vXMLRequest As IXMLDOMElement, _
                                    ByRef strApplicationNumber As String, _
                                    ByRef strApplicationFactFindNumber As String, _
                                    ByRef strLifeSubQuoteNumber As String, _
                                    ByRef strCustomerNumber1 As String, _
                                    ByRef strCustomerVersionNumber1 As String, _
                                    ByRef strCustomerNumber2 As String, _
                                    ByRef strCustomerVersionNumber2 As String, _
                                    ByRef strTotalLifeMonthlyCost As String)
' header ----------------------------------------------------------------------------------
' description:
'   Fetches values from iput xml (of SaveDetails mthod)  into local variables
' pass:
'   xmlRequest to the mehtod - SaveDetails
'   other local variables for ApplicationNumber, ApplicationFactFindNumber,
'   LifeSubQuoteNumber,CustomerNumber1,CustomerVersionNumber1,
'   CustomerNumber2, CustomerVersionNumber2
'Return : Null
' Error :
'-------------------------------------------------------------------------------------------
    
    Const strFunctionName As String = "GetValuesFromXmlInForSaveDetails"
    
    Dim objXmlAssist As New XMLAssist
    
'    Dim xmlDoc As DOMDocument
        
'    Set xmlDoc = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    strApplicationNumber = objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = _
                objXmlAssist.GetTagValue(vXMLRequest, "APPLICATIONFACTFINDNUMBER")
    strLifeSubQuoteNumber = _
                objXmlAssist.GetTagValue(vXMLRequest, "LIFESUBQUOTENUMBER")
    
    strCustomerNumber1 = objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERNUMBER1")
    strCustomerVersionNumber1 = _
            objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERVERSIONNUMBER1")
    strCustomerNumber2 = objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERNUMBER2")
    strCustomerVersionNumber2 = _
            objXmlAssist.GetTagValue(vXMLRequest, "CUSTOMERVERSIONNUMBER2")
            
    strTotalLifeMonthlyCost = _
            objXmlAssist.GetTagValue(vXMLRequest, "TOTALLIFEMONTHLYCOST")
            
End Sub

Private Sub ILifeSubQuoteDO_DeleteAllLifeBenefits(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the life benefits for a life subquote
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllLifeBenefitsVbErr

    Const strFunctionName As String = "ILifeSubQuoteDO_DeleteAllLifeBenefits"
    
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objXmlAssist As New XMLAssist
    Dim xmlRootElem As IXMLDOMElement
    Dim xmlInRootElem As IXMLDOMElement
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    If vxmlTableElement.nodeName = "LIFESUBQUOTE" Then
        Set xmlInRootElem = vxmlTableElement
    Else
        Set xmlInRootElem = vxmlTableElement.getElementsByTagName("LIFESUBQUOTE").Item(0)
    End If
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strSubQuoteNo As String
    
    strAppNo = objXmlAssist.GetTagValue(xmlInRootElem, "APPLICATIONNUMBER")
    strAppFFNo = objXmlAssist.GetTagValue(xmlInRootElem, "APPLICATIONFACTFINDNUMBER")
    strSubQuoteNo = objXmlAssist.GetTagValue(xmlInRootElem, "LIFESUBQUOTENUMBER")
    
    If Len(strAppNo) = 0 Or Len(strAppFFNo) = 0 Or Len(strSubQuoteNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "APPLICATIONNUMBER = " & strAppNo & _
                                "; APPLICATIONFACTFINDNUMBER = " & strAppFFNo & _
                                "; LIFESUBQUOTENUMBER = " & strSubQuoteNo
    End If
    
    Set xmlRootElem = xmlRequestDoc.createElement("LIFEBENEFIT")
    xmlRequestDoc.appendChild xmlRootElem
    
    Set xmlElement = xmlInRootElem.getElementsByTagName("APPLICATIONNUMBER").Item(0)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlRootElem.appendChild xmlElement
    
    Set xmlElement = xmlInRootElem.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlRootElem.appendChild xmlElement
          
    Set xmlElement = xmlInRootElem.getElementsByTagName("LIFESUBQUOTENUMBER").Item(0)
    Set xmlElement = xmlElement.cloneNode(True)
    xmlRootElem.appendChild xmlElement
           
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadLifeBenefitData
    
    Dim blnRaiseError As Boolean
    blnRaiseError = False
    objIDOAssist.DeleteAll xmlRootElem, xmlClassDefDoc
    blnRaiseError = True
 
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteAllLifeBenefitsVbErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffectedByDeleteAll Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function ILifeSubQuoteDO_GetData(ByVal vXMLRequest As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "ILifeSubQuoteDO_GetData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objXmlAssist As New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadLifeSubQuoteData
    Set ILifeSubQuoteDO_GetData = objIDOAssist.GetData(vXMLRequest, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'MS 19/07/00 Modified for performance enhancements
Private Sub ILifeSubQuoteDO_SaveDetails(ByVal vXMLRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with the table 'LifeBenefit'
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'   Input Format:
'<LIFESUBQUOTELIFEBENEFIT>
'   <APPLICATIONNUMBER> ... </APPLICATIONNUMBER>
'   <APPLICATIONFACTFINDNUMBER> ... </APPLICATIONFACTFINDNUMBER>
'   <LIFESUBQUOTENUMBER> ... </LIFESUBQUOTENUMBER>
'   <CUSTOMERNUMBER1> ... </CUSTOMERNUMBER1>
'   <CUSTOMERVERSIONNUMBER1> ... </CUSTOMERVERSIONNUMBER1>
'   <CUSTOMERNUMBER2> ... </CUSTOMERNUMBER2>
'   <CUSTOMERVERSIONNUMBER2> ... </CUSTOMERVERSIONNUMBER2>
'   <LIFEBENEFITS>
'           <LIFEBENEFIT>
'               <APPLICANTS> ... </APPLICANTS>
'               <BENEFITTYPE> ... </BENEFITTYPE>
'                <IPT> ... </IPT>
'                <MONTHLYCOST> ... </MONTHLYCOST>
'           </LIFEBENEFIT>
'           <LIFEBENEFIT>
'                <APPLICANTS> ... </APPLICANTS>
'                <BENEFITTYPE> ... </BENEFITTYPE>
'                <IPT> ... </IPT>
'                <MONTHLYCOST> ... </MONTHLYCOST>
'           </LIFEBENEFIT>
'           .
'           .
'           <TOTALLIFEMONTHLYCOST>...</TOTALLIFEMONTHLYCOST>
'    </LIFEBENEFITS>
'</LIFESUBQUOTELIFEBENEFIT>
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ILifeSubQuoteDO_SaveDetailsVbErr

    Const strFunctionName As String = "ILifeSubQuoteDO_SaveDetails"
        
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    
'    Dim xmlIn As DOMDocument
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    'Store the PK values and TotalLifeMonthlyCost, Customer numbers in local variables
    Dim strApplicationNumber As String, strApplicationFactFindNumber As String
    Dim strLifeSubQuoteNumber As String, strTotalLifeMonthlyCost As String
    
    Dim strCustomerNumber1 As String, strCustomerNumber2 As String
    Dim strCustomerVersionNumber1 As String, strCustomerVersionNumber2 As String
    
    Call GetValuesFromXmlInForSaveDetails(vXMLRequest, strApplicationNumber, _
                            strApplicationFactFindNumber, strLifeSubQuoteNumber, _
                            strCustomerNumber1, strCustomerVersionNumber1, _
                            strCustomerNumber2, strCustomerVersionNumber2, _
                            strTotalLifeMonthlyCost)
    
    If strApplicationNumber = "" Or strApplicationFactFindNumber = "" Or _
                                            strLifeSubQuoteNumber = "" Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                oeInvalidParameter, " For table LIFESUBQUOTE "
    End If
                    
    
    'Build XML Node with Key Elements of the table 'LIFEBENEFIT'
    Dim strKeyXml As String, xmlKeyElem As IXMLDOMElement
    
'    strKeyXml = "<KEYELEMENTS>"
     strKeyXml = "<LIFESUBQUOTE>"
    strKeyXml = strKeyXml & "<APPLICATIONNUMBER>" & strApplicationNumber & "</APPLICATIONNUMBER>"
    strKeyXml = strKeyXml & "<APPLICATIONFACTFINDNUMBER>" & strApplicationFactFindNumber _
                          & "</APPLICATIONFACTFINDNUMBER>"
    strKeyXml = strKeyXml & "<LIFESUBQUOTENUMBER>" & strLifeSubQuoteNumber & "</LIFESUBQUOTENUMBER>"
    'strKeyXml = strKeyXml & "<LIFEBENEFITSEQUENCENUMBER/>"
'    strKeyXml = strKeyXml & "</KEYELEMENTS>"
    
'    Set xmlDoc = objXmlAssist.load(strKeyXml, TypeName(Me), strFunctionName)
'    Set xmlKeyElem = xmlDoc.getElementsByTagName("KEYELEMENTS").Item(0)

    'Update the table 'LifeSubQuote' with TotalLifeMonthlyCost
    Dim strXML As String
    
'    strXML = "<LIFESUBQUOTE>" & strKeyXml
    strXML = strKeyXml & "<TOTALLIFEMONTHLYCOST>" & strTotalLifeMonthlyCost & "</TOTALLIFEMONTHLYCOST>"
    strXML = strXML & "</LIFESUBQUOTE>"
    
    Dim xmlUpdateDoc As FreeThreadedDOMDocument40
    Set xmlUpdateDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    'Update the SubQuote
    ILifeSubQuoteDO_Update xmlUpdateDoc.documentElement
    
    objXmlAssist.ChangeNodeName xmlUpdateDoc, "LIFESUBQUOTE", "KEYELEMENTS"
    Set xmlKeyElem = xmlUpdateDoc.getElementsByTagName("KEYELEMENTS").Item(0)
    ' Get a list of nodes that need to be passed down to the data object
'    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
'    Set xmlNodeList = xmlIn.getElementsByTagName("LIFEBENEFIT")
    Set xmlNodeList = vXMLRequest.getElementsByTagName("LIFEBENEFIT")
    
    ' For each relevant node, build the required xml string and update the table 'LifeBenefit'
    Dim intNodeIndex  As Integer
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        'MS     01/08/00    SYS1291
        objElement.appendChild xmlKeyElem
        'objElement.appendChild xmlUpdateDoc.documentElement
        Call BuildXmlForUpdateLifeBenefit(objElement, xmlKeyElem, strCustomerNumber1, _
                                          strCustomerVersionNumber1, strCustomerNumber2, _
                                          strCustomerVersionNumber2)
        ILifeSubQuoteDO_CreateLifeBenefit objElement
    Next
    
    ' Clear the objects
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    
'    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlKeyElem = Nothing
'    Set xmlDoc = Nothing
    
    Exit Sub
    
ILifeSubQuoteDO_SaveDetailsVbErr:
    
    Set objXmlAssist = Nothing
    
'    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    Set xmlKeyElem = Nothing
'    Set xmlDoc = Nothing
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub
'MS 19/07/00 Modified for performance enhancements
'Public Function CreateLifeBenefit(ByVal vstrXMLRequest As String) As String
Private Function ILifeSubQuoteDO_CreateLifeBenefit(ByVal vXMLRequest As IXMLDOMElement) 'As IXMLDOMElement

' header ----------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with the table 'LifeBenefit'
'   for each set of data in the request - Table (LifeSubQuote)
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------

On Error GoTo ILifeSubQuoteDO_CreateLifeBenefitVbErr

    Const strFunctionName As String = "ILifeSubQuoteDO_CreateLifeBenefit"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
'    Dim xmlIn As DOMDocument
    Dim xmlElement As IXMLDOMElement
    Dim strXmlOut As String
    Dim intNewSequenceNumber As Integer
'    Dim strLifeBenefitDef As String
    Dim xmlTmpElem As IXMLDOMElement
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
'    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
        
    ' Generate the next value for LifeSubQuoteNumber
    Dim objIomCMClassDef As IomCMClassDef
    Set objIomCMClassDef = New omCMClassDef
    Dim xmlLifeBenefitDef As FreeThreadedDOMDocument40
    
    Set xmlLifeBenefitDef = objIomCMClassDef.LoadLifeBenefitData
    
    intNewSequenceNumber = objIDOAssist.GetNextSequenceNumber(vXMLRequest, xmlLifeBenefitDef, _
                                        "LIFEBENEFIT", "LIFEBENEFITSEQUENCENUMBER")
    Dim xmlTempElement As IXMLDOMElement
    

    
    'Assign the sequence number to xmlIn
    Set xmlElement = vXMLRequest.getElementsByTagName("LIFEBENEFITSEQUENCENUMBER").Item(0)
    
    If xmlElement Is Nothing Then
                
        Set xmlTempElement = vXMLRequest.getElementsByTagName("KEYELEMENTS").Item(0)
        'Create it
        Set xmlTmpElem = xmlTempElement.ownerDocument.createElement("LIFEBENEFITSEQUENCENUMBER")
        xmlTmpElem.Text = intNewSequenceNumber
        xmlTempElement.appendChild xmlTmpElem
    Else
        xmlElement.Text = intNewSequenceNumber
    End If
    
    'Build xml string to be returned by this function
    strXmlOut = "<LIFEBENEFIT>"
    strXmlOut = strXmlOut & "<LIFEBENEFITSEQUENCENUMBER>" & intNewSequenceNumber _
                          & "</LIFEBENEFITSEQUENCENUMBER>"
    strXmlOut = strXmlOut & "</LIFEBENEFIT>"
    
    'GD SYS3092 22/04/2002 ----------------------start--------------------------
    'Temp node used to hold each primary Key
    Dim xmlPrimaryKey As IXMLDOMElement
    
    'Clone the request for use with GetMandatoryNode, because it removes the node
    Dim xmlRequestClone As IXMLDOMElement
    
    'xmlTempElement is our starting point for the xml that is passed into Create()
    Set xmlTempElement = vXMLRequest.cloneNode(True)
    Set xmlRequestClone = vXMLRequest.cloneNode(True)
    'APPLICATIONNUMBER
    Set xmlPrimaryKey = objXmlAssist.GetMandatoryNode(xmlRequestClone, "//KEYELEMENTS/APPLICATIONNUMBER")
    xmlTempElement.appendChild xmlPrimaryKey
    
    'APPLICATIONFACTFINDNUMBER
    Set xmlPrimaryKey = objXmlAssist.GetMandatoryNode(xmlRequestClone, "//KEYELEMENTS/APPLICATIONFACTFINDNUMBER")
    xmlTempElement.appendChild xmlPrimaryKey
    
    'LIFESUBQUOTENUMBER
    Set xmlPrimaryKey = objXmlAssist.GetMandatoryNode(xmlRequestClone, "//KEYELEMENTS/LIFESUBQUOTENUMBER")
    xmlTempElement.appendChild xmlPrimaryKey
    
    'LIFEBENEFITSEQUENCENUMBER
    Set xmlPrimaryKey = objXmlAssist.GetMandatoryNode(xmlRequestClone, "//KEYELEMENTS/LIFEBENEFITSEQUENCENUMBER")
    xmlTempElement.appendChild xmlPrimaryKey
    
    Set xmlTmpElem = xmlTempElement.cloneNode(True)
    
    'GD SYS3092 22/04/2002 ----------------------end--------------------------
    
    objIDOAssist.Create xmlTmpElem, xmlLifeBenefitDef
    
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

'    ILifeSubQuoteDO_CreateLifeBenefit = strXmlOut
    
    Exit Function

ILifeSubQuoteDO_CreateLifeBenefitVbErr:

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'MS 19/07/00 Modified for performance enhancements
Private Sub ILifeSubQuoteDO_Update(ByVal vXMLRequest As IXMLDOMElement)
'Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object - Table 'LifeSubQuote'
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo ILifeSubQuoteDO_UpdateVbErr

    Const strFunctionName As String = "ILifeSubQuoteDO_Update"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objIDOAssist As IDOAssist
    
    
    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim objIomCMClassDef As IomCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
    
    objIDOAssist.Update vXMLRequest, objIomCMClassDef.LoadLifeSubQuoteData
    
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objIDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

ILifeSubQuoteDO_UpdateVbErr:
    
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


