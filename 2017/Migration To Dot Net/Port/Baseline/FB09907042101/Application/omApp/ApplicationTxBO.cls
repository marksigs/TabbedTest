VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ApplicationTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Application Business Object which requires transactions.

'Dependencies:  ApplicationDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'----------------------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'AD     05/11/99 Created.
'RF     10/01/99 Added methods for Declaration table.
'RF     29/02/00 AQR SYS0378: Automation error was occuring but error reporting did not
'                show true cause.
'SR     10/03/00 Methods 'UpdataApplicantCounts' and 'UpdateCorrespondenceSalutation'
'                use 'ApplicationManagerDO' rather than 'CustomerRoleDO'
'SR     14/03/00 New method 'CreateApplicationAndCustomer'
'IW     17/03/00 New method 'CreateApplicationVerification
'IW     20/03/00 New method 'CreateApplicationCreditcard
'SR     24/03/00 Added MemoPad methods
'MC     26/04/00 Added OtherInsuranceCompany methods
'ASm    25/05/00 Change error handlers to only SetAbort
'SR     02/06/00 New method 'SaveNewLoan'
'JLD    06/06/00 Just send the APPLICATIONFACTFIND node to UpdateApplicationFactfind in SaveNewLoan
'SR     12/06/00  New Method 'CreateApplicationCustomeAndVerification'
'MC     20/06/00 Amend SaveNewLoan method for use by Import
'JLD    27/06/00 If there are no LOANSLIABILITIES don't call the save method in SaveFinancialSummaryAndLiability
'MC     13/07/00  SYS1168 - Update DeleteDependantsForCustomers
'SR     09/08/00 Modified method CreateApplicationCustomerAndVerification. create a record in ApplicationStage, if required
'SR     16/08/00 - SYS1174: Modified method CreateApplication. create a record in user history only
'       if UserId and UnitId are mentioned in the Request.
'APS    30/11/00 CORE000018
'APS    12/01/01 SYS1817: oeNoApplicantsFound, oeInvalidNumberOfApplicants, oeInvalidNumberOfCustomers, oeInvalidNumberOfGuarantors
'       error enums removed and replaced by there respective error numbers
'DJP    29/01/01 SYS1808 - CreateApplication, default StageSequenceNumber and CaseSequenceNo to "1"
'APS    28/02/01 SYS1993 - Added CreateApplicationPriority method call to CreateApplication
'PSC    09/03/01 SYS1879 - Added ApproveRecommendApplication
'PSC    12/03/01 SYS2024 - Amend CreateApplication so that the initial stage is not
'                          created if task manager is present
'JLD    23/10/01 SYS2839 Pass thru' LOCKING node if present to createApplication.
'BG     21/12/01 SYS3268 Added UpdateApplicationOffer and CreateApplicationOffer methods
'PSC    07/02/02 SYS3995 Amend CreateApplicationOffer to error if Global Parameter is missing
'DRC    23/04/02  SYS3017 Mod to UpdateCustomerRelationship & CreateCustomerRelationship  so that there will be no more than 1 record per app
'--------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------
' MARS Specific History
'
'MF     19/08/2005  MAR20       Added CreatePropertyDetails
'PSC    22/09/2005  MAR32       Added DeriveAndSetSLADate
'SR     20/10/2005  MAR24       New methods 'CreateApplicationUnderwriting' and 'UpdateApplicationUnderwriting'
'---------------------------------------------------------------------------------------------------------------


Option Explicit
Private Const cstrTABLE_NAME = "APPLICATION"
Private Const cstrUSERHISTORY_TABLE = "USERHISTORY"
Private Const cstrAPPLICATIONSTAGE_TABLE = "APPLICATIONSTAGE"
Private Const cstrAPPLICATIONPRIORITY_TABLE = "APPLICATIONPRIORITY"
Private Const cstrAPPLICATIONFACTFIND_TABLE = "APPLICATIONFACTFIND"
Private Const cstrMORTGAGECALC_TABLE = "MORTGAGECALCULATORDETAILS"
Private Const cstrLOANPROPERTY_TABLE = "LOANPROPERTYDETAILS"
Private Const cstrAPPLICATIONVERIFICATION_TABLE = "APPLICATIONVERIFICATION"
Private Const cstrAPPLICATIONCREDITCARD_TABLE = "APPLICATIONCREDITCARD"
Private Const cstrTABLE_MEMOPAD = "MEMOPAD"
Private Const cstrAPPLICATIONOFFER_TABLE = "APPLICATIONOFFER"
Private Const cstrAPPLICATIONUNDERWRITING_TABLE = "APPLICATIONUNDERWRITING"

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required

Implements IApplicationTxBO
Implements ObjectControl

Private Function IApplicationTxBO_Create(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateCustomerRelationship multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationTxBO_Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.Create xmlElement
    Else
        ' For each relevant node call the data object Create
        For Each xmlElement In xmlNodeList
            objIApplicationDO.Create xmlElement
        Next
    End If
    
    Set IApplicationTxBO_Create = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_Create = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationTxBO_CreateApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationUnderwriting"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONUNDERWRITING_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONUNDERWRITING_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONUNDERWRITING_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateApplicationUnderwriting xmlElement
    Else
        ' For each relevant node call the data object Create
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateApplicationUnderwriting xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateApplicationUnderwriting = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationUnderwriting = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IApplicationTxBO_UpdateApplicationUnderwriting(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationUnderwriting"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONUNDERWRITING_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONUNDERWRITING_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONUNDERWRITING_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationUnderwriting xmlElement
    Else
        ' For each relevant node call the data object UpdateApplication
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationUnderwriting xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplicationUnderwriting = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationUnderwriting = objErrAssist.CreateErrorResponseEx()
   
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function


Private Function IApplicationTxBO_CreateApplication( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header --------------------------------------------------------------------------------------------
' description:  Create Application for single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'----------------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplication"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlInputElem As IXMLDOMElement
    
    Dim objIApplicationDO As IApplicationDO
    Dim objApplicationManagerBO As ApplicationManagerBO
    Dim objIApplClassDef As IApplicationClassDef
    Dim objGlobalParamBO As GlobalParameterBO       ' PSC 12/03/01 SYS2024
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                            
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objApplicationManagerBO = New ApplicationManagerDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objApplicationManagerBO = objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
    End If
    
    '-------------------------------------------------------------------------------------------
    ' Create Application
    '-------------------------------------------------------------------------------------------
    Set xmlTempResponseNode = IApplicationTxBO_Create(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    '--------------------------------------------------------------------------------------------
    'Create UserHistory
    '--------------------------------------------------------------------------------------------
    ' Add UserHistoryDate as today's date to request
    ' SR 16/08/00 - SYS1174: create a record in user history only if UserId and UnitId are mentioned
    '                        in the Request.
    
    Dim strUserId As String, strUnitId As String
    
    strUserId = objXmlAssist.GetTagValue(vxmlRequest, "USERID", , True)
    strUnitId = objXmlAssist.GetTagValue(vxmlRequest, "UNITID", , True)
    
    If Len(strUserId) <> 0 And Len(strUnitId) <> 0 Then
    
        Set xmlInputElem = xmlOut.createElement("REQUEST")
        Call objXmlAssist.CopyNode("APPLICATION", "USERHISTORY", vxmlRequest, xmlInputElem, True)
    
        Set xmlElement = xmlOut.createElement("USERHISTORYDATE")
        xmlElement.Text = Format(Now, "dd-mm-yyyy hh:mm:ss")
        xmlInputElem.firstChild.appendChild xmlElement
    
        Call objIApplicationDO.CreateUserHistory(xmlInputElem)
    End If
    
    '-------------------------------------------------------------------------------------------
    'Create ApplicationFactFind
    '-------------------------------------------------------------------------------------------
    Set xmlInputElem = xmlOut.createElement("REQUEST")
    Call objXmlAssist.CopyNode("APPLICATION", "APPLICATIONFACTFIND", vxmlRequest, xmlInputElem, True)
    
    ' set the NUMBEROFAPPLICANTS and NUMBEROFGUARANTORS to 0
    Set xmlElement = xmlOut.createElement("NUMBEROFAPPLICANTS")
    xmlElement.Text = "0"
    xmlInputElem.firstChild.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("NUMBEROFGUARANTORS")
    xmlElement.Text = "0"
    xmlInputElem.firstChild.appendChild xmlElement
    
    Call objIApplicationDO.CreateApplicationFactFind(xmlInputElem)
    
    '-------------------------------------------------------------------------------------------
    'Create MortgageCalculator Details
    '-------------------------------------------------------------------------------------------
    'Check whether all the non-key values are not null; if so, create record in
    ' MortgagerCalcDetails
    Set objIApplClassDef = New ApplicationClassDef
    If Not objXmlAssist.AreAllNonKeyValuesNullEx(vxmlRequest.firstChild, _
                                objIApplClassDef.LoadMortgageCalcDetailsData) Then

        Set xmlInputElem = xmlOut.createElement("REQUEST")
        Call objXmlAssist.CopyNode("APPLICATION", "MORTGAGECALCULATORDETAILS", vxmlRequest, xmlInputElem, True)

        objIApplicationDO.CreateMortgageCalcDetails xmlInputElem
    End If
    
    '-----------------------------------------------------------------------------------------
    'Create Application Stage
    '-----------------------------------------------------------------------------------------
    ' PSC 12/03/01 SYS2024 - Start
    Set objGlobalParamBO = _
                m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    
    ' Only do this if task manager is not present as it will create the first stage
    If objGlobalParamBO.IsTaskManager() = False Then
    
        Set xmlInputElem = xmlOut.createElement("REQUEST")
        Call objXmlAssist.CopyNode("APPLICATION", "APPLICATIONSTAGE", vxmlRequest, xmlInputElem, True)
        
        ' Set the stage number to 10 and stage name to 'Customer Registration'
        
        Set xmlElement = xmlOut.createElement("STAGENUMBER")
        xmlElement.Text = "10"
        xmlInputElem.firstChild.appendChild xmlElement
        
        Set xmlElement = xmlOut.createElement("STAGENAME")
        xmlElement.Text = "Customer Registration"
        xmlInputElem.firstChild.appendChild xmlElement
        
        ' SYS1808 - Default StageSequenceNumber and CaseSequenceNo to "1"
        Set xmlElement = xmlOut.createElement("STAGESEQUENCENO")
        xmlElement.Text = "1"
        xmlInputElem.firstChild.appendChild xmlElement
        
        Set xmlElement = xmlOut.createElement("CASESEQUENCENO")
        xmlElement.Text = "1"
        xmlInputElem.firstChild.appendChild xmlElement
        
        objIApplicationDO.CreateApplicationStage xmlInputElem
    End If
    ' PSC 12/03/01 SYS2024 - End
    
    '-----------------------------------------------------------------------------------------
    'Create Application Priority
    '-----------------------------------------------------------------------------------------
    Set xmlInputElem = xmlOut.createElement("REQUEST")
    Call objXmlAssist.CopyNode("APPLICATION", "APPLICATIONPRIORITY", vxmlRequest, xmlInputElem, True)
        
    objIApplicationDO.CreateApplicationPriority xmlInputElem
    
    '------------------------------------------------------------------------------
    '   Create Application Lock
    '-----------------------------------------------------------------------------
    
    'RF 29/02/00 AQR SYS0378: Start - Incorrect "REQUEST" blocks being used.
    'Set xmlInputElem = xmlOut.createElement("REQUEST")
    
    Dim strLocking As String
    ' JLD Check LOCKING value on input xml. It is optional
    strLocking = objXmlAssist.GetTagValue(vxmlRequest, "LOCKING", , True)
    If Not strLocking = "0" Then
        Set xmlInputElem = objXmlAssist.GetRequestNodeEx(vxmlRequest)
        'RF 29/02/00 AQR SYS0378: End
        
        Call objXmlAssist.CopyNode("APPLICATION", "APPLICATIONLOCK", vxmlRequest, xmlInputElem, True)
        
        Set xmlElement = xmlOut.createElement("LOCKDATE")
        xmlElement.Text = Format(Now, "dd/mm/yyyy")
        xmlInputElem.firstChild.appendChild xmlElement
        
        Set xmlElement = xmlOut.createElement("MACHINEID")
        xmlElement.Text = objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "MACHINEID")
        xmlInputElem.firstChild.appendChild xmlElement
        
        Set xmlElement = xmlOut.createElement("TYPEOFLOCK")
        xmlElement.Text = "On"
        xmlInputElem.firstChild.appendChild xmlElement
        
        ' RF 29/02/00 AQR SYS0378 Start - Check response from BO.
        Dim strResponse As String
        strResponse = objApplicationManagerBO.CreateLock(xmlInputElem.xml)
        objErrAssist.CheckXMLResponse strResponse, True, xmlResponseElem
        ' RF 29/02/00 AQR SYS0378 End
    End If
    
    Set IApplicationTxBO_CreateApplication = xmlResponseElem.cloneNode(True)
    
CreateApplicationExit:
    
    '   Clear all the instances of the local objects
    Set objIApplicationDO = Nothing
    Set objIApplClassDef = Nothing
    Set objApplicationManagerBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlInputElem = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objGlobalParamBO = Nothing      ' PSC 12/03/01 SYS2024
    
    Exit Function
                        
CreateApplicationVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplication = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationExit

End Function

Private Function IApplicationTxBO_CreateApplicationAndCustomer( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create new application, customer and a link between application and customer
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationAndCustomer"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOutApplElem As IXMLDOMElement  ' Applicaion element in Response
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objIApplicationBO As IApplicationBO
    Dim objIApplManagerBO As IApplicationManagerBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"

    End If
   
    If objContext Is Nothing Then
        Set objIApplicationBO = New ApplicationBO
        Set objIApplManagerBO = New ApplicationManagerBO
    Else
        Set objIApplicationBO = objContext.CreateInstance(App.Title & ".ApplicationBO")
        Set objIApplManagerBO = objContext.CreateInstance(App.Title & ".ApplicationManagerBO")
    End If
    
    ' Create Application
    Set xmlTempResponseNode = objIApplicationBO.ValidateApplication(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIApplicationBO.CreateApplication(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Get values of ApplicationNumber and FactFindNumber - required later
    Dim strApplNumber As String, strAFFNumber As String
    strApplNumber = xmlTempResponseNode.selectSingleNode("APPLICATIONNUMBER").Text
    strAFFNumber = xmlTempResponseNode.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
    
    Set xmlTempResponseNode = _
                objIApplManagerBO.Validate(vxmlRequest, bomAddCustomerToApplication)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIApplManagerBO.AddCustomerToApplication(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add ApplicationNumber ApplicationFactFindNumber to response
    Set xmlOutApplElem = xmlOut.createElement("APPLICATION")
    xmlResponseElem.appendChild xmlOutApplElem
    
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlOutApplElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlOutApplElem.appendChild xmlElement
    
    Set IApplicationTxBO_CreateApplicationAndCustomer = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationBO = Nothing
    Set objIApplManagerBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationAndCustomer = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IApplicationTxBO_CreateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with th table
'               'ApplicationCreditCard'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationCreditCardVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationCreditCard"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONCREDITCARD_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONCREDITCARD_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONCREDITCARD_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateApplicationCreditCard xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateApplicationCreditCard xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateApplicationCreditCard = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateApplicationCreditCardExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateApplicationCreditCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationCreditCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationCreditCardExit


End Function

Private Function IApplicationTxBO_CreateApplicationCustomerAndVerification( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationCustomerAndVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlApplStageElem As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlOutApplElem As IXMLDOMElement  ' Applicaion element in Response
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim objIApplicationBO As IApplicationBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"

    End If
   
    If objContext Is Nothing Then
        Set objIApplicationBO = New ApplicationBO
    Else
        Set objIApplicationBO = objContext.CreateInstance(App.Title & ".ApplicationBO")
    End If
    
    ' Create Application and Customer
    Set xmlTempResponseNode = objIApplicationBO.ValidateCreateApplicationAndCustomer(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set xmlTempResponseNode = objIApplicationBO.CreateApplicationAndCustomer(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Get values of ApplicationNumber and FactFindNumber - required later
    Dim strApplNumber As String, strAFFNumber As String
    strApplNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = xmlTempResponseNode.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    
    ' Ceate ApplicationVerification
    CreateApplicationVerification vxmlRequest, strApplNumber, strAFFNumber
    
    '-------------------------------------------------------------------------------
    'SR 09/08/00 : Create ApplicationStage if the respective tag 'APPLICATIONSTAGE'
    '              is mentioned in the Request
    '-------------------------------------------------------------------------------
    Set xmlApplStageElem = vxmlRequest.selectSingleNode(".//APPLICATIONSTAGE")
    
    If Not xmlApplStageElem Is Nothing Then
        ' Assign the new application number and fact find number
        Set xmlElement = xmlApplStageElem.selectSingleNode(".//APPLICATIONNUMBER")
        If xmlElement Is Nothing Then
            Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
            xmlApplStageElem.appendChild xmlElement
        End If
        xmlElement.Text = strApplNumber
        
        Set xmlElement = xmlApplStageElem.selectSingleNode(".//APPLICATIONFACTFINDNUMBER")
        If xmlElement Is Nothing Then
            Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
            xmlApplStageElem.appendChild xmlElement
        End If
        xmlElement.Text = strAFFNumber
        
        Set xmlTempResponseNode = objIApplicationBO.ValidateApplicationStage(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        Set xmlTempResponseNode = objIApplicationBO.CreateApplicationStage(vxmlRequest)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    
    ' Add ApplicationNumber ApplicationFactFindNumber to response
    Set xmlOutApplElem = xmlOut.createElement("APPLICATION")
    xmlResponseElem.appendChild xmlOutApplElem
    
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlOutApplElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlOutApplElem.appendChild xmlElement
    
    
    Set IApplicationTxBO_CreateApplicationCustomerAndVerification = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:

    Set objIApplicationBO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationCustomerAndVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit


End Function

Private Sub CreateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement, _
                                          ByVal strApplNumber As String, _
                                          ByVal strAFFNumber As String)
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssit As New XMLAssist
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlElement As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim objIApplicationDO As IApplicationDO
    
    ' Build the required - REQUEST string
    Set xmlNewRequest = objXmlAssit.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlDoc.createElement("APPLICATIONVERIFICATION")
    xmlNewRequest.appendChild xmlTableElem
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONNUMBER")
    xmlElement.Text = strApplNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = strAFFNumber
    xmlTableElem.appendChild xmlElement
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//RESIDENCYOKINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//AGEOKINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//LOANAMOUNTINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//LOANTERMINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//PRIMARYRESIDENCEINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//KNOWNPROPERTYINDICATOR")
    If Not xmlElement Is Nothing Then
        xmlTableElem.appendChild xmlElement.cloneNode(True)
    End If
    
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    objIApplicationDO.CreateApplicationVerification xmlNewRequest
    
    ' Clear all the objects
    Set objXmlAssit = Nothing
    Set xmlNewRequest = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlDoc = Nothing
    Set objIApplicationDO = Nothing
    
End Sub

Private Function IApplicationTxBO_CreateApplicationPriority(vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with the table
'               'ApplicationPriority' invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationPriorityVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationPriority"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONPRIORITY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONPRIORITY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONPRIORITY_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateApplicationPriority xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateApplicationPriority xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateApplicationPriority = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateApplicationPriorityExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateApplicationPriorityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationPriority = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationPriorityExit
    
End Function

Private Function IApplicationTxBO_CreateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with the table
'               'ApplicationStage'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationStageVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationStage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONSTAGE_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONSTAGE_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONSTAGE_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateApplicationStage xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateApplicationStage xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateApplicationStage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateApplicationStageExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateApplicationStageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationStage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationStageExit
End Function

Private Function IApplicationTxBO_CreateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with th table
'               'ApplicationVerification'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONVERIFICATION_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONVERIFICATION_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONVERIFICATION_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateApplicationVerification xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateApplicationVerification xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateApplicationVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateApplicationVerificationExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateApplicationVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateApplicationVerificationExit

End Function

Private Function IApplicationTxBO_CreateCustomerRelationship(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateCustomerRelationship multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerRelationshipVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateCustomerRelationship"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
   'AQR SYS3017 -  Max of 1 relationship per app
    Set xmlElement = vxmlRequest.selectSingleNode(".//CUSTOMERRELATIONSHIP")
                
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateCustomerRelationship xmlElement
    Else
           objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CUSTOMERRELATIONSHIP" & " tag not found"
    End If
    
    Set IApplicationTxBO_CreateCustomerRelationship = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateCustomerRelationshipExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateCustomerRelationshipVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateCustomerRelationship = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateCustomerRelationshipExit

End Function

Private Function IApplicationTxBO_CreateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateFinancialSummary multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateFinancialSummaryVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateFinancialSummary"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "FINANCIALSUMMARY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("FINANCIALSUMMARY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "FINANCIALSUMMARY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateFinancialSummary xmlElement
    Else
        ' For each relevant node call the data object CreateFinancialSummary
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateFinancialSummary xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateFinancialSummary = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateFinancialSummaryExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateFinancialSummaryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateFinancialSummary = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateFinancialSummaryExit

End Function

Private Function IApplicationTxBO_CreateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with th table
'               'MortgageCalculatorDetails'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanPropertyVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateLoanProperty"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrLOANPROPERTY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrLOANPROPERTY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrLOANPROPERTY_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateLoanProperty xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateLoanProperty xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateLoanProperty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateLoanPropertyExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateLoanPropertyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateLoanProperty = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateLoanPropertyExit
End Function


Private Function IApplicationTxBO_CreatePropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Creates NEWPROPERTY & NEWLOAN records associated with the APLICATIONFACTFIND
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyDetailsVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreatePropertyDetails"
    
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrLOANPROPERTY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrLOANPROPERTY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrLOANPROPERTY_TABLE & " tag not found"
        End If
    End If
   
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreatePropertyDetails xmlElement
    Else
        ' For each relevant node call the data object CreatePropertyDetails
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreatePropertyDetails xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreatePropertyDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreatePropertyDetailsExit:

    Set objIApplicationDO = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreatePropertyDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreatePropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreatePropertyDetailsExit
End Function


Private Function IApplicationTxBO_CreateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with th table
'               'MortgageCalculatorDetails'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateMortgageCalcDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrMORTGAGECALC_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrMORTGAGECALC_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrMORTGAGECALC_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateMortgageCalcDetails xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateMortgageCalcDetails xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateMortgageCalcDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateMortgageCalcDetailsExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateMortgageCalcDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateMortgageCalcDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageCalcDetailsExit
End Function

Private Function IApplicationTxBO_CreateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherInsuranceCompanyVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateOtherInsuranceCompany"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "OTHERINSURANCECOMPANY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("OTHERINSURANCECOMPANY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "OTHERINSURANCECOMPANY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If

    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateOtherInsuranceCompany xmlElement
    Else
        ' For each relevant node call the data object CreateOtherInsuranceCompany
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateOtherInsuranceCompany xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateOtherInsuranceCompany = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOtherInsuranceCompanyExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateOtherInsuranceCompanyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateOtherInsuranceCompany = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOtherInsuranceCompanyExit


End Function

Private Function IApplicationTxBO_CreateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateOtherResident multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherResidentVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateOtherResident"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO

    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "OTHERRESIDENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("OTHERRESIDENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "OTHERRESIDENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If

    ' Creat a 'GENERATEDKEYS' child node to the main response XML
    Set xmlKeysNode = xmlResponseElem.appendChild(xmlOut.createElement("GENERATEDKEYS"))
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateOtherResident xmlElement, xmlKeysNode
    Else
        ' For each relevant node call the data object CreateOtherResident
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateOtherResident xmlElement, xmlKeysNode
        Next
    End If
    
    Set IApplicationTxBO_CreateOtherResident = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateOtherResidentExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateOtherResidentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateOtherResident = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateOtherResidentExit

End Function

Private Function IApplicationTxBO_CreateUserHistory _
                                 (ByVal vxmlRequest As IXMLDOMElement) _
                                 As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateUserHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateUserHistoryVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateUserHistory _"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrUSERHISTORY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrUSERHISTORY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrUSERHISTORY_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateUserHistory xmlElement
    Else
        ' For each relevant node call the data object CreateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateUserHistory xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateUserHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateUserHistoryExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateUserHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateUserHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateUserHistoryExit
End Function

Private Function IApplicationTxBO_CreateVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateVerification multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "VERIFICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("VERIFICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VERIFICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateVerification xmlElement
    Else
        ' For each relevant node call the data object CreateVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateVerification xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateVerificationExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateVerificationExit

End Function

Private Function IApplicationTxBO_DeleteApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the
'               table 'ApplicationStage'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteApplicationStageVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteApplicationStage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONSTAGE_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONSTAGE_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONSTAGE_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteApplicationStage xmlElement
    Else
        ' For each relevant node call the data object DeleteVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteApplicationStage xmlElement
        Next
    End If
    
    Set IApplicationTxBO_DeleteApplicationStage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteApplicationStageExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteApplicationStageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteApplicationStage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteApplicationStageExit
End Function

Private Function IApplicationTxBO_DeleteDependantForCustomer(ByVal vxmlRequest As IXMLDOMElement, Optional ByVal bRaiseErr As Boolean = True) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteDependantForCustomer multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDependantForCustomerVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteDependantForCustomer"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "DEPENDANT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("DEPENDANT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DEPENDANT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteDependantForCustomer xmlElement, bRaiseErr
    Else
        ' For each relevant node call the data object DeleteDependantForCustomer
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteDependantForCustomer xmlElement, bRaiseErr
        Next
    End If
    
    Set IApplicationTxBO_DeleteDependantForCustomer = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteDependantForCustomerExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteDependantForCustomerVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteDependantForCustomer = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteDependantForCustomerExit

End Function

Private Function IApplicationTxBO_DeleteMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the
'               table 'MortgageCalculatorDetails'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteMortgageCalcDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrMORTGAGECALC_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrMORTGAGECALC_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrMORTGAGECALC_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteMortgageCalcDetails xmlElement
    Else
        ' For each relevant node call the data object DeleteVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteMortgageCalcDetails xmlElement
        Next
    End If
    
    Set IApplicationTxBO_DeleteMortgageCalcDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteMortgageCalcDetailsExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteMortgageCalcDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteMortgageCalcDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteMortgageCalcDetailsExit
End Function

Private Function IApplicationTxBO_DeleteOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteOtherResident multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteOtherResidentVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteOtherResident"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "OTHERRESIDENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("OTHERRESIDENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "OTHERRESIDENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteOtherResident xmlElement
    Else
        ' For each relevant node call the data object DeleteOtherResident
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteOtherResident xmlElement
        Next
    End If
    
    Set IApplicationTxBO_DeleteOtherResident = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteOtherResidentExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteOtherResidentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteOtherResident = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteOtherResidentExit

End Function

Private Function IApplicationTxBO_DeleteUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteVerification multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteUserHistoryVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteUserHistory"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrUSERHISTORY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrUSERHISTORY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrUSERHISTORY_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteUserHistory xmlElement
    Else
        ' For each relevant node call the data object DeleteVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteUserHistory xmlElement
        Next
    End If
    
    Set IApplicationTxBO_DeleteUserHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteUserHistoryExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteUserHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteUserHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteUserHistoryExit
End Function

Private Function IApplicationTxBO_DeleteVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteVerification multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "VERIFICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("VERIFICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VERIFICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.DeleteVerification xmlElement
    Else
        ' For each relevant node call the data object DeleteVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.DeleteVerification xmlElement
        Next
    End If
    
    Set IApplicationTxBO_DeleteVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteVerificationExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteVerificationExit

End Function

Private Function IApplicationTxBO_SaveFinancialSummaryAndLiability( _
                ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo SaveFinancialSummaryAndLiabilityVbErr

    Const strFunctionName As String = "IApplicationTxBO_SaveFinancialSummaryAndLiability"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim objSqlAssist As SQLAssist
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    Set objSqlAssist = New SQLAssist
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim objCFBO As CustomerFinancialBO
    Dim objIADOAssist As IADOAssist
    
    Dim strApplNumber As String, strAFFNumber As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If vxmlRequest.nodeName = "FINANCIALSUMMARY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName("FINANCIALSUMMARY").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "FINANCIALSUMMARY" & " tag not found"
    End If
    
    If objContext Is Nothing Then
        Set objCFBO = New CustomerFinancialBO
        Set objIADOAssist = New ADOAssist
    Else
        Set objCFBO = objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    '---------------------------------------------------------------------------
    'SaveFinancialSummary - Decide whether a record is to be created or updated
    'and call appropriate method
    '---------------------------------------------------------------------------
    Dim strCondition As String
    strApplNumber = objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER")
    strAFFNumber = objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONFACTFINDNUMBER")
    
    If strApplNumber = "" Or strAFFNumber = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingKey, _
                            "Application number And/Or ApplicationFactFindNumber not found"
    End If
    
    strCondition = "APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplNumber) _
                  & " And APPLICATIONFACTFINDNUMBER = " & objSqlAssist.FormatString(strAFFNumber)
    
    If objIADOAssist.CheckRecordExists("FINANCIALSUMMARY", strCondition) Then
        Set xmlTempResponseNode = IApplicationTxBO_UpdateFinancialSummary(vxmlRequest)
    Else
        Set xmlTempResponseNode = IApplicationTxBO_CreateFinancialSummary(vxmlRequest)
    End If
    
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    '---------------------------------------------------------------------------
    'Save Liability - records in the table 'LoansLiabilities' if any passed in
    '---------------------------------------------------------------------------
    Dim xmlNodeList As IXMLDOMNodeList
    Set xmlNodeList = vxmlRequest.getElementsByTagName("LOANSLIABILITIES")
    If xmlNodeList.length > 0 Then
        Dim strResponse As String
        strResponse = objCFBO.SaveLiability(vxmlRequest.xml)
        Set xmlDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    End If
        
    Set IApplicationTxBO_SaveFinancialSummaryAndLiability = xmlResponseElem
    
SaveFinancialSummaryAndLiabilityExit:
    Set objCFBO = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlElement = Nothing
    
    Exit Function
SaveFinancialSummaryAndLiabilityVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_SaveFinancialSummaryAndLiability = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveFinancialSummaryAndLiabilityExit
    
End Function

Private Function IApplicationTxBO_SaveMenuItemStatus( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
'------------------------------------------------------------------------------------------
' NOTE: Look for tag 'UpdateAllowed' in the request. If the value with it is zero, do not
'       modify the existing record.
'------------------------------------------------------------------------------------------
On Error GoTo SaveMenuItemStatusVbErr

    Const strFunctionName As String = "IApplicationTxBO_SaveMenuItemStatus"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSqlAssist As New SQLAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objIADOAssist As IADOAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If vxmlRequest.nodeName = "MENUITEMSTATUS" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("MENUITEMSTATUS")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "MENUITEMSTATUS" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
   
    Dim strCondition As String
    Dim strApplNumber As String, strAFFNumber As String
    Dim strCustomerNumber As String, strCustomerVersionNumber As String
    Dim strMenuItemId As String, strUpdateAllowed As String
    
    strUpdateAllowed = objXmlAssist.GetTagValue(vxmlRequest, "UPDATEALLOWED")
    
    If Not (xmlElement Is Nothing) Then
        ' Find whether a record already with the values passed in. If so call update else create
        strApplNumber = xmlElement.selectSingleNode("APPLICATIONNUMBER").Text
        strAFFNumber = xmlElement.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
        strCustomerNumber = xmlElement.selectSingleNode("CUSTOMERNUMBER").Text
        strCustomerVersionNumber = xmlElement.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
        strMenuItemId = xmlElement.selectSingleNode("MENUITEMID").Text
        
        strCondition = " CUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) _
                       & " AND CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                       & " AND APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplNumber) _
                       & " AND APPLICATIONFACTFINDNUMBER = " & strAFFNumber _
                       & " MENUITEMID = " & strMenuItemId
                        
        If objIADOAssist.CheckRecordExists("MENUITEMSTATUS", strCondition) Then
            If strUpdateAllowed <> 0 Then
                objIApplicationDO.UpdateMenuItemStatus (xmlElement)
            End If
        Else
            objIApplicationDO.CreateMenuItemStatus (xmlElement)
        End If
    
    Else
        ' For each relevant node call the appropriate data object method
        For Each xmlElement In xmlNodeList
            strApplNumber = xmlElement.selectSingleNode("APPLICATIONNUMBER").Text
            strAFFNumber = xmlElement.selectSingleNode("APPLICATIONFACTFINDNUMBER").Text
            strCustomerNumber = xmlElement.selectSingleNode("CUSTOMERNUMBER").Text
            strCustomerVersionNumber = xmlElement.selectSingleNode("CUSTOMERVERSIONNUMBER").Text
            strMenuItemId = xmlElement.selectSingleNode("MENUITEMID").Text
        
            strCondition = " CUSTOMERNUMBER = " & objSqlAssist.FormatString(strCustomerNumber) _
                           & " AND CUSTOMERVERSIONNUMBER = " & strCustomerVersionNumber _
                           & " AND APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplNumber) _
                           & " AND APPLICATIONFACTFINDNUMBER = " & strAFFNumber _
                           & " AND MENUITEMID = " & strMenuItemId
                       
            If objIADOAssist.CheckRecordExists("MENUITEMSTATUS", strCondition) Then
                If strUpdateAllowed <> "0" Then
                    objIApplicationDO.UpdateMenuItemStatus xmlElement
                End If
            Else
                objIApplicationDO.CreateMenuItemStatus xmlElement
            End If
        Next
    End If
    
    Set IApplicationTxBO_SaveMenuItemStatus = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveMenuItemStatusExit:
    
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSqlAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveMenuItemStatusVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_SaveMenuItemStatus = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveMenuItemStatusExit
End Function

Private Function IApplicationTxBO_SaveNewLoan( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Save a single instance of the persistant data associated with the table
'               'NewLoan'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewLoanVbErr

    Const strFunctionName As String = "IApplicationTxBO_SaveNewLoan"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlAFFElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMElement
    
    Dim objIApplicationDO As IApplicationDO
    Dim objINewPropertyBO As INewPropertyBO
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As New SQLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objIApplicationClassDef As IApplicationClassDef
    
    Dim strApplNumber As String, strAFFNumber As String
    
    Dim xmlFindElement As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If vxmlRequest.nodeName = "NEWLOAN" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("NEWLOAN")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                                    "NEWLOAN" & " tag not found"
        Else
            Set xmlElement = xmlNodeList.Item(0)
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objINewPropertyBO = New NewPropertyBO
        Set objIADOAssist = New ADOAssist
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objINewPropertyBO = objContext.CreateInstance(App.Title & ".NewPropertyBO")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objIApplicationClassDef = New ApplicationClassDef
    
    '-----------------------------------------------------------------------------------------
    ' Check whether a corresponding record already exist in table 'NewLoan', if so update it
    ' else create one
    '----------------------------------------------------------------------------------------
    Dim strCondition As String
        
    strApplNumber = vxmlRequest.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlRequest.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    strCondition = " APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplNumber) & " AND " _
                   & " APPLICATIONFACTFINDNUMBER =" & strAFFNumber
                   
    If objIADOAssist.CheckRecordExists("NEWLOAN", strCondition) Then
        objIApplicationDO.UpdateNewLoan xmlElement
    Else
        objIApplicationDO.CreateNewLoan xmlElement
    End If

    '--------------------------------------------------------
    ' Save NewProperty And Deposit data
    '--------------------------------------------------------
    'Check if NewProperty element exists as a child of NewLoan
    Set xmlFindElement = xmlElement.selectSingleNode("NEWPROPERTY")
    If Not (xmlFindElement Is Nothing) Then
        Set xmlTempResponseNode = objINewPropertyBO.Validate(xmlElement, bomSaveNewPropertyAndDeposit)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
        Set xmlTempResponseNode = objINewPropertyBO.SaveNewPropertyAndDeposit(xmlElement)
        objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    End If
    
    '------------------------------------------------------------------
    'Update Application (Save details to Application Fact Find table)
    '------------------------------------------------------------------
    'Check if ApplicationFactFind element exists as a child of NewLoan
    Set xmlFindElement = xmlElement.selectSingleNode("APPLICATIONFACTFIND")
    If Not (xmlFindElement Is Nothing) Then
        Set xmlAFFElement = xmlElement.getElementsByTagName("APPLICATIONFACTFIND").Item(0)
        'objIApplicationDO.UpdateApplicationFactFind xmlElement
        objIApplicationDO.UpdateApplicationFactFind xmlAFFElement
    End If

    Set IApplicationTxBO_SaveNewLoan = xmlResponseElem
SaveNewLoanExit:
    
    Set objIApplicationDO = Nothing
    Set objIADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveNewLoanVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_SaveNewLoan = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewLoanExit

End Function

Private Function IApplicationTxBO_SaveVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveVerification multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_SaveVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objIApplicationClassDef As IApplicationClassDef
    Dim botOperationType As BOOPERATIONTYPE
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Initialise
    '
    If vxmlRequest.nodeName = "VERIFICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("VERIFICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VERIFICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
   
    '
    ' Main
    '
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadVerificationData

    If Not (xmlElement Is Nothing) Then
        ' Call the appropriate DO method for this element
        botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
        Select Case botOperationType
            Case booUpdate: objIApplicationDO.UpdateVerification xmlElement
            Case booCreate: objIApplicationDO.CreateVerification xmlElement
            Case booDelete: objIApplicationDO.DeleteVerification xmlElement
            Case Else
                ' N/A
        End Select
    Else
        ' For each relevant node call the appropriate data object method
        For Each xmlElement In xmlNodeList
            botOperationType = objXmlAssist.DetermineOperationEx(xmlElement, xmlClassDefDoc)
            Select Case botOperationType
                Case booUpdate: objIApplicationDO.UpdateVerification xmlElement
                Case booCreate: objIApplicationDO.CreateVerification xmlElement
                Case booDelete: objIApplicationDO.DeleteVerification xmlElement
                Case Else
                    ' N/A
            End Select
        Next
    End If
    
    Set IApplicationTxBO_SaveVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
SaveVerificationExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlClassDefDoc = Nothing
    
    Exit Function

SaveVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_SaveVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveVerificationExit

End Function

Private Function IApplicationTxBO_Update(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateApplication multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationTxBO_Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrTABLE_NAME)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrTABLE_NAME & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.Update xmlElement
    Else
        ' For each relevant node call the data object UpdateApplication
        For Each xmlElement In xmlNodeList
            objIApplicationDO.Update xmlElement
        Next
    End If
    
    Set IApplicationTxBO_Update = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_Update = objErrAssist.CreateErrorResponseEx()
   
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function

Private Function IApplicationTxBO_UpdateApplicantCounts( _
                 ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update the columns 'NumberOfGuarantors' and 'NumberOfApplicants' in the
'               table 'ApplicationFactFInd' based on CustomrRole data
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicantCountsVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicantCounts"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseElem As IXMLDOMElement
    Dim xmlInputElem As IXMLDOMElement  'Input node to CustomerRoleDO & AFFDO
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objIApplicationDO As IApplicationDO
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objIApplicationDO = New ApplicationDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objIApplManagerDO = _
            objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objIApplicationDO = _
            objContext.CreateInstance(App.Title & ".ApplicationDO")
        Set objGlobalParamDO = _
            objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    '------------------------------------------------------------------------------------------
    ' count the number of applicants and guarantors in the CustomerRole table
    ' for the supplied ApplicationFactFind details
    '------------------------------------------------------------------------------------------
    Set xmlInputElem = xmlOut.createElement("CUSTOMERROLE")
    Set xmlElement = xmlOut.createElement("APPLICATIONNUMBER")
    xmlElement.Text = objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONNUMBER")
    xmlInputElem.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    xmlElement.Text = objXmlAssist.GetTagValue(vxmlRequest, "APPLICATIONFACTFINDNUMBER")
    xmlInputElem.appendChild xmlElement
    
    Set xmlTempResponseElem = objIApplManagerDO.FindCustomerRoleList(xmlInputElem)
        
    Set xmlNodeList = xmlTempResponseElem.getElementsByTagName("CUSTOMERROLETYPE")
    
    Dim intNumApplicants, intNumCustomers, intNumGuarantors As Integer
    Dim intMaxApplicants, intMaxCustomers, intMaxGuarantors As Integer
    Dim intCustomer As Integer
    
    intNumCustomers = xmlNodeList.length
    
    For intCustomer = 0 To intNumCustomers - 1
        Select Case xmlNodeList.Item(intCustomer).Text
        Case "1"            ' Applicant role
            intNumApplicants = intNumApplicants + 1
        Case "2"            ' Guarantor role
            intNumGuarantors = intNumGuarantors + 1
        End Select
    Next
    
    '------------------------------------------------------------------------------------------
    ' Check there is at least one Applicant; check against global parameter values
    ' of MaximumApplicants, MaximumCustomers, and MaximumGuarantors
    '------------------------------------------------------------------------------------------
    
    If intNumApplicants < 1 Then
        '"An Application must have at least one applicant"
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 539
    End If
    
    Dim strResponseXml As String
    
    strResponseXml = objGlobalParamDO.GetCurrentParameter("MaximumApplicants")
    Set xmlDoc = objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
    intMaxApplicants = _
        CInt(objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT"))
    
    If intNumApplicants > intMaxApplicants Then
        '"Number of Applicants is greater than the maximum allowed"
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 540
    End If
    
    strResponseXml = objGlobalParamDO.GetCurrentParameter("MaximumGuarantors")
    Set xmlDoc = objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
    intMaxGuarantors = _
        CInt(objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT"))
    
    If intNumGuarantors > intMaxGuarantors Then
        '"Number of Guarantors is greater than the maximum allowed"
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 542
    End If
    
    strResponseXml = objGlobalParamDO.GetCurrentParameter("MaximumCustomers")
    Set xmlDoc = objXmlAssist.load(strResponseXml, TypeName(Me), strFunctionName)
    intMaxCustomers = _
        CInt(objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT"))
    
    If intNumCustomers > intMaxCustomers Then
        '"Number of Customers is greater than the maximum allowed"
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 541
    End If
        
    '----------------------------------------------------------------------------------------
    ' update ApplicationFactFind with the new values for
    ' NumberOfApplicants and NumberOfGuarantors
    '----------------------------------------------------------------------------------------
    Set xmlInputElem = xmlOut.createElement("REQUEST")
    objXmlAssist.CopyNode "APPLICATION", "APPLICATIONFACTFIND", vxmlRequest, xmlInputElem, True
    
    Set xmlElement = xmlOut.createElement("NUMBEROFAPPLICANTS")
    xmlElement.Text = CStr(intNumApplicants)
    xmlInputElem.firstChild.appendChild xmlElement
    
    Set xmlElement = xmlOut.createElement("NUMBEROFGUARANTORS")
    xmlElement.Text = CStr(intNumGuarantors)
    xmlInputElem.firstChild.appendChild xmlElement
    
    objIApplicationDO.UpdateApplicationFactFind xmlInputElem
        
    Set IApplicationTxBO_UpdateApplicantCounts = xmlResponseElem
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
UpdateApplicantCountsExit:
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlInputElem = Nothing
    Set xmlElement = Nothing
    Set xmlNodeList = Nothing
    
    Set objIApplicationDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objIApplicationDO = Nothing
    
    Exit Function

UpdateApplicantCountsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicantCounts = objErrAssist.CreateErrorResponseEx()
   
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Set objErrAssist = Nothing
    
    ' go to clean-up section
    Resume UpdateApplicantCountsExit
End Function

Private Function IApplicationTxBO_UpdateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateApplication multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplication"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "APPLICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplication xmlElement
    Else
        ' For each relevant node call the data object UpdateApplication
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplication xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplication = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplication = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationExit

End Function

Private Function IApplicationTxBO_UpdateApplicationFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update  multiple instances of the persistant data associated with the
'               table 'ApplicationFactFind'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationFactFind"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONFACTFIND_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONFACTFIND_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONFACTFIND_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationFactFind xmlElement
    Else
        ' For each relevant node call the data object
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationFactFind xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplicationFactFind = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationFactFind = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit
End Function

Private Function IApplicationTxBO_UpdateApplicationStage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateApplicationStage multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationStageVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationStage"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONSTAGE_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONSTAGE_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONSTAGE_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationStage xmlElement
    Else
        ' For each relevant node call the data object UpdateApplicationStage
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationStage xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplicationStage = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationStageExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationStageVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationStage = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationStageExit
End Function

Private Function IApplicationTxBO_UpdateApplicationVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with the table
'               'ApplicationVerification'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONVERIFICATION_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONVERIFICATION_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONVERIFICATION_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationVerification xmlElement
    Else
        ' For each relevant node call the data object UpdateApplicationVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationVerification xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplicationVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationVerificationExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationVerificationExit

End Function

Private Function IApplicationTxBO_UpdateApplicationCreditCard(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with the table
'               'ApplicationCreditCard'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationCreditCardVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationCreditCard"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrAPPLICATIONCREDITCARD_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrAPPLICATIONCREDITCARD_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrAPPLICATIONCREDITCARD_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationCreditCard xmlElement
    Else
        ' For each relevant node call the data object UpdateApplicationCreditCard
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationCreditCard xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateApplicationCreditCard = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationCreditCardExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationCreditCardVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationCreditCard = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationCreditCardExit

End Function

Private Function IApplicationTxBO_UpdateCorrespondenceSalutation( _
                    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update the column CorrespondenceSalutaton in the Table Application
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCorrespondenceSalutationVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateCorrespondenceSalutation"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNodeList As IXMLDOMNodeList
    
    
    Dim objIApplManagerDO As IApplicationManagerDO
    Dim objIApplicationDO As IApplicationDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplManagerDO = New ApplicationManagerDO
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplManagerDO = _
                objContext.CreateInstance(App.Title & ".ApplicationManagerDO")
        Set objIApplicationDO = _
                objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    '----------------------------------------------------------------------
    ' get CustomerNumber and CustomerVersionNumber for first 2 applicants
    '----------------------------------------------------------------------
    ' start with the table name
    Dim strTagName As String
    Set xmlElement = vxmlRequest.ownerDocument.createElement("CUSTOMERROLE")
    strTagName = "APPLICATIONNUMBER"
    Set xmlTempElem = vxmlRequest.ownerDocument.createElement(strTagName)
    xmlTempElem.Text = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    xmlElement.appendChild xmlTempElem

    strTagName = "APPLICATIONFACTFINDNUMBER"
    Set xmlTempElem = vxmlRequest.ownerDocument.createElement(strTagName)
    xmlTempElem.Text = objXmlAssist.GetTagValue(vxmlRequest, strTagName)
    xmlElement.appendChild xmlTempElem

    Set xmlTempElem = objIApplManagerDO.FindMainApplicants(xmlElement)
    Set xmlNodeList = xmlTempElem.getElementsByTagName("CUSTOMERROLE")

    Dim strSalutation As String
    Dim strTitle1 As String
    Dim strSurName1 As String
    Dim strTitle2 As String
    Dim strSurName2 As String

    Dim intNumCustomers As String
    intNumCustomers = xmlNodeList.length
    
    Call GetSurNameAndTitle(xmlNodeList, strSurName1, strTitle1, strSurName2, strTitle2)
   
    If intNumCustomers = 1 Then
        strSalutation = strTitle1 & " " & strSurName1
    Else
        ' check for "Mr and Mrs"
        If strSurName1 = strSurName2 And _
            ((strTitle1 = "Mr" And strTitle2 = "Mrs") Or _
            (strTitle2 = "Mr" And strTitle1 = "Mrs")) Then
            strSalutation = strTitle1 & " & " & strTitle2 & " " & strSurName1
        Else
            strSalutation = strTitle1 & " " & strSurName1 & " & " & _
                strTitle2 & " " & strSurName2
        End If
    End If
    
    '---------------------------------------------------------------------------
    ' update salutation
    '---------------------------------------------------------------------------
    ' Add salutaion node and value to Input
    Set xmlElement = vxmlRequest.getElementsByTagName("APPLICATION").Item(0)
    Set xmlTempElem = vxmlRequest.ownerDocument.createElement("CORRESPONDENCESALUTATION")
    xmlTempElem.Text = strSalutation
    xmlElement.appendChild xmlTempElem

    objIApplicationDO.Update vxmlRequest
    
    Set IApplicationTxBO_UpdateCorrespondenceSalutation = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateCorrespondenceSalutationExit:

    Set objIApplManagerDO = Nothing
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlElement = Nothing
    Set xmlTempElem = Nothing
    Set xmlNodeList = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateCorrespondenceSalutationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Set objErrAssist = Nothing
    
   ' go to clean-up section
    Resume UpdateCorrespondenceSalutationExit
   
End Function

Private Function IApplicationTxBO_UpdateCustomerRelationship( _
                            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateCustomerRelationship multiple instances of the persistant data
'               associated with this business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerRelationshipVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateCustomerRelationship"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'AQR SYS3017 -  Max of 1 relationship per app
    
    Set xmlElement = vxmlRequest.selectSingleNode(".//CUSTOMERRELATIONSHIP")
                
     
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
  'AQR SYS3017 -  Max of 1 relationship per app
  
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateCustomerRelationship xmlElement
    Else
         objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "CUSTOMERRELATIONSHIP" & " tag not found"
    End If
    
    Set IApplicationTxBO_UpdateCustomerRelationship = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateCustomerRelationshipExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateCustomerRelationshipVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateCustomerRelationship = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateCustomerRelationshipExit

End Function

Private Function IApplicationTxBO_UpdateFinancialSummary(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateFinancialSummary multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateFinancialSummaryVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateFinancialSummary"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "FINANCIALSUMMARY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("FINANCIALSUMMARY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "FINANCIALSUMMARY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateFinancialSummary xmlElement
    Else
        ' For each relevant node call the data object UpdateFinancialSummary
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateFinancialSummary xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateFinancialSummary = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateFinancialSummaryExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateFinancialSummaryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateFinancialSummary = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
       
    ' go to clean-up section
    Resume UpdateFinancialSummaryExit

End Function

Private Function IApplicationTxBO_UpdateLoanProperty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with the table
'               'MortgageCalculatorDetails'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanPropertyVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateLoanProperty"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objINewPropertyDO As INewPropertyDO
    Dim bNewPropertySummary As Boolean
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrLOANPROPERTY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrLOANPROPERTY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrLOANPROPERTY_TABLE & " tag not found"
        End If
    End If
       
    '~ MF MAR20 If NEWPROPERTYSUMMARY tag is present then save RoomType data also
    bNewPropertySummary = Not vxmlRequest.getElementsByTagName("NEWPROPERTYSUMMARY") Is Nothing
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        If bNewPropertySummary Then
            Set objINewPropertyDO = New NewPropertyDO
        End If
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
        If bNewPropertySummary Then
            Set objINewPropertyDO = objContext.CreateInstance(App.Title & ".NewPropertyDO")
        End If
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateLoanProperty xmlElement
    Else
        ' For each relevant node call the data object UpdateLoanProperty
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateLoanProperty xmlElement
            If bNewPropertySummary Then
                objINewPropertyDO.UpdateNewPropertyRoomDetails xmlElement
            End If
        Next
    End If
    
    Set IApplicationTxBO_UpdateLoanProperty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateLoanPropertyExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateLoanPropertyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateLoanProperty = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateLoanPropertyExit
End Function

Private Function IApplicationTxBO_UpdateMortgageCalcDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with the table
'               'MortgageCalculatorDetails'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateMortgageCalcDetails"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrMORTGAGECALC_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrMORTGAGECALC_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrMORTGAGECALC_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateMortgageCalcDetails xmlElement
    Else
        ' For each relevant node call the data object UpdateMortgageCalcDetails
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateMortgageCalcDetails xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateMortgageCalcDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateMortgageCalcDetailsExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateMortgageCalcDetailsVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateMortgageCalcDetails = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageCalcDetailsExit
End Function

Private Function IApplicationTxBO_UpdateOtherInsuranceCompany(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherInsuranceCompanyVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateOtherInsuranceCompany"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "OTHERINSURANCECOMPANY" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("OTHERINSURANCECOMPANY")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "OTHERINSURANCECOMPANY" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateOtherInsuranceCompany xmlElement
    Else
        ' For each relevant node call the data object UpdateOtherInsuranceCompany
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateOtherInsuranceCompany xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateOtherInsuranceCompany = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateOtherInsuranceCompanyExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateOtherInsuranceCompanyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateOtherInsuranceCompany = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateOtherInsuranceCompanyExit


End Function

Private Function IApplicationTxBO_UpdateOtherResident(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateOtherResident multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherResidentVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateOtherResident"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "OTHERRESIDENT" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("OTHERRESIDENT")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "OTHERRESIDENT" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateOtherResident xmlElement
    Else
        ' For each relevant node call the data object UpdateOtherResident
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateOtherResident xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateOtherResident = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateOtherResidentExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateOtherResidentVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateOtherResident = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateOtherResidentExit

End Function

Private Function IApplicationTxBO_UpdateUserHistory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateUserHistory multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateUserHistoryVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateUserHistory"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = cstrUSERHISTORY_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(cstrUSERHISTORY_TABLE)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    cstrUSERHISTORY_TABLE & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateUserHistory xmlElement
    Else
        ' For each relevant node call the data object UpdateUserHistory
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateUserHistory xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateUserHistory = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateUserHistoryExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateUserHistoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateUserHistory = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateUserHistoryExit

End Function

Private Function IApplicationTxBO_UpdateVerification(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateVerification multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVerificationVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateVerification"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "VERIFICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("VERIFICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "VERIFICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateVerification xmlElement
    Else
        ' For each relevant node call the data object UpdateVerification
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateVerification xmlElement
        Next
    End If
    
    Set IApplicationTxBO_UpdateVerification = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateVerificationExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVerificationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateVerification = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateVerificationExit

End Function

Private Function IApplicationTxBO_CreateDeclaration( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   create multiple instances of the persistant data associated with this
'   business object. Invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclarationVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateDeclaration"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Dim strTableName As String
    strTableName = "DECLARATION"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.CreateDeclaration xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objIApplicationDO.CreateDeclaration xmlElement
        Next
    End If
    
    Set IApplicationTxBO_CreateDeclaration = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateDeclarationExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateDeclarationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateDeclaration = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Resume CreateDeclarationExit

End Function

Private Function IApplicationTxBO_DeleteAllDeclaration( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Delete the data for all instances of the persistant data associated with
'   this data object
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllDeclarationVbErr
    
    Const strFunctionName As String = "IApplicationTxBO_DeleteAllDeclaration"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
        
    ' Pass through to the data object
    objIApplicationDO.DeleteAllDeclaration vxmlRequest
    
    Set IApplicationTxBO_DeleteAllDeclaration = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteAllDeclarationExit:

    Set objIApplicationDO = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function
    
DeleteAllDeclarationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteAllDeclaration = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Resume DeleteAllDeclarationExit

End Function

Private Function IApplicationTxBO_SaveDeclaration( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest  xml Request element
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveDeclarationVbErr

    Const strFunctionName As String = "IApplicationTxBO_SaveDeclaration"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim blnCreate As Boolean
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    Dim strTableName As String
    
    '------------------------------------------------------------------------------------------
    '   save any declaration records using deleteall and create
    '------------------------------------------------------------------------------------------
    
    strTableName = "DECLARATION"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If Not xmlElement Is Nothing Then
        blnCreate = False

        On Error GoTo UpdateError
        objIApplicationDO.UpdateDeclaration xmlElement
        On Error GoTo SaveDeclarationVbErr
        If blnCreate Then
            ' Record does not exist to update, so create it now
            objIApplicationDO.CreateDeclaration xmlElement
        End If
    Else
        For Each xmlElement In xmlNodeList
            blnCreate = False

            On Error GoTo UpdateError
            objIApplicationDO.UpdateDeclaration xmlElement
            On Error GoTo SaveDeclarationVbErr
            If blnCreate Then
                ' Record does not exist to update, so create it now
                objIApplicationDO.CreateDeclaration xmlElement
            End If
        Next
    End If
    
    '------------------------------------------------------------------------------------------
    '   update applicationfactfind (for declaration notes)
    '------------------------------------------------------------------------------------------
    
    strTableName = "APPLICATIONFACTFIND"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = strTableName Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName(strTableName)
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeMissingPrimaryTag, strTableName & " tag not found"
        End If
    End If
   
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplicationFactFind xmlElement
    Else
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplicationFactFind xmlElement
        Next
    End If
    
    Set IApplicationTxBO_SaveDeclaration = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateError:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        blnCreate = True
        Resume Next
    End If
    
SaveDeclarationExit:

    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

SaveDeclarationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_SaveDeclaration = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Resume SaveDeclarationExit

End Function

Private Sub GetSurNameAndTitle(ByVal xmlNodeList As IXMLDOMNodeList, _
                               ByRef strSurName1 As String, ByRef strTitle1 As String, _
                               ByRef strSurName2 As String, ByRef strTitle2 As String)
' header ----------------------------------------------------------------------------------
' description: Finds the Surname and title for the first two applicants and assigns them
'              to appropriate variables
'              Called from UpdateCorrespondenceSalutation
' pass:
' return:   n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetSurNameAndTitlevbErr

    Const strFunctionName As String = "GetSurNameAndTitle"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlDoc As FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim xmlCriteria As IXMLDOMElement
    Dim xmlOutput As IXMLDOMElement
    Dim objXmlTempElem As IXMLDOMElement
    Dim objXmlTableNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    
    Dim objICustomerBO As ICustomerBO
    
    If objContext Is Nothing Then
        Set objICustomerBO = New CustomerBO
    Else
        Set objICustomerBO = _
                objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If
    
    Dim strTagName As String

    Dim intNumCustomers As Integer
    Dim intCustomer As Integer

    Set objXmlDoc = New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    intNumCustomers = xmlNodeList.length
    Set xmlResponseNode = objXmlDoc.createElement("RESPONSE")

    ' Consider only first two applicants further
    If intNumCustomers > 2 Then intNumCustomers = 2

    For intCustomer = 1 To intNumCustomers
        Set objXmlElem = objXmlDoc.createElement("REQUEST")
        Set objXmlTableNode = objXmlDoc.createElement("TABLE")
        objXmlElem.appendChild objXmlTableNode
        objXmlTableNode.Text = "CUSTOMERVERSION"
    
        '   Build Criteria Nodes
        Set objXmlTempElem = xmlNodeList.Item(intCustomer - 1)
        
        Set xmlCriteria = objXmlDoc.createElement("CRITERIA")
        objXmlTableNode.appendChild xmlCriteria
        
        strTagName = "CUSTOMERNUMBER"
        Set objXmlNode = objXmlDoc.createElement(strTagName)
        objXmlNode.Text = objXmlTempElem.selectSingleNode(strTagName).Text
        xmlCriteria.appendChild objXmlNode
        
        strTagName = "CUSTOMERVERSIONNUMBER"
        Set objXmlNode = objXmlDoc.createElement(strTagName)
        objXmlNode.Text = objXmlTempElem.selectSingleNode(strTagName).Text
        xmlCriteria.appendChild objXmlNode
        
        '   Build the Output Nodes
        Set xmlOutput = objXmlDoc.createElement("OUTPUT")
        objXmlTableNode.appendChild xmlOutput
        
        strTagName = "FIELD"
        Set objXmlNode = objXmlDoc.createElement(strTagName)
        objXmlNode.Text = "SURNAME"
        xmlOutput.appendChild objXmlNode
        
        strTagName = "FIELD"
        Set objXmlNode = objXmlDoc.createElement(strTagName)
        objXmlNode.Text = "TITLE"
        xmlOutput.appendChild objXmlNode
        
        Set objXmlNode = objICustomerBO.GetComponentData(objXmlElem)
        objErrAssist.CheckResponse objXmlNode, xmlResponseNode, True
        
        If intCustomer = 1 Then
            strSurName1 = objXmlNode.selectSingleNode(".//SURNAME").Text
            strTitle1 = objXmlAssist.GetAttributeValue(objXmlNode, "TITLE", "TEXT")
        Else
            strSurName2 = objXmlNode.selectSingleNode(".//SURNAME").Text
            strTitle2 = objXmlAssist.GetAttributeValue(objXmlNode, "TITLE", "TEXT")
        End If
    Next
    
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set xmlCriteria = Nothing
    Set xmlOutput = Nothing
    Set objXmlTempElem = Nothing
    Set objXmlTableNode = Nothing
    
    Set objICustomerBO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

GetSurNameAndTitlevbErr:
        
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set xmlCriteria = Nothing
    Set xmlOutput = Nothing
    Set objXmlTempElem = Nothing
    Set objXmlTableNode = Nothing
    
    Set objICustomerBO = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Set objErrAssist = Nothing
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Function IApplicationTxBO_CreateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateMemoPad"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
        
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_MEMOPAD Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_MEMOPAD).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_MEMOPAD & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    objIApplicationDO.CreateMemoPad xmlElement
    
    Set IApplicationTxBO_CreateMemoPad = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CreateExit:
    Set objErrAssist = Nothing
    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

CreateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateMemoPad = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit

End Function

Private Function IApplicationTxBO_DeleteMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete multiple instances of the persistant data associated with the
'               table 'Memo Pad'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeleteMemoPad"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_MEMOPAD Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_MEMOPAD).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_MEMOPAD & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    objIApplicationDO.DeleteMemoPad xmlElement
    
    Set IApplicationTxBO_DeleteMemoPad = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
DeleteExit:
    Set objErrAssist = Nothing
    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
    Exit Function

DeleteVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeleteMemoPad = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteExit

End Function

Private Function IApplicationTxBO_UpdateMemoPad(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with the
'               table 'Memo Pad'
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateMemoPad"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIApplicationDO As IApplicationDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_MEMOPAD Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_MEMOPAD).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_MEMOPAD & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    objIApplicationDO.UpdateMemoPad xmlElement
    
    Set IApplicationTxBO_UpdateMemoPad = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateExit:
    Set objErrAssist = Nothing
    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateMemoPad = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateExit

End Function

Private Function IApplicationTxBO_ApproveRecommendApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ApproveRecommendApplicationVbErr

    Const strFunctionName As String = "IApplicationTxBO_ApproveRecommendApplication"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim xmlTempRequest  As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_NAME & " tag not found"
    End If
   
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    
    Set xmlTempRequest = xmlElement.cloneNode(True)
    
    m_objXmlAssist.ChangeNodeName xmlTempRequest, "APPLICATION", cstrTABLE_MEMOPAD
    
    objIApplicationDO.CreateMemoPad xmlTempRequest
    
    m_objXmlAssist.ChangeNodeName xmlTempRequest, cstrTABLE_MEMOPAD, cstrAPPLICATIONFACTFIND_TABLE
    
    objIApplicationDO.UpdateApplicationFactFind xmlTempRequest
        
    Set IApplicationTxBO_ApproveRecommendApplication = xmlResponseElem
    
    m_objContext.SetComplete
    
ApproveRecommendApplicationExit:

    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    
    Exit Function

ApproveRecommendApplicationVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_ApproveRecommendApplication = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume ApproveRecommendApplicationExit

End Function

Private Function IApplicationTxBO_UpdateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationOfferVbErr

    Const strFunctionName As String = "IApplicationTxBO_UpdateApplicationOffer"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlAddElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement

    Dim objIApplicationDO As IApplicationDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrAPPLICATIONOFFER_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrAPPLICATIONOFFER_TABLE).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrAPPLICATIONOFFER_TABLE & " tag not found"
    End If
   
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    
    'BG amend request to set SIGNEDOFFERRECEIVED
    Set xmlTempElem = vxmlRequest.selectSingleNode("APPLICATIONOFFER")
    Set xmlAddElem = xmlOut.createElement("SIGNEDOFFERRECEIVED")
    xmlAddElem.Text = Now()
    xmlTempElem.appendChild xmlAddElem
    
    
    
    objIApplicationDO.UpdateApplicationOffer xmlElement
    
    Set IApplicationTxBO_UpdateApplicationOffer = xmlResponseElem
    
    m_objContext.SetComplete
    
UpdateApplicationOfferExit:
    
    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    
    Exit Function

UpdateApplicationOfferVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_UpdateApplicationOffer = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume UpdateApplicationOfferExit

End Function

Private Function IApplicationTxBO_CreateApplicationOffer(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationOfferVbErr

    Const strFunctionName As String = "IApplicationTxBO_CreateApplicationOffer"
    
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlAddElem As IXMLDOMElement
    Dim xmlAddElemChild As IXMLDOMElement
    Dim strXmlResponse As String
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim dblExpiration As Double
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim lngSequenceNumber As Long

    Dim objIApplicationDO As IApplicationDO
    Dim objGlobalParamBO As GlobalParameterBO
    Dim objIApplicationClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrAPPLICATIONOFFER_TABLE Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrAPPLICATIONOFFER_TABLE).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrAPPLICATIONOFFER_TABLE & " tag not found"
    End If
    
    'BG Create amended request to send down to DO
    'first copy all of incoming request
    
    Set xmlTempElem = vxmlRequest.cloneNode(True)
    Set xmlTempElem = xmlTempElem.selectSingleNode("APPLICATIONOFFER")
    
    'BG now create OFFERISSUEDATE and OFFEREXPIRYDATE elements and append them to the request
    Set xmlAddElem = xmlOut.createElement("OFFERISSUEDATE")
    xmlAddElem.Text = Now()
    xmlTempElem.appendChild xmlAddElem
    
    'BG Get OfferExpirationDate globalparameter
    Set objGlobalParamBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
    strXmlResponse = objGlobalParamBO.GetCurrentParameter("OfferExpirationDate")
    Set xmlDoc = m_objXmlAssist.load(strXmlResponse, TypeName(Me), strFunctionName)
    
    ' PSC 07/02/02 SYS3995
    m_objErrAssist.CheckResponse xmlDoc.documentElement, xmlResponseElem, True
    
    dblExpiration = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "AMOUNT")
        
    Set xmlAddElem = xmlOut.createElement("OFFEREXPIRYDATE")
    xmlAddElem.Text = DateAdd("d", dblExpiration, Now())
    xmlTempElem.appendChild xmlAddElem
    
    'BG Set sequencenumber
    ' Get the next highest sequence number
                                  
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadApplicationOfferData()
    Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    lngSequenceNumber = objIDOAssist.GetNextSequenceNumber(vxmlRequest, xmlClassDefDoc, "APPLICATIONOFFER", "OFFERSEQUENCENO")
      
    Set xmlAddElem = xmlOut.createElement("OFFERSEQUENCENO")
    xmlAddElem.Text = lngSequenceNumber
    xmlTempElem.appendChild xmlAddElem
    
    'BG Call CreateApplicationOffer method
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    
    objIApplicationDO.CreateApplicationOffer xmlTempElem
    
    'BG  if it gets here everything is ok so return sequencenumber
    Set xmlAddElem = xmlOut.createElement("OFFER")
    xmlResponseElem.appendChild xmlAddElem
    Set xmlAddElemChild = xmlOut.createElement("OFFERSEQUENCENO")
    xmlAddElemChild.Text = lngSequenceNumber
    xmlAddElem.appendChild xmlAddElemChild
    
    Set IApplicationTxBO_CreateApplicationOffer = xmlResponseElem
    
    m_objContext.SetComplete
    
CreateApplicationOfferExit:

    Set objIApplicationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicationDO = Nothing
    Set objGlobalParamBO = Nothing
    Set objIApplicationClassDef = Nothing
    Set objIDOAssist = Nothing
    
    
    Exit Function

CreateApplicationOfferVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_CreateApplicationOffer = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume CreateApplicationOfferExit

End Function


' PSC 22/09/2005 MAR32 - Start
Private Function IApplicationTxBO_DeriveAndSetSLADate(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeriveAndSetSLADateVbErr

    Const strFunctionName As String = "IApplicationTxBO_DeriveAndSetSLADate"
    
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIApplicationDO As IApplicationDO
    Dim objGlobalParamDO As GlobalParameterDO
    Dim xmlSLADate As IXMLDOMNode
    Dim xmlLatestDetails As IXMLDOMNode
    
    Dim intApplicationType As Integer
    Dim dblSLAInterval As Double
    Dim strAppNo As String
    Dim strSLAExpiryDate As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrTABLE_NAME & " tag not found"
    End If
    
    Set xmlElement = xmlElement.cloneNode(True)
    
    strAppNo = xmlGetMandatoryNodeText(xmlElement, "APPLICATIONNUMBER")
    
    Set objIApplicationDO = m_objContext.CreateInstance(App.Title & ".ApplicationDO")
    Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    
    Set xmlLatestDetails = objIApplicationDO.GetApplicationData(xmlElement)
    
    strSLAExpiryDate = xmlGetNodeText(xmlLatestDetails, "APPLICATION/SLAEXPIRYDATE")
    
    ' Expiry date not already set
    If Len(strSLAExpiryDate) = 0 Then

        intApplicationType = xmlGetNodeText(xmlLatestDetails, "APPLICATION/APPLICATIONFACTFIND/TYPEOFAPPLICATION")
        
        If IsValidationType("TypeOfMortgage", intApplicationType, "N") Or _
           IsValidationType("TypeOfMortgage", intApplicationType, "R") Then
            dblSLAInterval = objGlobalParamDO.GetCurrentParameterByType("SLANewLoanRemort", "AMOUNT")
        ElseIf IsValidationType("TypeOfMortgage", intApplicationType, "F") Or _
                IsValidationType("TypeOfMortgage", intApplicationType, "M") Then
            dblSLAInterval = objGlobalParamDO.GetCurrentParameterByType("SLAAddBorrowing ", "AMOUNT")
        End If
        
        Dim dteSLADate As Date
        dteSLADate = Date
        
        dteSLADate = DateAdd("d", dblSLAInterval, dteSLADate)
        
        Set xmlSLADate = xmlOut.createElement("SLAEXPIRYDATE")
        xmlElement.appendChild xmlSLADate
        xmlSLADate.Text = Format(dteSLADate, "dd/mm/yyyy")
        
        objIApplicationDO.Update xmlElement
    End If
    
    Set IApplicationTxBO_DeriveAndSetSLADate = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
DeriveAndSetSLADateExit:

    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicationDO = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlSLADate = Nothing
    Set xmlLatestDetails = Nothing

    
    Exit Function

DeriveAndSetSLADateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationTxBO_DeriveAndSetSLADate = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetAbort
    
    ' go to clean-up section
    Resume DeriveAndSetSLADateExit

End Function
' PSC 22/09/2005 MAR32 - End



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



