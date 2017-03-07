Attribute VB_Name = "ODIGetUserMortgageAdminDetails"
'Workfile:      GetUserMortgageAdminDetails.bas
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:
'
'Dependencies:  List any other dependent components
'               e.g. BuildingsAndContentsSubQuoteTxBO, BuildingsAndContentsSubQuoteDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'STB    22/05/02    SYS4609 Implemented GetUserMortgageAdminDetails.
'------------------------------------------------------------------------------------------
Option Explicit


Public Sub GetUserMortgageAdminDetails( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Send a product update/creation through to Optimus.
' pass:
'   objODITransformerState
'       Stores ODI state information.
'   vxmlRequestNode
'       <REQUEST USERID="" UNITID="" MACHINEID="" CHANNELID="" OPERATION="GetUserMortgageAdminDetails" ADMINSYSTEMSTATE="">
'           <OMIGAUSER USERID=""/>
'       </REQUEST>
'   vxmlResponseNode
'       <RESPONSE TYPE="SUCCESS" ADMINSYSTEMSTATE="">
'           <OMIGAUSER USERID="" USERNAME="" JOBTITLE="" DEPARTMENT="" BRANCHNUMBER="" LANGUAGEPREFERENCE="" INITIALMENU="" TELEPHONENUMBER="" EXTENSIONNUMBER="">
'               <SUBSYSTEMDETAILS SUBSYSTEM="" USERAUTHORITYLEVEL="" GROUPPROFILENAME="" ACCESSOTHERBRANCHES="" ACCESSSTAFFACCOUNTS="" MAXIMUMREQUESTSCHEDULE="" BACKDATEDREPORTING=""/>
'           </OMIGAUSER>
'       </RESPONSE>
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim xmlElement As IXMLDOMNode
    Dim xmlOmigaUser As IXMLDOMNode
    Dim xmlSubsystem As IXMLDOMNode
    Dim xmlUserProfileKey As IXMLDOMNode
    Dim xmlUserProfileImpl As IXMLDOMNode
    Dim xmlApplicationIDImpl As IXMLDOMNode
    Dim xmlConverterResponse As IXMLDOMNode

    Const strFunctionName As String = "GetUserMortgageAdminDetails"

    On Error GoTo GetUserMortgageAdminDetailsExit

    'Get the USERPROFILEKEY element to be used in the Optmius search.
    Set xmlUserProfileKey = CreateUserProfileKey(xmlGetNodeText(vxmlRequestNode, "//OMIGAUSER/@USERID"))

    'Obtain the user details.
    Set xmlConverterResponse = PlexusHomeImpl_findByPrimaryKey_UserProfileKey(objODITransformerState, xmlUserProfileKey)
    CheckConverterResponse xmlConverterResponse, True

    'Handle the response.
    Set xmlUserProfileImpl = xmlGetMandatoryNode(xmlConverterResponse, "ARGUMENTS/OBJECT/USERPROFILEIMPL", oerecordnotfound)

    'The Optimus user will be in the format: -
    '    <USERPROFILEIMPL>
    '        <USERPROFILEDETAILIMPL>
    '            <USERCODE DATA=""/>
    '            <USERNAME DATA=""/>
    '            <TITLE DATA=""/>
    '            <DEPARTMENTCODE DATA=""/>
    '            <BRANCHNUMBER DATA=""/>
    '            <LANGUAGE DATA=""/>
    '            <INITIALMENUOPTION DATA=""/>
    '            <PHONENUMBER DATA=""/>
    '            <PHONEEXTENSION DATA=""/>
    '        </USERPROFILEDETAILIMPL>
    '        <APPLICATIONIDIMPL>
    '            <APPLICATIONTYPE DATA="">
    '            <AUTHORITYLEVEL DATA="">
    '            <GROUPPROFILENAME DATA="">
    '            <ACCESSOTHERBRANCHES DATA="">
    '            <ACCESSSTAFFACCOUNT DATA="">
    '            <MAXIMUMREQUESTSCHEDULE DATA="">
    '            <MAXIMUMBACKDATINGREPORTING DATA="">
    '        </APPLICATIONIDIMPL>
    '        <APPLICATIONIDIMPL>
    '           ...
    '        </APPLICATIONIDIMPL>
    '    </USERPROFILEIMPL>

    'Create the OMIGAUSER element and append it to the RESOPNSE node.
    Set xmlOmigaUser = vxmlRequestNode.ownerDocument.createElement("OMIGAUSER")
    vxmlResponseNode.appendChild xmlOmigaUser

    'USERID.
    xmlSetAttributeValue xmlOmigaUser, "USERID", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/USERCODE/@DATA")

    'USERNAME.
    xmlSetAttributeValue xmlOmigaUser, "USERNAME", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/USERNAME/@DATA")

    'JOBTITLE.
    xmlSetAttributeValue xmlOmigaUser, "JOBTITLE", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/TITLE/@DATA")

    'DEPARTMENT.
    xmlSetAttributeValue xmlOmigaUser, "DEPARTMENT", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/DEPARTMENTCODE/@DATA")

    'BRANCHNUMBER.
    xmlSetAttributeValue xmlOmigaUser, "BRANCHNUMBER", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/BRANCHNUMBER/@DATA")
    
    'LANGUAGEPREFERENCE.
    xmlSetAttributeValue xmlOmigaUser, "LANGUAGEPREFERENCE", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/LANGUAGE/@DATA")
    
    'INITIALMENU.
    xmlSetAttributeValue xmlOmigaUser, "INITIALMENU", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/INITIALMENUOPTION/@DATA")
    
    'TELEPHONENUMBER.
    xmlSetAttributeValue xmlOmigaUser, "TELEPHONENUMBER", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/PHONENUMBER/@DATA")
    
    'EXTENSIONNUMBER.
    xmlSetAttributeValue xmlOmigaUser, "EXTENSIONNUMBER", xmlGetNodeText(xmlUserProfileImpl, "USERPROFILEIMPL/USERPROFILEDETAILIMPL/PHONEEXTENSION/@DATA")
    
    'Copy each subsystem.
    For Each xmlApplicationIDImpl In xmlUserProfileImpl.selectNodes("//APPLICATIONIDIMPL")
        Set xmlSubsystem = vxmlResponseNode.ownerDocument.createElement("SUBSYSTEMDETAILS")
        xmlOmigaUser.appendChild xmlSubsystem
        
        'SUBSYSTEM.
        xmlSetAttributeValue xmlSubsystem, "SUBSYSTEM", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/APPLICATIONTYPE/@DATA")
        
        'USERAUTHORITYLEVEL.
        xmlSetAttributeValue xmlSubsystem, "USERAUTHORITYLEVEL", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/AUTHORITYLEVEL/@DATA")
        
        'GROUPPROFILENAME.
        xmlSetAttributeValue xmlSubsystem, "GROUPPROFILENAME", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/GROUPPROFILENAME/@DATA")
        
        'ACCESSOTHERBRANCHES.
        xmlSetAttributeValue xmlSubsystem, "ACCESSOTHERBRANCHES", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/ACCESSOTHERBRANCHES/@DATA")
        
        'ACCESSSTAFFACCOUNTS.
        xmlSetAttributeValue xmlSubsystem, "ACCESSSTAFFACCOUNTS", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/ACCESSSTAFFACCOUNT/@DATA")
        
        'MAXIMUMREQUESTSCHEDULE.
        xmlSetAttributeValue xmlSubsystem, "MAXIMUMREQUESTSCHEDULE", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/MAXIMUMREQUESTSCHEDULE/@DATA")
        
        'BACKDATEDREPORTING.
        xmlSetAttributeValue xmlSubsystem, "BACKDATEDREPORTING", xmlGetNodeText(xmlApplicationIDImpl, "APPLICATIONIDIMPL/MAXIMUMBACKDATINGREPORTING/@DATA")
    Next xmlApplicationIDImpl
    
    'Exceptions.
    AddExceptionsToResponse xmlConverterResponse, vxmlResponseNode
    
GetUserMortgageAdminDetailsExit:
    Set xmlElement = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName
End Sub
