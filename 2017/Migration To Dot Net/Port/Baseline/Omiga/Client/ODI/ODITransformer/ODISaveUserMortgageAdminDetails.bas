Attribute VB_Name = "ODISaveUserMortgageAdminDetails"
'Workfile:      ODISaveUserMortgageAdminDetails.bas
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
'STB    21/05/02    SYS4609 Implemented SaveUserMortgageAdminDetails.
'------------------------------------------------------------------------------------------
Option Explicit


Public Sub SaveUserMortgageAdminDetails( _
    ByVal objODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Sends the user data to Optimus.
' pass:
'   objODITransformerState
'       Stores ODI state information.
'   vxmlRequestNode
'       <REQUEST USERID="" UNITID="" MACHINEID="" CHANNELID="" OPERATION="SaveUserMortgageAdminDetails" MODE="CREATE|UPDATE" ADMINSYSTEMSTATE="">
'           <OMIGAUSER USERID="" UNITID="" USERNAME="" JOBTITLE="" DEPARTMENT="" BRANCHNUMBER="" LANGUAGEPREFERENCE="" INITIALMENU="" TELEPHONENUMBER="" EXTENSIONNUMBER="">
'               <SUBSYSTEMDETAILS SUBSYSTEM="" USERAUTHORITYLEVEL="" GROUPPROFILENAME="" ACCESSOTHERBRANCHES="" ACCESSSTAFFACCOUNTS="" MAXIMUMREQUESTSCHEDULE="" BACKDATEDREPORTING=""/>
'           </OMIGAUSER>
'       </REQUEST>
'   vxmlResponseNode
'       XML RESPONSE node (standard + ADMINSYSTEMSTATE).
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim sUserID As String
    Dim xmlElement As IXMLDOMNode
    Dim xmlOmigaUser As IXMLDOMNode
    Dim xmlSubsystem As IXMLDOMNode
    Dim xmlUserProfile As IXMLDOMNode
    Dim xmlApplicationID As IXMLDOMNode
    Dim xmlUserProfileDetail As IXMLDOMNode
    
    Const strFunctionName As String = "SaveUserMortgageAdminDetails"

    On Error GoTo SaveUserMortgageAdminDetailsExit
        
    'Get a reference to the omiga user element.
    Set xmlOmigaUser = xmlGetMandatoryNode(vxmlRequestNode, "OMIGAUSER")
        
    'Get the userid.
    sUserID = xmlGetNodeText(vxmlRequestNode, "//OMIGAUSER/@USERID")
        
    'TODO: Ensure this structure is correct.
    
    'The structure to create/update is as follows: -
'    <USERPROFILEIMPL>
'        <PRIMARYKEY DATA=""/>
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
    
    'USERPROFILEIMPL.
    Set xmlUserProfile = vxmlResponseNode.ownerDocument.createElement("USERPROFILEIMPL")

    'PRIMARYKEY.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("PRIMARYKEY")
    xmlUserProfile.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@USERID")
        
    'USERPROFILEDETAILIMPL.
    Set xmlUserProfileDetail = vxmlResponseNode.ownerDocument.createElement("USERPROFILEDETAILIMPL")
    xmlUserProfile.appendChild xmlUserProfileDetail
    
    'USERCODE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("USERCODE")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@USERID")
    
    'USERNAME.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("USERNAME")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@USERNAME")

    'TITLE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("TITLE")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@JOBTITLE")
    
    'DEPARTMENTCODE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("DEPARTMENTCODE")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@DEPARTMENT")
    
    'BRANCHNUMBER.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("BRANCHNUMBER")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@BRANCHNUMBER")
    
    'LANGUAGE.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("LANGUAGE")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@LANGUAGEPREFERENCE")
    
    'INITIALMENUOPTION.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("INITIALMENUOPTION")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@INITIALMENU")
    
    'PHONENUMBER.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("PHONENUMBER")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@TELEPHONENUMBER")
    
    'PHONEEXTENSION.
    Set xmlElement = vxmlResponseNode.ownerDocument.createElement("PHONEEXTENSION")
    xmlUserProfileDetail.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlOmigaUser, "OMIGAUSER/@EXTENSIONNUMBER")
    'End USERPROFILEDETAILIMPL.
    
    'For each subsystem suplpied create an APPLICATIONIDIMPL element.
    For Each xmlSubsystem In vxmlRequestNode.selectNodes("//SUBSYSTEMDETAILS")
        'APPLICATIONIDIMPL.
        Set xmlApplicationID = vxmlRequestNode.ownerDocument.createElement("APPLICATIONIDIMPL")
        xmlUserProfile.appendChild xmlApplicationID
    
        'APPLICATIONTYPE.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("APPLICATIONTYPE")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@SUBSYSTEM")
        
        'AUTHORITYLEVEL.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("AUTHORITYLEVEL")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@SUBSYSTEM")
        
        'GROUPPROFILENAME.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("GROUPPROFILENAME")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@GROUPPROFILENAME")
        
        'ACCESSOTHERBRANCHES.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("ACCESSOTHERBRANCHES")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@ACCESSOTHERBRANCHES")
        
        'ACCESSSTAFFACCOUNT.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("ACCESSSTAFFACCOUNT")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@ACCESSSTAFFACCOUNTS")
        
        'MAXIMUMREQUESTSCHEDULE.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("MAXIMUMREQUESTSCHEDULE")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@MAXIMUMREQUESTSCHEDULE")
        
        'MAXIMUMBACKDATINGREPORTING.
        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("MAXIMUMBACKDATINGREPORTING")
        xmlApplicationID.appendChild xmlElement
        xmlSetAttributeValue xmlElement, "DATA", xmlGetNodeText(xmlSubsystem, "SUBSYSTEMDETAILS/@BACKDATEDREPORTING")
    Next xmlSubsystem
    
    'If mode is create then create the user in Optimus.
    If xmlGetNodeText(vxmlRequestNode, "//REQUEST/@MODE") = "CREATE" Then
        'Supply the user key value.
        Set vxmlResponseNode = PlexusHomeImpl_create_UserProfileKey(objODITransformerState, sUserID)
        
        'Ensure no errors were returned.
        CheckConverterResponse vxmlResponseNode, True
    End If
    
     'Update the mortgage product with its name.
    Set vxmlResponseNode = UserProfileImpl_update(objODITransformerState, xmlUserProfile)

    'Ensure no errors were returned.
    CheckConverterResponse vxmlResponseNode, True

    'Ensure we have a response.
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oerecordnotfound
    End If
    
SaveUserMortgageAdminDetailsExit:
    Set xmlElement = Nothing
    Set xmlOmigaUser = Nothing
    Set xmlSubsystem = Nothing
    Set xmlUserProfile = Nothing
    Set xmlApplicationID = Nothing
    Set xmlUserProfileDetail = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub


