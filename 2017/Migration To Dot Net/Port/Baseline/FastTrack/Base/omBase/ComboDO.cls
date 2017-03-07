VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ComboDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ComboDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Code for omiga4 Data Object - ComboDO
'               contains implied standard interface
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLResquest As String)
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData as String) As String
'Dependencies:  SQLAssist, XMLAssist, errAssist
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' AS    01/07/99    Created.
' RF    20/09/99    Added IsItemInValidation.
' RF    06/10/99    DOs shouldn't call SetAbort except on system error.
' RF    17/12/99    Added GetFirstComboValueId.
' MS    21/07/00    Moved methods from MP as part of performance rework.
' MC    07/08/2000  SYS1409 Amend isolation mode for SPM to LockMethod as advised following load testing
' PSC   11/08/2000  SYS1430 Back out SYS1409
' LD    07/11/2000  Explicity close recordsets
' AS    20/11/2000  SYS1670: Explicity close recordsets
' DM    17/05/01    SYS2316
' DM    17/05/01    SYS2316 Move outer joins
' DM    31/07/01    SYS2539 Remove outer join code from Oracle side of Run Time switch.
'                   This is from GetComboValue, GetComboList.
' AS    13/11/03    CORE1 Removed GENERIC_SQL.
'------------------------------------------------------------------------------------------
'BMids History:
'
' Prog  Date        Description
' MDC   12/09/2002  BMIDS00336 Added GetComboValueIdFromValueName and GetFirstComboValueIdFromValueName
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
' Prog  Date     Description
' SR    13/04/04 BBG132 - New method GetRegSaleRelatedComboList
' MV    07/07/2004  BBG123 WP16 Add function to ascertain if an application is a TofE
' TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------

Option Explicit
Implements ObjectControl
Implements IComboDO ' MS    21/07/00    performance rework.
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

'------------------------------------------------------------------------------------------
'===========================================
'Table Format
'===========================================
' ComboGroup                Datatype        Primary key     Foreign key
'   GroupName               VARCHAR2(30)    Y               Y   (ComboValue.GroupName)
'   GroupNote               VARCHAR2(255)
'
' ComboValue                Datatype        Primary key     Foreign key
'   GroupName               VARCHAR2(30)    Y               Y   (ComboGroup.GroupName)
'   ValueId                 SHORT           Y
'   ValueName               VARCHAR2(30)
'
' ComboValidation           Datatype        Primary key     Foreign key
'   GroupName               VARCHAR2(30)    Y               Y   (ComboValue.GroupName)
'   ValueId                 SHORT           Y               Y   (ComboValue.ValueId)
'   Validation              VARCHAR2(1)     Y
'------------------------------------------------------------------------------------------

'===========================================
'Constant Declaration Section
'===========================================
'===========================================
'Variable Declaration Section
'===========================================
Public Function GetComboList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the combo data for all instances of the persistant data associated with
'               this data object matching the request criteria
'
' pass:         vstrXMLRequest  xml Request data stream containing list which identifies an
'                               instance of the persistant data to be retrieved. Received in
'                               the format:
' i.e.
'   <LIST><LISTNAME>Title</LISTNAME></LIST>
'
' return:       GetComboList    string containing XML data stream representation of
'                               data retrieved
' i.e.
' <LIST>
'    <LISTNAME NAME=Title>
'       <LISTENTRY>
'           <VALUEID>1</VALUEID>
'           <VALUENAME>Mr</VALUENAME>
'           <VALIDATIONTYPELIST>
'               <VALIDATIONTYPE>M</VALIDATIONTYPE>
'               <VALIDATIONTYPE>O</VALIDATIONTYPE>
'               ...
'           </VALIDATIONTYPELIST>
'       </LISTENTRY>
'       ...
'    </LISTNAME>
'    <LISTNAME>
'    ...
'    </LISTNAME>
' </LIST>
' Raise Errors:     if record not found, raise omiga4RecordNotFound
'                   if no combo group specified raise omiga4err108
' maintenance -----------------------------------------------------------------------------
' Date          Developer   Comments
' 30/06/1999    IK          Initial Creation
' 02/07/1999    AS          Initial GetComboList implementation
' 07/09/1999    AS          Changed GetComboList to read multiple combo groups in one call
'------------------------------------------------------------------------------------------
On Error GoTo GetComboListVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComboList"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlRow As New FreeThreadedDOMDocument40
        
    Dim objXmlNode As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    Dim objXmlListNameNode As IXMLDOMNode
    Dim objXmlRowNode As IXMLDOMElement
    Dim objXmlValList As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objADOAssist As ADOAssist
    Dim cmd As New ADODB.Command
    Dim rstComboRecSet As New ADODB.Recordset
        
    Dim strSQL As String
    Dim strComboList As String
        
    Dim strCurrentGroupName As String
    Dim strCurrentValueId As String
    Dim strPreviousGroupName As String
    Dim strPreviousValueId As String
    Dim strGroupName As String
    Dim blnSameGroupName As Boolean
    Dim blnSameValueId As Boolean
    Dim blnAddRow As Boolean
    Dim blnFoundInSPM As Boolean
    objXmlIn.async = False
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    strGroupName = m_objXmlAssist.GetTagValue(objXmlIn.documentElement, "LISTNAME")
    If Not Len(strGroupName) > 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err108
    End If
    ' get the list from the shared property manager
    If Not m_objContext Is Nothing Then
        strComboList = GetAllComboListsFromSPM(objXmlIn, blnFoundInSPM)
    End If
    If Not blnFoundInSPM Then   ' no entry in SPM?
        
        If m_objContext Is Nothing Then
            Set objADOAssist = New ADOAssist
        Else
            Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
        End If
        Dim strGroupNames As String
        strGroupNames = GetGroupNames(objXmlIn)
        ' DM 17/05/01 SYS2316 Move outer joins
        ' DM 31/07/01 SYS2539
        cmd.ActiveConnection = objADOAssist.GetConnStr
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "usp_GetComboList"
        cmd.Parameters.Append cmd.CreateParameter("p_ListNames", adBSTR, adParamInput, Len(strGroupNames), strGroupNames)
            
        Set rstComboRecSet = cmd.Execute
                
        If Not rstComboRecSet.EOF Then
            
            Set objXmlNode = objXmlOut.createElement("LIST")
            Set objXmlListNode = objXmlOut.appendChild(objXmlNode)
                                    
            ' AS 30/05/01 For SQL Server MoveFirst gives the error
            ' "Rowset position cannot be restarted". Assume already at start of
            ' recordset. Should also be true for Oracle.
            While Not rstComboRecSet.EOF
                
                strCurrentGroupName = rstComboRecSet!GROUPNAME
                strCurrentValueId = rstComboRecSet!VALUEID
                                
                blnSameGroupName = StrComp(strCurrentGroupName, strPreviousGroupName, _
                                            vbTextCompare) = 0
                blnSameValueId = StrComp(strCurrentValueId, strPreviousValueId, _
                                            vbTextCompare) = 0
                If Not blnSameGroupName Then
                                    
                    Set objXmlNode = objXmlOut.createElement("LISTNAME")
                    objXmlNode.setAttribute "NAME", strCurrentGroupName
                    Set objXmlListNameNode = objXmlListNode.appendChild(objXmlNode)
                    Set objXmlRow = m_objXmlAssist.load(Me.GetXMLFromRecordSet _
                                    (rstComboRecSet), TypeName(Me), strFunctionName)
                    blnAddRow = True
                ElseIf (blnSameGroupName And Not blnSameValueId) Then
                    Set objXmlRow = m_objXmlAssist.load(Me.GetXMLFromRecordSet _
                                    (rstComboRecSet), TypeName(Me), strFunctionName)
                    blnAddRow = True
                Else
                    Set objXmlValList = objXmlRowNode.getElementsByTagName _
                                            ("VALIDATIONTYPELIST").Item(0)
                    Set objXmlNode = objXmlOut.createElement("VALIDATIONTYPE")
                    objXmlNode.Text = rstComboRecSet!VALIDATIONTYPE
                    objXmlValList.appendChild objXmlNode
                                        
                    blnAddRow = False
                End If
                                                
                If blnAddRow Then
                    Set objXmlRowNode = objXmlListNameNode.appendChild _
                                        (objXmlRow.documentElement)
                    blnAddRow = False
                End If
                strPreviousGroupName = strCurrentGroupName
                strPreviousValueId = strCurrentValueId
                rstComboRecSet.MoveNext
                Set objXmlRow = Nothing
            Wend
        Else
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
        End If
        rstComboRecSet.Close
        If Not m_objContext Is Nothing Then
            AddAllComboListsToSPM objXmlOut
        End If
    Else
        Set objXmlRow = m_objXmlAssist.load(strComboList, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlOut.appendChild(objXmlRow.documentElement)
    End If
        'rstComboRecSet.Close
    GetComboList = objXmlOut.xml
    Set rstComboRecSet = Nothing
    Set cmd = Nothing
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set objXmlOut = Nothing
    Set objXmlRow = Nothing
    Set objXmlNode = Nothing
    Set objXmlIn = Nothing
    Set objXmlListNode = Nothing
    Set objXmlListNameNode = Nothing
    Set objXmlRowNode = Nothing
    Set objXmlValList = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetComboListVbErr:
        
    ' Close the recordset if still open
        If Not rstComboRecSet Is Nothing Then
                If rstComboRecSet.State = adStateOpen Then
                        rstComboRecSet.Close
                End If
        End If
    Set rstComboRecSet = Nothing
    Set cmd = Nothing
'    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set objXmlOut = Nothing
    Set objXmlRow = Nothing
    Set objXmlNode = Nothing
    Set objXmlIn = Nothing
    Set objXmlListNode = Nothing
    Set objXmlListNameNode = Nothing
    Set objXmlRowNode = Nothing
    Set objXmlValList = Nothing
    Set objSQLAssist = Nothing
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetGroupNames(ByVal vobjXMLDocIn As FreeThreadedDOMDocument40) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetGroupNamesVBErr
    Dim strFunctionName As String
    strFunctionName = "GetGroupNames"
    Dim objSQLAssist As New SQLAssist
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objXmlListNames As IXMLDOMNodeList
    Dim objXmlListNode As IXMLDOMNode
    Dim strListNames As String
    Dim blnOr As Boolean
        
    Set objXmlListNames = vobjXMLDocIn.getElementsByTagName("LISTNAME")
    blnOr = False
    strListNames = "("
    For Each objXmlListNode In objXmlListNames
        If blnOr Then
            strListNames = strListNames & " OR "
        End If
        strListNames = strListNames & "COMBOVALUE.GROUPNAME="
        strListNames = strListNames & objSQLAssist.FormatString(objXmlListNode.Text)
        blnOr = True
    Next
    strListNames = strListNames & ")"
    GetGroupNames = strListNames
    Set objSQLAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objXmlListNames = Nothing
    Set objXmlListNode = Nothing
'    Set objErrAssist = Nothing
        
    Exit Function
GetGroupNamesVBErr:
    Set objSQLAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objXmlListNames = Nothing
    Set objXmlListNode = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Sub AddComboTextToSPM(vstrComboGroup As String, _
                                vstrComboValue As String, _
                                vstrComboText As String)
' header ----------------------------------------------------------------------------------
' description:  Adds a combo text entry to SPM
' pass:         vstrComboGroup  - Combo Group Name
'               vstrComboValue  - Combo Value
'               vstrComboText   - Combo Text
' return:       n/a
' maintenance -----------------------------------------------------------------------------
' Date          Developer        Comments
' 06/12/1999    Ashley Smith     Initial Creation
'------------------------------------------------------------------------------------------
On Error GoTo AddComboTextToSPMVBErr
    
    Dim strFunctionName As String
    strFunctionName = "AddComboTextToSPM"
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim bExists As Boolean
    Dim lIsolationMode As Long
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
        
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("ComboLists", lIsolationMode, lReleaseMode, bExists)
                
    Dim strEntry As String
    ' Combo text entries in SPM in the format ComboGroup.ComboValue = ComboText
    strEntry = vstrComboGroup & "." & vstrComboValue
    Set spmProperty = spmGroup.CreateProperty(strEntry, bExists)
    spmProperty.Value = vstrComboText
                
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
            
    Exit Sub
AddComboTextToSPMVBErr:
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Dim objErrAssist As New ErrAssist
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function GetComboListFromSPM(strGroupName As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for the combo group held in the SPM
' pass:
'   strGroupName  String which identifies a combo group name
' return:       String containing XML data stream
'               representation of data
' maintenance -----------------------------------------------------------------------------
' Date          Developer        Comments
' 05/07/1999    Ashley Smith     Initial implementation
'------------------------------------------------------------------------------------------
On Error GoTo GetComboListFromSPMVBErr
    
    Dim strFunctionName As String
    strFunctionName = "GetComboListFromSPM"
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
'    Dim objErrAssist As New ErrAssist
    Dim bExists As Boolean
    Dim lIsolationMode As Long
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
      
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("ComboLists", lIsolationMode, lReleaseMode, bExists)
            
    Set spmProperty = spmGroup.CreateProperty(strGroupName, bExists)
            
    If bExists Then
        GetComboListFromSPM = spmProperty.Value
    Else
        spmProperty.Value = ""
    End If
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetComboListFromSPMVBErr:
    
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetAllComboListsFromSPM(ByVal vobjXMLDocIn As FreeThreadedDOMDocument40, _
                                        ByRef rblnFoundInSPM As Boolean) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves the combo list xml from SPM
' pass:         vobjXMLDocIn - Request XML FreeThreadedDOMDocument40
'               rblnFoundInSPM - Boolean which indicates success
' return:       Combo list xml as string
' Raise Errors: system errors
'------------------------------------------------------------------------------------------
On Error GoTo GetAllComboListsFromSPMVBErr
    Dim strFunctionName As String
    strFunctionName = "GetAllComboListsFromSPM"
        
'    Dim objXmlAssist As New XMLAssist
    Dim objXmlList As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
'    Dim objErrAssist As New ErrAssist
    Dim iLoop As Integer
    Dim objXmlListNames As IXMLDOMNodeList
    Dim objXmlListNode As IXMLDOMNode
    Dim objXmlListName As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim strComboList As String
            
    ' which combo groups do we require?
    Set objXmlListNames = vobjXMLDocIn.getElementsByTagName("LISTNAME")
    rblnFoundInSPM = True
    Set objXmlNode = objXmlOut.createElement("LIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlNode)
    ' search for each combo list in SPM separately
    While (iLoop < objXmlListNames.Length And rblnFoundInSPM = True)
        Set objXmlListName = objXmlListNames.Item(iLoop)
        ' get the combo list from SPM
        strComboList = GetComboListFromSPM(objXmlListName.Text)
        If strComboList <> "" Then  ' list found then add to response
            Set objXmlList = m_objXmlAssist.load(strComboList, TypeName(Me), strFunctionName)
            objXmlListNode.appendChild objXmlList.documentElement
        Else
            ' if we can't find one list then this is considered a failure
            rblnFoundInSPM = False
        End If
        iLoop = iLoop + 1
    Wend
    GetAllComboListsFromSPM = objXmlOut.xml
    Set objXmlList = Nothing
    Set objXmlOut = Nothing
'    Set objXmlAssist = Nothing
    Set objXmlListNames = Nothing
    Set objXmlListNode = Nothing
'    Set objErrAssist = Nothing
    Set objXmlListName = Nothing
    Set objXmlNode = Nothing
    Exit Function
GetAllComboListsFromSPMVBErr:
    
    Set objXmlList = Nothing
    Set objXmlOut = Nothing
'    Set objXmlAssist = Nothing
    Set objXmlListNames = Nothing
    Set objXmlListNode = Nothing
    Set objXmlListName = Nothing
    Set objXmlNode = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Sub AddAllComboListsToSPM(ByVal vobjXMLDocIn As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:  Adds all combo lists to SPM
' pass:         vobjXMLDocIn - FreeThreadedDOMDocument40
' return:       n/a
' Raise Errors: System errors
' maintenance -----------------------------------------------------------------------------
' Date          Developer       Comments
' 05/07/1999    Ashley Smith    Initial creation
' 06/12/1999    Ashley Smith    Added functionality to add combo text element to SPM
'------------------------------------------------------------------------------------------
On Error GoTo AddAllComboListsToSPMVBErr
    Dim strFunctionName As String
    strFunctionName = "AddAllComboListsToSPM"
        
    Dim objXmlListNames As IXMLDOMNodeList
    Dim objXmlListEntries As IXMLDOMNodeList
    Dim objXmlListName As IXMLDOMElement
    Dim objXmlListEntry As IXMLDOMElement
            
    Dim strComboGroupName As String
    Dim strComboText As String
    Dim strComboValue As String
    Set objXmlListNames = vobjXMLDocIn.getElementsByTagName("LISTNAME")
        
    For Each objXmlListName In objXmlListNames
        strComboGroupName = objXmlListName.getAttribute("NAME")
        AddComboListToSPM strComboGroupName, objXmlListName.xml
            
        Set objXmlListEntries = objXmlListName.getElementsByTagName("LISTENTRY")
        ' Loop through all combo list entries to get the combo text and then add that
        ' combo text to SPM
        For Each objXmlListEntry In objXmlListEntries
            strComboText = objXmlListEntry.getElementsByTagName("VALUENAME").Item(0).Text
            strComboValue = objXmlListEntry.getElementsByTagName("VALUEID").Item(0).Text
            AddComboTextToSPM strComboGroupName, strComboValue, strComboText
        Next
    Next
    Set objXmlListNames = Nothing
    Set objXmlListName = Nothing
    Set objXmlListEntries = Nothing
    Set objXmlListEntry = Nothing
        
    Exit Sub
AddAllComboListsToSPMVBErr:
        
    Set objXmlListNames = Nothing
    Set objXmlListName = Nothing
    Set objXmlListEntries = Nothing
    Set objXmlListEntry = Nothing
'    Dim objErrAssist As New ErrAssist
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub
Private Sub AddComboListToSPM(strGroupName As String, strComboListXML As String)
' header ----------------------------------------------------------------------------------
' description:  Get the data for the combo group held in the SPM
' pass:         strGroupName    String which identifies the combo group
'               strGroupListXML String (XML) which identifies all combo list entries
'                               in the the combo group as per the list returned from GetList
' maintenance -----------------------------------------------------------------------------
' Date          Developer        Comments
' 05/07/1999    Ashley Smith     Initial creation
'------------------------------------------------------------------------------------------
On Error GoTo AddComboListToSPMVBErr
    
    Dim strFunctionName As String
    strFunctionName = "AddComboListToSPM"
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim bExists As Boolean
'    Dim objErrAssist As New ErrAssist
    Dim lIsolationMode As Long
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
        
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("ComboLists", lIsolationMode, lReleaseMode, bExists)
            
    Set spmProperty = spmGroup.CreateProperty(strGroupName, bExists)
            
    'If spmProperty.Value = "" Then
        spmProperty.Value = strComboListXML
    'End If
                
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
            
    Exit Sub
AddComboListToSPMVBErr:
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetXMLFromRecordSet(ByVal vrstComboRecSet As ADODB.Recordset) As String
' header ----------------------------------------------------------------------------------
' description:  Create XML data stream from field elements in ADO Recordsets.
'               Add any values derived from field elements to XML.
'
' pass:         vrstComboRecSet         ADO ValueId Recordset positioned at
'                                       the required record
'
' return:       GetXMLFromRecordSet string containing XML data stream
'
' maintenance -----------------------------------------------------------------------------
' Date          Developer        Comments
' 30/06/1999    Ivan Kemp        Initial Creation
' 02/07/1999    Ashley Smith     GetXMLFromRecordSet initial implementation
'------------------------------------------------------------------------------------------
On Error GoTo GetXMLFromRecordSetVbErr
    Dim strFunctionName As String
    strFunctionName = "GetXMLFromRecordSet"
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlInsertNode As IXMLDOMNode
    Dim objXmlValListNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim int0 As Integer
'    Dim objErrAssist As New ErrAssist
    Set objXmlElem = objXmlOut.createElement("LISTENTRY")
    Set objXmlInsertNode = objXmlOut.appendChild(objXmlElem)
    ' ************************************************************************
    '   The following creates an xml element for each field returned in the
    '   Value recordset. The element is named from the field (column) name and the
    '   element text is set from the field value.
    ' *****************************************************************************
    For int0 = 0 To (vrstComboRecSet.Fields.Count - 1)
        If Not IsNull(vrstComboRecSet.Fields.Item(int0).Value) Then
            Set objXmlElem = objXmlOut.createElement(vrstComboRecSet.Fields.Item(int0).Name)
            objXmlElem.Text = vrstComboRecSet.Fields.Item(int0).Value
            Set objXmlNode = objXmlInsertNode.appendChild(objXmlElem)
        End If
    Next
        
    ' Add in the validationlist element and move the validationtype node to be
    ' a child of the validationlist element
    Dim objXmlValNode As IXMLDOMElement
    Set objXmlElem = objXmlOut.createElement("VALIDATIONTYPELIST")
    Set objXmlValNode = objXmlInsertNode.appendChild(objXmlElem)
        
    Set objXmlNode = objXmlOut.getElementsByTagName("VALIDATIONTYPE").Item(0)
    If Not objXmlNode Is Nothing Then
        objXmlValNode.appendChild objXmlNode
    End If
    GetXMLFromRecordSet = objXmlOut.xml
    Set objXmlValListNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlInsertNode = Nothing
    Set objXmlOut = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetXMLFromRecordSetVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    ' log any error
    m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlInsertNode = Nothing
    Set objXmlOut = Nothing
    Set objXmlValListNode = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetComboValue(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing list which identifies an
'                   instance of the persistant data to be retrieved. Received in
'                   the format:
'                   <LIST>
'                       <GROUPNAME>Title</GROUPNAME>
'                       <VALUEID>1</VALUEID>
'                   </LIST>
' return:
'   GetComboValue   string containing XML data stream representation of
'                   data retrieved in the format:
'                   <LISTENTRY>
'                       <VALUEID>1</VALUEID>
'                       <VALUENAME>Mr</VALUENAME>
'                       <VALIDATIONTYPELIST>
'                           <VALIDATIONTYPE>M</VALIDATIONTYPE>
'                           ...
'                           <VALIDATIONTYPE>O</VALIDATIONTYPE>
'                       <VALIDATIONTYPELIST>
'                   </LISTENTRY>
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'   if no group name specified raise omiga4err108
'   if no valueID specified raise omiga4err109
' maintenance -----------------------------------------------------------------------------
' Date          Developer   Comments
' 07/09/1999    AS          Changed GetComboValue due to GetComboList reading
'                           multiple combo groups in one call
'------------------------------------------------------------------------------------------
On Error GoTo GetComboValueVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComboValue"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlRow As New FreeThreadedDOMDocument40
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlValList As IXMLDOMElement
    Dim objXmlRowNode As IXMLDOMElement
'    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objADOAssist As ADOAssist
    Dim cmd As New ADODB.Command
    Dim rstComboRecSet As New ADODB.Recordset
    Dim strSQL As String
    Dim strGroupName As String
    Dim strValueId As String
    Dim strComboEntry As String
    Dim strCurrentValueId As String
    Dim strPreviousValueId As String
    Dim blnSameValueId As Boolean
    Dim blnAddRow As Boolean
        
    objXmlIn.async = False
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlElem = objXmlIn.getElementsByTagName("LIST").Item(0)
        
    strGroupName = m_objXmlAssist.GetTagValue(objXmlElem, "GROUPNAME")
    strValueId = m_objXmlAssist.GetTagValue(objXmlElem, "VALUEID")
            
    If Not Len(strGroupName) > 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err108
    End If
    If Not Len(strValueId) > 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err109
    End If
    ' get combo entry information from the shared property manager
    If Not m_objContext Is Nothing Then
        strComboEntry = GetComboValueFromSPM(strGroupName, strValueId)
    End If
    If Not Len(strComboEntry) > 0 Then
        If m_objContext Is Nothing Then
            Set objADOAssist = New ADOAssist
        Else
            Set objADOAssist = m_objContext.CreateInstance(App.Title & ".ADOAssist")
        End If
        'DM 17/05/01 SYS2316
        ' DM 31/07/01 SYS2539 taken out runtime switch.
        Dim strFormatedGroupName As String
        'strFormatedGroupName = objSQLAssist.FormatString(strGroupName)
        strFormatedGroupName = strGroupName ' Don't need to put quotes around the string when using an ADO parameter
        cmd.ActiveConnection = objADOAssist.GetConnStr
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "usp_GetComboValue"
        ' Need to set the precision of a numeric data type for SQLServer
        Dim param As ADODB.Parameter
        Set param = cmd.CreateParameter("p_ValueID", adNumeric, adParamInput, , CInt(strValueId))
        param.Precision = 5
        cmd.Parameters.Append param
        cmd.Parameters.Append cmd.CreateParameter("p_GroupName", adBSTR, adParamInput, Len(strFormatedGroupName), strFormatedGroupName)
        ' execute the cmd and store the results.
        Set rstComboRecSet = cmd.Execute
        If Not rstComboRecSet.EOF Then
            ' DM Removed this as it causes an error in SQL Server, tested against Oracle ok.
            'rstComboRecSet.MoveFirst
                        
            While Not rstComboRecSet.EOF
                                
                strCurrentValueId = rstComboRecSet!VALUEID
                                
                blnSameValueId = StrComp(strCurrentValueId, strPreviousValueId, _
                                            vbTextCompare) = 0
                If (Not blnSameValueId) Then
                    Set objXmlRow = m_objXmlAssist.load(Me.GetXMLFromRecordSet _
                                    (rstComboRecSet), TypeName(Me), strFunctionName)
                    blnAddRow = True
                Else
                    Set objXmlValList = objXmlRowNode.getElementsByTagName _
                                            ("VALIDATIONTYPELIST").Item(0)
                    Set objXmlNode = objXmlOut.createElement("VALIDATIONTYPE")
                    objXmlNode.Text = rstComboRecSet!VALIDATIONTYPE
                    objXmlValList.appendChild objXmlNode
                                        
                    blnAddRow = False
                End If
                                                
                If blnAddRow Then
                    Set objXmlRowNode = objXmlOut.appendChild _
                                        (objXmlRow.documentElement)
                    blnAddRow = False
                End If
                                
                strPreviousValueId = strCurrentValueId
                rstComboRecSet.MoveNext
                Set objXmlRow = Nothing
            Wend
            rstComboRecSet.Close
        Else
            m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
        End If
    Else
        Set objXmlRow = m_objXmlAssist.load(strComboEntry, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlOut.appendChild(objXmlRow.documentElement)
    End If
    'rstComboRecSet.Close
    GetComboValue = objXmlOut.xml
    Set rstComboRecSet = Nothing
    Set cmd = Nothing
'    Set objErrAssist = Nothing
    Set objADOAssist = Nothing
    Set objSQLAssist = Nothing
'    Set objXmlAssist = Nothing
        
    Set objXmlIn = Nothing
    Set objXmlRow = Nothing
    Set objXmlOut = Nothing
    Set objXmlElem = Nothing
    Set objXmlNode = Nothing
    Set objXmlRowNode = Nothing
    Set objXmlValList = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Function
GetComboValueVbErr:
    
    ' Close the recordset if still open
        If Not rstComboRecSet Is Nothing Then
                If rstComboRecSet.State = adStateOpen Then
                        rstComboRecSet.Close
                End If
        End If
    Set cmd = Nothing
    Set rstComboRecSet = Nothing
'    Set objErrAssist = Nothing
    Set objADOAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlRow = Nothing
    Set objXmlOut = Nothing
    Set objXmlNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlRowNode = Nothing
    Set objXmlValList = Nothing
        
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetComboValueFromSPM(strGroupName As String, strValueId As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetComboValueFromSPMVBErr:
    
    Dim strFunctionName As String
    strFunctionName = "GetComboValueFromSPM"
    Dim strComboList As String
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    strComboList = GetComboListFromSPM(strGroupName)
    If Len(strComboList) > 0 Then
        Set objXmlData = m_objXmlAssist.load(strComboList, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlData.selectSingleNode("//LISTENTRY[VALUEID=""" & strValueId & """]")
    End If
    If Not objXmlNode Is Nothing Then
        GetComboValueFromSPM = objXmlNode.xml
    End If
    Set objXmlNode = Nothing
    Set objXmlData = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetComboValueFromSPMVBErr:
    Set objXmlNode = Nothing
    Set objXmlData = Nothing
'    Set objXmlAssist = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:  Add any derived values to XML.
' pass:         vstrXMLData         base XML data stream
'                                   as:
'                                   <tablename>
'                                       <element1>element1 value</element1>
'                                       <elementn>elementn value</elementn>
' return:       AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
    Dim strFunctionName As String
    strFunctionName = "AddDerivedData"
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    'Dim objXmlElem As IXMLDOMElement
    objXmlDoc.setProperty "NewParser", True
    objXmlDoc.validateOnParse = False
    objXmlDoc.loadXML vstrXMLData
    If objXmlDoc.parseError.errorCode <> 0 Then
        RaiseParserError strFunctionName, objXmlDoc.parseError
    End If
                   
    AddDerivedData = objXmlDoc.xml
    Set objXmlDoc = Nothing
End Function
Private Sub RaiseParserError(ByVal strCallingFunction As String, _
                             ByVal objParseError As IXMLDOMParseError)
' header ----------------------------------------------------------------------------------
' description:  Raise XML parser error
' pass:         strCallingFunction
'               objParseError
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo RaiseParserErrorVbErr:
    
    Dim strFunctionName As String
    strFunctionName = "AddDerivedData"
'    Dim objErrAssist As New ErrAssist
    Dim strParserError As String
    strParserError = m_objErrAssist.FormatParserError(objParseError)
    m_objErrAssist.RaiseError TypeName(Me), strCallingFunction, omiga4XMLParserError, strParserError
'    Set objErrAssist = Nothing
    Exit Sub
RaiseParserErrorVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function GetKeyString(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'       Check all keys are present in the XML, and return a string in the
'       format "KeyFieldA = ValueX and KeyFieldB = ValueY and KeyFieldC = ValueZ"
'       Raise an error if not all keys have values specified
'       Should be used by Delete(), Update() and GetData()
' pass:
' Raise Errors:
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo errhandler
    
    Dim strFunctionName As String
    strFunctionName = "GetKeyString"
'    Dim objErrAssist As New ErrAssist
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.setProperty "NewParser", True
    objXmlDoc.validateOnParse = False
    objXmlDoc.loadXML vstrXMLRequest
    If objXmlDoc.parseError.errorCode <> 0 Then
        RaiseParserError strFunctionName, objXmlDoc.parseError
    End If
    Dim strKeys As String
' TODO ************************************************************************
'   Get the values for each key. If any are unspecified,
'   raise error omiga4InvalidKeyString
' *****************************************************************************
    
    GetKeyString = strKeys
    Set objXmlDoc = Nothing
'    Set objErrAssist = Nothing
    Exit Function
errhandler:
    
    Set objXmlDoc = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Public Function GetComboText(vstrComboName As String, vstrValueId As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the combo text (i.e. ValueName) for a given combo name and ValueId
' pass:
'   vstrComboName       value for ComboValue.ComboName
'   vstrValueId         value for ComboValue.ValueId
' return:
'   GetComboText        required value of ComboValue.ValueName
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetComboTextVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComboText"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
        
    Dim strComboText As String
    strComboText = GetComboTextFromSPM(vstrComboName, vstrValueId)
    If (Len(strComboText) = 0) Then
        Dim strComboList As String
        strComboList = "<LIST><LISTNAME>" & vstrComboName & "</LISTNAME></LIST>"
        Me.GetComboList strComboList
        strComboText = GetComboTextFromSPM(vstrComboName, vstrValueId)
    End If
        
    'Dim objXmlDocRequest As New FreeThreadedDOMDocument40
    'Dim objXmlRootNode As IXMLDOMNode
    'Dim objXmlElem As IXMLDOMElement
    'Set objXmlElem = objXmlDocRequest.createElement("LIST")
    'Set objXmlRootNode = objXmlDocRequest.appendChild(objXmlElem)
    'Set objXmlElem = objXmlDocRequest.createElement("GROUPNAME")
    'objXmlElem.Text = vstrComboName
    'objXmlRootNode.appendChild objXmlElem
    'Set objXmlElem = objXmlDocRequest.createElement("VALUEID")
    'objXmlElem.Text = vstrValueId
    'objXmlRootNode.appendChild objXmlElem
    'strReply = GetComboValue(objXmlDocRequest.xml)
    'Dim objXmlReply As New FreeThreadedDOMDocument40
    'Dim objXmlAssist As New XMLAssist
    'Set objXmlReply = objXmlAssist.load(strReply, TypeName(Me), strFunctionName)
    'Dim xmlNodeList As IXMLDOMNodeList
    'Set xmlNodeList = objXmlReply.getElementsByTagName("LISTENTRY")
    'If (xmlNodeList.length <> 1) Then
    '    objErrAssist.RaiseError TypeName(Me), strFunctionName, _
    '        omiga4RecordNotFound
    'End If
    'Dim xmlElem As IXMLDOMElement
    'Set xmlElem = xmlNodeList.Item(0)
    'Dim blnTagExists As Boolean
    'strComboText = objXmlAssist.GetTagValue(xmlElem, "VALUENAME", blnTagExists)
    If Len(strComboText) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4RecordNotFound
    End If
    GetComboText = strComboText
'    Set objContext = Nothing
    Exit Function
GetComboTextVbErr:
'    Set objContext = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function GetComboTextFromSPM(strGroupName As String, strValueId As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the combo text (i.e. ValueName) from the Shared Property Manager
' pass:
'   vstrComboName       value for ComboValue.ComboName
'   vstrValueId         value for ComboValue.ValueId
' return:
'   GetComboTextFromSPM required value of ComboValue.ValueName
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetComboTextFromSPMVBErr:
    
    Dim strFunctionName As String
    strFunctionName = "GetComboTextFromSPM"
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
'    Dim objErrAssist As New ErrAssist
    Dim bExists As Boolean
    Dim lIsolationMode As Long
    'PSC 11/08/2000  SYS1430 Back out SYS1409
    lIsolationMode = LockSetGet
    Dim lReleaseMode As Long
    lReleaseMode = Process
      
    Set spmMgr = CreateObject _
        (gstrMSGSPM)
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("ComboLists", lIsolationMode, lReleaseMode, bExists)
            
    Dim strEntry As String
    ' SPM Entry is held as GroupName.ValueID = ComboText
    strEntry = strGroupName & "." & strValueId
    Set spmProperty = spmGroup.CreateProperty(strEntry, bExists)
            
    If bExists Then
        GetComboTextFromSPM = spmProperty.Value
    Else
        spmProperty.Value = ""
    End If
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
    Exit Function
GetComboTextFromSPMVBErr:
    
    Set spmMgr = Nothing
    Set spmGroup = Nothing
    Set spmProperty = Nothing
'    Set objErrAssist = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function IsItemInValidation(ByVal vstrGroupName As String, _
    ByVal vstrValueId As String, ByVal vstrValidationType As String) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   This method establishes whether the input validation types matches the validation
'   type for the list of items in a combo list.  If it does then it returns a true
'   value.
' pass:
'   vstrXMLRequest      Format:
'                       <LIST>
'                           <GROUPNAME></GROUPNAME>
'                           <VALUEID></VALUEID>
'                           <VALIDATIONTYPELIST>
'                               <VALIDATIONTYPE></VALIDATIONTYPE>
'                               ...
'                               <VALIDATIONTYPE></VALIDATIONTYPE>
'                           </VALIDATIONTYPELIST>
'                       <LIST>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo IsItemInValidationVbErr
    Dim strFunctionName As String
    strFunctionName = "IsItemInValidation"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlComboNodeList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement
    Dim blnItemFound As Boolean
    Dim strResponse As String
    Dim intComboNode, intNumComboNodes As Integer
    If Len(vstrGroupName) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Group Name must be specified"
    End If
    If Len(vstrValueId) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Value Id must be specified"
    End If
    If Len(vstrValidationType) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Validation Type must be specified"
    End If
    Set xmlListElem = xmlDoc.createElement("LIST")
    xmlDoc.appendChild xmlListElem
    Set xmlElem = xmlDoc.createElement("GROUPNAME")
    xmlElem.Text = vstrGroupName
    xmlListElem.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("VALUEID")
    xmlElem.Text = vstrValueId
    xmlListElem.appendChild xmlElem
    strResponse = Me.GetComboValue(xmlDoc.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set xmlComboNodeList = xmlDoc.getElementsByTagName("VALIDATIONTYPE")
    intNumComboNodes = xmlComboNodeList.Length
    If intNumComboNodes = 0 Then
        blnItemFound = False
    Else
        Do While blnItemFound = False And intComboNode < intNumComboNodes
            If xmlComboNodeList.Item(intComboNode).Text = vstrValidationType Then
                blnItemFound = True
            End If
            intComboNode = intComboNode + 1
        Loop
    End If
    IsItemInValidation = blnItemFound
'    Set objContext = Nothing
    Set xmlDoc = Nothing
'    Set objErrAssist = Nothing
    Set xmlComboNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    Exit Function
IsItemInValidationVbErr:
'    Set objContext = Nothing
    Set xmlDoc = Nothing
    Set xmlComboNodeList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Function GetFirstComboValueId(ByVal vstrGroupName As String, _
    ByVal vstrValidationType As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a single combo value id as a simple string rather than as an xml response.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetFirstComboValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "GetFirstComboValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim xmlDoc As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
            
    Dim strResponse As String
    strResponse = GetComboValueId(vstrGroupName, vstrValidationType)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    GetFirstComboValueId = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "VALUEID")
    Set xmlDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
GetFirstComboValueIdVbErr:
    
    Set xmlDoc = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetComboValueId(ByVal vstrGroupName As String, _
                                ByVal vstrValidationType As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrGroupName       Combo group name to be searched
'   vstrValidationType  Validation type to search for in group
' return:
'   Response Format:
'        <VALUEIDLIST>
'            <VALUEID></VALUEID>
'            <VALUEID></VALUEID>
'                ...
'            <VALUEID></VALUEID>
'        </VALUEIDLIST>
'------------------------------------------------------------------------------------------
On Error GoTo GetComboValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComboValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlValueListElem As IXMLDOMElement
    Dim xmlValueIdElem As IXMLDOMElement
    Dim xmlEntryList As IXMLDOMNodeList
    Dim xmlValidationList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement
    Dim strResponse As String
    Dim strValueId As String
    If Len(vstrGroupName) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Group Name must be specified"
    End If
    If Len(vstrValidationType) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Validation Type must be specified"
    End If
    Set xmlListElem = xmlDoc.createElement("LIST")
    xmlDoc.appendChild xmlListElem
    Set xmlElem = xmlDoc.createElement("LISTNAME")
    xmlElem.Text = vstrGroupName
    xmlListElem.appendChild xmlElem
        
    strResponse = Me.GetComboList(xmlDoc.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set xmlEntryList = xmlDoc.getElementsByTagName("LISTENTRY")
    Dim blnIsMatched As Boolean
    Dim intNoOfEntries As Integer
    Dim intEntryIndex As Integer
    Dim intNoOfValidations As Integer
    Dim intValidationIndex As Integer
    Set xmlValueListElem = xmlOut.createElement("VALUEIDLIST")
    xmlOut.appendChild xmlValueListElem
    intEntryIndex = 0
    intNoOfEntries = xmlEntryList.Length
    ' Go through each validation type of each entry until we find a match
    Do While intEntryIndex < intNoOfEntries
        Set xmlElem = xmlEntryList.Item(intEntryIndex)
        Set xmlValidationList = xmlElem.getElementsByTagName("VALIDATIONTYPE")
        intValidationIndex = 0
        intNoOfValidations = xmlValidationList.Length
        blnIsMatched = False
        Do While intValidationIndex < intNoOfValidations And blnIsMatched = False
            If xmlValidationList.Item(intValidationIndex).Text = vstrValidationType Then
                Set xmlValueIdElem = xmlOut.createElement("VALUEID")
                xmlValueIdElem.Text = m_objXmlAssist.GetTagValue(xmlElem, "VALUEID")
                xmlValueListElem.appendChild xmlValueIdElem
                blnIsMatched = True
            End If
            intValidationIndex = intValidationIndex + 1
        Loop
        intEntryIndex = intEntryIndex + 1
    Loop
    GetComboValueId = xmlOut.xml
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlEntryList = Nothing
    Set xmlValidationList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlValueListElem = Nothing
    Set xmlValueIdElem = Nothing
    Exit Function
GetComboValueIdVbErr:
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlEntryList = Nothing
    Set xmlValidationList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlValueListElem = Nothing
    Set xmlValueIdElem = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
'BMIDS00336 MDC 12/09/2002
Function GetFirstComboValueIdFromValueName(ByVal vstrGroupName As String, _
    ByVal vstrValueName As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a single combo value id as a simple string rather than as an xml response.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetFirstComboValueIdFromValueNameVbErr
    Dim strFunctionName As String
    strFunctionName = "GetFirstComboValueIdFromValueName"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim xmlDoc As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
            
    Dim strResponse As String
    strResponse = GetComboValueIdFromValueName(vstrGroupName, vstrValueName)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    GetFirstComboValueIdFromValueName = m_objXmlAssist.GetTagValue(xmlDoc.documentElement, "VALUEID")
    Set xmlDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
GetFirstComboValueIdFromValueNameVbErr:
    
    Set xmlDoc = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetComboValueIdFromValueName(ByVal vstrGroupName As String, _
                                ByVal vstrValueName As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrGroupName       Combo group name to be searched
'   vstrValueName       Text for the combo value name
' return:
'   Response Format:
'        <VALUEIDLIST>
'            <VALUEID></VALUEID>
'            <VALUEID></VALUEID>
'                ...
'            <VALUEID></VALUEID>
'        </VALUEIDLIST>
'------------------------------------------------------------------------------------------
On Error GoTo GetComboValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "GetComboValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
'    Dim objXmlAssist As New XMLAssist
    Dim xmlDoc As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlValueListElem As IXMLDOMElement
    Dim xmlValueIdElem As IXMLDOMElement
    Dim xmlEntryList As IXMLDOMNodeList
    Dim xmlValidationList As IXMLDOMNodeList
    Dim xmlElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement
    Dim strResponse As String
    Dim strValueId As String
    If Len(vstrGroupName) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Group Name must be specified"
    End If
    If Len(vstrValueName) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4InvalidParameter, "Value Name must be specified"
    End If
    Set xmlListElem = xmlDoc.createElement("LIST")
    xmlDoc.appendChild xmlListElem
    Set xmlElem = xmlDoc.createElement("LISTNAME")
    xmlElem.Text = vstrGroupName
    xmlListElem.appendChild xmlElem
        
    strResponse = Me.GetComboList(xmlDoc.xml)
    Set xmlDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    Set xmlEntryList = xmlDoc.getElementsByTagName("LISTENTRY")
    Dim blnIsMatched As Boolean
    Dim intNoOfEntries As Integer
    Dim intEntryIndex As Integer
    Dim intNoOfValidations As Integer
    Dim intValidationIndex As Integer
    Set xmlValueListElem = xmlOut.createElement("VALUEIDLIST")
    xmlOut.appendChild xmlValueListElem
    intEntryIndex = 0
    intNoOfEntries = xmlEntryList.Length
    ' Go through each validation type of each entry until we find a match
    Do While intEntryIndex < intNoOfEntries
        Set xmlElem = xmlEntryList.Item(intEntryIndex)
        Set xmlValidationList = xmlElem.getElementsByTagName("VALUENAME")
        intValidationIndex = 0
        intNoOfValidations = xmlValidationList.Length
        blnIsMatched = False
        Do While intValidationIndex < intNoOfValidations And blnIsMatched = False
            If UCase$(xmlValidationList.Item(intValidationIndex).Text) = UCase$(vstrValueName) Then
                Set xmlValueIdElem = xmlOut.createElement("VALUEID")
                xmlValueIdElem.Text = m_objXmlAssist.GetTagValue(xmlElem, "VALUEID")
                xmlValueListElem.appendChild xmlValueIdElem
                blnIsMatched = True
            End If
            intValidationIndex = intValidationIndex + 1
        Loop
        intEntryIndex = intEntryIndex + 1
    Loop
    GetComboValueIdFromValueName = xmlOut.xml
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlEntryList = Nothing
    Set xmlValidationList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    Set xmlValueListElem = Nothing
    Set xmlValueIdElem = Nothing
    Exit Function
GetComboValueIdVbErr:
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlEntryList = Nothing
    Set xmlValidationList = Nothing
    Set xmlElem = Nothing
    Set xmlListElem = Nothing
    Set xmlValueListElem = Nothing
    Set xmlValueIdElem = Nothing
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    m_objErrAssist.AddToErrSource strFunctionName
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
'BMIDS00336 MDC 12/09/2002 - End
Public Function GetNewLoanValue() As String
' header ----------------------------------------------------------------------------------
' description:
'   Return the first combo value for an entry found to be a New Loan
' pass:
'
' return:       New Loan Value
'------------------------------------------------------------------------------------------
On Error GoTo GetNewLoanValueVbErr
    Dim strFunctionName As String
    strFunctionName = "GetNewLoanValue"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
            
    Dim strResponse As String
    strResponse = GetComboValueId("TypeOfMortgage", "N")
    Set xmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    GetNewLoanValue = _
                    m_objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "VALUEID")
    Set xmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
GetNewLoanValueVbErr:
    
    Set xmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function IsNewLoan(ByVal vstrTypeOfApplication) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if vstrTypeOfApplication corresponds to a New Loan.
' pass:
'   vstrTypeOfApplication
' return:
'   IsNewLoan
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsNewLoanVbErr
    Dim strFunctionName As String
    strFunctionName = "IsNewLoan"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
        
    IsNewLoan = IsItemInValidation("TypeOfMortgage", vstrTypeOfApplication, "N")
'    Set objErrAssist = Nothing
    Exit Function
IsNewLoanVbErr:
        
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
End Function
Public Function IsFurtherAdvance(ByVal vstrTypeOfApplication) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if vstrTypeOfApplication corresponds to a Further Advance
' pass:
'   vstrTypeOfApplication
' return:
'   IsFurtherAdvance
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsFurtherAdvanceVbErr
    Dim strFunctionName As String
    strFunctionName = "IsFurtherAdvance"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
        
    IsFurtherAdvance = IsItemInValidation("TypeOfMortgage", vstrTypeOfApplication, "F")
'    Set objErrAssist = Nothing
    Exit Function
IsFurtherAdvanceVbErr:
        
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
End Function

'BBG123 Start
Public Function IsTransferOfEquity(ByVal vstrTypeOfApplication) As Boolean

On Error GoTo IsTransferOfEquityVbErr
    Dim strFunctionName As String
    strFunctionName = "IsTransferOfEquity"
    IsTransferOfEquity = IsItemInValidation("TypeOfMortgage", vstrTypeOfApplication, "T")
    Exit Function
IsTransferOfEquityVbErr:
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
        
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
End Function
'BBG123 End

Public Function IsRemortgage(ByVal vstrTypeOfApplication) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if vstrTypeOfApplication corresponds to a Remortgage.
' pass:
'   vstrTypeOfApplication
' return:
'   IsRemortgage
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo IsRemortgageVbErr
    Dim strFunctionName As String
    strFunctionName = "IsRemortgage"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
        
    IsRemortgage = IsItemInValidation("TypeOfMortgage", vstrTypeOfApplication, "R")
'    Set objErrAssist = Nothing
    Exit Function
IsRemortgageVbErr:
        
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
End Function
Public Function IsReInspection(ByVal vstrValuationType) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if vstrValuationType corresponds to a Reinspection
'
' pass:
'   vstrValuationType       Valuation type
'
' return:                   True if Reinspection
'------------------------------------------------------------------------------------------
On Error GoTo IsReInspectionVbErr
    Dim strFunctionName As String
    strFunctionName = "IsReInspection"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
           
    IsReInspection = IsItemInValidation("ValuationType", vstrValuationType, "R")
'    Set objErrAssist = Nothing
    Exit Function
IsReInspectionVbErr:
        
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
End Function
Public Function GetFirstComboValidation(ByVal vstrComboName As String, _
                                        ByVal vstrValueId As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the combo text (i.e. ValueName) for a given combo name and ValueId
' pass:
'   vstrComboName       value for ComboValue.ComboName
'   vstrValueId         value for ComboValue.ValueId
' return:
'                       The first validtion for the valueid
'------------------------------------------------------------------------------------------
On Error GoTo GetFirstComboValidationVbErr
    Dim strFunctionName As String
    strFunctionName = "GetFirstComboValidation"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
    Dim objXmlDocRequest As New FreeThreadedDOMDocument40
    Dim objXmlRootNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Set objXmlElem = objXmlDocRequest.createElement("LIST")
    Set objXmlRootNode = objXmlDocRequest.appendChild(objXmlElem)
    Set objXmlElem = objXmlDocRequest.createElement("GROUPNAME")
    objXmlElem.Text = vstrComboName
    objXmlRootNode.appendChild objXmlElem
    Set objXmlElem = objXmlDocRequest.createElement("VALUEID")
    objXmlElem.Text = vstrValueId
    objXmlRootNode.appendChild objXmlElem
    Dim strReply As String
    strReply = GetComboValue(objXmlDocRequest.xml)
    Dim objXmlReply As New FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
    Set objXmlReply = m_objXmlAssist.load(strReply, TypeName(Me), strFunctionName)
    Dim xmlNodeList As IXMLDOMNodeList
    Set xmlNodeList = objXmlReply.getElementsByTagName("LISTENTRY")
    If (xmlNodeList.Length <> 1) Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4RecordNotFound
    End If
    Dim xmlElem As IXMLDOMElement
    Set xmlElem = xmlNodeList.Item(0)
    Dim blnTagExists As Boolean
    Dim strComboText As String
    strComboText = m_objXmlAssist.GetTagValue(xmlElem, "VALIDATIONTYPE ", blnTagExists)
    If Not blnTagExists Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4RecordNotFound
    End If
    GetFirstComboValidation = strComboText
'    Set objContext = Nothing
    Set objXmlDocRequest = Nothing
    Set objXmlRootNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlReply = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    Exit Function
GetFirstComboValidationVbErr:
'    Set objContext = Nothing
    Set objXmlDocRequest = Nothing
    Set objXmlRootNode = Nothing
    Set objXmlElem = Nothing
    Set objXmlReply = Nothing
    Set xmlNodeList = Nothing
    Set xmlElem = Nothing
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function
Private Function IComboDO_GetNewLoanValue() As String
    IComboDO_GetNewLoanValue = GetNewLoanValue
End Function
Private Function IComboDO_GetQuickQuoteLocationValueId() As String
'Public Function GetQuickQuoteLocationValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Return the combo value for Quick Quote location
' pass:
'
' return:   Quick Quote Location Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo IComboDO_GetQuickQuoteLocationValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "IComboDO_GetQuickQuoteLocationValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlResponseDoc As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objComboDO As ComboDO
    Dim objGlobalParamDO As GlobalParameterDO
        
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    Dim strResponse As String
    Dim strLocationCode As String
    Dim strLocationValue As String
    strResponse = objGlobalParamDO.GetCurrentParameter("QQLocation")
    Set objXmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strLocationCode = m_objXmlAssist.GetTagValue(objXmlResponseDoc.documentElement, "STRING")
    strResponse = objComboDO.GetComboValueId("PropertyLocation", strLocationCode)
    Set objXmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strLocationValue = m_objXmlAssist.GetTagValue(objXmlResponseDoc.documentElement, "VALUEID")
    IComboDO_GetQuickQuoteLocationValueId = strLocationValue
    Set objComboDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objXmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
IComboDO_GetQuickQuoteLocationValueIdVbErr:
    
    Set objComboDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objXmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IComboDO_GetQuickQuoteValuationTypeValueId() As String
'Public Function GetQuickQuoteValuationTypeValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Return the combo value for Quick Quote location
' pass:
'
' return:   Quick Quote Location Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo IComboDO_GetQuickQuoteValuationTypeValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "IComboDO_GetQuickQuoteValuationTypeValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objXmlResponseDoc As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objComboDO As ComboDO
    Dim objGlobalParamDO As GlobalParameterDO
        
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objGlobalParamDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    Dim strResponse As String
    Dim strValuationCode As String
    Dim strValuationValue As String
    strResponse = objGlobalParamDO.GetCurrentParameter("QQValuationType")
    Set objXmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strValuationCode = m_objXmlAssist.GetTagValue(objXmlResponseDoc.documentElement, "STRING")
    strResponse = objComboDO.GetComboValueId("ValuationType", strValuationCode)
    Set objXmlResponseDoc = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    strValuationValue = m_objXmlAssist.GetTagValue(objXmlResponseDoc.documentElement, "VALUEID")
    IComboDO_GetQuickQuoteValuationTypeValueId = strValuationValue
    Set objComboDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objXmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
IComboDO_GetQuickQuoteValuationTypeValueIdVbErr:
    
    Set objComboDO = Nothing
    Set objGlobalParamDO = Nothing
    Set objXmlResponseDoc = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IComboDO_GetDormantLegalFeeValueId() As String
'Public Function GetDormantLegalFeeValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Return the combo values for the entries found to be a dormant legal fee
' pass:
'
' return:       Dormant Legal Fee Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo IComboDO_GetDormantLegalFeeValueIdVbErr
    Dim strFunctionName As String
    strFunctionName = "IComboDO_GetDormantLegalFeeValueId"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim objComboDO As ComboDO
    Dim objXmlResponse As FreeThreadedDOMDocument40
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
        
    If m_objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("LegalFeeType", "D")
    Set objXmlResponse = m_objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
    IComboDO_GetDormantLegalFeeValueId = _
                    m_objXmlAssist.GetTagValue(objXmlResponse.documentElement, "VALUEID")
    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
     
    Exit Function
IComboDO_GetDormantLegalFeeValueIdVbErr:
    
    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
'    Set objXmlAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
        
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
'    Set objErrAssist = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'SR 07/04/2004 : BBG132 -
Private Function IComboDO_GetRegSaleRelatedComboList(ByVal vxmlElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo IComboDO_GetRegSaleRelatedComboListVbErr
    
    Dim strFunctionName As String
    strFunctionName = "IComboDO_GetRegSaleRelatedComboList"
    
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlRow As New FreeThreadedDOMDocument40
    Dim xmlListNode As IXMLDOMNode
    Dim xmlListNameNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMElement
    Dim xmlValList As IXMLDOMNode
    Dim xmlRowNode As IXMLDOMElement

'    Dim objXmlAssist As XMLAssist
    Dim objADOAssist As ADOAssist
'    Dim objErrAssist As ErrAssist
    Dim cmd As New ADODB.Command
    Dim rstComboRecSet As New ADODB.Recordset
    Dim strRegSaleQuestionType As String
    Dim strCurrentGroupName As String
    Dim strCurrentValueId As String
    Dim strPreviousGroupName As String
    Dim strPreviousValueId As String
    Dim strGroupName As String
    Dim blnSameGroupName As Boolean
    Dim blnSameValueId As Boolean
    Dim blnAddRow As Boolean
    
'    Set objXmlAssist = New XMLAssist
    Set objADOAssist = New ADOAssist
'    Set objErrAssist = New ErrAssist
    
    strRegSaleQuestionType = m_objXmlAssist.GetTagValue(vxmlElement, "REGULATEDSALEQUESTIONTYPE")
    cmd.ActiveConnection = objADOAssist.GetConnStr
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_GetRegSaleQuestionCombos"
    cmd.Parameters.Append cmd.CreateParameter("p_AdvisedStatus", adBSTR, adParamInput, Len(strRegSaleQuestionType), strRegSaleQuestionType)
    Set rstComboRecSet = cmd.Execute
    
    If Not rstComboRecSet.EOF Then
        Set xmlNode = xmlOut.createElement("LIST")
        Set xmlListNode = xmlOut.appendChild(xmlNode)
    
        While Not rstComboRecSet.EOF
            strCurrentGroupName = rstComboRecSet!GROUPNAME
            strCurrentValueId = rstComboRecSet!VALUEID
            
            blnSameGroupName = StrComp(strCurrentGroupName, strPreviousGroupName, vbTextCompare) = 0
            blnSameValueId = StrComp(strCurrentValueId, strPreviousValueId, vbTextCompare) = 0
            
            If Not blnSameGroupName Then
                Set xmlNode = xmlOut.createElement("LISTNAME")
                xmlNode.setAttribute "NAME", strCurrentGroupName
                Set xmlListNameNode = xmlListNode.appendChild(xmlNode)
                Set xmlRow = m_objXmlAssist.load(Me.GetXMLFromRecordSet _
                                    (rstComboRecSet), TypeName(Me), strFunctionName)
                blnAddRow = True
            ElseIf (blnSameGroupName And Not blnSameValueId) Then
                Set xmlRow = m_objXmlAssist.load(Me.GetXMLFromRecordSet _
                                    (rstComboRecSet), TypeName(Me), strFunctionName)
                blnAddRow = True
            Else
                Set xmlValList = xmlRowNode.getElementsByTagName("VALIDATIONTYPELIST").Item(0)
                Set xmlNode = xmlOut.createElement("VALIDATIONTYPE")
                xmlNode.Text = rstComboRecSet!VALIDATIONTYPE
                xmlValList.appendChild xmlNode
                                        
                blnAddRow = False
            End If
            
            If blnAddRow Then
                Set xmlRowNode = xmlListNameNode.appendChild(xmlRow.documentElement)
                blnAddRow = False
            End If
            
            strPreviousGroupName = strCurrentGroupName
            strPreviousValueId = strCurrentValueId
            rstComboRecSet.MoveNext
            Set xmlRow = Nothing
        Wend
    Else
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    rstComboRecSet.Close
    Set IComboDO_GetRegSaleRelatedComboList = xmlOut.documentElement

    Set rstComboRecSet = Nothing
    Set cmd = Nothing
'    Set objErrAssist = Nothing
'    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlRow = Nothing
    Set xmlNode = Nothing
    Set xmlListNode = Nothing
    Set xmlListNameNode = Nothing
    Set xmlValList = Nothing
    Set xmlRowNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
'    Set objContext = Nothing
    Exit Function

IComboDO_GetRegSaleRelatedComboListVbErr:
    ' Close the recordset if still open
    If Not rstComboRecSet Is Nothing Then
        If rstComboRecSet.State = adStateOpen Then
            rstComboRecSet.Close
        End If
    End If
     
    Set rstComboRecSet = Nothing
    Set cmd = Nothing
'    Set objXmlAssist = Nothing
    Set objADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlRow = Nothing
    Set xmlNode = Nothing
    Set xmlListNode = Nothing
    Set xmlListNameNode = Nothing
    Set xmlValList = Nothing
    Set xmlRowNode = Nothing
     
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
'    Set objErrAssist = Nothing
'    Set objContext = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
