VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuickQuoteApplicantDetailsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'QuickQuoteApplicantDetailsDO' with actual class name
' *****************************************************************************

'Workfile:      QuickQuoteApplicantDetailsDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String)
'               Public Sub Update(vstrXMLRequest As String)
'               Public Sub Delete(vstrXMLRequest As String)
'               Public Function GetData(vstrXMLRequest As String) As String
'               Public Function GetXMLFromRecordSet(vrstRecordSet As ADODB.Recordset) As String
'               Public Function AddDerivedData(vstrXMLData As String) As String
'Dependencies:  ADOAssist
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MCS    17/08/99 Created
'MCS    02/09/99 GetQuickQuoteData Delete Added. GetData, Update removed
'MCS    28/09/99    Updated following code reviews
'JLD    22/10/99    Change for issue 19, don't return an error unless all applicants have no data
'LD     07/11/00 Explicity close recordsets
'LD     07/11/00 Explicity destroy command objects
'LD     11/06/01 SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'------------------------------------------------------------------------------------------
Option Explicit
Implements IQuickQuoteApplicantDetailsDO
Implements ObjectControl

' Private data
Private m_objXMLAssist As XMLAssist
Private m_objContext As ObjectContext
Private m_objSQLAssist As SQLAssist
Private m_objErrAssist As ErrAssist
 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Const APPLICANT_SELF_EMPLOYED As String = "S"

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog  Date     Description
'DF    02/07/02 BMIDS00134 - Have applied a fix to GetQuickQuoteData proc to ensure 1st applicant
'                details are blanked out before we search for details of the 2nd applicant.
'MDC   01/11/02 BMIDS00654 Added AMOUNTREQUESTED
'SR    21/11/02 BMIDS01050 [JR SYSMCP1221]- modified method 'GetMainSelfEmployment'
'               - Remove UPPER function for SQLServer
'------------------------------------------------------------------------------------------

Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------

On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
        
    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
           
    objDOAssist.Create vstrXMLRequest, LoadData
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If

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

Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
        
    Dim objDOAssist As DOAssist
    
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    objDOAssist.Update vstrXMLRequest, LoadData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Sub

UpdateVbErr:
    
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

Public Sub Delete(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance to be deleted
' Raise Errors:
'       omiga4RecordNotFound
'       omiga4InvalidKeyString
'       parser errors
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "Delete"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
                
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
                
    Dim objDOAssist As DOAssist
    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    objDOAssist.Delete vstrXMLRequest, LoadData
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Sub

DeleteVbErr:

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

Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetData"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
        
    Dim objDOAssist  As DOAssist
                            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                            
    GetData = objDOAssist.GetData(vstrXMLRequest, LoadData)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objDOAssist = Nothing
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function
    
GetDataVbErr:
        
    GetData = ""
    
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

End Function


Public Function GetQuickQuoteData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetQuickQuoteDataVbErr

    Const strFunctionName As String = "GetQuickQuoteData"
    
    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    Dim objDOAssist   As DOAssist
    Dim objQQOutgoingsDO   As QuickQuoteOutgoingsDO
    Dim objQQPersonalDebtsDO   As QuickQuotePersonalDebtsDO
    Dim objSQLAssist As New SQLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
        Set objQQOutgoingsDO = New QuickQuoteOutgoingsDO
        Set objQQPersonalDebtsDO = New QuickQuotePersonalDebtsDO
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objQQOutgoingsDO = objContext.CreateInstance(App.Title & ".QuickQuoteOutgoingsDO")
        Set objQQPersonalDebtsDO = objContext.CreateInstance(App.Title & ".QuickQuotePersonalDebtsDO")
    End If
    
    Dim objXmlAssist   As New XMLAssist
        
    Dim objXmlQQApplicantDetailsDoc As New FreeThreadedDOMDocument40
    objXmlQQApplicantDetailsDoc.validateOnParse = False
    objXmlQQApplicantDetailsDoc.setProperty "NewParser", True
    Dim objXmlQQPersonalDebtsDoc As New FreeThreadedDOMDocument40
    objXmlQQPersonalDebtsDoc.validateOnParse = False
    objXmlQQPersonalDebtsDoc.setProperty "NewParser", True
    Dim objXmlQQOutgoingsDoc As New FreeThreadedDOMDocument40
    objXmlQQOutgoingsDoc.validateOnParse = False
    objXmlQQOutgoingsDoc.setProperty "NewParser", True
    
    Dim XmlOutDoc As New FreeThreadedDOMDocument40
    XmlOutDoc.validateOnParse = False
    XmlOutDoc.setProperty "NewParser", True
    Dim objXmlDoc  As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    
    Dim objQQApplicantDetailsNode As IXMLDOMNode
    Dim objQQPersonalDebtsNode As IXMLDOMNode
    Dim objQQOutgoingsNode As IXMLDOMNode
        
    Dim strQQApplicantDetails As String
    Dim strQQPersonalDebtsList As String
    Dim strQQOutgoingsList As String
        
    Dim objXmlOut As IXMLDOMNode
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim strAppNo As String
    Dim strAppFFNo As String
    Dim strCustNo As String
    Dim strCustVerNo As String
    Dim strFixedXML As String
    Dim strSourceTable As String
    Dim strXML As String
    Dim strTable As String
    Dim nloop As Integer
    Dim ERRNUMBER As Long
    'Dim blnRaiseError As Boolean
        
    Set objXmlDoc = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    'No of Elements that match these Keys
         
    Set xmlNodeList = objXmlDoc.getElementsByTagName("CUSTOMERROLE")

    If (xmlNodeList.length < 1) Then
         ' "Invalid parameter value"
         objErrAssist.RaiseError TypeName(Me), strFunctionName, _
             omiga4InvalidParameter, "No rows found"
    End If
            
    'This is the list we are going to return
    Set objXmlOut = XmlOutDoc.createElement("QUICKQUOTEAPPLICANTDETAILSLIST")
      
    Dim objXmlElem As IXMLDOMElement
    Dim blnDataFound As Boolean
    Dim blnSkipToNextApplicant As Boolean
    blnDataFound = False
    'Loop round all the Applicants
    For nloop = 0 To (xmlNodeList.length - 1)
                                
        blnSkipToNextApplicant = False
        strFixedXML = ""
        'DPF 02/07/02 - ensure all string variables are bank before we pull back data (BMIDS00134)
        strQQApplicantDetails = ""
        strQQPersonalDebtsList = ""
        strQQOutgoingsList = ""
        strCustNo = ""
        strCustVerNo = ""
        'End of BMIDS00134
        
        Set objXmlElem = xmlNodeList.Item(nloop)

        strAppNo = objXmlAssist.GetTagValue(objXmlElem, "APPLICATIONNUMBER")
        strAppFFNo = objXmlAssist.GetTagValue(objXmlElem, "APPLICATIONFACTFINDNUMBER")
        strCustNo = objXmlAssist.GetTagValue(objXmlElem, "CUSTOMERNUMBER")
        strCustVerNo = objXmlAssist.GetTagValue(objXmlElem, "CUSTOMERVERSIONNUMBER")
            

        strFixedXML = strFixedXML + "<APPLICATIONNUMBER>" & strAppNo & "</APPLICATIONNUMBER>"
        strFixedXML = strFixedXML + "<APPLICATIONFACTFINDNUMBER>" & strAppFFNo & "</APPLICATIONFACTFINDNUMBER>"
        strFixedXML = strFixedXML + "<CUSTOMERNUMBER>" & strCustNo & "</CUSTOMERNUMBER>"
        strFixedXML = strFixedXML + "<CUSTOMERVERSIONNUMBER>" & strCustVerNo & "</CUSTOMERVERSIONNUMBER>"
        
        'Call Get Data for QuickQuoteApplicantDetails
        
        strSourceTable = "QUICKQUOTEAPPLICANTDETAILS"

        strXML = "<" & strSourceTable & ">" & strFixedXML & "</" & strSourceTable & ">"
        'blnRaiseError = True
        On Error Resume Next
        strQQApplicantDetails = GetData(strXML)
        'blnRaiseError = False
        objErrAssist.SaveErr
        On Error GoTo GetQuickQuoteDataVbErr
        If objErrAssist.LastErr = 0 Then        'JLD 22/10/99
            blnDataFound = True
        Else
            If objErrAssist.LastErr = omiga4RecordNotFound Then
                blnSkipToNextApplicant = True
            Else
                objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                        objErrAssist.LastErr
            End If
        End If
        
        'carry on with outgoings and debts if we have some data for this applicant
        If blnSkipToNextApplicant = False Then
        
            strSourceTable = "QUICKQUOTEOUTGOINGS"
    
            strXML = "<" & strSourceTable & ">" & strFixedXML & "</" & strSourceTable & ">"
    
            'Call find list for QuickQuoteOutgoingsDO
            'If record was not found, this is not an error, so don't raise one JLD
            On Error Resume Next
            strQQOutgoingsList = objQQOutgoingsDO.FindList(strXML)
            objErrAssist.SaveErr
            On Error GoTo GetQuickQuoteDataVbErr
            If objErrAssist.LastErr <> omiga4RecordNotFound And objErrAssist.LastErr <> 0 Then
                objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                        objErrAssist.LastErr
            End If
            
            strSourceTable = "QUICKQUOTEPERSONALDEBTS"
    
            strXML = "<" & strSourceTable & ">" & strFixedXML & "</" & strSourceTable & ">"
            
            'Call  find list for QuickQuotePersonalDebtsDO
            'If record was not found, this is not an error, so don't raise one JLD
            On Error Resume Next
            strQQPersonalDebtsList = objQQPersonalDebtsDO.FindList(strXML)
            objErrAssist.SaveErr
            On Error GoTo GetQuickQuoteDataVbErr
            If objErrAssist.LastErr <> omiga4RecordNotFound And objErrAssist.LastErr <> 0 Then
                objErrAssist.RaiseError TypeName(Me), strFunctionName, _
                                        objErrAssist.LastErr
            End If
                
            'Get the XML for the Applicant details
            Set objXmlQQApplicantDetailsDoc = objXmlAssist.load(strQQApplicantDetails, TypeName(Me), strFunctionName)
            'get the parent node
            Set objQQApplicantDetailsNode = objXmlQQApplicantDetailsDoc.documentElement
            If Len(strQQPersonalDebtsList) > 0 Then
                'Get the XML Personal debts list details for this applicant
                Set objXmlQQPersonalDebtsDoc = objXmlAssist.load(strQQPersonalDebtsList, TypeName(Me), strFunctionName)
                'get the parent node
                Set objQQPersonalDebtsNode = objXmlQQPersonalDebtsDoc.documentElement
                'Append the personal debts list to the Applicant details
                objQQApplicantDetailsNode.appendChild objQQPersonalDebtsNode
            End If
            If Len(strQQOutgoingsList) > 0 Then
                'Get the XML Outgoings list details for this apllicant
                Set objXmlQQOutgoingsDoc = objXmlAssist.load(strQQOutgoingsList, TypeName(Me), strFunctionName)
                'get the parent node
                Set objQQOutgoingsNode = objXmlQQOutgoingsDoc.documentElement
                'Append the Outgoings personal debts to the applicant details
                objQQApplicantDetailsNode.appendChild objQQOutgoingsNode
            End If
                'now append the applicant details including all Outgoings and personal debts
                objXmlOut.appendChild objQQApplicantDetailsNode
            
        End If
        
    Next
    
    'if no customer data was found raise an error  JLD
    If blnDataFound = False Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    GetQuickQuoteData = objXmlOut.xml
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objQQOutgoingsDO = Nothing
    Set objQQPersonalDebtsDO = Nothing
    Set objXmlQQApplicantDetailsDoc = Nothing
    Set objXmlQQPersonalDebtsDoc = Nothing
    Set objXmlQQOutgoingsDoc = Nothing
    Set XmlOutDoc = Nothing
    Set objDOAssist = Nothing
    Set objSQLAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function

GetQuickQuoteDataVbErr:
     
    'If Not blnRaiseError Then
    '    If Err.Number = omiga4RecordNotFound Then
    '        ERRNUMBER = Err.Number
    '        Resume Next
    '    End If
    'End If
    
    Set objXmlAssist = Nothing
    Set objQQOutgoingsDO = Nothing
    Set objQQPersonalDebtsDO = Nothing
    Set objXmlQQApplicantDetailsDoc = Nothing
    Set objXmlQQPersonalDebtsDoc = Nothing
    Set objXmlQQOutgoingsDoc = Nothing
    Set XmlOutDoc = Nothing
    Set objDOAssist = Nothing
    Set objSQLAssist = Nothing
    
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
    
    GetQuickQuoteData = ""
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FindList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
' return:
'   FindList        string containing XML data stream representation of
'                   data retrieved
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindList"

    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    FindList = objDOAssist.FindList(vstrXMLRequest, LoadData)

    Set objErrAssist = Nothing
    Set objDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    Set objDOAssist = Nothing
    
    FindList = ""
    
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

Public Sub DeleteAll(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr

    Const strFunctionName As String = "DeleteAll"

    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    objDOAssist.DeleteAll vstrXMLRequest, LoadData
 
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If

    Exit Sub

DeleteAllVbErr:

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

Public Function AddDerivedData(ByVal vstrXMLData As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vstrXMLData         base XML data stream
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'   AddDerivedData      base XML data stream plus any derived values
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "AddDerivedData"

    Dim objErrAssist As New ErrAssist

    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    AddDerivedData = vstrXMLData
    
    Set objErrAssist = Nothing
    
    Exit Function
            
    Set objErrAssist = Nothing
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function
    
AddDerivedDataVBErr:

    Set objErrAssist = Nothing
    
    AddDerivedData = ""

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function LoadData()
'------------------------------------------------------------------------------------------
' XML Request Format:
'
'   APPLICATIONNUMBER               NOT NULL VARCHAR2(12)       Y               Y(ApplicationFactFind.ApplicationNumber)
'   APPLICATIONFACTFINDNUMBER       NOT NULL NUMBER(5)          Y               Y(ApplicationFactFind.ApplicationFactFindNumber)
'   CUSTOMERNUMBER                  NOT NULL VARCHAR2(12)       Y               Y(CustomerVersion.CustomerNumber)
'   CUSTOMERVERSIONNUMBER           NOT NULL NUMBER(5)          Y               Y(CustomerVersion.CustomerVersionNumber)
'   ANNUALINCOMEORNETPROFIT         Number(10)
'   CCJDETAILSSEQUENCENUMBER        Number(5)
'   EMPLOYMENTSTATUS  Number(5)
'   GOODHEALTH                      Number(1)
'   SMOKER                          Number(1)
'   TOTALMONTHLYOUTGOINGS           Number(5)
'   TOTALDEBTSOUTSTANDINGBALANCE    Number(10)
'   TOTALDEBTSMONTHLYPAYMENTS       Number(5)
'   UKTAXPAYER                      Number(1)
' *****************************************************************************

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "QUICKQUOTEAPPLICANTDETAILS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtDouble</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ANNUALINCOMEORNETPROFIT<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CCJDETAILSSEQUENCENUMBER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CCJHistory</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>EMPLOYMENTSTATUS<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>EmploymentStatus</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>GOODHEALTH<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SMOKER<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALMONTHLYOUTGOINGS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALDEBTSOUTSTANDINGBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALDEBTSMONTHLYPAYMENTS<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>UKTAXPAYER<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'BMIDS00654 MDC 05/11/2002
    strXML = strXML & "<OTHERS>AMOUNTREQUESTED<TYPE>dbdtLong</TYPE></OTHERS>"
    'BMIDS00654 MDC 05/11/2002 - End
    strXML = strXML & "</TABLENAME>"

LoadData = strXML

End Function

Private Function IQuickQuoteApplicantDetailsDO_GetMainSelfEmployment(ByVal vxmlRequest As IXMLDOMElement) As Boolean
    On Error GoTo GetEmploymentStatusVbErr
    Dim strSQL As String
    Const strFunctionName As String = "IQuickQuoteApplicantDetailsDO_GetMainSelfEmployment"
    Dim objIADOAssist As IADOAssist
    Dim rsEmpStatus As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim adoCommand As ADODB.Command
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlEmpNode As IXMLDOMNode
    Dim strAppNo As String
    Dim strAppFactFind As String
    Dim bSelfEmployed As Boolean
    Dim strEmpStatus As Variant
        
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    ' Get the customer number and customer verison number from the XML passed
    ' in and use that to access the QuickQuoteApplicantDetails table to locate
    ' the employment type.
    
    strAppNo = m_objXMLAssist.GetNodeValue(vxmlRequest, ".//APPLICATIONNUMBER")
    strAppFactFind = m_objXMLAssist.GetNodeValue(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER")
    
    'SR 21/11/2002: BMIDS01050 - [JR SYSMCP1221] - Remove UPPER function for SQLServer
    #If GENERIC_SQL Then
        If objIADOAssist.GetDBEngine = SQLServer Then
            strSQL = " SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE GROUPNAME = ? "
        ElseIf objIADOAssist.GetDBEngine = Oracle Then
            strSQL = " SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE UPPER(GROUPNAME) = ? "
        End If
    #Else
        strSQL = " SELECT VALIDATIONTYPE FROM COMBOVALIDATION WHERE UPPER(GROUPNAME) = ? "
    #End If
    
    strSQL = strSQL + " AND VALUEID IN (SELECT EMPLOYMENTSTATUS FROM QUICKQUOTEAPPLICANTDETAILS " + _
             " WHERE APPLICATIONNUMBER = ? " + _
             " AND APPLICATIONFACTFINDNUMBER = ? )"
    'SR 21/11/2002: BMIDS01050 - End
    
    Set adoCommand = New ADODB.Command
    
    ' Now the parameters
    Set adoParameter = adoCommand.CreateParameter("p_sGroupName", adBSTR, adParamInput, Len("EMPLOYMENTSTATUS"), "EMPLOYMENTSTATUS")
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("p_sAppNo", adBSTR, adParamInput, Len(strAppNo), strAppNo)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("p_sAppFactFind", adBSTR, adParamInput, Len(strAppFactFind), strAppFactFind)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    
    adoCommand.ActiveConnection = objIADOAssist.GetConnStr
    adoCommand.CommandText = strSQL

    Set rsEmpStatus = adoCommand.Execute(, , adCmdText)
    
    If rsEmpStatus.EOF Or rsEmpStatus Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    bSelfEmployed = False
    
    Do While rsEmpStatus.EOF = False And Not bSelfEmployed
        strEmpStatus = rsEmpStatus("VALIDATIONTYPE")
    
        If Not IsNull(strEmpStatus) And Not IsEmpty(strEmpStatus) Then
            If APPLICANT_SELF_EMPLOYED = UCase(strEmpStatus) Then
                bSelfEmployed = True
            End If
        End If
        rsEmpStatus.MoveNext
    Loop
    
    rsEmpStatus.Close
        Set rsEmpStatus = Nothing
    Set adoCommand = Nothing
    
    IQuickQuoteApplicantDetailsDO_GetMainSelfEmployment = bSelfEmployed
    Exit Function
    
GetEmploymentStatusVbErr:

    ' Close the recordset if still open
        If Not rsEmpStatus Is Nothing Then
                If rsEmpStatus.State = adStateOpen Then
                        rsEmpStatus.Close
                End If
        End If
        Set rsEmpStatus = Nothing
    Set adoCommand = Nothing

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

    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub ObjectControl_Activate()
    If m_objXMLAssist Is Nothing Then
        Set m_objXMLAssist = New XMLAssist
    End If

    If m_objContext Is Nothing Then
        Set m_objContext = GetObjectContext()
    End If

    If m_objSQLAssist Is Nothing Then
        Set m_objSQLAssist = New SQLAssist
    End If

    If m_objErrAssist Is Nothing Then
        Set m_objErrAssist = New ErrAssist
    End If
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objXMLAssist = Nothing
    Set m_objErrAssist = Nothing
    Set m_objSQLAssist = Nothing
End Sub

