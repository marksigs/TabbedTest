VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BankWizardDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Omiga4DO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for omiga4 Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MCS    17/08/99    Created
'MCS    30/08/99    Update following code review
'MCS    01/09/99    Optional timimg output added
'PSC    30/11/99    Amend to use DOMDocuments rather than strings
'PSC    14/12/99    Amend to not use the default interface so that scripting cannot access
'                   methods but internal Omiga components can
'PSC    16/12/99    Amend interfaces
'PSC    17/12/99    Amend to use new IDOAssist interface
'PSC    01/02/00    Make objContext, objErrAssist, objXmlAssist and objIDOAssist
'                   object level rather than procedure level
'IVW    10/07/00    Bank Wizard is now an EXE, removed MTS code.
'TJ	13/01/02    SYS3745 Removed comments causing errors in OOSS
'------------------------------------------------------------------------------------------

Option Explicit

' TODO ************************************************************************
'   Add interface definition class and amend this to implement it
' *****************************************************************************
Implements IBankWizardDO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist

Private Function IBankWizardDO_GetBankDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetBankDetailsVbErr

    Const strFunctionName As String = "IBankWizardDO_GetBankDetails"
      
    Dim strValidationGenOut As String
    Dim strElementText As String
    Dim xmlValidSORTOut As FreeThreadedDOMDocument40
    Dim xmlValidACCOUNTOut As FreeThreadedDOMDocument40
    Dim xmlReturnBlock As New FreeThreadedDOMDocument40
    xmlReturnBlock.validateOnParse = False
    xmlReturnBlock.setProperty "NewParser", True
    Dim xmlTableClone As IXMLDOMElement
    Dim xmlSORTErrNoElement As IXMLDOMElement
    Dim xmlACCOUNTErrNoElement As IXMLDOMElement
    Dim intNodesLength As Integer
    Dim strTempElem As String
    Dim xmlRoot As IXMLDOMElement
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlSortCodeNode As IXMLDOMNode
    Dim xmlAccountNumberNode As IXMLDOMNode
    Dim xmlChildElement As IXMLDOMElement
    Dim strErrorNumber As String
    Dim strSortNoErrorText As String
    Dim strAccountNoErrorText As String
                
    Set xmlChildElement = xmlReturnBlock.createElement("BANKWIZARD")
    Set xmlReturnNode = xmlReturnBlock.appendChild(xmlChildElement)
    
    Set xmlChildElement = xmlReturnBlock.createElement("SORTCODECHECK")
    Set xmlSortCodeNode = xmlReturnNode.appendChild(xmlChildElement)
    
    Set xmlChildElement = xmlReturnBlock.createElement("ACCOUNTNUMBERCHECK")
    Set xmlAccountNumberNode = xmlReturnNode.appendChild(xmlChildElement)
        
    ' Create the VC++ Component which calls the Egg Top End Functionality
    ' SYS3153  - TJ 04/12/01    Changed to late binding
    Dim objBankDetailsInterface As Object
    
    Set objBankDetailsInterface = CreateObject("BankDetailsInterface.BankDetailsInterface")
    ' End change
    
    Set xmlTableClone = vxmlTableElement.cloneNode(True)
    
    m_objXmlAssist.ChangeNodeName xmlTableClone, "BANKDETAILS", "BANKDETAILSINPUT"
    
    '1st VALIDATION CALL Checks SORT CODE's 6 digits(NO hyphens) authenticity
    
    strValidationGenOut = objBankDetailsInterface.ValidateSortCode(xmlTableClone.xml)
    Set xmlValidSORTOut = m_objXmlAssist.load(strValidationGenOut, TypeName(Me), strFunctionName)
    Set xmlSORTErrNoElement = xmlValidSORTOut.getElementsByTagName("STATUS").Item(0)
    
    ' If an error has occurred retrieve error string from the bank wizard component XML error block.
    ' This error must be serious because if there is no Status Tag the BankWizard returns a
    ' dedicated error XML block.
          
    If xmlSORTErrNoElement Is Nothing Then ' no status block - serious error
        
        Set xmlSORTErrNoElement = xmlValidSORTOut.getElementsByTagName("BANKDETAILSERRORTEXT").Item(0)
        
        strElementText = xmlSORTErrNoElement.Text
        
        If xmlSORTErrNoElement Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , "Sort Code: Missing STATUS and ERROR DETAILS elements"
        Else
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , strElementText
        End If
      
    Else
        ' check the status field return to ensure no error
        
        strElementText = xmlSORTErrNoElement.Text
        
        strErrorNumber = "0"
            
        Select Case Trim(strElementText)
        
            Case "2" ' Error in the sortcode block
                strErrorNumber = "544"
            Case "1" ' Warning in the sortcode block
                strErrorNumber = "545"
                    
        End Select
        
        If strErrorNumber <> "0" Then
            ' get the error text from the return block
            strSortNoErrorText = xmlValidSORTOut.getElementsByTagName("MESSAGETEXT").Item(0).Text
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, CLng(strErrorNumber), , strSortNoErrorText
        End If
            
    End If
    
    ' Remove trailing spaces from the returned strings.
    Set xmlRoot = xmlValidSORTOut.documentElement
        
    intNodesLength = xmlRoot.childNodes.length
    
    For intNodesLength = 0 To intNodesLength - 1
        strTempElem = Trim(xmlRoot.childNodes.Item(intNodesLength).Text)
        xmlRoot.childNodes.Item(intNodesLength).Text = strTempElem
    Next
    
    xmlSortCodeNode.appendChild xmlValidSORTOut.documentElement
    
    ' If the sort code is okay then we can check the account number.
    ' The MESSAGETEXT tag will contain either SUCCESS if everything was okay, or if there
    ' are problems it will contain an error string.
    ' The STATUS tag will contain either 0 - Success, 1 - Warning, 2 - Error
    
    If xmlSORTErrNoElement.Text = "0" Then
        ' Only if the sort code validation was successful can we do an account number validation
        strValidationGenOut = objBankDetailsInterface.ValidateAccountDetails(xmlTableClone.xml)
        Set xmlValidACCOUNTOut = m_objXmlAssist.load(strValidationGenOut, TypeName(Me), strFunctionName)
        Set xmlACCOUNTErrNoElement = xmlValidACCOUNTOut.getElementsByTagName("STATUS").Item(0)
        
        If xmlACCOUNTErrNoElement Is Nothing Then
            Set xmlACCOUNTErrNoElement = xmlValidACCOUNTOut.getElementsByTagName("BANKDETAILSERRORTEXT").Item(0)
            strElementText = xmlACCOUNTErrNoElement.Text
        
            If xmlACCOUNTErrNoElement Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , "Account Code: Missing STATUS and ERROR DETAILS elements"
            Else
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , strElementText
            End If
      
        Else
            ' check the status field return to ensure no error
        
            strElementText = xmlACCOUNTErrNoElement.Text
        
            strErrorNumber = "0"
            
            Select Case Trim(strElementText)
       
                Case "2" ' Error in the account number block
                    strErrorNumber = "544"
                Case "1" ' Warning in the account number block
                    strErrorNumber = "545"
                    
            End Select
        
            If strErrorNumber <> "0" Then
                ' get the error text from the return block
                strAccountNoErrorText = xmlValidACCOUNTOut.getElementsByTagName("MESSAGETEXT").Item(0).Text
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, CLng(strErrorNumber), , strAccountNoErrorText
            End If
        End If
    End If
        
    xmlAccountNumberNode.appendChild xmlValidACCOUNTOut.documentElement
        
    Set IBankWizardDO_GetBankDetails = xmlReturnBlock.documentElement
       
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlValidSORTOut = Nothing
    Set xmlValidACCOUNTOut = Nothing
    Set xmlReturnBlock = Nothing
    Set xmlTableClone = Nothing
    Set xmlSORTErrNoElement = Nothing
    Set xmlACCOUNTErrNoElement = Nothing
    Set xmlRoot = Nothing
    Set xmlReturnNode = Nothing
    Set xmlSortCodeNode = Nothing
    Set xmlAccountNumberNode = Nothing
    Set xmlChildElement = Nothing
    
    Set objBankDetailsInterface = Nothing
    
    Exit Function

GetBankDetailsVbErr:
    
    Set xmlValidSORTOut = Nothing
    Set xmlValidACCOUNTOut = Nothing
    Set xmlReturnBlock = Nothing
    Set xmlTableClone = Nothing
    Set xmlSORTErrNoElement = Nothing
    Set xmlACCOUNTErrNoElement = Nothing
    Set xmlRoot = Nothing
    Set xmlReturnNode = Nothing
    Set xmlSortCodeNode = Nothing
    Set xmlAccountNumberNode = Nothing
    Set xmlChildElement = Nothing
    Set objBankDetailsInterface = Nothing
            
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
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    
    If m_objContext Is Nothing Then
        Set m_objIDOAssist = New DOAssist
    Else
        Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing

End Sub





