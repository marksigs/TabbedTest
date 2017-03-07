VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ReferencesDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'History:
'
' Prog      Date     Description
' BG        10/11/01 SYS3384 Changed the way TotalRepayment variable was being used to create
'                    TOTALMONTHLYREPAYMENT - it now returns a value with 2 decimal places.
'------------------------------------------------------------------------------------------
Option Explicit

Implements IReferencesDO
Implements ObjectControl

Private gobjContext As ObjectContext

'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        AQR         Description
' MV    17/09/2002  BMIDS00468  Modified IReferencesDO_GetCurrEmployersRef
' MV    17/09/2002  BMIDS00471  Modified IReferencesDO_GetLoanDetailsForRef
' MV    06/02/2003  BM0312      Amended IReferencesDO_CreatePrevEmployersRef,
'                                       IReferencesDO_CreatePrevLendersRef,
'                                       IReferencesDO_CreatePrevLandLordsRef
' MV    25/02/2003  BM0312      Amended IReferencesDO_CreatePrevEmployersRef,
'                                       IReferencesDO_CreatePrevLendersRef,
'                                       IReferencesDO_CreatePrevLandLordsRef
' MV    26/07/2003  BM0312      Amended IReferencesDO_CreatePrevEmployersRef,
'                                       IReferencesDO_CreatePrevLendersRef,
'                                       IReferencesDO_CreatePrevLandLordsRef
'------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

'--------------------------------------------------------------------------------------------
'BMIDS Specific History:

'DPF     17/09/2002      Have added field 'VerificationType' to any CurrentLendersRef methods
'--------------------------------------------------------------------------------------------

Private Sub IReferencesDO_CreateEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_CreateEmployersReference"
    
    adoCreateFromNode vxmlRequestNode, "EMPLOYERSREFERENCE"
    
CreateEmployersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo CreateCurrEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreateCurrEmployersRef"
    
    'First call CreateEmployersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "JOBTITLE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATESTARTED"
     
    IReferencesDO_CreateEmployersReference xmlNode, vxmlResponseNode
    
    adoCreateFromNode vxmlRequestNode, "CURRENTEMPLOYERSREF"
    
    'Update the ConfirmedEarnedIncome
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem2 As IXMLDOMElement
    Dim xmlReqNode As IXMLDOMNode
    Dim strCustNo As String
    Dim strCustVerNo As String
    Dim strEmpSeqNo As String
    
    'Create a list of CONFIRMEDEARNEDINCOME nodes
    Set xmlNodeList = vxmlRequestNode.selectNodes("CONFIRMEDEARNEDINCOME")
    If xmlNodeList.length > 0 Then
        
        'Create an individual CONFIRMEDEARNEDINCOME node to pass in
        Set xmlElem2 = xmlDOM.createElement("CONFIRMEDEARNEDINCOME")
        Set xmlReqNode = vxmlResponseNode.appendChild(xmlElem2)
        Set xmlReqNode = xmlNodeList.Item(0)
                
        'Record some key data
        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
        strCustVerNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
        strEmpSeqNo = xmlGetAttributeText(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER")
                
        For Each xmlReqNode In xmlNodeList
            
            'For each CONFIRMEDINCOME node in request, update the individual node
            xmlSetAttributeValue xmlReqNode, "CUSTOMERNUMBER", strCustNo
            xmlSetAttributeValue xmlReqNode, "CUSTOMERVERSIONNUMBER", strCustVerNo
            xmlSetAttributeValue xmlReqNode, "EMPLOYMENTSEQUENCENUMBER", strEmpSeqNo
                        
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "PAYMENTFREQUENCYTYPE"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMETYPE"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMEAMOUNT"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMESEQUENCENUMBER"
           
            Debug.Print xmlReqNode.xml
                          
            adoCreateFromNode xmlReqNode, "CONFIRMEDEARNEDINCOME"
        
        Next
    End If
   
CreateCurrEmployersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    'SG 05/02/02 SYS3837
    Set xmlElem2 = Nothing
    Set xmlReqNode = Nothing
    Set xmlNodeList = Nothing
    'SG 05/02/02 SYS3837
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateCurrLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreateCurrLendersRef"
    
    'First call CreateLendersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYREPAYMENT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATEOFLOAN"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "LOANNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "GOODACCOUNTCONDUCT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "PREVIOUSARREARS"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "VERIFICATIONTYPE"
    
    IReferencesDO_CreateLendersReference xmlNode, vxmlResponseNode
    
    adoCreateFromNode vxmlRequestNode, "CURRENTLENDERSREF"
    
CreateCurrLendersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreatePrevLendersRef(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreatePrevLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreatePrevLendersRef"
    
    'First call CreateLendersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    
    IReferencesDO_GetLendersReference xmlNode, vxmlResponseNode
    
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYREPAYMENT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATEOFLOAN"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "LOANNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "GOODACCOUNTCONDUCT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "PREVIOUSARREARS"
    
    If vxmlResponseNode.hasChildNodes = False Then
        IReferencesDO_CreateLendersReference xmlNode, vxmlResponseNode
    Else
        IReferencesDO_UpdateLendersReference xmlNode, vxmlResponseNode
    End If
    
    adoCreateFromNode vxmlRequestNode, "PREVIOUSLENDERSREF"
    
    
CreatePrevLendersRefExit:
    
    gobjContext.SetComplete
    
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreatePrevEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreatePrevEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreatePrevEmployersRef"
    
    'First call CreateEmployersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    
    IReferencesDO_GetEmployersReference xmlNode, vxmlResponseNode
    
    xmlCopyAttribute vxmlRequestNode, xmlNode, "JOBTITLE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATESTARTED"
    
    If vxmlResponseNode.hasChildNodes = False Then
        IReferencesDO_CreateEmployersReference xmlNode, vxmlResponseNode
    Else
        IReferencesDO_UpdateEmployersReference xmlNode, vxmlResponseNode
    End If
    
    adoCreateFromNode vxmlRequestNode, "PREVIOUSEMPLOYERSREF"
    
CreatePrevEmployersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateAccountantsReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateAccountantsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_CreateAccountantsReference"
    
    adoCreateFromNode vxmlRequestNode, "ACCOUNTANTSREFERENCE"
    
CreateAccountantsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateLandlordsReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_CreateLandlordsReference"
    
    adoCreateFromNode vxmlRequestNode, "LANDLORDSREFERENCE"
    
CreateLandlordsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateCurrLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreateCurrLandlordsRef"
    
    'First call CreateLandlordsReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "TENANCYSTARTDATE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYRENTAL"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "SATISFACTORYCONDUCT"
    
    IReferencesDO_CreateLandlordsReference xmlNode, vxmlResponseNode
    
    adoCreateFromNode vxmlRequestNode, "CURRENTLANDLORDSREF"
    
CreateCurrLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreatePrevLandlordsRef(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo CreatePrevLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_CreatePrevLandlordsRef"
    
    'First call CreateLandlordsReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    
    IReferencesDO_GetLandlordsRef xmlNode, vxmlResponseNode
    
    xmlCopyAttribute vxmlRequestNode, xmlNode, "TENANCYSTARTDATE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYRENTAL"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "SATISFACTORYCONDUCT"
    
    If vxmlResponseNode.hasChildNodes = False Then
        IReferencesDO_CreateLandlordsReference xmlNode, vxmlResponseNode
    Else
        IReferencesDO_UpdateLandlordsReference xmlNode, vxmlResponseNode
    End If
    
    adoCreateFromNode vxmlRequestNode, "PREVIOUSLANDLORDSREF"
       
CreatePrevLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_CreateLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_CreateLendersReference"
    
    adoCreateFromNode vxmlRequestNode, "LENDERSREFERENCE"
    
CreateLendersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_GetEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_GetEmployersReference"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "EMPLOYERSREFERENCE"
    
GetEmployersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetCurrEmployersRef(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetCurrEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetCurrEmployersRef"
    
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCurrEmpRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    Dim xmlDOM2 As FreeThreadedDOMDocument40
    Dim xmlResponseDom2 As FreeThreadedDOMDocument40
    Dim xmlElem2 As IXMLDOMElement
    Dim xmlNode2 As IXMLDOMNode
    Dim xmlRespElem2 As IXMLDOMElement
    Dim xmlRespNode2 As IXMLDOMNode
    
    ' create a response just from GetEmployersReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    
    IReferencesDO_GetEmployersReference xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("EMPLOYERSREFERENCE")
    
    'Create some objects to work with
    Set xmlResponseDom2 = New FreeThreadedDOMDocument40
    xmlResponseDom2.validateOnParse = False
    xmlResponseDom2.setProperty "NewParser", True
    Set xmlRespElem2 = xmlResponseDom2.createElement("RESPONSE")
    Set xmlRespNode2 = xmlResponseDom2.appendChild(xmlRespElem2)
    Set xmlDOM2 = New FreeThreadedDOMDocument40
    xmlDOM2.validateOnParse = False
    xmlDOM2.setProperty "NewParser", True
    Set xmlElem2 = xmlDOM2.createElement("REQUEST")
    Set xmlNode2 = xmlDOM2.appendChild(xmlElem2)
    Set xmlNode2 = xmlNode2.appendChild(xmlDOM2.createElement("COMFIRMEDEARNEDINCOME"))
    xmlCopyAttribute vxmlRequestNode, xmlNode2, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode2, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode2, "EMPLOYMENTSEQUENCENUMBER"
    'Use them to get the ConfirmedEarnedIncome information
    adoGetAsXML xmlNode2, xmlRespNode2, "CONFIRMEDEARNEDINCOME"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CURRENTEMPLOYERSREF"
    Set xmlCurrEmpRefNode = vxmlResponseNode.selectSingleNode("CURRENTEMPLOYERSREF")
    
    'Append the ConfirmedEarnedIncome data to the ResponseNode to be passed back.
    xmlRespNode2.selectSingleNode ("Response")
    If xmlRespNode2.hasChildNodes Then
        While xmlRespNode2.childNodes.length > 0
            xmlCurrEmpRefNode.appendChild xmlRespNode2.childNodes(0)
        Wend
    End If
    
    If Not xmlCurrEmpRefNode Is Nothing Then
        If Not xmlNode Is Nothing Then
            xmlCopyAttribute xmlNode, xmlCurrEmpRefNode, "JOBTITLE"
            xmlCopyAttribute xmlNode, xmlCurrEmpRefNode, "DATESTARTED"
        End If
    End If
    
    gobjContext.SetComplete
    
GetCurrEmployersRefExit:
    
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDom2 = Nothing
    Set xmlRespElem2 = Nothing
    Set xmlRespNode2 = Nothing
    Set xmlDOM2 = Nothing
    Set xmlElem2 = Nothing
    Set xmlNode2 = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub


Private Sub IReferencesDO_GetCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetCurrLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetCurrLendersRef"
    
    ' call GetLendersReference first
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCurrLendersRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    
    ' create a response just from GetLendersReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    
    IReferencesDO_GetLendersReference xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("LENDERSREFERENCE")
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CURRENTLENDERSREF"
    Set xmlCurrLendersRefNode = vxmlResponseNode.selectSingleNode("CURRENTLENDERSREF")
    
    If Not xmlCurrLendersRefNode Is Nothing Then
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "MONTHLYREPAYMENT"
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "DATEOFLOAN"
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "LOANNUMBER"
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "GOODACCOUNTCONDUCT"
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "PREVIOUSARREARS"
        xmlCopyAttribute xmlNode, xmlCurrLendersRefNode, "VERIFICATIONTYPE"
    End If
    
GetCurrLendersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDom = Nothing
    Set xmlCurrLendersRefNode = Nothing
    Set xmlRespElem = Nothing
    Set xmlRespNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetPrevLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetPrevLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetPrevLendersRef"
    
    ' call GetLendersReference first
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlPrevLendersRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    
    ' create a response just from GetLendersReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    
    IReferencesDO_GetLendersReference xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("LENDERSREFERENCE")
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "PREVIOUSLENDERSREF"
    Set xmlPrevLendersRefNode = vxmlResponseNode.selectSingleNode("PREVIOUSLENDERSREF")
    
    If Not xmlPrevLendersRefNode Is Nothing Then
        xmlCopyAttribute xmlNode, xmlPrevLendersRefNode, "MONTHLYREPAYMENT"
        xmlCopyAttribute xmlNode, xmlPrevLendersRefNode, "DATEOFLOAN"
        xmlCopyAttribute xmlNode, xmlPrevLendersRefNode, "LOANNUMBER"
        xmlCopyAttribute xmlNode, xmlPrevLendersRefNode, "GOODACCOUNTCONDUCT"
        xmlCopyAttribute xmlNode, xmlPrevLendersRefNode, "PREVIOUSARREARS"
    End If
    
GetPrevLendersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDom = Nothing
    Set xmlPrevLendersRefNode = Nothing
    Set xmlRespElem = Nothing
    Set xmlRespNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetPrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetPrevEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetPrevEmployersRef"
    
    ' call GetEmployersReference first
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlPrevEmpRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    
    ' create a response just from GetEmployersReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    
    IReferencesDO_GetEmployersReference xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("EMPLOYERSREFERENCE")
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "PREVIOUSEMPLOYERSREF"
    Set xmlPrevEmpRefNode = vxmlResponseNode.selectSingleNode("PREVIOUSEMPLOYERSREF")
    
    If Not xmlPrevEmpRefNode Is Nothing Then
        xmlCopyAttribute xmlNode, xmlPrevEmpRefNode, "JOBTITLE"
        xmlCopyAttribute xmlNode, xmlPrevEmpRefNode, "DATESTARTED"
    End If
    
GetPrevEmployersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetAccountantsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetAccountantsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_GetAccountantsRef"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ACCOUNTANTSREFERENCE"
    
GetAccountantsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_GetLandlordsRef"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "LANDLORDSREFERENCE"
    
GetLandlordsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetCurrLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetCurrLandlordsRef"
    
    ' call GetLendersReference first
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlCurrLandlordsRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    
    ' create a response just from GetLandlordsReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    
    IReferencesDO_GetLandlordsRef xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("LANDLORDSREFERENCE")
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "CURRENTLANDLORDSREF"
    Set xmlCurrLandlordsRefNode = vxmlResponseNode.selectSingleNode("CURRENTLANDLORDSREF")
    
    If Not xmlCurrLandlordsRefNode Is Nothing Then
        xmlCopyAttribute xmlNode, xmlCurrLandlordsRefNode, "TENANCYSTARTDATE"
        xmlCopyAttribute xmlNode, xmlCurrLandlordsRefNode, "MONTHLYRENTAL"
        xmlCopyAttribute xmlNode, xmlCurrLandlordsRefNode, "SATISFACTORYCONDUCT"
    End If
    
GetCurrLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDom = Nothing
    Set xmlCurrLandlordsRefNode = Nothing
    Set xmlRespElem = Nothing
    Set xmlRespNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetPrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetPrevLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_GetPrevLandlordsRef"
    
    ' call GetLendersReference first
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlResponseDom As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlPrevLandlordsRefNode As IXMLDOMNode
    Dim xmlRespElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    
    ' create a response just from GetLandlordsReference and copy required attributes to final output
    Set xmlResponseDom = New FreeThreadedDOMDocument40
    xmlResponseDom.validateOnParse = False
    xmlResponseDom.setProperty "NewParser", True
    Set xmlRespElem = xmlResponseDom.createElement("RESPONSE")
    Set xmlRespNode = xmlResponseDom.appendChild(xmlRespElem)
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    
    IReferencesDO_GetLandlordsRef xmlNode, xmlRespNode
    Set xmlNode = xmlRespNode.selectSingleNode("LANDLORDSREFERENCE")
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "PREVIOUSLANDLORDSREF"
    Set xmlPrevLandlordsRefNode = vxmlResponseNode.selectSingleNode("PREVIOUSLANDLORDSREF")
    
    If Not xmlPrevLandlordsRefNode Is Nothing Then
        xmlCopyAttribute xmlNode, xmlPrevLandlordsRefNode, "TENANCYSTARTDATE"
        xmlCopyAttribute xmlNode, xmlPrevLandlordsRefNode, "MONTHLYRENTAL"
        xmlCopyAttribute xmlNode, xmlPrevLandlordsRefNode, "SATISFACTORYCONDUCT"
    End If
    
GetPrevLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlResponseDom = Nothing
    Set xmlPrevLandlordsRefNode = Nothing
    Set xmlRespElem = Nothing
    Set xmlRespNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_GetLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_GetLendersReference"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "LENDERSREFERENCE"
    
GetLendersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_UpdateEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateEmployersReference"
    
    adoUpdateFromNode vxmlRequestNode, "EMPLOYERSREFERENCE"
    
UpdateEmployersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_UpdateCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

On Error GoTo UpdateCurrEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateCurrEmployersRef"
    
    'First call UpdateEmployersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "JOBTITLE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATESTARTED"
    
    IReferencesDO_UpdateEmployersReference xmlNode, vxmlResponseNode
    
    'Update the ConfirmedEarnedIncome
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElem2 As IXMLDOMElement
    Dim xmlReqNode As IXMLDOMNode
    Dim strCustNo As String
    Dim strCustVerNo As String
    Dim strEmpSeqNo As String
    
    'Create a list of CONFIRMEDEARNEDINCOME nodes
    Set xmlNodeList = vxmlRequestNode.selectNodes("CONFIRMEDEARNEDINCOME")
    If xmlNodeList.length > 0 Then
        
        'Create an individual CONFIRMEDEARNEDINCOME node to pass in
        Set xmlElem2 = xmlDOM.createElement("CONFIRMEDEARNEDINCOME")
        Set xmlReqNode = vxmlResponseNode.appendChild(xmlElem2)
        Set xmlReqNode = xmlNodeList.Item(0)
                
        'Record some key data
        strCustNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERNUMBER")
        strCustVerNo = xmlGetAttributeText(vxmlRequestNode, "CUSTOMERVERSIONNUMBER")
        strEmpSeqNo = xmlGetAttributeText(vxmlRequestNode, "EMPLOYMENTSEQUENCENUMBER")
                
        For Each xmlReqNode In xmlNodeList
            
            'For each CONFIRMEDINCOME node in request, update the individual node
            xmlSetAttributeValue xmlReqNode, "CUSTOMERNUMBER", strCustNo
            xmlSetAttributeValue xmlReqNode, "CUSTOMERVERSIONNUMBER", strCustVerNo
            xmlSetAttributeValue xmlReqNode, "EMPLOYMENTSEQUENCENUMBER", strEmpSeqNo
                        
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "PAYMENTFREQUENCYTYPE"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMETYPE"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMEAMOUNT"
            xmlCopyAttribute vxmlRequestNode, xmlReqNode, "EARNEDINCOMESEQUENCENUMBER"
           
            Debug.Print xmlReqNode.xml
                          
            adoUpdateFromNode xmlReqNode, "CONFIRMEDEARNEDINCOME"
        
        Next
    End If
        
    adoUpdateFromNode vxmlRequestNode, "CURRENTEMPLOYERSREF"
    
    
    
UpdateCurrEmployersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    'SG 05/02/02 SYS3837
    Set xmlElem2 = Nothing
    Set xmlReqNode = Nothing
    Set xmlNodeList = Nothing
    'SG 05/02/02 SYS3837
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_UpdatePrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdatePrevEmployersRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdatePrevEmployersRef"
    
    'First call UpdateEmployersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("EMPLOYERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "JOBTITLE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATESTARTED"
    
    IReferencesDO_UpdateEmployersReference xmlNode, vxmlResponseNode
    
    adoUpdateFromNode vxmlRequestNode, "PREVIOUSEMPLOYERSREF"
    
UpdatePrevEmployersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_UpdateAccountantsReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateAccountantsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateAccountantsReference"
    
    adoUpdateFromNode vxmlRequestNode, "ACCOUNTANTSREFERENCE"
    
UpdateAccountantsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_UpdateLandlordsReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateLandlordsReference"
    
    adoUpdateFromNode vxmlRequestNode, "LANDLORDSREFERENCE"
    
UpdateLandlordsReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_UpdateCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdateCurrLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateCurrLandlordsRef"
    
    'First call UpdateLandlordsReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "TENANCYSTARTDATE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYRENTAL"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "SATISFACTORYCONDUCT"
    
    IReferencesDO_UpdateLandlordsReference xmlNode, vxmlResponseNode
    
    adoUpdateFromNode vxmlRequestNode, "CURRENTLANDLORDSREF"
    
UpdateCurrLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_UpdatePrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdatePrevLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdatePrevLandlordsRef"
    
    'First call UpdateLandlordsReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LANDLORDSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "TENANCYSTARTDATE"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYRENTAL"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "SATISFACTORYCONDUCT"
    
    IReferencesDO_UpdateLandlordsReference xmlNode, vxmlResponseNode
    
    adoUpdateFromNode vxmlRequestNode, "PREVIOUSLANDLORDSREF"
    
UpdatePrevLandlordsRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_UpdateLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateLendersReference"
    
    adoUpdateFromNode vxmlRequestNode, "LENDERSREFERENCE"
    
UpdateLendersReferenceExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesDO_UpdateCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdateCurrLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdateCurrLendersRef"
    
    'First call UpdateLendersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYREPAYMENT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATEOFLOAN"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "LOANNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "GOODACCOUNTCONDUCT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "PREVIOUSARREARS"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "VERIFICATIONTYPE"
    
    IReferencesDO_UpdateLendersReference xmlNode, vxmlResponseNode
    
    adoUpdateFromNode vxmlRequestNode, "CURRENTLENDERSREF"
    
UpdateCurrLendersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_UpdatePrevLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdatePrevLendersRefExit
    
    Const strFunctionName As String = "IReferencesDO_UpdatePrevLendersRef"
    
    'First call UpdateLendersReference
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("REQUEST")
    Set xmlNode = xmlDOM.appendChild(xmlElem)
    Set xmlNode = xmlNode.appendChild(xmlDOM.createElement("LENDERSREFERENCE"))
    xmlCopyAttribute vxmlRequestNode, xmlNode, "ACCOUNTGUID"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "MONTHLYREPAYMENT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "DATEOFLOAN"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "LOANNUMBER"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "GOODACCOUNTCONDUCT"
    xmlCopyAttribute vxmlRequestNode, xmlNode, "PREVIOUSARREARS"
    
    IReferencesDO_UpdateLendersReference xmlNode, vxmlResponseNode
    
    adoUpdateFromNode vxmlRequestNode, "PREVIOUSLENDERSREF"
    
UpdatePrevLendersRefExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesDO_GetEmpDetailsForRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetEmpDetailsExit
    
    Const strFunctionName As String = "IReferencesDO_GetEmpDetailsForRef"
    
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlIncomeNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlRefNode As IXMLDOMNode
    Dim xmlAttr As IXMLDOMAttribute
    
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("RESPONSE")
    Set xmlRespNode = xmlDOM.appendChild(xmlElem)
    
    adoGetAsXML vxmlRequestNode, xmlRespNode, "EMPDETAILSFORREF"
    
    'now doctor the returned data to look like the DTD
    Set xmlNodeList = xmlRespNode.selectNodes("EMPDETAILSFORREF")
    If xmlNodeList.length > 0 Then
        Set xmlElem = xmlDOM.createElement("EMPLOYMENTDETAIL")
        Set xmlNode = vxmlResponseNode.appendChild(xmlElem)
        Set xmlRefNode = xmlNodeList.Item(0)
        
        xmlCopyAttribute xmlRefNode, xmlNode, "CUSTOMERNUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "CUSTOMERVERSIONNUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "EMPLOYMENTSEQUENCENUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "EMPLOYMENTTYPE"
        xmlCopyAttribute xmlRefNode, xmlNode, "EMPLOYMENTTYPE_TEXT"
        xmlCopyAttribute xmlRefNode, xmlNode, "INDUSTRYTYPE"
        xmlCopyAttribute xmlRefNode, xmlNode, "INDUSTRYTYPE_TEXT"
        xmlCopyAttribute xmlRefNode, xmlNode, "OCCUPATIONTYPE"
        xmlCopyAttribute xmlRefNode, xmlNode, "OCCUPATIONTYPE_TEXT"
        xmlCopyAttribute xmlRefNode, xmlNode, "JOBTITLE"
        xmlCopyAttribute xmlRefNode, xmlNode, "DATESTARTEDORESTABLISHED"
        xmlCopyAttribute xmlRefNode, xmlNode, "NATIONALINSURANCENUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "TAXOFFICE"
        xmlCopyAttribute xmlRefNode, xmlNode, "TAXREFERENCENUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "POSTCODE"
        xmlCopyAttribute xmlRefNode, xmlNode, "BUILDINGORHOUSENAME"
        xmlCopyAttribute xmlRefNode, xmlNode, "BUILDINGORHOUSENUMBER"
        xmlCopyAttribute xmlRefNode, xmlNode, "STREET"
        xmlCopyAttribute xmlRefNode, xmlNode, "DISTRICT"
        xmlCopyAttribute xmlRefNode, xmlNode, "TOWN"
        xmlCopyAttribute xmlRefNode, xmlNode, "COUNTY"
        xmlCopyAttribute xmlRefNode, xmlNode, "COUNTRY"
        xmlCopyAttribute xmlRefNode, xmlNode, "COUNTRY_TEXT"
        If xmlAttributeValueExists(xmlRefNode, "SELFEMPSHARES") Then
            xmlCopyAttributeValue xmlRefNode, xmlNode, "SELFEMPSHARES", "PERCENTSHARESHELD"
        Else
            xmlCopyAttributeValue xmlRefNode, xmlNode, "EMPLSHARES", "PERCENTSHARESHELD"
        End If
        If xmlAttributeValueExists(xmlRefNode, "THIRDPARTYCOMPANYNAME") Then
            xmlCopyAttributeValue xmlRefNode, xmlNode, "THIRDPARTYCOMPANYNAME", "COMPANYNAME"
        Else
            xmlCopyAttributeValue xmlRefNode, xmlNode, "NAMEADDRESSCOMPANYNAME", "COMPANYNAME"
        End If

        For Each xmlRefNode In xmlNodeList
            'Add an earnedincome element
            Set xmlElem = xmlDOM.createElement("EARNEDINCOME")
            Set xmlIncomeNode = xmlNode.appendChild(xmlElem)
            xmlCopyAttribute xmlRefNode, xmlIncomeNode, "EARNEDINCOMEAMOUNT"
            xmlCopyAttribute xmlRefNode, xmlIncomeNode, "EARNEDINCOMETYPE"
            xmlCopyAttribute xmlRefNode, xmlIncomeNode, "PAYMENTFREQUENCYTYPE"
        Next
    End If
    
GetEmpDetailsExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlRespNode = Nothing
    Set xmlNode = Nothing
    Set xmlIncomeNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlRefNode = Nothing
    Set xmlAttr = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IReferencesDO_GetLoanDetailsForRef(ByVal vxmlRequestNode As IXMLDOMNode, _
                                                ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetLoanDetailsExit
    
    Const strFunctionName As String = "IReferencesDO_GetLoanDetailsForRef"
    
    Dim xmlDOM As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlRespNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlRefNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    Dim dblTotalRepayment As Double
    Dim nTotalBalance As Double
        
    Set xmlDOM = New FreeThreadedDOMDocument40
    xmlDOM.validateOnParse = False
    xmlDOM.setProperty "NewParser", True
    Set xmlElem = xmlDOM.createElement("RESPONSE")
    Set xmlRespNode = xmlDOM.appendChild(xmlElem)
    
    adoGetAsXML vxmlRequestNode, xmlRespNode, "LOANDETAILSFORREF"
    
    'Sum the monthly repayments and outstanding balances for each mortgageLoan returned
    Set xmlNodeList = xmlRespNode.selectNodes("LOANDETAILSFORREF")
    
    If xmlNodeList.length > 0 Then
        
        Set xmlElem = xmlDOM.createElement("LOANDETAIL")
        Set xmlNode = vxmlResponseNode.appendChild(xmlElem)
        Set xmlRefNode = xmlNodeList.Item(0)
        
        xmlCopyAttribute xmlRefNode, xmlNode, "ACCOUNTGUID"
        xmlCopyAttribute xmlRefNode, xmlNode, "ACCOUNTNUMBER"
        
        If xmlAttributeValueExists(xmlRefNode, "THIRDPARTYCOMPANYNAME") Then
            xmlCopyAttributeValue xmlRefNode, xmlNode, "THIRDPARTYCOMPANYNAME", "COMPANYNAME"
        Else
            xmlCopyAttributeValue xmlRefNode, xmlNode, "NAMEANDADDRESSCOMPANYNAME", "COMPANYNAME"
        End If
        
        dblTotalRepayment = 0
        nTotalBalance = 0
        
        For Each xmlRefNode In xmlNodeList
            dblTotalRepayment = dblTotalRepayment + CSafeDbl(xmlGetAttributeText(xmlRefNode, "MONTHLYREPAYMENT"))
            nTotalBalance = nTotalBalance + CSafeDbl(xmlGetAttributeText(xmlRefNode, "OUTSTANDINGBALANCE"))
        Next
        
        Set xmlAttrib = xmlNode.ownerDocument.createAttribute("TOTALMONTHLYREPAYMENT")
        xmlAttrib.Text = Format(dblTotalRepayment, "0.00")
        xmlNode.Attributes.setNamedItem xmlAttrib
        
        Set xmlAttrib = xmlNode.ownerDocument.createAttribute("TOTALOUTSTANDINGBALANCE")
        xmlAttrib.Text = CStr(nTotalBalance)
        xmlNode.Attributes.setNamedItem xmlAttrib

    End If
    
GetLoanDetailsExit:
    
    gobjContext.SetComplete
    Set xmlDOM = Nothing
    Set xmlElem = Nothing
    Set xmlRespNode = Nothing
    Set xmlNode = Nothing
    Set xmlRefNode = Nothing
    Set xmlNodeList = Nothing
    Set xmlAttrib = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub IReferencesDO_GetTenancyDetailsForRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo GetTenancyDetailsExit
    
    Const strFunctionName As String = "IReferencesDO_GetTenancyDetailsForRef"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "TENANCYDETAILSFORREF"
    
GetTenancyDetailsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub

