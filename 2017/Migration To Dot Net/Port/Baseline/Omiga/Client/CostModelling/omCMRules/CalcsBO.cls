VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CalcsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'********************************************************************************
'General Declarations Section
'********************************************************************************
'Workfile:      CalcsBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Redemption Fee Calculations Business Object
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'********************************************************************************
'History:
'
'Prog   Date        Description
'------------------------------
'DPF    15/07/2002  Created
'DPF    15/07/2002  Added method called CalculateRedemptionFees, takes in Loan
'                   component XML, calculates Redemption Fees and passes back XML
'
'********************************************************************************

Implements ObjectControl

Public Function CalculateRedemptionFees(vxmlRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  calculates Redemption Fees
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
'
' return:       xml Response data stream containing results of operation
'               either: TYPE="SUCCESS"
'               or: TYPE="SYSERR" and <ERROR> element
'
' Request:      xml Request data stream containing loan component details passed in from
'               the calcsBO
'------------------------------------------------------------------------------------------
On Error GoTo CalculateRedemptionFeesExit
    
    Const strFunctionName As String = "CalculateRedemptionFees"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

'**************************************
'Main Body Of Code

    'Variable declaration
    Dim objXmlFeesIn As IXMLDOMNode
    Dim objDOMFeesIn As FreeThreadedDOMDocument40
    Dim objXmlOut As FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlFeeListElem As IXMLDOMElement
    Dim objXmlFeeElem As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objFeeBandNode As IXMLDOMNode
    Dim lngStep As Long
    Dim lngPeriod As Long
    Dim dteEndDate As Date
    Dim dblFeeAmount As Double
    Dim dblPercentage As Double
    Dim dblLoanAmount As Double
    
    dblFeeAmount = 0
        
    Set objDOMFeesIn = New FreeThreadedDOMDocument40
    objDOMFeesIn.validateOnParse = False
    objDOMFeesIn.setProperty "NewParser", True
    Set objXmlOut = New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
        
    'build response block
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'locate necessary data within request block
    If objDOMFeesIn.LoadXML(vxmlRequest) Then
        
        dblLoanAmount = CSafeDbl(GetMandatoryAttributeText(objDOMFeesIn.documentElement, "LOANAMOUNT"))
        Set objXmlFeesIn = GetNode(objDOMFeesIn.documentElement, "//REDEMPTIONFEEBANDLIST")
        
        'build XML response block
        Set objXmlFeeListElem = objXmlOut.createElement("REDEMPTIONFEEAMOUNTLIST")
        
        If Not objXmlFeesIn Is Nothing Then
                
            'Loop through Redemption Fee records passed in
            For Each objFeeBandNode In objXmlFeesIn.childNodes
                If Not objFeeBandNode Is Nothing Then
                    
                    'assign values
                    lngStep = CSafeLng(objFeeBandNode.Attributes.getNamedItem("REDEMPTIONFEESTEPNUMBER").Text)
                    If AttributeValueExists(objFeeBandNode, "PERIOD") Then
                        lngPeriod = CSafeLng(objFeeBandNode.Attributes.getNamedItem("PERIOD").Text)
                    Else
                        dteEndDate = objFeeBandNode.Attributes.getNamedItem("PERIODENDDATE").Text
                    End If
                    dblPercentage = CSafeDbl(objFeeBandNode.Attributes.getNamedItem("FEEPERCENTAGE").Text)
                    
                    'calculate fee percentage
                    If dblPercentage > 0 Then
                        dblFeeAmount = dblLoanAmount * dblPercentage / 100
                    End If
                    
                    'build XML response block (Individual fee amounts)
                    Set objXmlFeeElem = objXmlOut.createElement("REDEMPTIONFEEAMOUNT")
                    objXmlFeeElem.setAttribute "STEP", lngStep
                    
                    'if we have a value for Fee Period / Period End Date assign it as an attribute
                    If Len(lngPeriod) > 0 Then
                        objXmlFeeElem.setAttribute "FEEPERIOD", lngPeriod
                        objXmlFeeElem.setAttribute "FEEPERIODENDDATE", ""
                    Else
                        objXmlFeeElem.setAttribute "FEEPERIODENDDATE", dteEndDate
                        objXmlFeeElem.setAttribute "FEEPERIOD", ""
                    End If
                    
                    '*** what should this be if we haven't set it above? ***
                    objXmlFeeElem.setAttribute "AMOUNT", dblFeeAmount
                    
                    'add individual record to fees block.
                    objXmlFeeListElem.appendChild objXmlFeeElem
                    
                Else
                    Err.Raise eXMLMISSINGELEMENT, strFunctionName, "tag not found:  REDEMPTIONFEEBAND"
                End If
            Next
        End If
        
        'add fee details to the response block
        objXmlResponseElem.appendChild objXmlFeeListElem
        
        'send response back to calling function
        CalculateRedemptionFees = objXmlOut.xml
    Else
        Err.Raise eFAILEDTOLOADREQUEST, TypeName(Me), "Failed to load XML for Redemption Fee calculation: " & objDOMFeesIn.parseError.reason
    End If
    

'**************************************
'Close down function

CalculateRedemptionFeesExit:

    Set objXmlFeeListElem = Nothing
    Set objXmlFeeElem = Nothing
    Set objXmlFeesIn = Nothing
    Set objFeeBandNode = Nothing
    Set objXmlOut = Nothing
    Set objXmlNode = Nothing
    Set objXmlResponseElem = Nothing
  
    CheckError strFunctionName, TypeName(Me)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
End Function

Private Sub ObjectControl_Activate()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()

End Sub
