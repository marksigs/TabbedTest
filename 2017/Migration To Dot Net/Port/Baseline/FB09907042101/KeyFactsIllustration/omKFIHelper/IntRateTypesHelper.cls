VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntRateTypesHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mxmlLoanComponent As IXMLDOMNode
Private mxmlInterestRates As IXMLDOMNodeList
Private mintRatesCount As Integer
Private mintVariableRates As Integer
Private mintFixedRates As Integer
Private mintCappedCollaredRates As Integer

Private Sub Class_Terminate()
    Set mxmlLoanComponent = Nothing
    Set mxmlInterestRates = Nothing
End Sub

'********************************************************************************
'** Function:       Initialise
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Initialises the class with the loan component and interest
'**                 rate type list for it.
'** Parameters:     vxmlLoanComponent - the loan component.
'**                 vxmlInterestRates - the list of interest rate type records
'**                 for the specified loan component.
'** Returns:        N/A
'** Errors:         None Expected
'********************************************************************************
Public Sub Initialise(ByVal vxmlLoanComponent As IXMLDOMNode, _
        ByVal vxmlInterestRates As IXMLDOMNodeList)
    Dim xmlInterestRate As IXMLDOMNode
    Dim eType As MortgageInterestRateType
    
    '*-store the reference to the loan component
    Const cstrFunctionName As String = "Initialise"

    On Error GoTo ErrHandler

    Set mxmlLoanComponent = vxmlLoanComponent
    '*-store the reference to the list
    Set mxmlInterestRates = vxmlInterestRates
    
    '*-initialise the member variables
    mintRatesCount = 0
    mintVariableRates = 0
    mintFixedRates = 0
    mintCappedCollaredRates = 0
        
    '*-count the different types of interest rates
    For Each xmlInterestRate In vxmlInterestRates
        mintRatesCount = mintRatesCount + 1
        eType = GetInterestRateType(xmlInterestRate)
        Select Case eType
            Case mrtStandardVariableRate, mrtDiscountedRate, mrtTrackerAbove
                mintVariableRates = mintVariableRates + 1
            Case mrtFixedRate
                mintFixedRates = mintFixedRates + 1
            Case Else
                mintCappedCollaredRates = mintCappedCollaredRates + 1
        End Select
    Next xmlInterestRate

    Set xmlInterestRate = Nothing
Exit Sub
ErrHandler:
    Set xmlInterestRate = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Sub

Public Property Get InterestRates() As IXMLDOMNodeList
    Set InterestRates = mxmlInterestRates
End Property

Public Property Get LoanComponent() As IXMLDOMNode
    Set LoanComponent = mxmlLoanComponent
End Property

Public Property Get RatesCount() As Integer
    RatesCount = mintRatesCount
End Property

Public Property Get VariableRates() As Integer
    VariableRates = mintVariableRates
End Property

Public Property Get FixedRates() As Integer
    FixedRates = mintFixedRates
End Property

Public Property Get CappedCollaredRates() As Integer
    CappedCollaredRates = mintCappedCollaredRates
End Property

Public Property Get HasMixedRates() As Boolean
    HasMixedRates = (mintVariableRates < mintRatesCount) _
            And (mintFixedRates < mintRatesCount) _
            And (mintCappedCollaredRates < mintRatesCount)
End Property

'********************************************************************************
'** Function:       InterestRateTypeAsText
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Returns the text for an interest rate type value.
'** Parameters:     veType - the interest rate type.
'** Returns:        The text for the specified rate.
'** Errors:         None Expected
'********************************************************************************
Public Function InterestRateTypeAsText(ByVal veType As MortgageInterestRateType) As String
    Select Case veType
        Case mrtStandardVariableRate
            InterestRateTypeAsText = "VARIABLE"
        Case mrtFixedRate
            InterestRateTypeAsText = "FIXED"
        Case mrtDiscountedRate
            InterestRateTypeAsText = "DISCOUNT"
        Case mrtCappedRate
            InterestRateTypeAsText = "CAPPED"
        Case mrtCollaredRate
            InterestRateTypeAsText = "COLLARED"
        Case mrtCappedAndCollaredRate
            InterestRateTypeAsText = "CAPPEDANDCOLLARED"
        Case mrtTrackerAbove
            InterestRateTypeAsText = "TRACKERABOVE"
    End Select
End Function

'********************************************************************************
'** Function:       GetInterestRateWithSequenceNum
'** Created by:     Andy Maggs
'** Date:           19/07/2004
'** Description:    Attempts to get and return the interest rate type record with
'**                 the specified interest rate.
'** Parameters:     vintSeqNum - the interest rate type sequence number to get.
'** Returns:        The interest rate type record if found, or Nothing if not.
'** Errors:         None Expected
'********************************************************************************
Public Function GetInterestRateWithSequenceNum(ByVal vintSeqNum As Integer) As IXMLDOMNode
    Const cstrFunctionName As String = "GetInterestRateWithSequenceNum"
    
    Dim xmlIntRate As IXMLDOMNode
    Dim intSeqNum As Integer

    On Error GoTo ErrHandler

    For Each xmlIntRate In mxmlInterestRates
        intSeqNum = xmlGetAttributeAsInteger(xmlIntRate, "INTERESTRATETYPESEQUENCENUMBER")
        If intSeqNum = vintSeqNum Then
            '*-this is the rate we are looking for, so quit
            Set GetInterestRateWithSequenceNum = xmlIntRate
            Exit Function
        End If
    Next xmlIntRate

    Set xmlIntRate = Nothing
Exit Function
ErrHandler:
    Set xmlIntRate = Nothing
    '*-check the error and throw it on
    errCheckError cstrFunctionName, TypeName(Me)
End Function
