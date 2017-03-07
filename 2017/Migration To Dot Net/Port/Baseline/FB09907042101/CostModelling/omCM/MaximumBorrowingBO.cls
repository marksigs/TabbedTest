VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MaximumBorrowingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'QuotationBO' with actual class name
' *****************************************************************************

'Workfile:      QuotationBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Quotation Business Object which 'supports transactions' only.
'               Any transactioned requests will
'               be delegated to QuotationTxBO

'Dependencies:  QuotationTxBO, QuotationDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    09/11/99    Created
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================




 Public Function CalculateMaximumBorrowing(ByVal vstrXMLRequest As String) As Double
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateMaximumBorrowingVbErr
    
    Const strFunctionName As String = "CalculateMaximumBorrowing"


    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
        
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    'BG SYS0501 23/08/00 changed dblMax1 and 2 data type to Double from long to handle larger numbers.
    Dim dblMax1 As Double
    Dim dblMax2 As Double
    Dim dblEarnedIncomeA As Double
    Dim dblEarnedIncomeB As Double
    dblEarnedIncomeA = 0
    dblEarnedIncomeB = 0

    Dim xmlElement As IXMLDOMElement
            
    Set xmlElement = xmlIn.getElementsByTagName("CUSTOMERA").Item(0)
        
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                   oeMissingPrimaryTag, "CUSTOMERA tag not found"
    End If

    dblEarnedIncomeA = objXmlAssist.GetTagValue(xmlElement, "EARNEDINCOME")
    
    'raise error if lngEarnedIncomeA = 0 ???

    Set xmlElement = xmlIn.getElementsByTagName("CUSTOMERB").Item(0)
        
    Dim dblHighestIncomeMultiple As Double
    Dim dblSecondIncomeMultiple As Double
    Dim dblJointIncomeMultiple As Double
    Dim dblSingleIncomeMultiple As Double
        
    Dim strHighestIncomeMultiple As String
    Dim strSecondIncomeMultiple As String
    Dim strJointIncomeMultiple As String
    Dim strSingleIncomeMultiple As String
        
        
        
    If Not xmlElement Is Nothing Then
        
        dblEarnedIncomeB = objXmlAssist.GetTagValue(xmlElement, "EARNEDINCOME")
        
        'raise error if lngEarnedIncomeB = 0 ???
    
        'check here that we have got the correct multipliers for this calculation
        strHighestIncomeMultiple = objXmlAssist.GetTagValue(xmlIn.documentElement, "HIGHESTEARNINCMULTIPLE")
        strSecondIncomeMultiple = objXmlAssist.GetTagValue(xmlIn.documentElement, "SECONDEARNINCMULTIPLE")
        strJointIncomeMultiple = objXmlAssist.GetTagValue(xmlIn.documentElement, "JOINTINCOMEMULTIPLE")
        
        If Len(strHighestIncomeMultiple) = 0 Or _
               Len(strJointIncomeMultiple) = 0 Or _
               Len(strSecondIncomeMultiple) = 0 Then
        
        'raise error missing income multiples ????
        'Dont do a calculation
            Exit Function
        
        End If
        dblHighestIncomeMultiple = CDbl(strHighestIncomeMultiple)
        dblJointIncomeMultiple = CDbl(strJointIncomeMultiple)
        dblSecondIncomeMultiple = CDbl(strSecondIncomeMultiple)
        
        dblMax1 = (dblEarnedIncomeA * dblHighestIncomeMultiple) + _
                           (dblEarnedIncomeB * dblSecondIncomeMultiple)
        dblMax2 = (dblEarnedIncomeA + dblEarnedIncomeB) * dblJointIncomeMultiple

        
                
    Else
                
        strSingleIncomeMultiple = objXmlAssist.GetTagValue(xmlIn.documentElement, "SINGLEINCOMEMULTIPLE")
        
        If Len(strSingleIncomeMultiple) = 0 Then
        'raise error missing income multiples ????
            Exit Function
        End If
        
        dblSingleIncomeMultiple = CDbl(strSingleIncomeMultiple)
        
        dblMax1 = (dblEarnedIncomeA * dblSingleIncomeMultiple)
        
    End If
    
    'Calculate once here
    Dim dblMaximumBorrowing As Double
    If dblMax1 >= dblMax2 Then
        dblMaximumBorrowing = dblMax1
    Else
        dblMaximumBorrowing = dblMax2
    End If
        
                                                                                                                                    
    CalculateMaximumBorrowing = dblMaximumBorrowing
    

CalculateMaximumBorrowingExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing

    Exit Function

CalculateMaximumBorrowingVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateMaximumBorrowing = objErrAssist.CreateErrorResponse
           
    '   go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function
