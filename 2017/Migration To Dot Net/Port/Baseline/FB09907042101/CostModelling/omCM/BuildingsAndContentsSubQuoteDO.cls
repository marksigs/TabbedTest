VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BuildingsAndContentsSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    
'Workfile:      BuildingsAndContentsSubQuoteDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Data Object
'               contains (implied) standard interface:
'               Public Sub Create(vxmlTableElement As IXMLDOMElement)
'               Public Sub Update(vxmlTableElement As IXMLDOMElement)
'               Public Sub Delete(vxmlTableElement As IXMLDOMElement)
'               Public Function GetData(vxmlTableElement As IXMLDOMElement,
'                                       vxmlInNode As IXMLDOMNode) As IXMLDOMNode
'               Public Function AddDerivedData(vxmlData As IXMLDOMNode)
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'MCS    17/08/99 Created
'MCS    30/08/99 Update following code review
'MCS    01/09/99 Optional timimg output added
'PSC    30/11/99 Amend to use DOMDocuments rather than strings
'JLD    09/02/2000 UpdateSubQuoteDetails() now also saves BCNOTES data.
'JLD    29/03/00    Amended FindList search criteria.
'JLD    17/04/00    Added product data to the return of FindList() also restructured returned
'                   XML to conform to standard
'IVW    09/11/00 Amended to disconnect record sets more efficiently
'DM     17/05/01    SYS2315
'PSC    01/08/01 SYS2546 Disconnect recordset befor closing it in FindList
'AD     04/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
'------------------------------------------------------------------------------------------

Option Explicit

Implements IBuildingsAndContentsSubQuoteDO


 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================

'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date     Description
'GD     16/08/02 BMIDS00312 - altered SQL in _FindList to exclude BUILDINGANDCONTENTSPRODUCT.
'MC     16/06/2004  BMIDS763    - INSURANCEADMIN FEE METHOD IMPLEMENTATION
'------------------------------------------------------------------------------------------
Public Sub Create(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlClassDefDoc = LoadData()
    objDOAssist.CreateEx vxmlTableElement, xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:

    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description


End Sub

Public Sub Delete(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "Delete"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
               
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
            
    Dim objDOAssist As DOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set xmlClassDefDoc = LoadData()
    objDOAssist.DeleteEx vxmlTableElement, xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

DeleteVbErr:

    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Public Sub DeleteAll(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr

    Const strFunctionName As String = "DeleteAll"

    Dim objErrAssist As New ErrAssist
    
    
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
        
    Dim objDOAssist As DOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set xmlClassDefDoc = LoadData()
    objDOAssist.DeleteAllEx vxmlTableElement, xmlClassDefDoc
 
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Sub

DeleteAllVbErr:

    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Public Function GetData(ByVal vxmlTableElement As IXMLDOMElement, _
                        ByVal vxmlInNode As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
'   vxmlInNode      xml node to which data should be attached
' return:
'                   xml node that has been attached
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "GetData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlClassDefDoc = LoadData()
    Set GetData = objDOAssist.GetDataEx(vxmlTableElement, LoadData, vxmlInNode)

    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
        
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Public Function FindList(ByVal vxmlTableElement As IXMLDOMElement, _
                         ByVal vxmlInNode As IXMLDOMNode) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
'   vxmlInNode      xml node to which data should be attached
' return:
'                   xml node that has been attached
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindList"

    Dim objErrAssist As New ErrAssist
    
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented
    
    Dim objDOAssist As DOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set xmlClassDefDoc = LoadData()
    Set FindList = objDOAssist.FindListEx(vxmlTableElement, xmlClassDefDoc, vxmlInNode)

    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************

    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Public Sub Update(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objDOAssist As DOAssist
    
    ' TODO ************************************************************************
    '   Remove omiga4NotImplemented error when this function has been implemented
    ' *****************************************************************************
    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented

    If objContext Is Nothing Then
      Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlClassDefDoc = LoadData()
    objDOAssist.UpdateEx vxmlTableElement, xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    Set objDOAssist = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    
    Exit Sub

UpdateVbErr:
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function LoadData() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
Dim XMLstr As String

    Const strFunctionName As String = "LoadData"

    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "MORTGAGESUBQUOTE"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>BCSUBQUOTENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DATEANDTIMEGENERATED<TYPE>dbdtDateTime</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>QUOTATIONTYPE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TOTALBCMONTHLYCOST<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BCPRODUCTNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"

    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set LoadData = xmlClassDefDoc
    
    Set xmlClassDefDoc = Nothing
    Set objXmlAssist = Nothing

End Function

Public Sub AddDerivedData(ByVal vxmlData As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vxmlData          base XML node
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "AddDerivedData"

    Dim objErrAssist As New ErrAssist

    Set objErrAssist = Nothing
    
    Exit Sub
            
    Set objErrAssist = Nothing
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Sub
    
AddDerivedDataVBErr:

    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_AddDerivedData(ByVal vxmlData As IXMLDOMNode)

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_CreateBCNotes(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_CreateBCNotes"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCNotesData()
    'objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "BCSUBQUOTENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description
    

End Sub

Private Function IBuildingsAndContentsSubQuoteDO_CreateSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubQuoteVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_CreateSubQuoteDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefSubQuoteDoc As FreeThreadedDOMDocument40
    Dim xmlClassDefDetailsDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objXmlAssist As XMLAssist
    Dim strBCSubQuoteNumber As String
    Dim xmlTmpNode As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Set key for SubQuote and create it
    '
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefSubQuoteDoc = objIClassDef.LoadBCSubQuoteData()
    'If a sequence number was passed in then use that, else generate a new one
    Set xmlTmpNode = vxmlTableElement.selectSingleNode(".//BUILDINGSANDCONTENTSSUBQUOTE/BCSUBQUOTENUMBER")
    If Not xmlTmpNode Is Nothing Then
        strBCSubQuoteNumber = xmlTmpNode.Text
        If strBCSubQuoteNumber = "" Then
            objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefSubQuoteDoc, "BCSUBQUOTENUMBER"
        End If
    Else
        objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefSubQuoteDoc, "BCSUBQUOTENUMBER"
    End If
    objIDOAssist.Create vxmlTableElement, xmlClassDefSubQuoteDoc
    
    
    '
    ' Set key for Details and create it
    '
    Set xmlClassDefDetailsDoc = objIClassDef.LoadBCDetailsData()
    'If a sequence number was passed in then use that, else generate a new one
    Set xmlTmpNode = vxmlTableElement.selectSingleNode(".//BUILDINGSANDCONTENTSDETAILS/BCSUBQUOTENUMBER")
    If Not xmlTmpNode Is Nothing Then
        strBCSubQuoteNumber = xmlTmpNode.Text
        If strBCSubQuoteNumber = "" Then
            objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDetailsDoc, "BCSUBQUOTENUMBER"
        End If
    Else
        objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDetailsDoc, "BCSUBQUOTENUMBER"
    End If
    objIDOAssist.Create vxmlTableElement, xmlClassDefDetailsDoc
    
    Set IBuildingsAndContentsSubQuoteDO_CreateSubQuoteDetails = vxmlTableElement.cloneNode(True)
    
    Set objIClassDef = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

CreateSubQuoteVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IBuildingsAndContentsSubQuoteDO_CreateValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateValuablesVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_CreateValuablesOverLimit"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim xmlTmpNode As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Set key for valuable (if necessary) and create it
    '
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCValuablesOLimitData()
    'SR 03/07/01 : SYS2433 - Search in the current context
    If vxmlTableElement.nodeName = "VALUABLESOVERLIMIT" Then
        Set xmlTmpNode = vxmlTableElement.selectSingleNode(".//VALUABLESLIMITSEQUENCENUMBER")
    Else
        Set xmlTmpNode = vxmlTableElement.selectSingleNode("./VALUABLESOVERLIMIT/VALUABLESLIMITSEQUENCENUMBER")
    End If
    
    If Not xmlTmpNode Is Nothing Then
        If xmlTmpNode.Text = "" Then
            objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "VALUABLESLIMITSEQUENCENUMBER"
        End If
    Else
        objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "VALUABLESLIMITSEQUENCENUMBER"
    End If
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set IBuildingsAndContentsSubQuoteDO_CreateValuablesOverLimit = vxmlTableElement.cloneNode(True)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlTmpNode = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

CreateValuablesVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlTmpNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IBuildingsAndContentsSubQuoteDO_DeleteAllValuables(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) to be deleted
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllValuablesVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_DeleteAllValuables"

    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
        
    Dim objDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim objIClassDef As IomCMClassDef
    Set objIClassDef = New omCMClassDef
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Set xmlClassDefDoc = objIClassDef.LoadBCValuablesOLimitData
    objDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteAllValuablesVbErr:

    Set xmlClassDefDoc = Nothing
    Set objIClassDef = Nothing
    Set objDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_DeleteValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteContact a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteValuablesVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_DeleteValuablesOverLimit"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCValuablesOLimitData
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteValuablesVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_DeleteAll(ByVal vxmlTableElement As IXMLDOMElement)

End Sub


Private Function IBuildingsAndContentsSubQuoteDO_FindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_FindList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist  As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist, _
        objSQLAssist As New SQLAssist
    Dim objComboDO As ComboDO

    Dim xmlTopElement As IXMLDOMElement
    Dim xmlListNode As IXMLDOMNode
    Dim xmlClassDefSubQuoteDoc As FreeThreadedDOMDocument40
    Dim xmlClassDefDetailsDoc As FreeThreadedDOMDocument40
    Dim xmlClassDefProductDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Dim strSQLString As String
    Dim strTagValue As String
    Dim blnTagExists As Boolean
    Dim blnDummy As Boolean

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    Set objIClassDef = New omCMClassDef

    ' Get WHERE criteria from XML
    '
    strSQLString = ""
    ' application number
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER", blnTagExists)
    If blnTagExists Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONNUMBER = '" & strTagValue & "')"
    End If

    ' Application Fact find number
    strTagValue = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER", blnTagExists)
    If blnTagExists Then
        strSQLString = strSQLString & IIf(strSQLString = "", "", " AND ") & _
            "(BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONFACTFINDNUMBER = " & strTagValue & " )"
    End If
    
    '
    ' Build the SQL query and execute
    '
    'GD BMIDS00312 strSQLString = "SELECT * FROM BUILDINGSANDCONTENTSSUBQUOTE,BUILDINGSANDCONTENTSDETAILS,BUILDINGANDCONTENTSPRODUCT WHERE " _
                   & strSQLString
    strSQLString = "SELECT * FROM BUILDINGSANDCONTENTSSUBQUOTE,BUILDINGSANDCONTENTSDETAILS WHERE " _
                   & strSQLString
                   
    strSQLString = strSQLString & " AND BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONNUMBER = BUILDINGSANDCONTENTSDETAILS.APPLICATIONNUMBER "
    strSQLString = strSQLString & " AND BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONFACTFINDNUMBER = BUILDINGSANDCONTENTSDETAILS.APPLICATIONFACTFINDNUMBER "
    strSQLString = strSQLString & " AND BUILDINGSANDCONTENTSSUBQUOTE.BCSUBQUOTENUMBER = BUILDINGSANDCONTENTSDETAILS.BCSUBQUOTENUMBER "
    'GD BMIDS00312 strSQLString = strSQLString & " AND BUILDINGSANDCONTENTSSUBQUOTE.BCPRODUCTNUMBER = BUILDINGANDCONTENTSPRODUCT.BCPRODUCTNUMBER "

   'IVW - Disconnecting record sets
   
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    cmd.CommandText = strSQLString
           
    rst.CursorLocation = adUseClient
    rst.CursorType = adOpenForwardOnly
    rst.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rst.Source = cmd
    
    rst.Open
    
    'PSC 01/08/01 - SYS2546
    Set rst.ActiveConnection = Nothing
    
    'IVW - END

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If

    '
    ' Convert the recordset into XML
    '
    Set xmlListNode = xmlOut.appendChild(xmlOut.createElement("BCSUBQUOTEDETAILSLIST"))
    Set xmlClassDefSubQuoteDoc = objIClassDef.LoadBCSubQuoteData()
    Set xmlClassDefDetailsDoc = objIClassDef.LoadBCDetailsData
    'GD BMIDS00312 Set xmlClassDefProductDoc = objIClassDef.LoadBCProductData

    ' AD 04/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    Dim xmlSubQuoteNode As IXMLDOMNode
    Do While Not rst.EOF
        Set xmlSubQuoteNode = objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefSubQuoteDoc, xmlListNode)
        objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDetailsDoc, xmlSubQuoteNode
        'GD BMIDS00312 objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefProductDoc, xmlSubQuoteNode
        rst.MoveNext
    Loop
    
    rst.Close
    conn.Close


    Set IBuildingsAndContentsSubQuoteDO_FindList = xmlOut.documentElement

    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    
    Set cmd = Nothing
    Set rst = Nothing
    Set conn = Nothing
    
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlOut = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    ' Close the recordset if still open
    If Not rst Is Nothing Then
            If rst.State = adStateOpen Then
                    rst.Close
            End If
    End If

    ' Close the connection if still open
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
    End If
    
    Set objComboDO = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    
    Set cmd = Nothing
    Set rst = Nothing
    Set conn = Nothing
    
    Set xmlTopElement = Nothing
    Set xmlListNode = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlOut = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function


Private Function IBuildingsAndContentsSubQuoteDO_FindValuablesOverLimitList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_FindValuablesOverLimitList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCValuablesOLimitData()
    Set IBuildingsAndContentsSubQuoteDO_FindValuablesOverLimitList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description
    
End Function

Private Function IBuildingsAndContentsSubQuoteDO_GetBCNotes(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_GetBCNotes"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCNotesData()
    Set IBuildingsAndContentsSubQuoteDO_GetBCNotes = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

'*=BMIDS763
Private Sub IBuildingsAndContentsSubQuoteDO_GetInsuranceAdminFee(ByVal vXMLRequest As MSXML2.IXMLDOMNode, ByVal vXMLResponse As MSXML2.IXMLDOMNode)
    
    Dim objCommand          As Command
    Dim objDBConn           As Connection
    Dim objParam            As Parameter
    Dim sApplicationNumber  As String
    Dim lAppFactFindNumber  As Long
    Dim rstData             As Recordset
    Dim TempXML             As FreeThreadedDOMDocument40
    Dim sTemp               As String
    Dim xmlEleTemp          As IXMLDOMNode
    Dim dAmount             As Double
    
    On Error GoTo ErrorGetInsuranceAdminFee
    
    Set TempXML = New FreeThreadedDOMDocument40
    TempXML.validateOnParse = False
    TempXML.setProperty "NewParser", True
    Set objCommand = New Command
    
    If xmlGetNode(vXMLResponse, "INSURANCEADMINFEE") Is Nothing Then
        Set xmlEleTemp = vXMLResponse.appendChild(TempXML.createElement("INSURANCEADMINFEE"))
    Else
        Set xmlEleTemp = vXMLResponse.selectSingleNode("INSURANCEADMINFEE")
    End If
    
    If xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
        Call xmlEleTemp.appendChild(TempXML.createElement("AMOUNT"))
    End If

    sApplicationNumber = xmlGetNodeText(vXMLRequest, "APPLICATIONNUMBER")
    lAppFactFindNumber = Val(xmlGetNodeText(vXMLRequest, "APPLICATIONFACTFINDNUMBER"))
        
    With objCommand
        .CommandText = "USP_GetInsuranceAdminFeeAmount"
        .CommandType = adCmdStoredProc
        
        Set objParam = objCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(sApplicationNumber), sApplicationNumber)
        Call .Parameters.Append(objParam)
        
        Set objParam = objCommand.CreateParameter("APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, 5, lAppFactFindNumber)
        Call .Parameters.Append(objParam)
        
    End With
    
    Set rstData = executeGetRecordSet(objCommand)
    
    If Not rstData Is Nothing Then
        If Not rstData.EOF And Not rstData.BOF Then
            If Not xmlGetNode(xmlEleTemp, "AMOUNT") Is Nothing Then
                dAmount = Val(rstData.Fields("Amount").Value & "")
                xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount 'rstData.Fields("Amount").Value
            End If
        End If
    Else
        xmlGetNode(xmlEleTemp, "AMOUNT").Text = dAmount
    End If
    
    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            rstData.Close
        End If
    End If
    
ExitGetInsuranceAdminFee:
    
    Set rstData = Nothing
    Set objParam = Nothing
    Set objCommand = Nothing
    Set objDBConn = Nothing
    
    Exit Sub
    
ErrorGetInsuranceAdminFee:
    
    'log error
    Resume ExitGetInsuranceAdminFee
    
    Exit Sub
    
    Resume
    
End Sub

Private Function IBuildingsAndContentsSubQuoteDO_GetPaymentFrequency(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetFreqVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_GetPaymentFrequency"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objComboDO As ComboDO
    Dim xmlBCDetailsClassDefDoc As FreeThreadedDOMDocument40
    Dim objBCDetailsIClassDef As IomCMClassDef
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlListElem As IXMLDOMElement
    Dim xmlTmpElem As IXMLDOMElement
    Dim objSQLAssist As New SQLAssist
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strBCSubQuoteNumber As String
    Dim strSQL As String
    
    strApplicationNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFFNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strBCSubQuoteNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "BCSUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFFNumber) = 0 Or _
       Len(strBCSubQuoteNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "APPLICATIONNUMBER = " & strApplicationNumber & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNumber & _
                                "; BCSUBQUOTENUMBER = " & strBCSubQuoteNumber
    End If
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
        Set objComboDO = New ComboDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    strSQL = "SELECT FREQUENCY FROM BUILDINGSANDCONTENTSDETAILS WHERE"
    strSQL = strSQL & " APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber) & " AND"
    strSQL = strSQL & " APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & strApplicationFFNumber & " AND"
    strSQL = strSQL & " BCSUBQUOTENUMBER = "
    strSQL = strSQL & strBCSubQuoteNumber
    
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
   
    'IVW
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If

    
    Set objBCDetailsIClassDef = New omCMClassDef
    Set xmlBCDetailsClassDefDoc = objBCDetailsIClassDef.LoadBCDetailsData
    
    ' AD 04/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    Set xmlListElem = xmlOut.createElement("BUILDINGSANDCONTENTSDETAILS")
    xmlOut.appendChild xmlListElem

    Set xmlTmpElem = xmlOut.createElement("FREQUENCY")
    'need to get the combo validation type for the valueId returned
    Dim strValidationType As String
    strValidationType = objComboDO.GetFirstComboValidation("BCPaymentFreq", rstThisRecSet.Fields.Item(0).Value)
    xmlTmpElem.Text = strValidationType
    xmlListElem.appendChild xmlTmpElem
    ' loop through the record set
    'While Not rstThisRecSet.EOF
    '    objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlBCDetailsClassDefDoc, xmlListElem
    '    rstThisRecSet.MoveNext
    'Wend
        
   
        rstThisRecSet.Close
        Set rstThisRecSet.ActiveConnection = Nothing
    
    conn.Close
         
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    Set IBuildingsAndContentsSubQuoteDO_GetPaymentFrequency = xmlListElem.cloneNode(True)

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objComboDO = Nothing
    Set objErrAssist = Nothing
    Set xmlListElem = Nothing
    Set objBCDetailsIClassDef = Nothing
    Set xmlBCDetailsClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetFreqVbErr:
    
        ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
        
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objComboDO = Nothing
    Set xmlListElem = Nothing
    Set objBCDetailsIClassDef = Nothing
    Set xmlBCDetailsClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
   
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IBuildingsAndContentsSubQuoteDO_GetSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'BG SYS1945 13/02/01 Changed the way the activeconnection was set to nothing
'------------------------------------------------------------------------------------------
On Error GoTo GetSubQuoteDetailsVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_GetSubQuoteDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim xmlBCSubQuoteClassDefDoc As FreeThreadedDOMDocument40
    Dim objBCSubQuoteIClassDef As IomCMClassDef
    Dim xmlBCDetailsClassDefDoc As FreeThreadedDOMDocument40
    Dim objBCDetailsIClassDef As IomCMClassDef
    'Dim xmlBCNotesClassDefDoc As DOMDocument
    'Dim objBCNotesIClassDef As IomCMClassDef
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlListElem As IXMLDOMElement
    Dim objSQLAssist As New SQLAssist
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strBCSubQuoteNumber As String
    Dim strSQL As String
    
    strApplicationNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFFNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strBCSubQuoteNumber = _
        objXmlAssist.GetTagValue(vxmlTableElement, "BCSUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFFNumber) = 0 Or _
       Len(strBCSubQuoteNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "APPLICATIONNUMBER = " & strApplicationNumber & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNumber & _
                                "; BCSUBQUOTENUMBER = " & strBCSubQuoteNumber
    End If
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    'DM 17/05/01 SYS2315
    #If GENERIC_SQL Then
    #Else
        strSQL = "SELECT * FROM BUILDINGSANDCONTENTSNOTES, BUILDINGSANDCONTENTSDETAILS, BUILDINGSANDCONTENTSSUBQUOTE WHERE"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONNUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSDETAILS.APPLICATIONNUMBER(+) AND"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONNUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSNOTES.APPLICATIONNUMBER(+) AND"
        
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSDETAILS.APPLICATIONFACTFINDNUMBER(+) AND"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSNOTES.APPLICATIONFACTFINDNUMBER(+) AND"
        
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.BCSUBQUOTENUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSDETAILS.BCSUBQUOTENUMBER(+) AND"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.BCSUBQUOTENUMBER = "
        strSQL = strSQL & " BUILDINGSANDCONTENTSNOTES.BCSUBQUOTENUMBER(+) AND"
        
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONNUMBER = "
        strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber) & " AND"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
        strSQL = strSQL & strApplicationFFNumber & " AND"
        strSQL = strSQL & " BUILDINGSANDCONTENTSSUBQUOTE.BCSUBQUOTENUMBER = "
        strSQL = strSQL & strBCSubQuoteNumber
    #End If
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset
    
    'DM 17/05/01 SYS2315
    #If GENERIC_SQL Then
        cmd.CommandText = "usp_GetSubQuoteDetails"
        cmd.CommandType = adCmdStoredProc
        ' Set up the parameters
        cmd.Parameters.Append cmd.CreateParameter("p_ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        Dim param As ADODB.Parameter
        
        Set param = cmd.CreateParameter("p_ApplicationFFNumber", adNumeric, adParamInput, , CInt(strApplicationFFNumber))
        param.Precision = 5
        cmd.Parameters.Append param

        Set param = cmd.CreateParameter("p_BCSubQuoteNumber", adNumeric, adParamInput, , CInt(strBCSubQuoteNumber))
        param.Precision = 5
        cmd.Parameters.Append param

    #Else
        cmd.CommandText = strSQL
    #End If
    
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    Set rstThisRecSet.ActiveConnection = Nothing
     
    'IVW
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
    
    Set objBCSubQuoteIClassDef = New omCMClassDef
    Set xmlBCSubQuoteClassDefDoc = objBCSubQuoteIClassDef.LoadBCSubQuoteData()
    Set objBCDetailsIClassDef = New omCMClassDef
    Set xmlBCDetailsClassDefDoc = objBCSubQuoteIClassDef.LoadBCDetailsData
    'Set objBCNotesIClassDef = New omCMClassDef
    'Set xmlBCNotesClassDefDoc = objBCSubQuoteIClassDef.LoadBCNotesData
        
    ' AD 04/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
        
    Set xmlListElem = xmlOut.createElement("BCSUBQUOTEDETAILSNOTES")
    xmlOut.appendChild xmlListElem

    ' loop through the record set
    While Not rstThisRecSet.EOF
        objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlBCSubQuoteClassDefDoc, xmlListElem
        objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlBCDetailsClassDefDoc, xmlListElem
        'objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlBCNotesClassDefDoc, xmlListElem
        rstThisRecSet.MoveNext
    Wend

        rstThisRecSet.Close
    
    
    conn.Close

            
    Set IBuildingsAndContentsSubQuoteDO_GetSubQuoteDetails = xmlListElem.cloneNode(True)

    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlListElem = Nothing
    Set objBCSubQuoteIClassDef = Nothing
    Set xmlBCSubQuoteClassDefDoc = Nothing
    Set objBCDetailsIClassDef = Nothing
    Set xmlBCDetailsClassDefDoc = Nothing
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    'Set objBCNotesIClassDef = Nothing
    'Set xmlBCNotesClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetSubQuoteDetailsVbErr:
    
        ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
        
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlListElem = Nothing
    Set objBCSubQuoteIClassDef = Nothing
    Set xmlBCSubQuoteClassDefDoc = Nothing
    Set objBCDetailsIClassDef = Nothing
    Set xmlBCDetailsClassDefDoc = Nothing
    'Set objBCNotesIClassDef = Nothing
    'Set xmlBCNotesClassDefDoc = Nothing
    Set objSQLAssist = Nothing
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IBuildingsAndContentsSubQuoteDO_UpdateBCNotes(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_UpdateBCNotes"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist
    
    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCNotesData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_UpdateSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubQuoteVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_UpdateSubQuoteDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefSubQuoteDoc As FreeThreadedDOMDocument40
    Dim xmlClassDefDetailsDoc As FreeThreadedDOMDocument40
    'Dim xmlClassDefNotesDoc As DOMDocument
    Dim objIClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefSubQuoteDoc = objIClassDef.LoadBCSubQuoteData()
    Set xmlClassDefDetailsDoc = objIClassDef.LoadBCDetailsData
    'Set xmlClassDefNotesDoc = objIClassDef.LoadBCNotesData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefSubQuoteDoc
    objIDOAssist.Update vxmlTableElement, xmlClassDefDetailsDoc
    'objIDOAssist.Update vxmlTableElement, xmlClassDefNotesDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    'Set xmlClassDefNotesDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateSubQuoteVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefSubQuoteDoc = Nothing
    Set xmlClassDefDetailsDoc = Nothing
    'Set xmlClassDefNotesDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IBuildingsAndContentsSubQuoteDO_UpdateValuablesOverLimit(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateValuablesVbErr

    Const strFunctionName As String = "IBuildingsAndContentsSubQuoteDO_UpdateValuablesOverLimit"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCValuablesOLimitData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateValuablesVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

