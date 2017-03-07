VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuickQuotePersonalDebtsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuotePersonalDebtsDO.cls
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
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MS     17/08/99    Created
'RF     01/09/99    GetPersonalDebts renamed to GetData
'MCS    28/09/99    Updated following code reviews
'------------------------------------------------------------------------------------------
Option Explicit

 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================

Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
' Raise Errors:
'------------------------------------------------------------------------------------------

' *****************************************************************************
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"
   
    Dim objErrAssist As New ErrAssist
    Dim objDOAssist As DOAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    objDOAssist.Create vstrXMLRequest, LoadData
    
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

    Dim objErrAssist As New ErrAssist
    
    #If TIME_DO Then
        Dim sngStartTime As Single
        Dim sngEndTime As Single
        sngStartTime = Timer
    #End If

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
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
'Public Function GetPersonalDebts(ByVal vstrXMLRequest As String) As String
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
            
    GetData = objDOAssist.GetData(vstrXMLRequest, LoadData)
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function

GetDataVbErr:
        
    Set objDOAssist = Nothing
    
    GetData = ""
    
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
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
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
    
    objDOAssist.Update vstrXMLRequest, LoadData
    
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

Public Function LoadData()
'------------------------------------------------------------------------------------------
'
'APPLICATIONNUMBER              NOT NULL VARCHAR2(12)
'APPLICATIONFACTFINDNUMBER      NOT NULL NUMBER(5)
'CUSTOMERNUMBER                 NOT NULL VARCHAR2(12)
'CUSTOMERVERSIONNUMBER          NOT NULL NUMBER(5)
'QUICKQUOTEPERSONALDEBTSTYPE    NOT NULL NUMBER(5)
'OUTSTANDINGBALANCE Number(10)
'MONTHLYPAYMENT Number(5)
'
' *****************************************************************************
Dim XMLstr As String

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "QUICKQUOTEPERSONALDEBTS"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>QUICKQUOTEPERSONALDEBTSTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>PersonalDebtsType</COMBO></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>OUTSTANDINGBALANCE<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MONTHLYPAYMENT<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

LoadData = strXML

End Function

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

