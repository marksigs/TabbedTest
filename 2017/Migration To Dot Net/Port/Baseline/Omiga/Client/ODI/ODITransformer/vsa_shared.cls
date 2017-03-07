VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vsa_shared"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      VSA_Shared.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   Omiga support for Visual Studio Analyzer
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     26/09/00    Created
'------------------------------------------------------------------------------------------
Option Explicit

Private Type Guid
    D1       As Long
    D2       As Integer
    D3       As Integer
    D4(8)    As Byte
End Type
Private Declare Function VSAWinCoCreateGuid Lib "OLE32.DLL" Alias "CoCreateGuid" (guidNewGuid As Guid) As Long

'------------------------------------------------------------------------------------------

#If USING_VSA Then
    'These assignments allow you to create the ESI (Event Source Installer)
    'and the IEC (InProc Event Creator) objects so that you can begin a VSA
    'session and fire VSA events.  You will need to select the following
    'type libraries in the References dialog box.
    '   MSVSA Automatable Inproc Event Creator 1.0 Type Library
    Private IEC As MSVSAInprocEventCreator
    
    ' Correleation IDs for Call/Return and Enter/Leave pairs
    Dim m_VSACallReturn As String
    Dim m_VSAEnterLeave As String
    
    'These variables define the event parameters and values
    Dim Keys(1) As String
    Dim Values(1) As String
    Dim lCount As Long
    Dim lFlags As Long
    
#End If

Public Sub Initialize(ByVal strFunctionName As String)
    #If USING_VSA Then
        'Begin the VSA session
        Set IEC = New MSVSAInprocEventCreator
        IEC.BeginSession gstrVSAEventSource, gstrVSASessionPrefix & "." & strFunctionName & " " & VSACreateGuid
        'Fire event
        Fire_OnCall
    #End If
End Sub

Private Sub Class_Terminate()
    #If USING_VSA Then
        Fire_OnReturn
        
        'Now that you are finished, it is time to end the VSA session
        IEC.EndSession
        Set IEC = Nothing
    #End If
End Sub

Public Sub Fire_OnCall()
    #If USING_VSA Then
        'First you need to initialize the array that you will use to pass
        'data. Since you are passing some data, you will need an array of as
        'many elements as you are passing.
        Const MAXVALUE = 1
        Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
        lCount = MAXVALUE
       
        'Use the cVSAEventDefaultSource flag for Call/Return events from the
        'caller.
        lFlags = cVSAEventDefaultSource
    
        'The CorrelationID needs to be the same value for pairs of the
        'Call/Enter/Leave/Return events. In this case, the Call/Return pair
        'has the same CorrelationID value.
        Keys(1) = PARAM_CorrelationID
        m_VSACallReturn = VSACreateGuid ' save for later use by Fire_OnReturn
        Values(1) = m_VSACallReturn
          
        'fire the CALL event
        IEC.FireEvent DEBUG_EVENT_CALL, Keys, Values, lCount, lFlags
    #End If
End Sub

Public Sub Fire_OnReturn()
    #If USING_VSA Then
        'First you need to initialize the array that you will use to pass
        'data. Since you are passing some data, you will need an array of as
        'many elements as you are passing.
        Const MAXVALUE = 1
        Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
        lCount = MAXVALUE
       
        'Use the cVSAEventDefaultSource flag for Call/Return events from the
        'caller.
        lFlags = cVSAEventDefaultSource
    
        'The CorrelationID needs to be the same value for pairs of the
        'Call/Enter/Leave/Return events. In this case, the Call/Return pair
        'has the same CorrelationID value.
        Keys(1) = PARAM_CorrelationID
        ' m_VSACallReturn initialised in Fire_OnCall
        Values(1) = m_VSACallReturn
          
        'fire the RETURN event
        IEC.FireEvent DEBUG_EVENT_RETURN, Keys, Values, lCount, lFlags
    #End If
End Sub

Public Sub Fire_OnEnter()
    #If USING_VSA Then
        'First you need to initialize the array that you will use to pass
        'data. Since you are passing some data, you will need an array of as
        'many elements as you are passing.
        Const MAXVALUE = 1
        Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
        lCount = MAXVALUE
       
        'Use the cVSAEventdefaultTarget flag for the Enter/Leave events from
        'the callee.
        lFlags = cVSAEventDefaultTarget
    
        'The CorrelationID needs to be the same value for pairs of the
        'Call/Enter/Leave/Return events. In this case, the Enter/Leave pair
        'has the same CorrelationID value.
        Keys(1) = PARAM_CorrelationID
        m_VSAEnterLeave = VSACreateGuid ' save for later use by Fire_OnLeave
        Values(1) = m_VSAEnterLeave
          
        'fire the ENTER event
        IEC.FireEvent DEBUG_EVENT_ENTER, Keys, Values, lCount, lFlags
    #End If
End Sub

Public Sub Fire_OnLeave()
    #If USING_VSA Then
        'First you need to initialize the array that you will use to pass
        'data. Since you are passing some data, you will need an array of as
        'many elements as you are passing.
        Const MAXVALUE = 1
        Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
        lCount = MAXVALUE
       
        'Use the cVSAEventdefaultTarget flag for the Enter/Leave events from
        'the callee.
        lFlags = cVSAEventDefaultTarget
    
        'The CorrelationID needs to be the same value for pairs of the
        'Call/Enter/Leave/Return events. In this case, the Enter/Leave pair
        'has the same CorrelationID value.
        Keys(1) = PARAM_CorrelationID
        ' m_VSAEnterLeave initialised in Fire_OnEnter
        Values(1) = m_VSAEnterLeave
          
        'fire the LEAVE event
        IEC.FireEvent DEBUG_EVENT_LEAVE_NORMAL, Keys, Values, lCount, lFlags
    #End If
End Sub

Public Sub Fire_DEBUG_EVENT_CALL(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_CALL, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_RETURN(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_RETURN, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_COMPONENT_START(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_COMPONENT_START, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_COMPONENT_STOP(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_COMPONENT_STOP, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_CALL_DATA(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_CALL_DATA, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_ENTER(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_ENTER, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_ENTER_DATA(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_ENTER_DATA, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_LEAVE_NORMAL(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_LEAVE_NORMAL, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_LEAVE_EXCEPTION(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_LEAVE_EXCEPTION, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_LEAVE_DATA(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_LEAVE_DATA, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_RETURN_DATA(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_RETURN_DATA, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_RETURN_NORMAL(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_RETURN_NORMAL, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_RETURN_EXCEPTION(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_RETURN_EXCEPTION, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_QUERY_SEND(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_QUERY_SEND, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_QUERY_ENTER(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_QUERY_ENTER, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_QUERY_LEAVE(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_QUERY_LEAVE, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_QUERY_RESULT(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_QUERY_RESULT, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_TRANSACTION_START(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_TRANSACTION_START, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_TRANSACTION_COMMIT(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_TRANSACTION_COMMIT, strText
    #End If
End Sub
Public Sub Fire_DEBUG_EVENT_TRANSACTION_ROLLBACK(Optional ByVal strText As String = "")
    #If USING_VSA Then
        InternalFireEvent1 DEBUG_EVENT_TRANSACTION_ROLLBACK, strText
    #End If
End Sub
Public Sub Fire_TRACING_GENERAL(ByVal strText As String)
    #If USING_VSA Then
        InternalFireEvent1 OMIGA4_EVENT_TRACING_GENERAL, strText
    #End If
End Sub
Public Sub Fire_ERRORS_SYSERR(ByVal strText As String)
    #If USING_VSA Then
        InternalFireEvent1 OMIGA4_EVENT_ERRORS_SYSERR, strText
    #End If
End Sub
Public Sub Fire_ERRORS_APPERR(ByVal strText As String)
    #If USING_VSA Then
        InternalFireEvent1 OMIGA4_EVENT_ERRORS_APPERR, strText
    #End If
End Sub
Private Sub InternalFireEvent(ByVal strEvent As String)
    #If USING_VSA Then
        'This is one of the functions that fires the VSA events. This function
        'fires the events with default source information and no other info.
        
        'First you need to initialize the array that you will use to pass
        'data. Since you are not passing any data this time, the array will be
        'one element; that is, an empty string.
        Const MAXVALUE = 1
        Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
        Keys(1) = ""
        Values(1) = ""
        lCount = MAXVALUE
        
        'The cVSAEventDefaultSource flag causes source information to be
        'stamped into the event by the LEC.
        lFlags = cVSAEventDefaultSource
        
        'This is the statement that fires the events.
        IEC.FireEvent strEvent, Keys, Values, lCount, lFlags
    #End If
End Sub

Private Sub InternalFireEvent1(ByVal strEvent As String, ByVal sText As String)
    #If USING_VSA Then
       'This is another function that fires the VSA events. This function
       'fires the events with default parameter information and other info
       'as well.
       
       'First you need to initialize the array that you will use to pass
       'data. Since you are passing some data, you will need an array of as
       'many elements as you are passing.
       Const MAXVALUE = 1
       Dim Keys(1 To MAXVALUE) As Variant, Values(1 To MAXVALUE) As Variant
       lCount = MAXVALUE
       
       'Using the cVSAEventDefaultSource or CVSAEventDefaultTarget
       'flags allow the IEC and LEC to stamp all default data into the event.
       'Use the cVSAEventDefaultSource flag for Call/Return events from the
       'caller.
       'Use the cVSAEventdefaultTarget flag for the Enter/Leave events from
       'the callee.
       lFlags = cVSAEventDefaultSource
    
       'The following are custom fields that are being added to the stock
       'events. The string in the Keys argument is used as the field name.
       Keys(1) = PARAM_Arguments
       
       Values(1) = sText
       
       'This is the statement that fires the events.
       IEC.FireEvent strEvent, Keys, Values, lCount, lFlags
    #End If
End Sub


Private Function VSACreateGuid() As String
    Dim guidNewGuid  As Guid
    Dim strBuffer    As String

    Call VSAWinCoCreateGuid(guidNewGuid)

    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D1), 8)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D2), 4)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D3), 4)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(0)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(1)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(2)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(3)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(4)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(5)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(6)), 2)
    VSACreateGuid = PadRight0(VSACreateGuid, Hex$(guidNewGuid.D4(7)), 2)
End Function

Private Function PadRight0(ByVal vstrBuffer As String, _
                           ByVal vstrBit As String, _
                           ByVal intLenRequired As Integer, _
                           Optional bHyp As Boolean _
                         ) As String
' -----------------------------------------------------------------------------------------
' description:
' return:
'------------------------------------------------------------------------------------------
    PadRight0 = vstrBuffer & _
                vstrBit & _
                String$(Abs(intLenRequired - Len(vstrBit)), "0")
End Function


