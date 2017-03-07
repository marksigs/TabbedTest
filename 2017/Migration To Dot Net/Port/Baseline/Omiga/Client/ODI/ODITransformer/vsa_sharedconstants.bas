Attribute VB_Name = "vsa_sharedconstants"
'Workfile:      VSA_Common.bas
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

#If USING_VSA Then
    'These are the definitions of the VSA Event field names.
    'They are used in the Key array when passing data. The
    'difference between default and non default parameters is
    'that the IEC and LEC provide information in these parameters
    'if the cVSAEventDefaultSource and the cVSAEventDefaultTarget
    'flags are specified in the FireEvent call
    
    ' Default Parameters
    Public Const PARAM_SourceMachine = "SourceMachine"
    Public Const PARAM_SourceProcessID = "SourceProcessId"
    Public Const PARAM_SourceThread = "SourceThread"
    Public Const PARAM_SourceComponent = "SourceComponent"
    Public Const PARAM_SourceSession = "SourceSession"
    Public Const PARAM_TargetMachine = "TargetMachine"
    Public Const PARAM_TargetProcessID = "TargetProcessId"
    Public Const PARAM_TargetThread = "TargetThread"
    Public Const PARAM_TargetComponent = "TargetComponent"
    Public Const PARAM_TargetSession = "TargetSession"
    Public Const PARAM_SecurityIdentity = "SecurityIdentity"
    Public Const PARAM_CausalityID = "CausalityID"
    Public Const PARAM_SourceProcessName = "SourceProcessName"
    Public Const PARAM_TargetProcessName = "TargetProcessName"
    
    ' Non Default Parameters
    Public Const PARAM_SourceHandle = "SourceHandle"
    Public Const PARAM_TargetHandle = "TargetHandle"
    Public Const PARAM_Arguments = "Arguments"
    Public Const PARAM_ReturnValue = "ReturnValue"
    Public Const PARAM_Exception = "Exception"
    Public Const PARAM_CorrelationID = "CorrelationID"
    
    'the null guid
    Public Const GUID_NULL = "{00000000-0000-0000-0000-000000000000}"
    
    'The 'user' event source, which is autoregistered for all stock events
    Public Const DEBUG_EVENT_SOURCE_USER = "{6C736D00-BCBF-11D0-8A23-00AA00B58E10}"
    
    'These are the definitions of the stock events that are available in VSA
    'These are defined in detail in the MSDN documentation
    Public Const DEBUG_EVENT_CALL = "{6C736D61-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_RETURN = "{6C736D62-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_COMPONENT_START = "{6C736D63-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_COMPONENT_STOP = "{6C736D64-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_CALL_DATA = "{6C736D65-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_ENTER = "{6C736D66-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_ENTER_DATA = "{6C736D67-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_LEAVE_NORMAL = "{6C736D68-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_LEAVE_EXCEPTION = "{6C736D69-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_LEAVE_DATA = "{6C736D6A-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_RETURN_DATA = "{6C736D6B-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_RETURN_NORMAL = "{6C736D6C-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_RETURN_EXCEPTION = "{6C736D6D-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_QUERY_SEND = "{6C736D6E-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_QUERY_ENTER = "{6C736D6F-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_QUERY_LEAVE = "{6C736D70-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_QUERY_RESULT = "{6C736D71-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_TRANSACTION_START = "{6C736D72-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_TRANSACTION_COMMIT = "{6C736D73-BCBF-11D0-8A23-00AA00B58E10}"
    Public Const DEBUG_EVENT_TRANSACTION_ROLLBACK = "{6C736D74-BCBF-11D0-8A23-00AA00B58E10}"
    
    ' Event types for RegisterCustomEvent
    Public Const DEBUG_EVENT_TYPE_OUTBOUND = 0
    Public Const DEBUG_EVENT_TYPE_INBOUND = 1
    Public Const DEBUG_EVENT_TYPE_GENERIC = 2
    Public Const DEBUG_EVENT_TYPE_DEFAULT = 2
    Public Const DEBUG_EVENT_TYPE_MEASURED = 3
    Public Const DEBUG_EVENT_TYPE_BEGIN = 4
    Public Const DEBUG_EVENT_TYPE_END = 5
    Public Const DEBUG_EVENT_TYPE_LAST = 5
    
    'OMIGA4 specific
    Public Const OMIGA4_EVENT_SOURCE = "{5C23F280-9463-11d4-8263-005004E8D1A7}"
    
    '... custom categories
    Public Const OMIGA4_EVENT_CATEGORY = "{0AD27AB9-951F-11d4-8264-005004E8D1A7}"
    Public Const OMIGA4_EVENT_CATEGORY_TRACING = "{72915661-9485-11d4-8263-005004E8D1A7}"
    Public Const OMIGA4_EVENT_CATEGORY_ERRORS = "{E3159859-9526-11d4-8264-005004E8D1A7}"
    '... TRACE custom event type
    Public Const OMIGA4_EVENT_TRACING_TYPE = DEBUG_EVENT_TYPE_DEFAULT
    Public Const OMIGA4_EVENT_ERRORS_TYPE = DEBUG_EVENT_TYPE_DEFAULT
    '... TRACE events
    Public Const OMIGA4_EVENT_TRACING_GENERAL = "{72915662-9485-11d4-8263-005004E8D1A7}"
    Public Const OMIGA4_EVENT_ERRORS_SYSERR = "{1898A9B9-9549-11d4-8264-005004E8D1A7}"
    Public Const OMIGA4_EVENT_ERRORS_APPERR = "{3DF979D9-9549-11d4-8264-005004E8D1A7}"


#End If



