VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LogAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'IK     28/08/99 Created
'MCS    29/09/99 Extend functionality
'RF     08/11/99 Added omAU, omOrg; removed omSS
'MS     23/06/00 Added omAQ,omBC,omCust,omAIP,omCE,omTP removed omCR
'------------------------------------------------------------------------------------------
Option Explicit

Public Enum OBJECT_TYPE
        otDO
        otBO
        otBoth
End Enum

Public Enum COMPONENT_NAME
        cnomAU
        cnomBase
        cnomCF
        cnomCM
'        cnomCR
        cnomLC
        cnomMP
        cnomOrg
        cnomPP
        cnomQQ
        cnomApp
        cnomAQ
        cnomBC
        cnomCust
        cnomAIP
        cnomCE
        cnomTP
End Enum

Private mtmStart As Single, mtmStop As Single
Private mstrSource As String

Public Sub StartTimerEx(ByVal vstrSource As String, _
    Optional ByVal blnLogStart As Boolean = False, _
    Optional ByVal votObjectType As OBJECT_TYPE = otDO)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    If ProfilingOn Then
        
        Dim blnOn As Boolean
        blnOn = False

        Select Case App.Title
            Case "omAU"
                blnOn = IsComponentProfilingOn(cnomAU, votObjectType)
            Case "omBase"
                blnOn = IsComponentProfilingOn(cnomBase, votObjectType)
            Case "omCF"
                blnOn = IsComponentProfilingOn(cnomCF, votObjectType)
            Case "omCM"
                blnOn = IsComponentProfilingOn(cnomCM, votObjectType)
'            Case "omCR"
'                blnOn = IsComponentProfilingOn(cnomCR, votObjectType)
            Case "omMP"
                blnOn = IsComponentProfilingOn(cnomMP, votObjectType)
            Case "omQQ"
                blnOn = IsComponentProfilingOn(cnomQQ, votObjectType)
            Case "omPP"
                blnOn = IsComponentProfilingOn(cnomPP, votObjectType)
            Case "omLC"
                blnOn = IsComponentProfilingOn(cnomLC, votObjectType)
            Case "omOrg"
                blnOn = IsComponentProfilingOn(cnomOrg, votObjectType)
            Case "omApp"
                blnOn = IsComponentProfilingOn(cnomApp, votObjectType)
            Case "omAQ"
                blnOn = IsComponentProfilingOn(cnomAQ, votObjectType)
            Case "omBC"
                blnOn = IsComponentProfilingOn(cnomBC, votObjectType)
            Case "omCust"
                blnOn = IsComponentProfilingOn(cnomCust, votObjectType)
            Case "omAIP"
                blnOn = IsComponentProfilingOn(cnomAIP, votObjectType)
            Case "omCE"
                blnOn = IsComponentProfilingOn(cnomCE, votObjectType)
            Case "omTP"
                blnOn = IsComponentProfilingOn(cnomTP, votObjectType)
            Case Else
                ' fixme
                'raise warning profiling not set for this component
        End Select
            
        If blnOn Then
            
            mstrSource = vstrSource
            mtmStart = Timer
        
            If blnLogStart Then
                App.LogEvent mstrSource & " [started]", vbLogEventTypeInformation
            End If
        
        End If
    End If
End Sub

Public Sub StopTimerEx(Optional ByVal votObjectType As OBJECT_TYPE = otDO)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
'   votObjectType
' return:
'-----------------------------------------------------------------------------------------------------

    If ProfilingOn Then
        
        Dim blnOn As Boolean
        blnOn = False
        
        Select Case App.Title
            Case "omAU"
                blnOn = IsComponentProfilingOn(cnomAU, votObjectType)
            Case "omBase"
                blnOn = IsComponentProfilingOn(cnomBase, votObjectType)
            Case "omCM"
                blnOn = IsComponentProfilingOn(cnomCM, votObjectType)
'            Case "omCR"
'                blnOn = IsComponentProfilingOn(cnomCR, votObjectType)
            Case "omMP"
                blnOn = IsComponentProfilingOn(cnomMP, votObjectType)
            Case "omQQ"
                blnOn = IsComponentProfilingOn(cnomQQ, votObjectType)
            Case "omPP"
                blnOn = IsComponentProfilingOn(cnomPP, votObjectType)
            Case "omLC"
                blnOn = IsComponentProfilingOn(cnomLC, votObjectType)
            Case "omCF"
                blnOn = IsComponentProfilingOn(cnomCF, votObjectType)
            Case "omOrg"
                blnOn = IsComponentProfilingOn(cnomOrg, votObjectType)
            Case "omApp"
                blnOn = IsComponentProfilingOn(cnomApp, votObjectType)
            Case "omAQ"
                blnOn = IsComponentProfilingOn(cnomAQ, votObjectType)
            Case "omBC"
                blnOn = IsComponentProfilingOn(cnomBC, votObjectType)
            Case "omCust"
                blnOn = IsComponentProfilingOn(cnomCust, votObjectType)
            Case "omAIP"
                blnOn = IsComponentProfilingOn(cnomAIP, votObjectType)
            Case "omCE"
                blnOn = IsComponentProfilingOn(cnomCE, votObjectType)
            Case "omTP"
                blnOn = IsComponentProfilingOn(cnomTP, votObjectType)

           Case Else
                ' fixme
                'raise warning profiling not set for this component
        End Select
        
        If blnOn Then
            mtmStop = Timer
            App.LogEvent mstrSource & " [elapsed] " & Format(mtmStop - mtmStart, "0.00000000"), vbLogEventTypeInformation
        End If
        
    End If
End Sub

Public Sub LogEventEx(vstrSource As String, vstrMessage As String)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    App.LogEvent "[" & vstrSource & "] " & vstrMessage, vbLogEventTypeInformation
End Sub


Public Sub StartTimer(vstrSource As String, Optional vblnLogStart As Boolean = False)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    If ProfilingOn Then
        mstrSource = vstrSource
        mtmStart = Timer
    
        If vblnLogStart Then
            App.LogEvent mstrSource & " [started]", vbLogEventTypeInformation
        End If
    End If
End Sub

Public Sub StopTimer()
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    If ProfilingOn Then
    
        mtmStop = Timer
        App.LogEvent mstrSource & " [elapsed] " & Format(mtmStop - mtmStart, "0.00000000"), vbLogEventTypeInformation
    
    End If
End Sub

Public Sub LogEvent(strSource As String, strMessage As String)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    App.LogEvent "[" & strSource & "] " & strMessage, vbLogEventTypeInformation
End Sub


Private Function IsComponentProfilingOn(ByVal vcnComponentName As COMPONENT_NAME, _
    ByVal votObjectType As OBJECT_TYPE) As Boolean
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim spmProperty2 As SharedProperty
    Dim blnGroupExists As Boolean
    Dim blnPropertyExists As Boolean
    Dim blnBoth As Boolean
    IsComponentProfilingOn = False
    blnBoth = False
    Set spmMgr = CreateObject _
        ("MTxSpm.SharedPropertyGroupManager.1")
            
    Select Case vcnComponentName
        Case cnomAU
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AULogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomBase
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("BaseLogOptions", LockSetGet, Process, blnGroupExists)
'         Case cnomCR
'            Set spmGroup = spmMgr.CreatePropertyGroup _
'                ("CRLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomQQ
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("QQLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomMP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("MPLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomCM
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CMLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomPP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("PPLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomLC
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("LCLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomCF
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CFLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomOrg
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("OrgLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomApp
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AppLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomAQ
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AQLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomBC
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("BCLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomCust
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CustLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomAIP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AIPLogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomCE
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CELogOptions", LockSetGet, Process, blnGroupExists)
        Case cnomTP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("TPLogOptions", LockSetGet, Process, blnGroupExists)
        Case Else
            ' fixme
            'raise warning error ?
    End Select
    
    Select Case votObjectType
        Case otDO
            Set spmProperty = _
                spmGroup.CreateProperty("DOProfiling", blnPropertyExists)
        Case otBO
            Set spmProperty = _
                spmGroup.CreateProperty("BOProfiling", blnPropertyExists)
        Case otBoth
            blnBoth = True
            Set spmProperty = _
                spmGroup.CreateProperty("DOProfiling", blnPropertyExists)
            Set spmProperty2 = _
                spmGroup.CreateProperty("BOProfiling", blnPropertyExists)
        Case Else
            ' fixme
            'raise warning error ?
    End Select
    
    If blnGroupExists And blnPropertyExists Then
        If spmProperty.Value = "ON" Then
            If blnBoth Then
                If spmProperty2.Value = "ON" Then
                    IsComponentProfilingOn = True
                End If
             Else
                IsComponentProfilingOn = True
             End If
        End If
    Else
        spmProperty.Value = "OFF"
        If blnBoth Then
            spmProperty2.Value = "OFF"
        End If
    End If
    
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing

End Function

Private Sub SetComponentProfilingFlag(ByVal vblnFlag As Boolean, _
    ByVal vcnComponentName As COMPONENT_NAME, _
    Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim spmProperty2 As SharedProperty
    Dim blnExists As Boolean
    Dim blnBoth As Boolean
    blnBoth = False
    
    Set spmMgr = CreateObject _
        ("MTxSpm.SharedPropertyGroupManager.1")
                  
     Select Case vcnComponentName
        Case cnomAU
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AULogOptions", LockSetGet, Process, blnExists)
        Case cnomBase
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("BaseLogOptions", LockSetGet, Process, blnExists)
'        Case cnomCR
'            Set spmGroup = spmMgr.CreatePropertyGroup _
'                ("CRLogOptions", LockSetGet, Process, blnExists)
        Case cnomMP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("MPLogOptions", LockSetGet, Process, blnExists)
        Case cnomQQ
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("QQLogOptions", LockSetGet, Process, blnExists)
        Case cnomCM
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CMLogOptions", LockSetGet, Process, blnExists)
        Case cnomPP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("PPLogOptions", LockSetGet, Process, blnExists)
        Case cnomLC
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("LCLogOptions", LockSetGet, Process, blnExists)
        Case cnomCF
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CFLogOptions", LockSetGet, Process, blnExists)
        Case cnomOrg
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("OrgLogOptions", LockSetGet, Process, blnExists)
        Case cnomApp
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AppLogOptions", LockSetGet, Process, blnExists)
        Case cnomAQ
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AQLogOptions", LockSetGet, Process, blnExists)
        Case cnomBC
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("BCLogOptions", LockSetGet, Process, blnExists)
        Case cnomCust
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CustLogOptions", LockSetGet, Process, blnExists)
        Case cnomAIP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("AIPLogOptions", LockSetGet, Process, blnExists)
        Case cnomCE
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("CELogOptions", LockSetGet, Process, blnExists)
        Case cnomTP
            Set spmGroup = spmMgr.CreatePropertyGroup _
                ("TPLogOptions", LockSetGet, Process, blnExists)
        Case Else
            ' fixme
            'raise warning ?
    End Select
    
    Select Case votObjectType
        Case otDO
            Set spmProperty = _
                spmGroup.CreateProperty("DOProfiling", blnExists)
        Case otBO
            Set spmProperty = _
                spmGroup.CreateProperty("BOProfiling", blnExists)
        Case otBoth
            blnBoth = True
            Set spmProperty = _
                spmGroup.CreateProperty("DOProfiling", blnExists)
            Set spmProperty2 = _
                spmGroup.CreateProperty("BOProfiling", blnExists)
        Case Else
            ' fixme
            'raise warning ?
    End Select
    
    If vblnFlag = True Then
        spmProperty.Value = "ON"
        If blnBoth Then
            spmProperty2.Value = "ON"
        End If
    Else
        spmProperty.Value = "OFF"
        If blnBoth Then
            spmProperty2.Value = "OFF"
        End If
    End If
    
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing

End Sub


Private Function ProfilingOn() As Boolean
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnExists As Boolean
    ProfilingOn = False
        
    Set spmMgr = CreateObject _
        ("MTxSpm.SharedPropertyGroupManager.1")
        
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("LogOptions", LockSetGet, Process, blnExists)
        
    Set spmProperty = _
        spmGroup.CreateProperty("Profiling", blnExists)
    
    If blnExists Then
        If spmProperty.Value = "ON" Then
            ProfilingOn = True
        End If
    Else
        spmProperty.Value = "OFF"
    End If
    
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing

End Function

Private Sub SetProfilingFlag(vblnFlag As Boolean)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    Dim spmMgr As SharedPropertyGroupManager
    Dim spmGroup As SharedPropertyGroup
    Dim spmProperty As SharedProperty
    Dim blnExists As Boolean
        
    Set spmMgr = CreateObject _
        ("MTxSpm.SharedPropertyGroupManager.1")
        
    Set spmGroup = spmMgr.CreatePropertyGroup _
        ("LogOptions", LockSetGet, Process, blnExists)
        
    Set spmProperty = _
        spmGroup.CreateProperty("Profiling", blnExists)
    
    If vblnFlag = True Then
        spmProperty.Value = "ON"
    Else
        spmProperty.Value = "OFF"
    End If
    
    Set spmProperty = Nothing
    Set spmGroup = Nothing
    Set spmMgr = Nothing

End Sub

Public Sub SetProfilingOn()
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetProfilingFlag True
End Sub

Public Sub SetProfilingOff()
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetProfilingFlag False
End Sub


'Public Sub SetCRProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
'    SetComponentProfilingFlag True, cnomCR, votObjectType
'End Sub

'Public Sub SetCRProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
'    SetComponentProfilingFlag False, cnomCR, votObjectType
'End Sub

Public Sub SetQQProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomQQ, votObjectType
End Sub

Public Sub SetQQProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomQQ, votObjectType
End Sub
Public Sub SetMPProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomMP, votObjectType
End Sub

Public Sub SetMPProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomMP, votObjectType
End Sub
Public Sub SetCMProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomCM, votObjectType
End Sub

Public Sub SetCMProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomCM, votObjectType
End Sub
Public Sub SetPPProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomPP, votObjectType
End Sub

Public Sub SetPPProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomPP, votObjectType
End Sub

Public Sub SetLCProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomLC, votObjectType
End Sub

Public Sub SetLCProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomLC, votObjectType
End Sub

Public Sub SetCFProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomCF, votObjectType
End Sub

Public Sub SetCFProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomCF, votObjectType
End Sub
Public Sub SetBaseProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomBase, votObjectType
End Sub

Public Sub SetBaseProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomBase, votObjectType
End Sub

Public Sub SetAppProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomApp, votObjectType
End Sub

Public Sub SetAppProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomApp, votObjectType
End Sub
Public Sub SetAQProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomAQ, votObjectType
End Sub

Public Sub SetAQProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomAQ, votObjectType
End Sub
Public Sub SetBCProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomBC, votObjectType
End Sub

Public Sub SetBCProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomBC, votObjectType
End Sub

Public Sub SetCustProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomCust, votObjectType
End Sub

Public Sub SetCustProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomCust, votObjectType
End Sub

Public Sub SetAIPProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomAIP, votObjectType
End Sub

Public Sub SetAIPProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomAIP, votObjectType
End Sub

Public Sub SetCEProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomCE, votObjectType
End Sub

Public Sub SetCEProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomCE, votObjectType
End Sub

Public Sub SetTPProfilingOn(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag True, cnomTP, votObjectType
End Sub

Public Sub SetTPProfilingOff(Optional ByVal votObjectType As OBJECT_TYPE = otBoth)
' header ---------------------------------------------------------------------------------------------
' description:
' pass:
' return:
'-----------------------------------------------------------------------------------------------------
    SetComponentProfilingFlag False, cnomTP, votObjectType
End Sub

