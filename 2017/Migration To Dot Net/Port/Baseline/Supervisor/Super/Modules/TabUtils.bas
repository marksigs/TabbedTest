Attribute VB_Name = "TabUtils"
Option Explicit
' This is here to workaround a problem with tabs that have ActiveX controls on them. If you tab
' to the last control on the current tab, the focus disappears onto the next tab. The workaround
' is to set the tabstop property of every control on every other tab to false. There may be controls
' that have one or more frames as their parents (container) so we need to start at the control and
' work our way back up until we find an sstab. If we do, then we check the Left property of the
' first frame (or control if it's place directly onto the sstab) and if it's more than 0, we're on
' the current tab and we want to set tabstop = true. Otherwise, set it to false.
Public Sub SetTabstops(frmToSet As Form)
    Dim c As Control
    Dim nLeft As Long
    Dim objParent As Object
    Dim bDone As Boolean
    
    On Error Resume Next
    For Each c In frmToSet.Controls
        If c.Name = "cmdInterestRateAdd" Then

        End If
        ' Is the control placed directly onto the tab?
        If TypeOf c.Container Is SSTab Then
            'Not all controls have the TabStop property

            c.TabStop = c.Left > 0
        ElseIf TypeOf c.Container Is Frame Then
            ' The control is held inside a frame (one or more)
            Set objParent = c.Container
            Dim bSetTab As Boolean
            bDone = False
            
            ' Loop finding the frame directly beneath the sstab.
            While bDone = False And Not objParent Is Nothing
                If TypeOf objParent Is Frame Then
                    If TypeOf objParent.Container Is SSTab Then
                        bDone = True
                        If TypeOf c Is OptionButton Then
                                Dim a As Boolean
                                a = True
                                Dim b As String
                                Dim d As Integer
                                b = c.Name
                                d = c.Index
'                            MsgBox "option button: " & c.Name
                        End If
                        Err.Clear
                        nLeft = objParent.Left
                        
                        If Err.Number = 0 Then
                            Err.Clear
                            bSetTab = nLeft > 0
                            
                            If bSetTab = True Then
                                ' Need to set tabstop to true, if it was originally
                                If Len(c.Tag) > 0 Then
                                    If CBool(c.Tag) = True Then
                                        bSetTab = True
                                    Else
                                        bSetTab = False
                                    End If
                                ElseIf c.TabStop = True Then
                                    bSetTab = True
                                Else
                                    bSetTab = False
                                End If
                            Else
                                bSetTab = False
                            End If
                            
                            If c.TabStop = True Then
                                c.Tag = True
                            End If
                                                        
                            c.TabStop = bSetTab
                        End If
                    Else
                        Set objParent = objParent.Container
                    End If
                Else
                    Set objParent = objParent.Container
                End If
            Wend
        ElseIf TypeOf objParent Is Form Then
            If c.Left > 0 And c.TabStop = True Then
                c.Tag = c.TabStop
            ElseIf c.Left > 0 And CBool(c.Tag) = True Then
                c.TabStop = True
            End If
        End If
    Next
End Sub


