VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityResourceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module: SecurityResourceList
' Description : A collection-class which can contain SecurityResource objects.
'
' Change history
' Prog      Date        Description
' STB       15/01/02    Created class.
' STB       22/01/02    SYS2957 Added the ObjectID property to SecurityResource
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit


'The aggregated collection of SecurityResource objects.
Private m_colSecurityResourceList As Collection


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : NewEnum
' Description   : Iterator routine for this class.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colSecurityResourceList.[_NewEnum]
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Count
' Description   : Returns the count of the collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Count() As Long
Attribute Count.VB_Description = "The number of SecurityResource objects in the list."
    Count = m_colSecurityResourceList.Count
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Item
' Description   : Returns the indexed SecurityResource item.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get Item(ByVal Index As Long) As SecurityResource
Attribute Item.VB_Description = "Returns a specific SecurityResource object by numerical index."
Attribute Item.VB_UserMemId = 0
    Set Item = m_colSecurityResourceList.Item(Index)
End Property


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Add
' Description   : Add the SecurityResource specified to the end of the
'                 collection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Add(ByRef clsSecurityResource As SecurityResource)
    m_colSecurityResourceList.Add clsSecurityResource
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Remove
' Description   : Remove the indexed item and return it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function Remove(ByVal lIndex As Long) As SecurityResource
    
    Dim clsSecurityResource As SecurityResource
    
    'Get a reference to the item being removed.
    Set clsSecurityResource = m_colSecurityResourceList.Item(lIndex)
    
    'Remove the item from the underlying collection.
    m_colSecurityResourceList.Remove lIndex
    
    'Return the item to the caller.
    Set Remove = clsSecurityResource
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Initialize
' Description   : Instantiate any object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    Set m_colSecurityResourceList = New Collection
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Class_Terminate
' Description   : Release any object references.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    Set m_colSecurityResourceList = Nothing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : LoadFromRecordset
' Description   : Populates itself from the recordset specified and the master
'                 list of securable items.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub LoadFromRecordset(ByRef rsAgentResources As ADODB.Recordset, ByRef clsMasterResourceList As SecurityResourceList)

    Dim clsResource As SecurityResource
    Dim clsMasterRsc As SecurityResource

    On Error GoTo Err_Handler

    'Iterate through each record and add a SecurityResource.
    If Not (rsAgentResources.BOF And rsAgentResources.EOF) Then
        Do While Not rsAgentResources.EOF
            'Create a new resource object.
            Set clsResource = New SecurityResource

            'Set the object and operation properties.
            clsResource.ObjectID = rsAgentResources.Fields.Item("ITEMID").Value

            'Get the objectname from the master-list list.
            Set clsMasterRsc = clsMasterResourceList.ItemByID(clsResource.ObjectID)
            
            'If the master resource record doesn't exist in the virtual table, then raise an error.
            If clsMasterRsc Is Nothing Then
                Err.Raise 445, "LoadFromRecordset", "A security resource record was found with no master resource in the virtual table."
            Else
                clsResource.ObjectName = clsMasterRsc.ObjectName
            End If

            'Add the resource to the internal collection.
            m_colSecurityResourceList.Add clsResource

            'Move onto the next permission.
            rsAgentResources.MoveNext
        Loop
    End If
    
Err_Handler:
    If Err Then
        Err.Raise Err.Number, "LoadFromRecordset", "Unable to populate a SecurityResourceList from the specified ADO Recordset. The reported error was: " & Err.Description
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ItemByID
' Description   : Return the first SecurityResource item which matches the
'                 ObjectID specified. This routine is only used internally
' within this component to perform lookups on the virtual table of Security-
' resources.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Function ItemByID(ByVal lObjectID As Long) As SecurityResource

    Dim clsResource As SecurityResource
    
    'Iterate through each item until the first matching one is found.
    For Each clsResource In m_colSecurityResourceList
        If clsResource.ObjectID = lObjectID Then
            Exit For
        End If
    Next clsResource
    
    'Return the resource to the caller (or nothing).
    Set ItemByID = clsResource

End Function
