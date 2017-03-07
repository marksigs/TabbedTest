VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConditionsBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'BMIDS History:
'
'Prog   Date        Description
'GD     09/09/02    BMIDS00313 - APWP2 - BM004 Added FindNonProductConditionsList,FindAppProductConditionsList
'GD     17/10/02    BMIDS00650 - Added GetApplicationData
'------------------------------------------------------------------------------------------
Public Sub DeleteApplicationConditions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindApplicationConditionsList( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindConditionsList( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub UpdateApplicationConditions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub DeleteAllApplicationConditions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateApplicationConditions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub CreateDefaultApplicationConditions( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
End Sub

Public Sub FindNonProductConditionsList( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub FindAppProductConditionsList( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

End Sub

Public Sub GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, _
                               ByVal vxmlParentNode As IXMLDOMNode)
                               
End Sub

