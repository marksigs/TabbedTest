Attribute VB_Name = "Constants_Column"
'Workfile:      Constants_Column.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'LD     10/01/01    Created
'------------------------------------------------------------------------------------------

Option Explicit

' StaticNode (lvStaticNode => (COLUMNKEY_ENTERPRISESERVERNAME = Name = 1)
'   EnterpriseServer (lvEnterpriseServer => (COLUMNKEY_COMPUTERNAME = Computer Pathname = 1), (COLUMNKEY_COMPUTERSTATUS = Computer Status = 2))
'       Computer (lvComputer => (COLUMNKEY_GROUPCOMPUTER = Computer Folder Names = 14)
'           QueuesFolder (lvQueuesFolder => (COLUMNKEY_QUEUENAME = Queue Name = 15), (COLUMNKEY_QUEUETYPE = Queue Type = 16), (COLUMNKEY_QUEUESTATUS = Queue Staus = 17))
'               Queue (lvQueue => (COLUMNKEY_GROUPQUEUE = Group Name = 6))
'                   StalledComponents (lvStalledComponents => (COLUMNKEY_STALLEDCOMPONENTSPROGID = COM/COM+ ProgID = 7))
'                   QueueEvents (lvQueueEvents => (COLUMNKEY_QUEUEEVENTSDAYTIME = Day / Time = 8), (COLUMNKEY_QUEUEEVENTSNAME = Event Name = 9))
'           NTEventLog (lvNTEventLog => (COLUMNKEY_NTLOGEVENTSTYPE = Type = 11), (COLUMNKEY_NTLOGEVENTSDATETIME = Day/Time = 12), (COLUMNKEY_NTLOGEVENTSMESSAGE / Message = 13))
'
Public Const COLUMNDISPLAY_ENTERPRISESERVERNAME = "Name"
Public Const COLUMNDISPLAY_COMPUTERNAME = "Computer PathName"
Public Const COLUMNDISPLAY_COMPUTERSTATUS = "Status"
Public Const COLUMNDISPLAY_GROUPQUEUE = "Queue Group"
Public Const COLUMNDISPLAY_STALLEDCOMPONENTSPROGID = "COM/COM+ ProgID"
Public Const COLUMNDISPLAY_QUEUEEVENTSDAYTIME = "Day / Time"
Public Const COLUMNDISPLAY_QUEUEEVENTSNAME = "Event Name"
Public Const COLUMNDISPLAY_NTLOGEVENTSGROUP = "NT Log Event Group"
Public Const COLUMNDISPLAY_NTLOGEVENTSTYPE = "Type"
Public Const COLUMNDISPLAY_NTLOGEVENTSDATETIME = "Date / Time"
Public Const COLUMNDISPLAY_NTLOGEVENTSMESSAGE = "Message"
Public Const COLUMNDISPLAY_GROUPCOMPUTER = "Computer Group"
Public Const COLUMNDISPLAY_QUEUENAME = "Queue Name"
Public Const COLUMNDISPLAY_QUEUETYPE = "Queue Type"
Public Const COLUMNDISPLAY_QUEUESTATUS = "Queue Status"


' must match keys defined in the designer control
Public Const COLUMNKEY_ENTERPRISESERVERNAME = "1"
Public Const COLUMNKEY_COMPUTERNAME = "1"
Public Const COLUMNKEY_COMPUTERSTATUS = "2"
Public Const COLUMNKEY_GROUPQUEUE = "6" ' group for children (QueueEvents, StalledComponents)
Public Const COLUMNKEY_STALLEDCOMPONENTSPROGID = "7"
Public Const COLUMNKEY_QUEUEEVENTSDAYTIME = "8"
Public Const COLUMNKEY_QUEUEEVENTSNAME = "9"
Public Const COLUMNKEY_NTLOGEVENTSTYPE = "11"
Public Const COLUMNKEY_NTLOGEVENTSDATETIME = "12"
Public Const COLUMNKEY_NTLOGEVENTSMESSAGE = "13"
Public Const COLUMNKEY_GROUPCOMPUTER = "14" ' group for children (Queues, NTLogEvent)
Public Const COLUMNKEY_QUEUENAME = "15"
Public Const COLUMNKEY_QUEUETYPE = "16"
Public Const COLUMNKEY_QUEUESTATUS = "17"


