Attribute VB_Name = "modADOHelper"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module: modADOHelper
' Description : ADO helper routines.
'
' Change history
' Prog      Date        Description
' STB       21/01/02    Created.
' STB       22/01/02    SYS2957 Supervisor connection passed as reference.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS
' RF        18/01/2006  MAR1132     Error when adding user in Supervisor -
'                                   Max Userid length now 64
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Global variable to hold stored proc parameter prefix ('p' or '@').
Public g_sParamPrefix As String

'The stored procedures used by this component.
Public Const USP_GET_AGENT_RESOURCE    As String = "USP_GETUSERITEM"
Public Const USP_SET_AGENT_RESOURCE    As String = "USP_SETUSERITEM"
Public Const USP_DELETE_AGENT_RESOURCE As String = "USP_DELETEUSERITEM"
Private Const cintLEN_USERID_COLUMN As Integer = 64 ' MAR1000
        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ADOExecuteSQL
' Description   : Simple routine to execute the SQL given and returns a
'                 recordset of results.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ADOExecuteSQL(ByVal sSQL As String, ByRef clsConnection As ADODB.Connection) As ADODB.Recordset

    Dim clsCommand As ADODB.Command
    
    'Create a command object to return a list of permissions.
    Set clsCommand = New ADODB.Command
    clsCommand.CommandType = adCmdText
    clsCommand.CommandText = sSQL
    
    'Associate the command with the connection.
    Set clsCommand.ActiveConnection = clsConnection
    
    'Return the resultant recordset.
    Set ADOExecuteSQL = clsCommand.Execute
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ADOExecuteCmd
' Description   : Simple routine to execute the SQL given and returns a
'                 recordset of results.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ADOExecuteCmd(ByRef clsCommand As ADODB.Command, ByRef clsConnection As ADODB.Connection) As ADODB.Recordset
       
    'Associate the command with the connection.
    Set clsCommand.ActiveConnection = clsConnection
    
    'Return the resultant recordset.
    Set ADOExecuteCmd = clsCommand.Execute
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ADOFormatSQLString
' Description   : Ensure no ' characters corrupt the string.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ADOFormatSQLString(ByVal sSQLString As String) As String
    
    Dim iPos As Integer
    Dim sOutput As String
    Dim iLength As Integer
    Dim sChar As String
    Dim i As Integer
    
    Const APOSTROPHE As String = "'"
    
    iLength = Len(sSQLString)
    
    For iPos = 1 To iLength Step 1
        sChar = Mid$(sSQLString, iPos, 1)
        sOutput = sOutput + sChar
        
        If sChar = APOSTROPHE Then
            ' add an extra apostrophe
            sOutput = sOutput + APOSTROPHE
        End If
    Next iPos
    
    ADOFormatSQLString = sOutput
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ADOCreateParameter
' Description   : Create and return a parameter to be appended to a command
'                 object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ADOCreateParameter(ByVal sName As String, ByVal uType As ADODB.DataTypeEnum, Optional ByVal lSize As Long, Optional ByVal lPrescision As Long, Optional uDirection As ADODB.ParameterDirectionEnum = adParamInput) As ADODB.Parameter

    Dim clsParameter As ADODB.Parameter
    
    'Create the parameter.
    Set clsParameter = New ADODB.Parameter
    
    'Setup the parameter.
    clsParameter.Name = sName
    clsParameter.Type = uType
    clsParameter.Direction = uDirection
    
    'If the type requires it, size and precsion can be specified here.
    Select Case uType
        Case adVarChar, adChar
            clsParameter.Size = lSize
    End Select
    
    'Return the parameter object to the caller.
    Set ADOCreateParameter = clsParameter

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ADOGetCommand
' Description   : Returns an instantiated command object for the stored proc
'                 name specified. The parameters collection will also be
'                 constructed (if relevant).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ADOGetCommand(ByVal sCommandName As String) As ADODB.Command

    Dim clsCommand As ADODB.Command

    'Create the command object.
    Set clsCommand = New ADODB.Command
    clsCommand.CommandType = adCmdStoredProc
    clsCommand.CommandText = sCommandName

    Select Case sCommandName
        Case USP_GET_AGENT_RESOURCE
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "USERID", adVarChar, cintLEN_USERID_COLUMN) ' MAR1000
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "ITEMID", adInteger)
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "ALLOW", adSmallInt, , , adParamOutput)

        Case USP_SET_AGENT_RESOURCE
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "USERID", adVarChar, cintLEN_USERID_COLUMN) 'MAR1000
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "ITEMID", adInteger)
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "ALLOW", adSmallInt)

        Case USP_DELETE_AGENT_RESOURCE
            clsCommand.Parameters.Append _
                ADOCreateParameter(g_sParamPrefix & "USERID", adVarChar, cintLEN_USERID_COLUMN) ' MAR1000

        'Add additional commands/stored procedures here. Alternatively, use the
        'Parameters.Refresh() method although this may have a performance impact.

    End Select

    'Return the command to the caller.
    Set ADOGetCommand = clsCommand

End Function

Public Sub ADOSetParameterPrefix(ByRef clsConnection As ADODB.Connection)

    'Interagate the command string and ascertain if we're talking to SQL-Server
    'or Oracle. The stored-proc parameter prefix is dependant upon this.
    If InStr(1, clsConnection.ConnectionString, "MSDAORA", vbTextCompare) > 0 Then
        g_sParamPrefix = "p"
    Else
        g_sParamPrefix = "@"
    End If
    
End Sub
