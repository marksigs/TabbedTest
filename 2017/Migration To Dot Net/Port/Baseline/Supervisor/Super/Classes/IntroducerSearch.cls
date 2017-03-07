VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntroducerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class Module  : IntermediarySearch
' Description   : Stores all search criteria and performs the search against the
'                 intermediary table
'
' Change history
' Prog      Date        Description
' TW 	    17/10/2006  EP2_15 - Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Stores the search criteria specified in the gui
Private Type SearchDetails
    sForename As String
    sSurname As String
    sCompany As String
    sTown As String
    nType As Integer
End Type

Private m_SearchDetails As SearchDetails

Public Sub SetForeName(ByVal sName As String)
    m_SearchDetails.sForename = sName
End Sub

Public Sub SetSurName(ByVal sName As String)
    m_SearchDetails.sSurname = sName
End Sub

Public Sub SetTown(ByVal sTown As String)
    m_SearchDetails.sTown = sTown
End Sub

Public Sub SetCompanyName(ByVal sName As String)
    m_SearchDetails.sCompany = sName
End Sub
Public Sub SetType(ByVal nType As Integer)
    m_SearchDetails.nType = nType
End Sub

Public Function GetForeName() As String
    GetForeName = m_SearchDetails.sForename
End Function
Public Function GetSurName() As String
    GetSurName = m_SearchDetails.sSurname
End Function
Public Function GetTown() As String
    GetTown = m_SearchDetails.sTown
End Function
Public Function GetCompanyName() As String
    GetCompanyName = m_SearchDetails.sCompany
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : FindIntroducers
' Description   : Using the criteria set in SearchDetails, performs the search against the table
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FindIntroducers(clsIntroducerTable As IntroducerTable)
    On Error GoTo Failed

    Dim sSQL As String
    Dim sCompanyNameSearch As String
    Dim sForenameSearch As String
    Dim sSurnameSearch As String
    Dim sTownSearch As String
    Dim sIntTypeSearch As String
    Dim rs As ADODB.Recordset
    
    ' Company
    If Len(m_SearchDetails.sCompany) > 0 Then
        sCompanyNameSearch = " Name LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sCompany)
    End If
    
    'Forename
    If Len(m_SearchDetails.sForename) > 0 Then
        If Len(sCompanyNameSearch) > 0 Then
            sForenameSearch = " and Forename LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sForename)
        Else
            sForenameSearch = " Forename LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sForename)
        End If
    End If
    
    'Surname
    If Len(m_SearchDetails.sSurname) > 0 Then
        If Len(sCompanyNameSearch) > 0 Or Len(sForenameSearch) > 0 Then
            sSurnameSearch = " and Surname LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sSurname)
        Else
            sSurnameSearch = " Surname LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sSurname)
        End If
    End If
    
    'Town
    If Len(m_SearchDetails.sTown) > 0 Then
        If Len(sCompanyNameSearch) > 0 Or Len(sForenameSearch) > 0 Or Len(sSurnameSearch) > 0 Then
            sTownSearch = " and AD.Town LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sTown)
        Else
            sTownSearch = " AD.Town LIKE " & g_clsFormProcessing.FormatWildcardString(m_SearchDetails.sTown)
        End If
    End If
           
    'Intermediary Type
    If m_SearchDetails.nType > 0 Then
        If Len(sCompanyNameSearch) > 0 Or Len(sForenameSearch) > 0 Or Len(sSurnameSearch) > 0 Or Len(sTownSearch) > 0 Then
            sIntTypeSearch = " and Int.IntermediaryType = " & m_SearchDetails.nType
        Else
            sIntTypeSearch = " Int.IntermediaryType = " & m_SearchDetails.nType
        End If
    End If
        
    'SB: Format SQL appropriate to specific DB server version.
    If g_clsDataAccess.GetDatabaseType = INDEX_ORACLE Then
        sSQL = "Select " & _
                    "intermediaryorganisation.name NAME, " & _
                    "intermediaryindividual.forename ForeName, " & _
                    "intermediaryindividual.surname Surname, " & _
                    "Int.IntermediaryGuid, " & _
                    "int.intermediarypanelid, " & _
                    "AD.TOWN " & _
              "From " & _
                    "Address AD, " & _
                    "INTERMEDIARY INT, " & _
                    "Intermediaryorganisation,  " & _
                    "intermediaryindividual " & _
            "Where " & _
                    "AD.ADDRESSGUID(+) = Int.IntermediaryGuid and " & _
                    "intermediaryorganisation.INTERMEDIARYGUID(+) = Int.INTERMEDIARYGUID AND " & _
                    "intermediaryindividual.INTERMEDIARYGUID(+) = Int.INTERMEDIARYGUID " & _
                    " and " & sCompanyNameSearch & _
                    sForenameSearch & _
                    sSurnameSearch & _
                    sTownSearch & _
                    sIntTypeSearch
    Else
        sSQL = "Select " & "intermediaryorganisation.name NAME, " & _
                "intermediaryindividual.forename ForeName, " & _
                "intermediaryindividual.surname Surname, " & _
                "Int.IntermediaryGuid, " & _
                "int.intermediarypanelid, " & _
                "AD.TOWN " & _
          "From " & "INTERMEDIARY INT LEFT OUTER JOIN " & _
                "Address AD ON AD.ADDRESSGUID = Int.IntermediaryGuid LEFT OUTER JOIN " & _
                "Intermediaryorganisation ON IntermediaryOrganisation.INTERMEDIARYGUID = Int.INTERMEDIARYGUID LEFT OUTER JOIN " & _
                "IntermediaryIndividual ON IntermediaryIndividual.INTERMEDIARYGUID = Int.INTERMEDIARYGUID " & _
        "Where " & sCompanyNameSearch & _
                sForenameSearch & _
                sSurnameSearch & _
                sTownSearch & _
                sIntTypeSearch
    End If
                
    Set rs = g_clsDataAccess.GetTableData(, sSQL)
    TableAccess(clsIntroducerTable).SetRecordSet rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
