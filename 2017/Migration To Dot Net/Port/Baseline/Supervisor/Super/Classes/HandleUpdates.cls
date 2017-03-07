VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : HandleUpdates
' Description   : User has changed an area of Supervisor and we need to record that fact in the
'                  Supervisor change log table
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port
' STB       02/01/02    SYS2551 - GetObjectTableClass() also handles admin centres.
' STB       07/01/02    SYS3581 - SaveObjectInstance() now stores the key value datatype to
'                                 allow GUIDs to be read-back correctly in frmManageUpdates.
' STB       04/02/02    SYS3857 - Base Rates can now be promoted.
' DJP       13/02/02    SYS4052 Save all change requests, whether user is logged in or not.
' DJP       20/02/02    SYS4121 Updated to allow client variants
' DJP       20/02/02    SYS4121 Updated to allow client variants - when promoting, allow additional
'                               items to be promoted.
' STB       13/05/02    SYS4417 Added AllowableIncomeFactors.
' CL        22/05/02    SYS4521 Resolve promotion functionality problems
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'
'   AW      14/05/02    BM088       Added INCOME_MULTIPLE processing
'   AW      16/05/02    BM087       Added MIG_RATE_SETS processing
'   AW      16/05/02    BM017       Added REDEM_SETS processing
'   AW      27/05/02    BM017       AMmendnents
'   GD      05/06/02    BMIDS00016  Promotions
'   AW      17/06/02    BMIDS00055  Added BASERATEBAND
'   GD      15/07/02    BMIDS00170  Changes for Income factors promotions
'   AW      29/07/02    BMIDS00237  Do not delete target when promoting an edit
'   CL      16/10/02    BMIDS00170  Modificationt to PromoteObjectTo handle
'                                   promotion of Income Factors
'   DB      06/01/03    SYS5457     Modify promotion of Intermediaries
'   DJP     05/03/03    BM0423      Allow saving of promotion request to a different use to the one logged on.
'   MC      02/06/2004  BMIDS763    BMIDS763 - CC075 - FEESETS New Class refs added
'   JD      15/06/04    BMIDS765    CC076 added Rental Income Rate Set
' [MC]      25/04/2004  BMIDS763    Data Promotion code added for three new sets
'                                   (Product switch fee, Insurance Fee set, TT FeeSet)
' JD        12/07/2004  BMIDS765    Promote rental income rate sets.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply (BBG1370) New screen for print configuration
' GHun      14/10/2005  MAR202 Promote Packs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom History
' Prog      Date        Description
' TW        09/10/2006  EP2_7 - Added handling for AdditionalBorrowingFee and CreditLimitIncreaseFee
' TW        17/10/2006  EP2_15 - Added MortgageClubNetworkAssociation processing
' TW        18/11/2006  EP2_132 ECR20/21 Further promotions for FSA related changes
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        30/11/2006  EP2_253 - Changes related to Mortgage Product Application Eligibility
' TW        11/12/2006  EP2_20 - Added handling for TransferOfEquityFee
' TW        14/12/2006  EP2_518 Added handling for Default Procuration Fees, Loan Amount Adjustments, LTV Amount Adjustments
' TW        15/01/2007  EP2_826 - Rationalisation of pop-up menus and actions to improve consistency and usability
' TW        05/02/2007  EP2_706 - Should  network be mandatory data for ar firms
' TW        23/03/2007  EP2_1942 - Promotion of Income Factors does not work
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SavePromotionRequest
' Description   : Records that something in Supervisor has changed. A form will call this method
'                 to record the current database connected to, the date the update was made, the
'                 ID of the object to save and the type of update (add/delete/edit)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SavePromotionRequest(sServiceName As String, _
                                sDatabaseUserID As String, _
                                dtCreationDate As Date, _
                                vObjectInstanceID As Variant, _
                                enumOperation As SupervisorPromoteType, _
                                Optional sTargetUser As String)
    On Error GoTo Failed
    Dim vConnection As Variant
    Dim clsSupervisorConnection As SupervisorConnectionTable
    
    Set clsSupervisorConnection = New SupervisorConnectionTable
    
    vConnection = SaveConnection(clsSupervisorConnection, sServiceName, sDatabaseUserID)

    ' Now the Update
    ' DJP BM0423 pass the target user to SaveSupervisorUpdate
    SaveSupervisorUpdate vObjectInstanceID, _
                         vConnection, _
                         DIRECTION_DEST, _
                         dtCreationDate, _
                         enumOperation, _
                         sTargetUser

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' TW 23/03/2007 EP2_1942
'Public Sub SaveChangeRequestFromName(sObjectName As String, colMatchValues As Collection, Optional sDescription As String)
Public Sub SaveChangeRequestFromName(sObjectName As String, colMatchValues As Collection, Optional sDescription As String, Optional intPromotedCount As Integer)
' TW 23/03/2007 EP2_1942 End
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = g_clsHandleUpdates.GetObjectTableClass(sObjectName)
    
    clsTableAccess.SetKeyMatchValues colMatchValues
    
' TW 23/03/2007 EP2_1942
'    SaveChangeRequest clsTableAccess, sDescription
    SaveChangeRequest clsTableAccess, sDescription, , , intPromotedCount
' TW 23/03/2007 EP2_1942 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' TW 23/03/2007 EP2_1942
'Public Sub SaveChangeRequest(clsSrcTable As TableAccess, Optional sDescription As String, Optional enumOperation As SupervisorPromoteType = PromoteEdit, Optional ByVal sPromoteUser As String)
Public Sub SaveChangeRequest(clsSrcTable As TableAccess, Optional sDescription As String, Optional enumOperation As SupervisorPromoteType = PromoteEdit, Optional ByVal sPromoteUser As String, Optional intPromotedCount As Integer)
' TW 23/03/2007 EP2_1942 End
    On Error GoTo Failed
    Dim sObjectDescription As String
    Dim clsSupervisorConnection As SupervisorConnectionTable
    Dim clsTableAccess As TableAccess
    Dim colKeys As Collection
    Dim colKeyValues As Collection
    Dim vInstance As Variant
    Dim vConnection As Variant
    Dim sServiceName As String
    Dim sDatabaseUserID As String
    Dim sDirection As String
    Dim dtSave As Date
    
    ' Get the date.
' TW 23/03/2007 EP2_1942 - Artificially modify the change time to ensure it is unique for group promotion items
    dtSave = Now + (intPromotedCount / 86400#)
' TW 23/03/2007 EP2_1942 End
    
    ' Only do this if the tables exist for it
    Set clsSupervisorConnection = New SupervisorConnectionTable
    Set clsTableAccess = clsSupervisorConnection
    
    If g_clsDataAccess.DoesTableExist(clsTableAccess.GetTable()) Then
        sDatabaseUserID = g_clsDataAccess.GetDatabaseUserID()
        sServiceName = g_clsDataAccess.GetDatabaseName()
            
        sObjectDescription = GetObjectDescription(clsSrcTable)
        sDirection = DIRECTION_SOURCE
        
        Set colKeys = clsSrcTable.GetKeyMatchFields()
        Set colKeyValues = clsSrcTable.GetKeyMatchValues()
        
        If colKeys.Count > 0 And (colKeys.Count = colKeyValues.Count) Then
            ' We now have the date, keys, key values and object description, so save it away!
            ' First the connection
            vConnection = SaveConnection(clsSupervisorConnection, sServiceName, sDatabaseUserID)
        
            ' Next, the instance
            vInstance = SaveObjectInstance(colKeys, colKeyValues, sObjectDescription, sDescription)
        
            ' Now the Update
            SaveSupervisorUpdate vInstance, vConnection, sDirection, dtSave, enumOperation, sPromoteUser
    
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "SaveChangeRequest: Key value and field count invalid, " & colKeys.Count & ", " & colKeyValues.Count
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveSupervisorUpdate(vInstance As Variant, _
                                 vConnection As Variant, _
                                 sDirection As String, _
                                 dtSave As Date, _
                                 enumOperation As SupervisorPromoteType, _
                                 Optional sTargetUser As String)
    On Error GoTo Failed
    Dim clsUpdate As SupervisorUpdateTable
    
    Set clsUpdate = New SupervisorUpdateTable
    
    ' Check to see if this instance exists already. If it does, use that. If not, create
    ' a new row

    TableAccess(clsUpdate).AddRow
    
    clsUpdate.SetConnectionID vConnection
    clsUpdate.SetInstanceID vInstance
    clsUpdate.SetCreationDate dtSave
    
    ' DJP BM0423 If a user has been passed in, use that, otherwise use the logged on user
    If Len(sTargetUser) > 0 Then
        clsUpdate.SetSupervisorUserID sTargetUser
    Else
        clsUpdate.SetSupervisorUserID g_sSupervisorUser
    End If
    
    clsUpdate.SetDirection sDirection
    clsUpdate.SetOperation enumOperation
    
    TableAccess(clsUpdate).Update
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function SaveObjectInstance(colKeys As Collection, colValues As Collection, sName As String, sDescription As String) As Variant
    On Error GoTo Failed

    ' Need to loop through all keys, saving each row
    Dim sKey As String
    ' DJP SQL Server port
    Dim vValue As Variant
    Dim nCount As Integer
    Dim nThisKey As Integer
    Dim clsInstance As SupervisorObjInstTable
    Dim clsSupervisorObject As SupervisorObjectTable
    Dim clsTableAccess As TableAccess
    Dim vInstance As Variant

    nCount = colKeys.Count

    
    ' First update the  SupervisorObjectInstance table
    Set clsSupervisorObject = New SupervisorObjectTable
    ' Check to see if a record exists already for the object name and keys we have passed in.
    clsSupervisorObject.FindInstance colKeys, colValues, sName
    
    If TableAccess(clsSupervisorObject).RecordCount > 0 Then
        vInstance = clsSupervisorObject.GetInstanceID()
            
        Dim colMatchValues As Collection
        Set colMatchValues = New Collection
        colMatchValues.Add vInstance
        
        TableAccess(clsSupervisorObject).SetKeyMatchValues colMatchValues
        TableAccess(clsSupervisorObject).DeleteRecords
    
    End If
    Set clsSupervisorObject = New SupervisorObjectTable

    TableAccess(clsSupervisorObject).AddRow
    vInstance = clsSupervisorObject.SetInstanceID()

    ' Now update all the keys associated with this change
    Set clsInstance = New SupervisorObjInstTable
    Set clsTableAccess = clsInstance

    nThisKey = 1
    Do
        sKey = colKeys(nThisKey)
        vValue = colValues(nThisKey)


        clsTableAccess.AddRow
        clsInstance.SetInstanceID vInstance
        clsInstance.SetKey sKey
        ' DJP SQL Server port - need to force a string so it doesn't get written as a date and come back as a string
        ' as we have no idea what locale it was
        If TypeName(vValue) = "Date" Then
            clsInstance.SetValue CStr(vValue)
        Else
            clsInstance.SetValue vValue
        End If
        
        clsInstance.SetName sName
        clsInstance.SetDescription sDescription
        
        'The data-type sould be stored along with the key value. If it isn't,
        'then GUIDs are read-back as strings when data is promoted and wont
        'work in SQL-Server.
        clsInstance.SetValueType TypeName(vValue)
        
        nThisKey = nThisKey + 1
    Loop While nThisKey <= nCount

    ' Do the updates, first the SupervisorObjectInstance table
    Set clsTableAccess = clsSupervisorObject
    clsTableAccess.Update

    ' Now the SupervisorObjectInstanceDet table
    Set clsTableAccess = clsInstance
    clsTableAccess.Update

    SaveObjectInstance = vInstance
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function SaveConnection(clsSupervisorConnection As SupervisorConnectionTable, sServiceName As String, sDatabaseUserID As String) As Variant
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim vConnectionID As Variant
    
    Set clsTableAccess = clsSupervisorConnection

    clsSupervisorConnection.FindConnection sServiceName, sDatabaseUserID
    Set clsTableAccess = clsSupervisorConnection

    If clsTableAccess.RecordCount() = 0 Then
        ' New entry required
        clsTableAccess.AddRow
        vConnectionID = clsSupervisorConnection.SetConnectionID()
        
        clsSupervisorConnection.SetServiceName sServiceName
        clsSupervisorConnection.SetUserID sDatabaseUserID
        clsTableAccess.Update
    Else
        vConnectionID = clsSupervisorConnection.GetConnectionID()
    End If
    
    SaveConnection = vConnectionID
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
Private Function GetObjectDescription(clsTableAccess As TableAccess) As String
    On Error GoTo Failed
    Dim sTable As String
    Dim sDescription As String
     
    sTable = UCase(clsTableAccess.GetTable())
    
    Select Case sTable
        Case "MORTGAGEPRODUCT"
            sDescription = MORTGAGE_PRODUCTS
        
        Case "COMBOGROUP"
            sDescription = COMBOBOX_ENTRIES
        
        Case "DISTRIBUTIONCHANNEL"
            sDescription = DISTRIBUTION_CHANNELS
                
        Case "MORTGAGELENDER"
            sDescription = LENDERS
        
' TW 09/10/2006 EP2_7
        Case "ADDITIONALBORROWINGFEE"
            sDescription = ADDITIONAL_BORROWING_FEES
        
        Case "CREDITLIMITINCREASEFEE"
            sDescription = CREDIT_LIMIT_INCREASE_FEES
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case "TRANSFEROFEQUITYFEE"
            sDescription = TRANSFER_OF_EQUITY_FEES
' TW 11/12/2006 EP2_20 End
        
        Case "ADMINISTRATIONFEE"
            sDescription = ADMIN_FEES
        
        '*=[MC]BMIDS763 - CC075 FEESETS
        Case "PRODUCTSWITCHFEEBAND"
            sDescription = PRODUCT_SWITCH_FEESETS
        Case "INSURANCEADMINFEEBAND"
            sDescription = INSURANCE_ADMIN_FEESETS
        Case "TTFEEBAND"
            sDescription = TT_FEESETS
        '*=[MC]BMIDS763-SECTION END
        
        Case "RENTALINCOMERATESETBAND"  'JD BMIDS765
            sDescription = RENTAL_INCOME_RATES
            
        Case "VALUATIONFEE"
            sDescription = VALUATION_FEES
        
        Case "BASERATE"
            sDescription = BASE_RATE
        
        '   AW  17/06/02    BMIDS00055
        Case "BASERATEBAND"
            sDescription = BASE_RATES
            
        '   AW  14/05/02    BM088
        Case "INCOMEMULTIPLESET"
            sDescription = INCOME_MULTIPLE
                
        '   AW  16/05/02    BM087
        Case "MPMIGRATESET"
            sDescription = MP_MIG_RATE_SETS
            
        '   AW  16/05/02    BM087
        Case "MPMIGRATEBAND"
            sDescription = MP_MIG_RATE_BAND
            
        '   AW  16/05/02    BM017
        Case "REDEMPTIONFEESET"
            sDescription = REDEM_FEE_SETS
            
        '   AW  16/05/02    BM017

        Case "REDEMPTIONFEEBAND"
            sDescription = REDEM_FEE_BAND
            
        Case "COUNTRY"
            sDescription = COUNTRIES
        
        Case "LIFECOVERRATES"
            sDescription = LIFE_COVER_RATES
        
        Case "BUILDINGANDCONTENTSPRODUCT"
            sDescription = BUILDINGS_AND_CONTENTS_PRODUCTS
        
        Case "PAYMENTPROTECTIONPRODUCT"
            sDescription = PAYMENT_PROTECTION_PRODUCTS
        
        Case "PAYMENTPROTECTIONRATES"
            sDescription = PAYMENT_PROTECTION_RATES
        
        Case "DEPARTMENT"
            sDescription = DEPARTMENTS
        
        Case "UNIT"
            sDescription = UNITS
        
        Case "OMIGAUSER"
            sDescription = USERS
        
        Case "COMPETENCY"
            sDescription = COMPETENCIES
                
        Case "GLOBALPARAMETER"
            sDescription = GLOBAL_PARAM_FIXED
            
        Case "GLOBALBANDEDPARAMETER"
            sDescription = GLOBAL_PARAM_BANDED
        
        Case "REGION"
            sDescription = REGIONS
        
        Case "WORKINGHOURS"
            sDescription = WORKING_HOURS
        
        Case "MESSAGE"
            sDescription = ERROR_MESSAGES
        
        Case "NAMEANDADDRESSDIRECTORY", "PANEL"
            sDescription = NAMES_AND_ADDRESSES
        
        ' DJP Phase 2 Task Management
        Case "TASK"
            sDescription = TASK_MANAGEMENT_TASKS
        
        Case "STAGE"
            sDescription = TASK_MANAGEMENT_STAGES
        
        Case "ACTIVITY"
            sDescription = TASK_MANAGEMENT_ACTIVITIES
        
        Case "CASETRACK"
            sDescription = BUSINESS_GROUPS
            
        Case "ADDITIONALQUESTIONS"
            sDescription = ADDITIONAL_QUESTIONS
        
        Case "HOSTTEMPLATE"
            sDescription = PRINTING_TEMPLATE
        
        Case "CONDITIONS"
            sDescription = CONDITIONS
        
        Case "BATCH"
            sDescription = BATCH_SCHEDULER
        
        Case "CURRENCY"
            sDescription = CURRENCIES
        
        Case "INTERMEDIARY"
            'sDescription = INTERMEDIARIES
            
            'DB SYS5457 06/01/03 - If we are promoting an Intermediary, we need to find out the type of
            ' intermediary we are promoting
            sDescription = GetIntermediaryDescriptionFromType(clsTableAccess)
            
        Case "ALLOWABLEINCOMEFACTORS"
            sDescription = INCOME_FACTORS
        
        'MAR45 GHun
        Case "TEMPLATE"
            sDescription = PRINTING_DOCUMENT
        'MAR45 End
        
        'MAR202 GHun
        Case "PACKCONTROL"
            sDescription = PRINTING_PACK
        'MAR202 End
' TW 17/10/2006 EP2_15
        Case "MORTGAGECLUBNETWORKASSOCIATION"
' TW 18/11/2006 EP2_132
            sDescription = CLUBSANDASSOCIATIONS
' TW 18/11/2006 EP2_132
            
        Case "PRINCIPALFIRM"
' TW 18/11/2006 EP2_132
'            sDescription = frmMain.tvwDB.SelectedItem
            sDescription = PRINCIPALFIRMPACKAGER
        Case "INTRODUCERFIRM"
            sDescription = INTRODUCERFIRM
' TW 23/11/2006 EP2_172
        Case "MORTGAGEPRODUCTEXCLUSIVITY"
            sDescription = EXCLUSIVEPRODUCTS
' TW 23/11/2006 EP2_172 End

' TW 18/11/2006 EP2_132
        Case "ARFIRM"
'            sDescription = frmMain.tvwDB.SelectedItem
            sDescription = ARFIRMS
        Case "FIRMTRADINGNAME"
            sDescription = FIRMTRADINGNAME
        Case "INTRODUCER"
            sDescription = "Introducer"
        Case "FIRMPERMISSIONS"
            sDescription = FIRMPERMISSIONS
        Case "FIRMCLUBNETWORKASSOCIATION"
            sDescription = NETWORKASSOCIATIONS
' TW 17/10/2006 EP2_15 End
' TW 30/11/2006 EP2_253
        Case "MORTGAGEPRODUCTPRODUCTCLASS"
            sDescription = MORTGAGEPRODUCTPRODUCTCLASS
        Case "MORTGAGEPRODUCTNATUREOFLOAN"
            sDescription = MORTGAGEPRODUCTNATUREOFLOAN
        Case "MORTGAGEPRODUCTINCOMESTATUS"
            sDescription = MORTGAGEPRODUCTINCOMESTATUS
' TW 30/11/2006 EP2_253 End
' TW 14/12/2006 EP2_518
        Case "PROCFEEDEFAULT"
            sDescription = DEFAULTPROCURATIONFEES
        Case "PROCFEEADJBYLOAN"
            sDescription = LOANAMOUNTADJUSTMENTS
        Case "PROCFEEADJBYLTV"
            sDescription = LTVAMOUNTADJUSTMENTS
' TW 14/12/2006 EP2_518 End
' TW 05/02/2007 EP2_706
        Case "APPOINTMENTS"
            sDescription = APPOINTMENTS
' TW 05/02/2007 EP2_706 End

        Case Else
            ' SYS4121 Support client variants
            Dim clsPromote As PromoteCS
            Set clsPromote = New PromoteCS
            sDescription = clsPromote.GetObjectDescription(sTable)
            If Len(sDescription) = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unknown object type: " & sTable
            End If
    End Select
    
    GetObjectDescription = sDescription
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub DeleteObject(clsSupervisorUpdateDetails As SupervisorUpdateDetails, _
                        clsTableAccess As TableAccess, _
                        sTargetDatabase As String)
    
    Dim colMatchValues As Collection
    Dim sObjectName As String
    
    sObjectName = clsSupervisorUpdateDetails.GetObjectName()
    Set colMatchValues = clsSupervisorUpdateDetails.GetKeyMatchValues()
    'use the same method we use in main super to delete these, but use the
    'target database too.
    If Not colMatchValues Is Nothing Then
        g_clsMainSupport.DeleteObject sObjectName, colMatchValues, sTargetDatabase, False
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "DeleteObject - key values are empty"
    End If

    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteObject(clsSupervisorUpdateDetails As SupervisorUpdateDetails, _
                         clsTableAccess As TableAccess, _
                         sTargetDatabase As String)
    ' Need to know the target database, identified by the key. Then need
    ' to know the keys, and the key match values
    On Error GoTo Failed
    Dim sTable As String
    Dim sFunctionName As String
    Dim clsPromote As SupervisorObjectCopy
    Dim clsProgress As MSGProgress
    Dim sObjectName As String
    Dim enumOperation As SupervisorPromoteType
    Dim clsPromoteCS As PromoteCS
    Dim bPromoted As Boolean
    
    Set clsPromoteCS = New PromoteCS
    
    sFunctionName = "PromoteObject"
    sTable = UCase(clsTableAccess.GetTable())
    sObjectName = clsSupervisorUpdateDetails.GetObjectName()
    
    Set clsPromote = New SupervisorObjectCopy
    
    Set clsProgress = New MSGProgress
    
    clsProgress.StartProgress frmManageUpdates, "Opening target database: " + sTargetDatabase + "..."
    clsProgress.StartProgress frmManageUpdates, "Updating: " + sObjectName + "..."
    
    clsPromote.SetProgress clsProgress
    
    enumOperation = clsSupervisorUpdateDetails.GetOperation()
    
    bPromoted = True
    
    If enumOperation = PromoteEdit Then
        Select Case sTable
            Case "MORTGAGEPRODUCT"
                clsPromote.PromoteProduct clsTableAccess, sTargetDatabase
            
            Case "COMBOGROUP"
                clsPromote.PromoteComboGroup clsTableAccess, sTargetDatabase
            
            Case "MORTGAGELENDER"
                clsPromote.PromoteLender clsTableAccess, sTargetDatabase
            
            Case "COUNTRY"
                clsPromote.PromoteCountry clsTableAccess, sTargetDatabase
            
            Case "DISTRIBUTIONCHANNEL"
                clsPromote.PromoteDistributionChannel clsTableAccess, sTargetDatabase
            
            Case "ADMINISTRATIONFEE", "VALUATIONFEE", "PAYMENTPROTECTIONRATES", _
                    "REDEMPTIONFEEBAND", "MPMIGRATEBAND", _
                    "ADDITIONALBORROWINGFEE", "CREDITLIMITINCREASEFEE", "TRANSFEROFEQUITYFEE"
                '   AW  29/07/02    BMIDS00237
                clsPromote.PromoteBandedFee clsTableAccess, sTargetDatabase, False
            '*=[MC]BMIDS763 - FEESETS PROMOTION CODE  JD ADDED RENTALINCOMERATES JD
            Case "TTFEEBAND", "PRODUCTSWITCHFEEBAND", "INSURANCEADMINFEEBAND", "RENTALINCOMERATESETBAND"
                clsPromote.PromoteBandedFee clsTableAccess, sTargetDatabase, False
            'BMIDS763 - SECTION END
            
'            Case "PAYMENTPROTECTIONRATES"
'                clsPromote.PromotePayProtRates clsTableAccess, sTargetDatabase
            
            Case "LIFECOVERRATES"
                clsPromote.PromoteLifeCoverRates clsTableAccess, sTargetDatabase
            
            Case "BUILDINGANDCONTENTSPRODUCT"
                clsPromote.PromoteBAndCProduct clsTableAccess, sTargetDatabase
            
            Case "PAYMENTPROTECTIONPRODUCT"
                clsPromote.PromotePayProtProducts clsTableAccess, sTargetDatabase
    
            Case "DEPARTMENT"
                clsPromote.PromoteDepartment clsTableAccess, sTargetDatabase
    
            Case "UNIT"
                clsPromote.PromoteUnit clsTableAccess, sTargetDatabase
            
            Case "OMIGAUSER"
                clsPromote.PromoteUser clsTableAccess, sTargetDatabase
            
            Case "COMPETENCY"
                clsPromote.PromoteCompetency clsTableAccess, sTargetDatabase
            
            Case "GLOBALPARAMETER"
                clsPromote.PromoteGlobalParameter clsTableAccess, sTargetDatabase
                
            Case "GLOBALBANDEDPARAMETER"
                clsPromote.PromoteBandedParameter clsTableAccess, sTargetDatabase
            
            Case "REGION"
                clsPromote.PromoteRegion clsTableAccess, sTargetDatabase
            
            Case "WORKINGHOURS"
                clsPromote.PromoteWorkingHours clsTableAccess, sTargetDatabase
            
            Case "MESSAGE"
                clsPromote.PromoteMessages clsTableAccess, sTargetDatabase
            
            Case "NAMEANDADDRESSDIRECTORY"
                clsPromote.PromoteThirdParty clsTableAccess, sTargetDatabase
            
            Case "ADDITIONALQUESTIONS"
                clsPromote.PromoteQuestion clsTableAccess, sTargetDatabase
            
            Case "TASK"
                clsPromote.PromoteTask clsTableAccess, sTargetDatabase
            
            Case "ACTIVITY"
                clsPromote.PromoteActivity clsTableAccess, sTargetDatabase
            
            Case "STAGE"
                clsPromote.PromoteStage clsTableAccess, sTargetDatabase
            
            Case "CONDITIONS"
                clsPromote.PromoteCondition clsTableAccess, sTargetDatabase
            'GD BMIDS00016
            Case "MORTGAGEPRODUCTCONDITION"
                clsPromote.PromoteMortgageProductCondition clsTableAccess, sTargetDatabase
                  
            Case "CASETRACK"
                clsPromote.PromoteBusinessGroup clsTableAccess, sTargetDatabase
            
            Case "CURRENCY"
                clsPromote.PromoteCurrency clsTableAccess, sTargetDatabase
            
            Case "HOSTTEMPLATE"
                clsPromote.PromotePrintingTemplate clsTableAccess, sTargetDatabase
            
            Case "INTERMEDIARY"
                clsPromote.PromoteIntermediary clsTableAccess, sTargetDatabase
                
            Case "BASERATE"
                clsPromote.PromoteBaseRate clsTableAccess, sTargetDatabase
            
            Case "BASERATEBAND"
                clsPromote.PromoteBaseRateSet clsTableAccess, sTargetDatabase
            'GD BMIDS00170
            'Case INCOME_FACTORS
            Case "ALLOWABLEINCOMEFACTORS"
                'BMIDS00170 16/10/02
                ' Set clsTableAccess = New AllowableIncomeFactorTable
                clsPromote.PromoteIncomefactor clsTableAccess, sTargetDatabase
                'BMIDS00170 16/10/02
            
            'MAR45 GHun
            Case "TEMPLATE"
                clsPromote.PromoteDocumentLocations clsTableAccess, sTargetDatabase
            'MAR45 End
            
            'MAR202 GHun
            Case "PACKCONTROL"
                clsPromote.PromotePackControl clsTableAccess, sTargetDatabase
            'MAR202 End

' TW 17/10/2006 EP2_15
            Case "MORTGAGECLUBNETWORKASSOCIATION"
                clsPromote.PromoteMortgageClubNetworkAssociation clsTableAccess, sTargetDatabase
            
            Case "PRINCIPALFIRM"
                clsPromote.PromotePrincipalFirm clsTableAccess, sTargetDatabase
            
            Case "INTRODUCER"
                clsPromote.PromoteIntroducer clsTableAccess, sTargetDatabase
            
            Case "INTRODUCERFIRM"
                clsPromote.PromoteIntroducerFirm clsTableAccess, sTargetDatabase
            
            Case "FIRMPERMISSIONS"
                clsPromote.PromoteFirmPermissions clsTableAccess, sTargetDatabase
' TW 17/10/2006 EP2_15 End
' TW 18/11/2006 EP2_132
            Case "FIRMTRADINGNAME"
                clsPromote.PromoteFirmTradingName clsTableAccess, sTargetDatabase
            Case "ARFIRM"
                clsPromote.PromoteARFirm clsTableAccess, sTargetDatabase
            Case "FIRMCLUBNETWORKASSOCIATION"
                clsPromote.PromoteFirmNetAssoc clsTableAccess, sTargetDatabase
' TW 18/11/2006 EP2_132 End
' TW 23/11/2006 EP2_172
            Case "MORTGAGEPRODUCTEXCLUSIVITY"
                clsPromote.PromoteMortgageProductExclusivity clsTableAccess, sTargetDatabase
' TW 23/11/2006 EP2_172 End
' TW 30/11/2006 EP2_253
'            Case "MORTGAGEPRODUCTPRODUCTCLASS"
'                clsPromote.PromoteMortgageProductProductClass clsTableAccess, sTargetDatabase
'            Case "MORTGAGEPRODUCTNATUREOFLOAN"
'                clsPromote.PromoteMortgageNatureOfLoanClass clsTableAccess, sTargetDatabase
'            Case "MORTGAGEPRODUCTINCOMESTATUS"
'                clsPromote.PromoteMortgageIncomeStatusClass clsTableAccess, sTargetDatabase
' TW 30/11/2006 EP2_253 End
' TW 14/12/2006 EP2_518
            Case "PROCFEEDEFAULT"
                clsPromote.PromoteProcFeeDefault clsTableAccess, sTargetDatabase
            Case "PROCFEEADJBYLOAN"
                clsPromote.PromoteProcFeeAdjByLoan clsTableAccess, sTargetDatabase
            Case "PROCFEEADJBYLTV"
                clsPromote.PromoteProcFeeAdjByLTV clsTableAccess, sTargetDatabase
' TW 14/12/2006 EP2_518 End
' TW 05/02/2007 EP2_706
            Case "APPOINTMENTS"
                clsPromote.PromoteAppointments clsTableAccess, sTargetDatabase
' TW 05/02/2007 EP2_706 End

            Case Else
                ' DJP SYS4121
                bPromoted = False
        
        End Select
        
        ' If we get this far, a promotion should have successfully taken place, so allow any additional
        ' client specific promotion to take place. This is for additional tables that a client may have added,
        ' for Mortgage Products, for example.
        bPromoted = clsPromoteCS.PromoteObject(clsPromote, clsTableAccess, sTargetDatabase) Or bPromoted
        
        If Not bPromoted Then
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unknown object, " + sTable
        End If
        
    ElseIf enumOperation = PromoteDelete Then
        DeleteObject clsSupervisorUpdateDetails, clsTableAccess, sTargetDatabase
    End If
    
    clsProgress.ClearProgress
    Exit Sub
Failed:
    clsProgress.ClearProgress
    g_clsErrorHandling.RaiseError , "Failed to promote one or more of the selected items. Check to ensure all related items are promoted onto the target system first. The error reported was: " & Err.DESCRIPTION
End Sub
Public Function GetObjectTableClass(sObjectName As String, Optional bTableForEdit As Boolean = True) As TableAccess
    On Error GoTo Failed
    Dim sFunctionName As String
    Dim clsTableAccess As TableAccess
    
    sFunctionName = "GetObjectTableClass"
    
    Select Case sObjectName
        Case LOCKS_ONLINE_APPLICATION, LOCKS_OFFLINE_APPLICATION
            Set clsTableAccess = New ApplicationLockTable
        
        Case LOCKS_ONLINE_CUSTOMER, LOCKS_OFFLINE_CUSTOMER
            Set clsTableAccess = New CustomerLockTable
        
        Case MORTGAGE_PRODUCTS
            Set clsTableAccess = New MortgageProductTable
        
        Case COMBOBOX_ENTRIES
            Set clsTableAccess = New ComboValueGroupTable
        
        Case LENDERS
            Set clsTableAccess = New MortgageLendersTable
        
' TW 09/10/2006 EP2_7
        Case ADDITIONAL_BORROWING_FEES
            Set clsTableAccess = New AdditionalBorrowingFeeTable
            
        Case CREDIT_LIMIT_INCREASE_FEES
            Set clsTableAccess = New CreditLimitIncreaseFeeTable
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case TRANSFER_OF_EQUITY_FEES
            Set clsTableAccess = New TransferOfEquityFeeTable
' TW 11/12/2006 EP2_20 End
        
' TW 17/10/2006 EP2_15
' TW 15/01/2007 EP2_826
'        Case CLUBSANDASSOCIATIONS
        Case CLUBSANDASSOCIATIONS, ASSOCIATIONS, CLUBS
' TW 15/01/2007 EP2_826 End
            Set clsTableAccess = New MortgageClubNetAssocTable
' TW 15/01/2007 EP2_826
'        Case INTRODUCER
        Case INTRODUCER, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS
' TW 15/01/2007 EP2_826 End
            Set clsTableAccess = New IntroducerTable
' TW 18/11/2006 EP2_132
'        Case PACKAGERS
' TW 15/01/2007 EP2_826
'        Case PACKAGERS, PRINCIPALFIRMPACKAGER
        Case PACKAGERS, PRINCIPALFIRMPACKAGER, PRINCIPALS
' TW 15/01/2007 EP2_826 End
            Set clsTableAccess = New PrincipalFirmTable
        Case INTRODUCERFIRM
            Set clsTableAccess = New IntroducerFirmTable
        Case FIRMPERMISSIONS
            Set clsTableAccess = New FirmPermissionsTable
        Case NETWORKASSOCIATIONS
            Set clsTableAccess = New FirmClubNetAssocTable
' TW 18/11/2006 EP2_132 End
' TW 17/10/2006 EP2_15 End
' TW 18/11/2006 EP2_132
        Case FIRMTRADINGNAME
            Set clsTableAccess = New FirmTradingNameTable
        Case ARFIRMS
            Set clsTableAccess = New ARFirmTable
' TW 18/11/2006 EP2_132 End
' TW 23/11/2006 EP2_172
        Case EXCLUSIVEPRODUCTS
            Set clsTableAccess = New MortProdExclusivityTable
' TW 23/11/2006 EP2_172 End
' TW 30/11/2006 EP2_253
        Case MORTGAGEPRODUCTPRODUCTCLASS
            Set clsTableAccess = New MortProdProdClassTable
        Case MORTGAGEPRODUCTNATUREOFLOAN
            Set clsTableAccess = New MortProdNatureOfLoanTable
        Case MORTGAGEPRODUCTINCOMESTATUS
            Set clsTableAccess = New MortProdIncomeStatusTable
' TW 30/11/2006 EP2_253 End
' TW 14/12/2006 EP2_518
        Case DEFAULTPROCURATIONFEES
            Set clsTableAccess = New ProcFeeDefaultTable
        Case LOANAMOUNTADJUSTMENTS
            Set clsTableAccess = New ProcFeeAdjByLoanTable
        Case LTVAMOUNTADJUSTMENTS
            Set clsTableAccess = New ProcFeeAdjByLTVTable
' TW 14/12/2006 EP2_518 End
' TW 05/02/2007 EP2_706
        Case APPOINTMENTS
            Set clsTableAccess = New AppointmentsTable
' TW 05/02/2007 EP2_706 End

        Case ADMIN_FEES
            Set clsTableAccess = New AdminFeeTable
        
        '*=[MC]BMIDS763 - CC075 - FEESETS
        Case PRODUCT_SWITCH_FEESETS
            Set clsTableAccess = New ProductSwitchFeeTable
        Case INSURANCE_ADMIN_FEESETS
            Set clsTableAccess = New InsuranceAdminFeeBand
        Case TT_FEESETS
            Set clsTableAccess = New TTFeeBand
        '*=[MC]BMIDS763 - CC075 - FEESETS
        
        Case RENTAL_INCOME_RATES  'JD BMIDS765
            Set clsTableAccess = New RentalIncomeRateSetBandTable
            
        Case VALUATION_FEES
            Set clsTableAccess = New ValuationFeeTable
        
        Case BASE_RATE
            Set clsTableAccess = New RateTable
            
        Case BASE_RATES
            Set clsTableAccess = New BaseRateTable
        
        '   AW  14/05/02    BM088
        Case INCOME_MULTIPLE
            Set clsTableAccess = New IncomeMultipleSetTable
           
        '   AW  16/05/02    BM087
        Case MP_MIG_RATE_SETS
            Set clsTableAccess = New MPMigRateTable
            
        '   AW  16/05/02    BM017
        Case REDEM_FEE_SETS
            Set clsTableAccess = New RedemptionFeeTable
        '   GD  16/05/02    AQR : BMIDS00016
        
        Case MORTGAGEPRODUCTCONDITION
            Set clsTableAccess = New MortProdProdCondTable
            
        Case COUNTRIES
            Set clsTableAccess = New CountryTable
        
        Case DISTRIBUTION_CHANNELS
            Set clsTableAccess = New DistributionChannelTable
        
        Case BUILDINGS_AND_CONTENTS_PRODUCTS
            Set clsTableAccess = New BuildingAndContentsTable
        
        Case LIFE_COVER_RATES
            Set clsTableAccess = New LifeCoverRatesTable
        
        Case PAYMENT_PROTECTION_RATES
            Set clsTableAccess = New PayProtRatesTable
        
        Case PAYMENT_PROTECTION_PRODUCTS
            Set clsTableAccess = New PayProtProductsTable
        
        Case DEPARTMENTS
            Set clsTableAccess = New DepartmentTable
        
        Case UNITS
            Set clsTableAccess = New UnitTable
        
        Case USERS
            Set clsTableAccess = New OmigaUserTable
        
        Case COMPETENCIES
            Set clsTableAccess = New CompetencyTable
        
        Case GLOBAL_PARAM_FIXED
            Set clsTableAccess = New FixedParametersTable
        
        Case GLOBAL_PARAM_BANDED
            Set clsTableAccess = New BandedGlobalParametersTable
        
        Case REGIONS
            Set clsTableAccess = New RegionTable
        
        Case WORKING_HOURS
            Set clsTableAccess = New WorkingHoursTable
        
        Case ERROR_MESSAGES
            Set clsTableAccess = New ErrorMessageTable
        
        Case NAMES_AND_ADDRESSES, LEGAL_REP_ADDRESS, VALUER_ADDRESS, LOCAL_ADDRESS
            If bTableForEdit Then
                Set clsTableAccess = New NameAndAddressDirTable
            Else
                Set clsTableAccess = New PanelTable
            End If
        
        Case TASK_MANAGEMENT_TASKS
            Set clsTableAccess = New TaskTable
        
        Case TASK_MANAGEMENT_STAGES
            Set clsTableAccess = New StageTable
        
        Case TASK_MANAGEMENT_ACTIVITIES
            Set clsTableAccess = New ActivityTable
        
        Case PRINTING_TEMPLATE
            Set clsTableAccess = New PrintingTemplateTable
        
        Case ADDITIONAL_QUESTIONS
            Set clsTableAccess = New AdditionalQuestionsTable
        
        Case BUSINESS_GROUPS
            Set clsTableAccess = New BusinessGroupTable
        
        Case CONDITIONS
            Set clsTableAccess = New ConditionsTable
        
        Case BATCH_SCHEDULER
            Set clsTableAccess = New BatchTable
        
        Case CURRENCIES
            Set clsTableAccess = New CurrencyTable
        
        'DB SYS5457 06/01/03 - Instead of just one type of Intermediary, we now have four, although the INTERMEDIARIES type is still used
        Case INTERMEDIARIES, LEADAGENT_DESCRIPTION, INDIVIDUAL_DESCRIPTION, INTERMEDIARY_COMPANY_DESCRIPTION, ADMIN_CENTRE_DESCRIPTION, _
            LEADAGENT, INDIVIDUAL, INTERMEDIARY_COMPANY, ADMIN_CENTRE, BROKER, PACKAGER ' EP15 pct
            Set clsTableAccess = New IntermediaryTable
            
            'Convert the object name to a single value.
            sObjectName = INTERMEDIARIES
        
        'GD BMIDS00170 start
        Case INCOME_FACTORS
            Set clsTableAccess = New AllowableIncomeFactorTable
        'GD BMIDS00170 end
        
        'MAR45 GHun
        Case PRINTING_DOCUMENT
            Set clsTableAccess = New TemplateTable
        'MAR45 End
        
        'MAR202 GHun
        Case PRINTING_PACK
            Set clsTableAccess = New PackControlTable
        'MAR202 End
        
        
        
        Case Else
            ' SYS4121 DJP Updated to allow client variant to return its table class.
            Dim clsPromote As PromoteCS
            Set clsPromote = New PromoteCS
            Set clsTableAccess = clsPromote.GetObjectTableClass(sObjectName, bTableForEdit)
            
            If clsTableAccess Is Nothing Then
                g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unlnown object name - " & sObjectName
            End If
            
    End Select
    
    Set GetObjectTableClass = clsTableAccess
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'DB SYS5457 06/01/03 - Taken from MCAP.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryDescriptionFromType
' Description   : Given the Intermediary passed in, read the type of Intermediary, and depending on
'                 the type, return a textual description of that type.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetIntermediaryDescriptionFromType(clsIntermediary As IntermediaryTable) As String
    On Error GoTo Failed
    Dim nType As Variant
    Dim sIntermediary As String
    Dim nRecordCount As Long
    
    ' First, find out if we have actually been to the database to retrieve the record - there should be only 1.
    nRecordCount = TableAccess(clsIntermediary).RecordCount
    
    ' If we don't have the record, go to the database and retrieve it.
    If nRecordCount = 0 Then
        TableAccess(clsIntermediary).GetTableData
    End If
    
    nType = clsIntermediary.GetIntermediaryType
    
    ' Return the description for the type retrieved.
    If Not IsNull(nType) And Not IsEmpty(nType) Then
        Select Case nType
        Case LeadAgentType
            sIntermediary = LEADAGENT_DESCRIPTION
        
        Case CompanyType
            sIntermediary = INTERMEDIARY_COMPANY_DESCRIPTION
        
        Case AdminCentreType
            sIntermediary = ADMIN_CENTRE_DESCRIPTION
        
        Case IndividualType
            sIntermediary = INDIVIDUAL_DESCRIPTION
                    
        Case Else
            sIntermediary = INTERMEDIARIES
        End Select
    Else
        sIntermediary = INTERMEDIARIES
    End If

    GetIntermediaryDescriptionFromType = sIntermediary
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'DB End
Private Sub Class_Initialize()

End Sub


