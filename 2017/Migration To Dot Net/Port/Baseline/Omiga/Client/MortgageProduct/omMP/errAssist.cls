VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Workfile:      ErrAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Helper object for handling errors
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'IK     30/06/99 Created
'RF     05/11/99 Added optional parameter to IsWarning
'RF     17/11/99 Indicate which errors (if between 1 and 499) have been hardcoded
'SR     24/03/00 New method ReRaiseResponseError
'SR     28/03/00 New method 'GetErrorNumberFromResponse'
'MH     14/07/00 Added version number info into XML Error string
'PSC    07/08/00 Change version to return the App.Comments as this is where the
'                Build No is held
'DJP    12/09/00 Added oeMTSNotFound for when there is no object context.
'CL     23/10/00 CORE000001 Added error 556 Cannot connect to database
'MV     29/11/00 Added omiga4NoDataForParameter ( Err Number 291)
'ASm    30/11/00 CORE000022: Added oeMissingParameter error number 557
'ASm    11/01/01 SYS1817: OMIGAERROR moved to errConstants.bas
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Enumeration Declaration Section
 '=============================================

'***********************************************************************************
'* THIS FOLLOWING ENUM IS HERE FOR BACKWARD COMPATIBILITY ONLY.
'* PLEASE DO NOT ADD TO THIS.
'***********************************************************************************
Public Enum ERRNUMBER
    
    omiga4Err001 = vbObjectError + 512 + 1
    omiga4Err002 = vbObjectError + 512 + 2
    
    omiga4Err100 = vbObjectError + 512 + 100
    omiga4Err101 = vbObjectError + 512 + 101
    omiga4Err102 = vbObjectError + 512 + 102
    omiga4Err103 = vbObjectError + 512 + 103
    omiga4err104 = vbObjectError + 512 + 104
    omiga4Err105 = vbObjectError + 512 + 105
    omiga4err106 = vbObjectError + 512 + 106
    omiga4err107 = vbObjectError + 512 + 107
    omiga4err108 = vbObjectError + 512 + 108
    omiga4err109 = vbObjectError + 512 + 109
    omiga4err110 = vbObjectError + 512 + 110
    omiga4err111 = vbObjectError + 512 + 111
    omiga4err112 = vbObjectError + 512 + 112
    omiga4err113 = vbObjectError + 512 + 113
    omiga4err114 = vbObjectError + 512 + 114
    omiga4err115 = vbObjectError + 512 + 115
    omiga4err116 = vbObjectError + 512 + 116
    omiga4err117 = vbObjectError + 512 + 117
    omiga4err118 = vbObjectError + 512 + 118
    omiga4err119 = vbObjectError + 512 + 119
    omiga4err120 = vbObjectError + 512 + 120
    omiga4err121 = vbObjectError + 512 + 121
    omiga4err122 = vbObjectError + 512 + 122
    omiga4err123 = vbObjectError + 512 + 123
    omiga4err124 = vbObjectError + 512 + 124
    omiga4err125 = vbObjectError + 512 + 125
    omiga4err126 = vbObjectError + 512 + 126
    omiga4err127 = vbObjectError + 512 + 127
    omiga4Err128 = vbObjectError + 512 + 128
    omiga4Err129 = vbObjectError + 512 + 129
    omiga4Err130 = vbObjectError + 512 + 130
    omiga4Err131 = vbObjectError + 512 + 131
    omiga4Err132 = vbObjectError + 512 + 132
    omiga4Err133 = vbObjectError + 512 + 133
    omiga4Err134 = vbObjectError + 512 + 134
    omiga4Err135 = vbObjectError + 512 + 135
    omiga4err136 = vbObjectError + 512 + 136
    omiga4err137 = vbObjectError + 512 + 137
    omiga4err138 = vbObjectError + 512 + 138
    omiga4err139 = vbObjectError + 512 + 139
    
    omiga4err147 = vbObjectError + 512 + 147
    omiga4err148 = vbObjectError + 512 + 148
    omiga4err149 = vbObjectError + 512 + 149
    omiga4err150 = vbObjectError + 512 + 150
    omiga4err151 = vbObjectError + 512 + 151
    omiga4err152 = vbObjectError + 512 + 152
    omiga4err153 = vbObjectError + 512 + 153
    omiga4err154 = vbObjectError + 512 + 154
    omiga4err155 = vbObjectError + 512 + 155
    omiga4err156 = vbObjectError + 512 + 156
    omiga4err157 = vbObjectError + 512 + 157
    omiga4err158 = vbObjectError + 512 + 158
    omiga4err159 = vbObjectError + 512 + 159
    
    omiga4err163 = vbObjectError + 512 + 163
    omiga4err164 = vbObjectError + 512 + 164
    omiga4err165 = vbObjectError + 512 + 165
    omiga4err166 = vbObjectError + 512 + 166
    omiga4err167 = vbObjectError + 512 + 167
    omiga4err168 = vbObjectError + 512 + 168
    omiga4err169 = vbObjectError + 512 + 169
    omiga4err170 = vbObjectError + 512 + 170
    omiga4err171 = vbObjectError + 512 + 171
    omiga4err172 = vbObjectError + 512 + 172
    omiga4err173 = vbObjectError + 512 + 173
    omiga4err174 = vbObjectError + 512 + 174
    omiga4err175 = vbObjectError + 512 + 175
    omiga4err176 = vbObjectError + 512 + 176
    
    omiga4err182 = vbObjectError + 512 + 182
    omiga4err183 = vbObjectError + 512 + 183
    
    omiga4err190 = vbObjectError + 512 + 190
    omiga4err191 = vbObjectError + 512 + 191
    omiga4err192 = vbObjectError + 512 + 192
    
    omiga4err207 = vbObjectError + 512 + 207
    omiga4err208 = vbObjectError + 512 + 208
    omiga4err209 = vbObjectError + 512 + 209
    omiga4err210 = vbObjectError + 512 + 210
    omiga4err211 = vbObjectError + 512 + 211
    omiga4err212 = vbObjectError + 512 + 212
    omiga4err213 = vbObjectError + 512 + 213
    
    omiga4err222 = vbObjectError + 512 + 222
    
    omiga4err228 = vbObjectError + 512 + 228
    omiga4err229 = vbObjectError + 512 + 229
    omiga4err230 = vbObjectError + 512 + 230
    omiga4err231 = vbObjectError + 512 + 231
    omiga4err232 = vbObjectError + 512 + 232
    omiga4err233 = vbObjectError + 512 + 233
    omiga4err234 = vbObjectError + 512 + 234
    
    omiga4Err262 = vbObjectError + 512 + 262
    
    omiga4RecordNotFound = vbObjectError + 512 + 500
    omiga4SurnameChange = vbObjectError + 512 + 501
    omiga4XMLParserError = vbObjectError + 512 + 502
    omiga4InvalidKeyString = vbObjectError + 512 + 503
    omiga4NoAfterImagePresent = vbObjectError + 512 + 504
    omiga4CommandFailed = vbObjectError + 512 + 505
    omiga4MissingPrimaryTag = vbObjectError + 512 + 506
    omiga4InvalidParameter = vbObjectError + 512 + 507
    omiga4DuplicateKey = vbObjectError + 512 + 508
    omiga4NoDataForCreate = vbObjectError + 512 + 509
    omiga4NoDataForUpdate = vbObjectError + 512 + 510
    omiga4ArrayLimitExceeded = vbObjectError + 512 + 511
    omiga4NoRowsAffected = vbObjectError + 512 + 512
    omiga4InvalidNoOfRows = vbObjectError + 512 + 513
    omiga4internalerror = vbObjectError + 512 + 514
    omiga4MissingFieldDesc = vbObjectError + 512 + 515
    omiga4MissingTableDesc = vbObjectError + 512 + 516
    omiga4MissingTypeDesc = vbObjectError + 512 + 517
    omiga4MissingElementValue = vbObjectError + 512 + 518
    omiga4MissingElement = vbObjectError + 512 + 519
    omiga4MissingKeyDesc = vbObjectError + 512 + 520
    omiga4MissingUpdateSet = vbObjectError + 512 + 521
    omiga4MissingTableName = vbObjectError + 512 + 522
    omiga4MissingXMLTableName = vbObjectError + 512 + 523
    omiga4MissingKey = vbObjectError + 512 + 524
    omiga4NoRowsAffectedByDeleteAll = vbObjectError + 512 + 525
    omiga4InValidKeyValue = vbObjectError + 512 + 526
    omiga4InValidDataTypeValue = vbObjectError + 512 + 527
    omiga4InValidKey = vbObjectError + 512 + 528
    omiga4NoFieldsFound = vbObjectError + 512 + 529
    omiga4NoFieldItemFound = vbObjectError + 512 + 530
    omiga4NoFieldItemName = vbObjectError + 512 + 531
    omiga4NoComboTagValue = vbObjectError + 512 + 532
    omiga4InvalidDateTimeFormat = vbObjectError + 512 + 533
    omiga4NoBeforeImagePresent = vbObjectError + 512 + 534
    omiga4MissingLoadDataEnum = vbObjectError + 512 + 535
    omiga4MissingOrInvalidEnumForValidation = vbObjectError + 512 + 536
    
    omiga4NotImplemented = vbObjectError + 512 + 900
    omiga4InvalidMessageNo = 901

    omiga4UnspecifiedError = vbObjectError + 512 + 999
    
End Enum

'local variable(s) to hold property value(s)
Private mvarNumber As Long 'local copy
Private mvarSource As String 'local copy
Private mvarDescription As String 'local copy
Private mvarHelpContext As String 'local copy
Private mvarHelpFile As String 'local copy

Private mstrMessageText As String
Private mstrMessageType As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BMIDS Specific History
'
'Prog   Date        Description
'DB     21/03/2003  BM0483  Performance upgrades - generate correct omiga error numbers
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function CreateErrorResponse() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlReponseElem As IXMLDOMElement
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlDescriptionElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlReponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlReponseElem
    
    Set xmlErrorElem = xmlDoc.createElement("ERROR")
    xmlReponseElem.appendChild xmlErrorElem
    
    Set xmlElement = xmlDoc.createElement("NUMBER")
    xmlElement.Text = Err.Number
    xmlErrorElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("SOURCE")
    xmlElement.Text = Err.Source
    xmlErrorElem.appendChild xmlElement
    
    Set xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION")
    xmlDescriptionElem.Text = Err.Description
    xmlErrorElem.appendChild xmlDescriptionElem
    
    Set xmlElement = xmlDoc.createElement("VERSION")
    xmlElement.Text = App.Comments
    xmlErrorElem.appendChild xmlElement
    
    If Me.IsApplicationError = True Then
        xmlReponseElem.setAttribute "TYPE", "APPERR"
    
        If Len(xmlDescriptionElem.Text) = 0 Then
            xmlDescriptionElem.Text = Me.GetMessageText(Err.Number)
        End If
    Else
        xmlReponseElem.setAttribute "TYPE", "SYSERR"
    End If
        
    CreateErrorResponse = xmlDoc.xml
    
    Set xmlDoc = Nothing
    Set xmlReponseElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlDescriptionElem = Nothing
    Set xmlElement = Nothing


End Function

Public Function FormatParserError(ByVal objParseError As IXMLDOMParseError) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Dim strErrDesc As String    ' formatted parser error
    
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str(objParseError.Line) & vbCr & _
        "Character: " & Str(objParseError.linepos) & vbCr & _
        "Source text: " & objParseError.srcText
    
    FormatParserError = strErrDesc

End Function

Public Function GetMessageText(lngMessageNo As Long) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
'RF 17/11/99 Indicate which errors (if between 1 and 499) have been hardcoded
'------------------------------------------------------------------------------------------

'***********************************************************************************
'* THIS METHOD IS HERE FOR BACKWARD COMPATIBILITY ONLY.
'* PLEASE DO NOT ADD TO THIS.
'***********************************************************************************
    
    Dim strMsgText As String
    
    Select Case lngMessageNo
        '------------------------------------------------------------------------------------------
        ' err nos 1 - 99
        '------------------------------------------------------------------------------------------
        Case omiga4Err001
            strMsgText = "Invalid number to generate a check digit"
        Case omiga4Err002
            strMsgText = "Request is for invalid number type"
        '------------------------------------------------------------------------------------------
        ' err nos 100 - 499
        '------------------------------------------------------------------------------------------
        Case omiga4Err100
            strMsgText = "Post Code must be entered for Address Search"
        Case omiga4Err101
            strMsgText = "Surname and Forename or Initials required for Search"
        Case omiga4Err102
            strMsgText = "Surname must be entered with Date of Birth"
        Case omiga4Err103
            strMsgText = "Invalid lock type"
        Case omiga4err104
            strMsgText = "Invalid search string"
        Case omiga4Err105
            strMsgText = "Customer Number and Customer Version Number must be specified"
        Case omiga4err106
            strMsgText = "Invalid parameter value: userid is null"
        Case omiga4err107
            strMsgText = "Cannot interpret database error"
        Case omiga4err108
            strMsgText = "Combo group must be specified"
        Case omiga4err109
            strMsgText = "ValueID for combo group must be specified"
        Case omiga4err110
            strMsgText = "PAF Address unsuccessful"
        Case omiga4err111
            strMsgText = "You are not authorised to logon to this unit.  Please contact your System Supervisor."
        Case omiga4err112
            strMsgText = "You cannot be logged on.  Please check your user id. is correct and retype your password before retrying.  If you fail logon again, please contact your System Supervisor."
        Case omiga4err113
            strMsgText = "New password is invalid. Please retype your new password " & _
                "and confirmed password, ensuring that the password conforms with the " & _
                "password guidelines and is not the same as your previous passwords. " & _
                "If your new password fails again, please contact your System Supervisor."
        Case omiga4err114
            strMsgText = "You have insufficient authority to access this function/area. Please contact your System Supervisor."
        Case omiga4err115
            strMsgText = "You have no access assigned to you.  Please contact your System Supervisor."
        Case omiga4err116
            strMsgText = "You are not authorised access outside of your assigned working hours. Please contact your System Supervisor."
        Case omiga4err117
            strMsgText = "Maximum number of access attempts has been exceeded. Please contact your System Supervisor."
        Case omiga4err118
            strMsgText = "Your Password has expired. You must set a new password to continue."
        Case omiga4err119
            strMsgText = "Completeness Check failed for Application"
        Case omiga4err120
            strMsgText = "Insufficient address criteria to perform postal address file (P.A.F.) search"
        Case omiga4err121
            strMsgText = "Quick Quote is not available, remodel via Cost Modelling"
        Case omiga4err122
            strMsgText = "Quick Quote Details - mandatory field missing"
        Case omiga4err123
            strMsgText = "Quick Quote Applicant Details - mandatory field missing"
        Case omiga4err124
            strMsgText = "Quick Quote Outgoings - mandatory field missing"
        Case omiga4err125
            strMsgText = "Quick Quote Personal Debts - mandatory field missing"
        Case omiga4err126
            strMsgText = "Attitude To Borrowing - mandatory field missing."
        Case omiga4err127
            strMsgText = "Missing Applicants for Application"
        Case omiga4Err128
            strMsgText = "Telephone Number and Usage must be entered"
            
        Case omiga4Err130
            strMsgText = "Term of Loan out of range"
        Case omiga4Err131
            strMsgText = "Invalid Amount Requested"
        Case omiga4Err132
            strMsgText = "Further Filtering Criteria invalid or incomplete"
        Case omiga4Err133
            strMsgText = "Customer already linked to package"
        Case omiga4Err134
            strMsgText = "Record for this addresss type already exists"
        Case omiga4Err135
            strMsgText = "Date moved in must be before date moved out"
        Case omiga4err136
            strMsgText = "CompanyID must be specified"
        Case omiga4err137
            strMsgText = "Surname and First Forename must be entered"
        Case omiga4err138
            strMsgText = "Incorrect parameters to search for customer roles"
        Case omiga4err139
            strMsgText = "Customer already linked to application"
        
        Case omiga4err147
            strMsgText = "Mandatory Bank/Credit Card data missing"
        Case omiga4err148
            strMsgText = "Mandatory Life Product data missing"
        Case omiga4err149
            strMsgText = "Mandatory Declined Mortgage data missing"
        Case omiga4err150
            strMsgText = "Mandatory Arrears History data missing"
        Case omiga4err151
            strMsgText = "Mandatory Loan/Liability data missing"
        Case omiga4err152
            strMsgText = "Mandatory CCJ History data missing"
        Case omiga4err153
            strMsgText = "Mandatory Bankruptcy History data missing"
        Case omiga4err154
            strMsgText = "No details entered for additional card holders"
        Case omiga4err155
            strMsgText = "No details entered for others linked to Loan/Liability"
        Case omiga4err156
            strMsgText = "No details entered for others linked to Life Product"
        Case omiga4err157
            strMsgText = "No details entered for others linked to Declined Mortgage"
        Case omiga4err158
            strMsgText = "No details entered for others linked to Arrears History"
        Case omiga4err159
            strMsgText = "No details entered for others linked to CCJ History"
        
        Case omiga4err163
            strMsgText = "Both Amount Requested and Purchase Price are required to calculate LTV"
        Case omiga4err164
            strMsgText = "Customer cannot be younger than minimum age"
        Case omiga4err165
            strMsgText = "Customer cannot be older than maximum age"
        Case omiga4err166
            strMsgText = "First character of surname and first forename must be alpha"
        Case omiga4err167
            strMsgText = "Purchase Price must be entered for One-Off Costs to be calculated"
        Case omiga4err168
            strMsgText = "A customer may not be reordered within, assigned to, or removed from this type of mortgage application"
        Case omiga4err169
            strMsgText = "A customer may not be removed where they are the sole applicant"
        Case omiga4err170
            strMsgText = "Customer Number must be specified"
        Case omiga4err171
            strMsgText = "A customer may not be removed from an ApplicationPackage where they are the sole applicant"
        Case omiga4err172
            strMsgText = "Loan split must be entered for part & part"
        Case omiga4err173
            strMsgText = "An Application must have at least one applicant"
        Case omiga4err174
            strMsgText = "Number of Applicants is greater than the maximum allowed"
        Case omiga4err175
            strMsgText = "Number of Customers is greater than the maximum allowed"
        Case omiga4err176
            strMsgText = "Number of Guarantors is greater than the maximum allowed"

        Case omiga4err182
            strMsgText = "This application is currently being accessed by another user. Do you wish to continue in read-only mode?"
        Case omiga4err183
            strMsgText = "Customer is currently being accessed by another user. Do you wish to continue in read-only mode?"

        Case omiga4err190
            strMsgText = "Customer is already locked to another application. Do you wish to continue in read-only mode?"
        Case omiga4err192
            strMsgText = "Outstanding Loan Amount and Total Monthly Payment Amount required for Further Advance/Transfer of Equity"

        Case omiga4err207
            strMsgText = "Invalid Repayment Type"
        Case omiga4err208
            strMsgText = "The total Loan amount, and/or the loan amounts for the multiple loans, is zero. Cannot issue payments on a zero loan amount."
        Case omiga4err209
            strMsgText = "Charges greater than Loan"
        Case omiga4err210
            strMsgText = "Sum of Charges and Incentive greater than Total Loan"
        Case omiga4err211
            strMsgText = "APR has generated an invalid figure. Please check parameters or retry"
        Case omiga4err212
            strMsgText = "The Mortgage Product contains Upper Band Limits with invalid null or zero"
        Case omiga4err213
            strMsgText = "The Mortgage Product contains Rounding Factors with invalid null or zero"

        Case omiga4err222
            strMsgText = "Calculations Engine could not be Interfaced"
        
        Case omiga4err228
            strMsgText = "Mandatory payment protection sub-quote details missing"
        Case omiga4err229
            strMsgText = "Total cover amount must be divisible by %s"
        Case omiga4err230
            strMsgText = "Total cover amount has been limited to maximum allowed and applicant cover amounts have been reset"
        Case omiga4err231
            strMsgText = "Sum of applicant cover amounts exceeds total payment protection cover"
        Case omiga4err232
            strMsgText = "Payment protection rates cannot be found. Please check that applicants are eligible for the selected product."
        Case omiga4err233
            strMsgText = "You have no current password registered.  Please contact your System Supervisor."
        Case omiga4err234
            strMsgText = "New password and confirmed password must be the same.  Please retype the new and confirmed passwords and retry."
        
        Case omiga4Err262
            strMsgText = "Additional details field must be empty where no additional others are indicated"
        
        '------------------------------------------------------------------------------------------
        ' err nos >=500
        '------------------------------------------------------------------------------------------
        Case omiga4RecordNotFound
            strMsgText = "Record not found"
        Case omiga4XMLParserError
            ' Corresponds to XML DOMDocument.parseError.errorCode <> 0
            strMsgText = "Invalid XML format"
        Case omiga4InvalidKeyString
            ' Null key value found in parameter, e.g. to delete function
            strMsgText = "All keys must be specified for this operation"
        Case omiga4NoAfterImagePresent
            ' The after image is not present for an update request
            strMsgText = "The after image is not present for an update request"
        Case omiga4MissingPrimaryTag
            strMsgText = "XML Request Error: Primary data tag not present"
        Case omiga4CommandFailed
            strMsgText = "Database operation failed"
        Case omiga4InvalidParameter
            strMsgText = "Invalid parameter value"
        Case omiga4DuplicateKey
            strMsgText = "Database operation failed: Duplicate key"
        Case omiga4NoDataForCreate
            strMsgText = "Database operation failed: No Data for Create"
        Case omiga4NoDataForUpdate
            strMsgText = "Database operation failed: No Data for Update"
        Case omiga4ArrayLimitExceeded
            strMsgText = "Array limit exceeded"
        Case omiga4NoRowsAffected
            strMsgText = "No rows were affected by the request"
        Case omiga4InvalidNoOfRows
            strMsgText = "Invalid number of rows: "
        Case omiga4MissingFieldDesc
            strMsgText = "Couldn't a this field element in the class description"
        Case omiga4MissingTableDesc
            strMsgText = "Missing Table Description in XMLClass Definition"
        Case omiga4MissingTypeDesc
            strMsgText = "Missing Type Description Value in XMLClass Definition"
        Case omiga4MissingElementValue
            strMsgText = "No Value for the Element of the Key supplied"
        Case omiga4MissingElement
            strMsgText = "Not ALL the Elements for the Keys specified have been supplied"
        Case omiga4MissingKeyDesc
            strMsgText = "Missing Key Description in XMLClass Definition"
        Case omiga4MissingUpdateSet
            strMsgText = "No Value for Set specified in SQL Update"
        Case omiga4MissingTableName
            strMsgText = "No Table Name Found in XML class Definition"
        Case omiga4MissingXMLTableName
            strMsgText = "No Table Name specified in XML class Definition"
        Case omiga4MissingKey
            strMsgText = "No Key specified in XML class Definition"
        Case omiga4NoRowsAffectedByDeleteAll
            strMsgText = "No rows were affected by the Delete All request"
        Case omiga4InValidKeyValue
            strMsgText = "Invalid Enumerated Type for Key Value"
        Case omiga4InValidDataTypeValue
            strMsgText = "Invalid Enumerated Type Data Type Value"
        Case omiga4InValidKey
            strMsgText = "Invalid Enumerated Type Key"
        Case omiga4NoFieldsFound
            strMsgText = "No fields found in XML class def"
        Case omiga4NoFieldItemFound
            strMsgText = "No field  Item found in XML class def"
        Case omiga4NoFieldItemName
            strMsgText = "No field Item Name in XML class def"
        Case omiga4NoComboTagValue
            strMsgText = "No Combo Tag Value in XML class def"
        Case omiga4InvalidDateTimeFormat
            strMsgText = "Invalid DateTime format specified"
        Case omiga4NoBeforeImagePresent
            ' The before image is not present for an update request
            strMsgText = "The before image is not present for an update request"
       Case omiga4MissingLoadDataEnum
            strMsgText = "An enumeration value has not been specified for Load Data "
       Case omiga4internalerror
            ' Extra error description text should always be used with this message number
            strMsgText = "Internal error"
        '------------------------------------------------------------------------------------------
        ' err nos >=900
        '------------------------------------------------------------------------------------------
        Case omiga4NotImplemented
            strMsgText = "Function not implemented"
        Case omiga4InvalidMessageNo
            strMsgText = "Unspecified message"
       Case Else
            strMsgText = "Unspecified application error"
    End Select
    
    If lngMessageNo >= omiga4Err001 And lngMessageNo < omiga4Err262 Then
        strMsgText = "Hardcoded message: " & strMsgText
    End If
    
    GetMessageText = strMsgText

End Function

Public Sub RaiseError(ByVal strObjectName As String, _
    ByVal strFunctionName As String, lngErrNo As Long, _
    ParamArray strAdditionalOptions())
' header ----------------------------------------------------------------------------------
' description:
'   Raise a VB error
' pass:
' return:   n/a
'------------------------------------------------------------------------------------------
    Dim strSource As String

    Dim strText As String
    Dim strLeftHandSide As String
    Dim strMessage As String
    Dim lngPosition As Long
    Dim intIndex As Integer
    
    strText = Me.GetMessageText(lngErrNo)
    
    ' Find %s and substitute it with the substitution parameters
    lngPosition = InStr(1, strText, "%s", vbTextCompare)
    
    intIndex = 1
    
    Do While lngPosition <> 0 And intIndex <= UBound(strAdditionalOptions)
        strLeftHandSide = Left$(strText, lngPosition - 1)
        strMessage = strMessage & strLeftHandSide
        
        ' Substitute parameter if present
        If Not IsMissing(strAdditionalOptions(intIndex)) Then
            strMessage = strMessage & strAdditionalOptions(intIndex)
        Else
            strMessage = strMessage & "*** MISSING PARAMETER ***"
        End If
        
        strText = Mid$(strText, lngPosition + 2)
        lngPosition = InStr(1, strText, "%s", vbTextCompare)
        intIndex = intIndex + 1
    Loop
    
    strMessage = strMessage & strText
    
    ' If we have additional parameters
    If UBound(strAdditionalOptions) >= 0 Then
        
        ' First parameter is the additional error text. Add it if present
        If Not IsMissing(strAdditionalOptions(0)) Then
            strMessage = strMessage & vbCr & "Details: " & strAdditionalOptions(0)
        End If
        
    End If
            
    strSource = App.Title & "." & strObjectName & "." & strFunctionName
    
    Err.Raise lngErrNo, strSource, strMessage
End Sub

Public Sub LogError(ByVal strObjectName As String, _
    ByVal strFunctionName As String, lngErrNo As Long, _
    Optional ByVal strErrText As String, _
    Optional ByVal eventType As Long = vbLogEventTypeError)
' header ----------------------------------------------------------------------------------
' description:
'   Write an error to the application log
' pass:
' return:   n/a
'------------------------------------------------------------------------------------------
    Dim strMsg As String
    Dim strDesc As String
    Dim strSource As String
    
    If IsApplicationError(lngErrNo) Then
        If strErrText = "" Then
            strDesc = Me.GetMessageText(lngErrNo)
        Else
            strDesc = Me.GetMessageText(lngErrNo) & "; Details: " & strErrText
        End If
    Else
        strDesc = strErrText
    End If
    
    strSource = App.Title & "." & strObjectName & "." & strFunctionName
    
    strMsg = _
        "Error number: " & Str(lngErrNo) & vbCr & _
        "Description: " & strDesc & vbCr & _
        "Source: " & strSource & vbCr
        
    App.LogEvent strMsg, eventType
    
    ' FIXME Remove this once writing to event viewer works correctly
    Debug.Print strMsg

End Sub

Public Function CheckXMLResponse( _
    ByVal vstrXmlReponse As String, Optional ByVal vblnRaiseError As Boolean = False, _
    Optional ByVal vxmlInResponseElement As IXMLDOMElement = Nothing) As Long
' header ----------------------------------------------------------------------------------
' description:
'   takes an xml response and returns the error number. Re raises the error if
'   vblnRaiseError is set to true
' pass:
'   vstrXmlReponse        xml Response data stream containing data to be checked
'   vblnRaiseError        Determines whether to raise the error
'   vxmlInResponseElement xml Element to add the warnings to
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "CheckXMLResponse"
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objResponseElement As IXMLDOMElement
    Dim objErrNumElement As IXMLDOMElement
    Dim objErrSourceElement As IXMLDOMElement
    Dim objErrDescElement As IXMLDOMElement
    Dim objFirstChildElem As IXMLDOMNode
    Dim lngErrorNumber As Long
    
    Dim xmlMessageList As IXMLDOMNodeList
    Dim xmlMessageElem As IXMLDOMElement

    Set objXmlIn = objXmlAssist.load(vstrXmlReponse, TypeName(Me), strFunctionName)

    Set objResponseElement = objXmlIn.getElementsByTagName("RESPONSE").Item(0)
    
    If objResponseElement Is Nothing Then
        Me.RaiseError TypeName(Me), strFunctionName, _
            omiga4MissingPrimaryTag, "Expected RESPONSE tag"
    End If
    
    ' If the error is to be re-raised get the appropriate data out of the response
    ' and raise the error
    If objResponseElement.getAttribute("TYPE") = "WARNING" Then
        If Not vxmlInResponseElement Is Nothing Then
            vxmlInResponseElement.setAttribute "TYPE", "WARNING"
            
            Set objFirstChildElem = vxmlInResponseElement.firstChild
            Set xmlMessageList = objXmlIn.getElementsByTagName("MESSAGE")
    
            ' insert messages at the top of the response
            For Each xmlMessageElem In xmlMessageList
                vxmlInResponseElement.insertBefore xmlMessageElem.cloneNode(True), _
                                                   objFirstChildElem
            Next
        End If
    ElseIf objResponseElement.getAttribute("TYPE") <> "SUCCESS" Then
        Set objErrNumElement = objResponseElement.getElementsByTagName("NUMBER").Item(0)
        lngErrorNumber = CLng(objErrNumElement.Text)
        
        If vblnRaiseError = True Then
            Set objErrSourceElement = objResponseElement.getElementsByTagName("SOURCE").Item(0)
            Set objErrDescElement = objResponseElement.getElementsByTagName("DESCRIPTION").Item(0)
            Err.Raise lngErrorNumber, objErrSourceElement.Text, objErrDescElement.Text
        End If
    End If

    CheckXMLResponse = lngErrorNumber
    
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objResponseElement = Nothing
    Set objErrNumElement = Nothing
    Set objErrSourceElement = Nothing
    Set objErrDescElement = Nothing
    Set objFirstChildElem = Nothing

    
End Function

Public Function IsWarning(Optional ByVal lngErrNo As Long = oeUnspecifiedError) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if error is a warning.
' pass:
'   lngErrNo
'       Error number to test. This is an optional parameter. It defaults to
'       oeUnspecifiedError, in which case the Number property of the VB Err object is used.
' return:
'------------------------------------------------------------------------------------------
    
    Dim strErrorType As String
    Dim lngOmigaErrorNo As Long
    Dim blnIsWarning As Boolean
    
    blnIsWarning = False
    
    If IsApplicationError(Err.Number) Then
        lngOmigaErrorNo = GetOmigaErrorNumber(Err.Number)
        strErrorType = GetMessageDetails(lngOmigaErrorNo, "MESSAGETYPE")
        
        If StrComp(strErrorType, "Warning", vbTextCompare) = 0 Or _
           StrComp(strErrorType, "Prompt", vbTextCompare) = 0 Then
            blnIsWarning = True
        End If
    End If
    
    IsWarning = blnIsWarning
    
End Function

Public Function IsApplicationError(Optional ByVal lngErrNo As Long = oeUnspecifiedError) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Check if error is an application error.
' pass:
'   lngErrNo
'       Error number to test. This is an optional parameter. It defaults to
'       oeUnspecifiedError, in which case the Number property of the VB Err object is used.
' return:
'------------------------------------------------------------------------------------------
    Dim lngOmigaErrorNo As Long
    Dim blnIsApplicationError As Boolean
    
    blnIsApplicationError = False
    
    If lngErrNo = oeUnspecifiedError Then
        lngErrNo = Err.Number
    End If
        
'    If lngErrNo <> 0 Then
'        lngOmigaErrorNo = GetOmigaErrorNumber(lngErrNo)
'
'        ' AS 23/11/99 For the error to be an Omiga4 Application error it must also
'        ' not be in the ADO error number range see ADODB.ErrorValueEnum
'        If (lngOmigaErrorNo > 0 And _
'            lngOmigaErrorNo <= clngMAX_ERROR_NO And _
'            (lngOmigaErrorNo < clngADO_START_ERROR_NO Or _
'            lngOmigaErrorNo > clngADO_END_ERROR_NO)) Then
'
'            blnIsApplicationError = True
'
'        End If
'    End If
    
    'DB BM0483 - Re-worked to subtract 512 from the error number.
    If lngErrNo <> 0 Then
        lngOmigaErrorNo = GetOmigaErrorNumber(lngErrNo)

        If (lngOmigaErrorNo > 0 And _
            lngOmigaErrorNo <= clngMAX_ERROR_NO) Then
            lngErrNo = lngErrNo - vbObjectError
            If (lngErrNo < clngADO_START_ERROR_NO Or _
                lngOmigaErrorNo > clngADO_END_ERROR_NO) Then
     
                blnIsApplicationError = True
     
            End If
        End If
    End If
    'DB End
    
    IsApplicationError = blnIsApplicationError
    
End Function


Public Function IsSystemError(Optional ByVal lngErrNo As Long = oeUnspecifiedError) As Boolean
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    If lngErrNo = oeUnspecifiedError And Err.Number = 0 Then
        IsSystemError = False
    Else
        IsSystemError = Not IsApplicationError(lngErrNo)
    End If
End Function

Public Function FormatMessageNode() As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    
    Dim xmlMessageDoc As New FreeThreadedDOMDocument40
    xmlMessageDoc.validateOnParse = False
    xmlMessageDoc.setProperty "NewParser", True
    Dim xmlMessageElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim lngOmigaErrorNo As Long
        
    Set xmlMessageElement = xmlMessageDoc.createElement("MESSAGE")
    xmlMessageDoc.appendChild xmlMessageElement
    Set xmlElement = xmlMessageDoc.createElement("TEXT")
    xmlElement.Text = Err.Description
    xmlMessageElement.appendChild xmlElement
    
    lngOmigaErrorNo = GetOmigaErrorNumber(Err.Number)
    Set xmlElement = xmlMessageDoc.createElement("TYPE")
    xmlElement.Text = "Warning"
    xmlMessageElement.appendChild xmlElement
    
    FormatMessageNode = xmlMessageDoc.xml
    
    Set xmlMessageDoc = Nothing
    Set xmlMessageElement = Nothing
    Set xmlElement = Nothing
        
End Function

Public Sub SaveErr()
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    mvarNumber = Err.Number
    mvarSource = Err.Source
    mvarDescription = Err.Description
    mvarHelpContext = Err.HelpContext
    mvarHelpFile = Err.HelpFile
End Sub

Public Sub ReRaise()
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Err.Raise mvarNumber, mvarSource, mvarDescription, mvarHelpFile, mvarHelpContext
End Sub

Public Function LastErr() As Long
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    LastErr = mvarNumber
End Function

Public Sub AddToErrSource(ByVal vstrDescription As String)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Err.Source = vstrDescription & ", " & Err.Source
End Sub

Private Function GetErrorType(ByVal lngErrNo As Long)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------

    Select Case lngErrNo
        Case omiga4err230
            GetErrorType = "Warning"
    End Select

End Function

Public Sub AddWarning(ByVal xmlResponse As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------

    Dim xmlMessageElement As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlFirstChild As IXMLDOMNode
    Dim lngOmigaErrorNo As Long
    
    xmlResponse.setAttribute "TYPE", "WARNING"
    
    Set xmlFirstChild = xmlResponse.firstChild
    
    Set xmlMessageElement = xmlResponse.ownerDocument.createElement("MESSAGE")
    
    xmlResponse.insertBefore xmlMessageElement, xmlFirstChild
        
    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETEXT")
    xmlElement.Text = Err.Description
    xmlMessageElement.appendChild xmlElement
    
    lngOmigaErrorNo = GetOmigaErrorNumber(Err.Number)
    Set xmlElement = xmlResponse.ownerDocument.createElement("MESSAGETYPE")
    xmlElement.Text = GetMessageDetails(lngOmigaErrorNo, "MESSAGETYPE")
    xmlMessageElement.appendChild xmlElement
    
    Set xmlMessageElement = Nothing
    Set xmlElement = Nothing
    Set xmlFirstChild = Nothing
    
End Sub

Private Function GetMessageDetails(ByVal lngErrorNo As Long, ByVal strFieldName) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "GetMessageDetails"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objMessageDO As MessageDO
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlResponseDoc As FreeThreadedDOMDocument40
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    
    Dim strText As String
    Dim strType As String
    Dim blnIsHardCoded As Boolean

    blnIsHardCoded = GetHardcodedDetails(lngErrorNo, strText, strType)
    
    ' Only go to the database for details if it is not a hard coded message
    If blnIsHardCoded = False Then
        If objContext Is Nothing Then
            Set objMessageDO = New MessageDO
        Else
            Set objMessageDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".MessageDO")
        End If
    
        Set xmlTableElem = xmlRequestDoc.createElement("MESSAGE")
        xmlRequestDoc.appendChild xmlTableElem
    
        Set xmlElement = xmlRequestDoc.createElement("MESSAGENUMBER")
        xmlTableElem.appendChild xmlElement
        xmlElement.Text = CStr(lngErrorNo)
    
        Dim strResponse As String
        strResponse = objMessageDO.GetMessageDetails(xmlRequestDoc.xml)
        Set xmlResponseDoc = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)
        strText = objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "MESSAGETEXT")
        strType = objXmlAssist.GetTagValue(xmlResponseDoc.documentElement, "MESSAGETYPE")
    End If
    
    If strFieldName = "MESSAGETEXT" Then
        GetMessageDetails = strText
    ElseIf strFieldName = "MESSAGETYPE" Then
        GetMessageDetails = strType
    End If
    
    Set objMessageDO = Nothing
    Set xmlRequestDoc = Nothing
    Set xmlResponseDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing

End Function

Public Sub ThrowError(ByVal strObjectName As String, _
                      ByVal strFunctionName As String, lngOmigaErrorNo As Long, _
                      ParamArray strAdditionalOptions())
' header ----------------------------------------------------------------------------------
' description:  Raise a VB error
' pass:
' return:   n/a
'------------------------------------------------------------------------------------------
    Dim strSource As String
    Dim strText As String
    Dim strLeftHandSide As String
    Dim strMessage As String
    Dim lngPosition As Long
    Dim intIndex As Integer
    Dim lngErrNo As Long
        
    lngErrNo = vbObjectError + 512 + lngOmigaErrorNo
    
    strText = GetMessageDetails(lngOmigaErrorNo, "MESSAGETEXT")
    
    ' Find %s and substitute it with the substitution parameters
    lngPosition = InStr(1, strText, "%s", vbTextCompare)
    
    intIndex = 1
    
    Do While lngPosition <> 0 And intIndex <= UBound(strAdditionalOptions)
        strLeftHandSide = Left$(strText, lngPosition - 1)
        strMessage = strMessage & strLeftHandSide
        
        ' Substitute parameter if present
        If Not IsMissing(strAdditionalOptions(intIndex)) Then
            strMessage = strMessage & strAdditionalOptions(intIndex)
        Else
            strMessage = strMessage & "*** MISSING PARAMETER ***"
        End If
        
        strText = Mid$(strText, lngPosition + 2)
        lngPosition = InStr(1, strText, "%s", vbTextCompare)
        intIndex = intIndex + 1
    Loop
    
    strMessage = strMessage & strText
    
    ' If we have additional parameters
    If UBound(strAdditionalOptions) >= 0 Then
        
        ' First parameter is the additional error text. Add it if present
        If Not IsMissing(strAdditionalOptions(0)) Then
            strMessage = strMessage & vbCr & "Details: " & strAdditionalOptions(0)
        End If
        
    End If
            
    strSource = App.Title & "." & strObjectName & "." & strFunctionName
    
    Err.Raise lngErrNo, strSource, strMessage
End Sub

Public Function GetOmigaErrorNumber(ByVal lngErrorNo As Long) As Long
' header ----------------------------------------------------------------------------------
' description: Converts an  error number to an omiga number. As vbObjectError and 512
'              are subtracted when raising an omiga error, system numbers will end up
'              much larger than proper omiga error numbers
' pass:
'------------------------------------------------------------------------------------------
    GetOmigaErrorNumber = lngErrorNo - vbObjectError - 512
End Function

Private Function GetHardcodedDetails(ByVal vlngErrorNo As Long, _
                                     ByRef rstrText, _
                                     ByRef rtrType) As Boolean
' header ----------------------------------------------------------------------------------
' description: Check if details are hardcoded and return them if they are
' pass:
'------------------------------------------------------------------------------------------
    Dim blnFound As Boolean
    
    blnFound = True
    rstrText = ""
    rtrType = ""
    
    Select Case vlngErrorNo
        Case 556
             rstrText = "Unable to establish database connection"
             rtrType = "Error"
        Case 901
            rstrText = "Error message not found"
            rtrType = "Error"
        Case Else
            blnFound = False
                
    End Select
                                         
    GetHardcodedDetails = blnFound

End Function

Public Function CreateErrorResponseEx() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'------------------------------------------------------------------------------------------
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlReponseElem As IXMLDOMElement
    Dim xmlErrorElem As IXMLDOMElement
    Dim xmlDescriptionElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlReponseElem = xmlDoc.createElement("RESPONSE")
    xmlDoc.appendChild xmlReponseElem
    
    Set xmlErrorElem = xmlDoc.createElement("ERROR")
    xmlReponseElem.appendChild xmlErrorElem
    
    Set xmlElement = xmlDoc.createElement("NUMBER")
    xmlElement.Text = Err.Number
    xmlErrorElem.appendChild xmlElement
    
    Set xmlElement = xmlDoc.createElement("SOURCE")
    xmlElement.Text = Err.Source
    xmlErrorElem.appendChild xmlElement
    
    Set xmlDescriptionElem = xmlDoc.createElement("DESCRIPTION")
    xmlDescriptionElem.Text = Err.Description
    xmlErrorElem.appendChild xmlDescriptionElem
    
    Set xmlElement = xmlDoc.createElement("VERSION")
    xmlElement.Text = App.Comments
    xmlErrorElem.appendChild xmlElement
    
    If Me.IsApplicationError = True Then
        xmlReponseElem.setAttribute "TYPE", "APPERR"
    
        If Len(xmlDescriptionElem.Text) = 0 Then
            xmlDescriptionElem.Text = Me.GetMessageText(Err.Number)
        End If
    Else
        xmlReponseElem.setAttribute "TYPE", "SYSERR"
    End If
        
    Set CreateErrorResponseEx = xmlReponseElem.cloneNode(True)
    
    Set xmlDoc = Nothing
    Set xmlReponseElem = Nothing
    Set xmlErrorElem = Nothing
    Set xmlDescriptionElem = Nothing
    Set xmlElement = Nothing

End Function

Public Function CheckResponse(ByVal vxmlResponseToCheck As IXMLDOMElement, _
                              ByVal vxmlResponseToAddTo As IXMLDOMElement, _
                              Optional ByVal vblnRaiseError As Boolean = False) As Long
' header ----------------------------------------------------------------------------------
' description:
'   takes an xml response and returns the error number. Re raises the error if
'   vblnRaiseError is set to true
' pass:
'   vxmlResponseToCheck   xml Response to check for errors
'   vxmlResponseToAddTo   xml Element to add the warnings to
'   vblnRaiseError        Determines whether to raise the error
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "CheckResponse"
    
    Dim objXmlAssist As New XMLAssist
    Dim xmlErrNumElement As IXMLDOMElement
    Dim xmlErrSourceElement As IXMLDOMElement
    Dim xmlErrDescElement As IXMLDOMElement
    Dim xmlFirstChildElem As IXMLDOMNode
    Dim lngErrorNumber As Long
    
    Dim xmlMessageList As IXMLDOMNodeList
    Dim xmlMessageElem As IXMLDOMElement
    
    If vxmlResponseToCheck Is Nothing Or vxmlResponseToAddTo Is Nothing Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Response to Check or Response to add to missing"
    End If
    
    If vxmlResponseToCheck.nodeName <> "RESPONSE" Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "RESPONSE must be top level tag"
    End If
    
    ' If the error is to be re-raised get the appropriate data out of the response
    ' and raise the error
    If vxmlResponseToCheck.getAttribute("TYPE") = "WARNING" Then
        vxmlResponseToAddTo.setAttribute "TYPE", "WARNING"
        
        Set xmlFirstChildElem = vxmlResponseToCheck.firstChild
        Set xmlMessageList = vxmlResponseToCheck.getElementsByTagName("MESSAGE")

        ' insert messages at the top of the response
        For Each xmlMessageElem In xmlMessageList
            vxmlResponseToAddTo.appendChild xmlMessageElem.cloneNode(True)
        Next
    ElseIf vxmlResponseToCheck.getAttribute("TYPE") <> "SUCCESS" Then
        Set xmlErrNumElement = vxmlResponseToCheck.getElementsByTagName("NUMBER").Item(0)
        lngErrorNumber = CLng(xmlErrNumElement.Text)
        
        If vblnRaiseError = True Then
            Set xmlErrSourceElement = vxmlResponseToCheck.getElementsByTagName("SOURCE").Item(0)
            Set xmlErrDescElement = vxmlResponseToCheck.getElementsByTagName("DESCRIPTION").Item(0)
            Err.Raise lngErrorNumber, xmlErrSourceElement.Text, xmlErrDescElement.Text
        End If
    End If

    CheckResponse = lngErrorNumber
    
    Set objXmlAssist = Nothing
    Set xmlErrNumElement = Nothing
    Set xmlErrSourceElement = Nothing
    Set xmlErrDescElement = Nothing
    Set xmlFirstChildElem = Nothing
    
End Function

Public Function ReRaiseResponseError(ByVal vxmlResponseToCheck As IXMLDOMElement) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Takes an xml response and returns the error number. re-raises the error, if found
' pass:
'   vxmlResponseToCheck   xml Response to check for errors
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "ReRaiseResponseError"
    
    Dim lngErrorNumber As Long
    
    If vxmlResponseToCheck Is Nothing Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Response to Check missing"
    End If
    
    If vxmlResponseToCheck.nodeName <> "RESPONSE" Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "RESPONSE must be top level tag"
    End If
    
    ' Check for error and re-raise, if found
    Dim strResponseType As String, strErrorSource As String, strErrDesc As String
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNode = vxmlResponseToCheck.getAttributeNode("TYPE")
    If Not xmlNode Is Nothing Then
    
        strResponseType = vxmlResponseToCheck.getAttribute("TYPE")
        If strResponseType <> "SUCCESS" And strResponseType <> "WARNING" Then
            lngErrorNumber = vxmlResponseToCheck.getElementsByTagName("NUMBER").Item(0).Text
            strErrorSource = vxmlResponseToCheck.getElementsByTagName("SOURCE").Item(0).Text
            strErrDesc = vxmlResponseToCheck.getElementsByTagName("DESCRIPTION").Item(0).Text
                
            Err.Raise lngErrorNumber, strErrorSource, strErrDesc
        End If
    
    End If
    

    ReRaiseResponseError = lngErrorNumber
    
End Function

Public Function GetErrorNumberFromResponse(ByVal vxmlResponseToCheck As IXMLDOMElement, _
                                           Optional ByVal blnOmigaError As Boolean = False) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Takes an xml response and returns the error number. If blnOmigaError is True, then finds
'   corresponding Omiga Error number and returns it
' pass:
'   vxmlResponseToCheck   xml Response to check for errors
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "GetErrorNumberFromResponse"
    
    Dim lngErrorNumber As Long
    
    If vxmlResponseToCheck Is Nothing Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Response to Check missing"
    End If
    
    If vxmlResponseToCheck.nodeName <> "RESPONSE" Then
        Me.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "RESPONSE must be top level tag"
    End If
    
    ' Check for error and re-raise, if found
    Dim strResponseType As String, strErrorSource As String, strErrDesc As String
    
    strResponseType = vxmlResponseToCheck.getAttribute("TYPE")
    If strResponseType <> "SUCCESS" And strResponseType <> "WARNING" Then
        lngErrorNumber = vxmlResponseToCheck.getElementsByTagName("NUMBER").Item(0).Text
        If blnOmigaError Then
            lngErrorNumber = GetOmigaErrorNumber(lngErrorNumber)
        End If
    End If

    GetErrorNumberFromResponse = lngErrorNumber

End Function

