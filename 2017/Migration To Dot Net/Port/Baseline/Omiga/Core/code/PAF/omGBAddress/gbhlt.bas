Attribute VB_Name = "HLT1"
' ******************************************************************************
' *                             GB MAILING SYSTEMS Ltd.                        *
' *                             =======================                        *
' *                       Copyright 1995 All rights reserved                   *
' *                                                                            *
' * Function                hlt1.bas                                           *
' *                         object - general, procedure - declarations         *
' * Purpose:                Declares GB HLT functions from DLL                 *
' *                         Defines several global variables                   *
' *                                                                            *
' *                               ** IMPORTANT **                              *
' *             This code contains calls which requires the                    *
' *             Postcoder level of data, and the gbdlpost Tools DLL            *
' *             For premise level data, you will need to to remove             *
' *             the GBPostcodeAddress functionality, and for thorofare         *
' *             level data you will also need to remove the SELBUI             *
' *             form and functionality                                         *
' *                                                                            *
' * Author:                 Duncan Edmonstone                                  *
' * Date:                   18/10/95                                           *
' * Version:                1.0                                                *
' *                                                                            *
' *                                                                            *
' *                              Modification History                          *
' *                              ====================                          *
' * Date       Intials             Comments                                    *
' *                                                                            *
' * 18/10/95   DE                  Module Created                              *
' *                                                                            *
' ******************************************************************************

Option Explicit

' Declare all GB functions required from DLL
' HLT Primary Addressing Functions
' --------------------------------
Declare Function GBOpen Lib "gbhltv32.dll" Alias "_GBOpen@0" () As Integer
Declare Function GBClose Lib "gbhltv32.dll" Alias "_GBClose@0" () As Integer
Declare Function GBGetAddress Lib "gbhltv32.dll" Alias "_GBGetAddress@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetAddressFromDPS Lib "gbhltv32.dll" Alias "_GBGetAddressFromDPS@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetAddressFromAKey Lib "gbhltv32.dll" Alias "_GBGetAddressFromAKey@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetAKeyFromDPS Lib "gbhltv32.dll" Alias "_GBGetAKeyFromDPS@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetError Lib "gbhltv32.dll" Alias "_GBGetError@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBPostcodeAddress Lib "gbhltv32.dll" Alias "_GBPostcodeAddress@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetNext Lib "gbhltv32.dll" Alias "_GBGetNext@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetPrevious Lib "gbhltv32.dll" Alias "_GBGetPrevious@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBUpdatePassword Lib "gbhltv32.dll" Alias "_GBUpdatePassword@8" (ByVal szUserBuff As String, ByVal iBuffLen As Integer) As Integer

' HLT GetInformation() type functions
' -----------------------------------
Declare Function GBGetAbbrCounty Lib "gbhltv32.dll" Alias "_GBGetAbbrCounty@8" (ByVal szAbbrCounty As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetAddressKey Lib "gbhltv32.dll" Alias "_GBGetAddressKey@8" (ByVal szAddressKey As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetBuilding Lib "gbhltv32.dll" Alias "_GBGetBuilding@8" (ByVal szBuilding As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetCountry Lib "gbhltv32.dll" Alias "_GBGetCountry@8" (ByVal szCountry As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetCounty Lib "gbhltv32.dll" Alias "_GBGetCounty@8" (ByVal szCounty As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetCountyRequired Lib "gbhltv32.dll" Alias "_GBGetCountyRequired@8" (ByVal szCntyReq As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDefine Lib "gbhltv32.dll" Alias "_GBGetDefine@8" (ByVal szDefine As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDepartment Lib "gbhltv32.dll" Alias "_GBGetDepartment@8" (ByVal szDepartment As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDepLocality Lib "gbhltv32.dll" Alias "_GBGetDepLocality@8" (ByVal szDLocality As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDepThorofare Lib "gbhltv32.dll" Alias "_GBGetDepThorofare@8" (ByVal szDThorofare As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDHACode Lib "gbhltv32.dll" Alias "_GBGetDHACode@8" (ByVal szDHACode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDoubleLocality Lib "gbhltv32.dll" Alias "_GBGetDoubleLocality@8" (ByVal szDDLocality As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetDPS Lib "gbhltv32.dll" Alias "_GBGetDPS@8" (ByVal szDPS As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetEasting Lib "gbhltv32.dll" Alias "_GBGetEasting@8" (ByVal szEasting As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetEstGridRef Lib "gbhltv32.dll" Alias "_GBGetEstGridRef@8" (ByVal szEstGridRef As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetHLTVersion Lib "gbhltv32.dll" Alias "_GBGetHLTVersion@8" (ByVal szHltVer As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetLocality Lib "gbhltv32.dll" Alias "_GBGetLocality@8" (ByVal szLocality As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetLocalityDetails Lib "gbhltv32.dll" Alias "_GBGetLocalityDetails@8" (ByVal szFullLoc As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetLAAreaCode Lib "gbhltv32.dll" Alias "_GBGetLAAreaCode@8" (ByVal szLAAreaCode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetLACode Lib "gbhltv32.dll" Alias "_GBGetLACode@8" (ByVal szLACode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetLAWardCode Lib "gbhltv32.dll" Alias "_GBGetLAWardCode@8" (ByVal szLAWardCode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetMailSort Lib "gbhltv32.dll" Alias "_GBGetMailSort@8" (ByVal szMailSort As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetMosaic Lib "gbhltv32.dll" Alias "_GBGetMosaic@8" (ByVal szMosaic As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetNHSCode Lib "gbhltv32.dll" Alias "_GBGetNHSCode@8" (ByVal szNHSCode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetNHSRegion Lib "gbhltv32.dll" Alias "_GBGetNHSRegion@8" (ByVal szNHSRegion As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetNorthing Lib "gbhltv32.dll" Alias "_GBGetNorthing@8" (ByVal szNorthing As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetOldMailSort Lib "gbhltv32.dll" Alias "_GBGetOldMailSort@8" (ByVal szOldMailSort As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetOrganisation Lib "gbhltv32.dll" Alias "_GBGetOrganisation@8" (ByVal szOrganisation As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetPAFVersion Lib "gbhltv32.dll" Alias "_GBGetPAFVersion@8" (ByVal szPafVer As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetPOBox Lib "gbhltv32.dll" Alias "_GBGetPOBox@8" (ByVal szPoBox As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetPostcode Lib "gbhltv32.dll" Alias "_GBGetPostcode@8" (ByVal szPostcode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetSecurityCode Lib "gbhltv32.dll" Alias "_GBGetSecurityCode@8" (ByVal szSecCode As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetSubBuilding Lib "gbhltv32.dll" Alias "_GBGetSubBuilding@8" (ByVal szSubBuilding As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetThorofare Lib "gbhltv32.dll" Alias "_GBGetThorofare@8" (ByVal szThorofare As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetThorofareDetails Lib "gbhltv32.dll" Alias "_GBGetThorofareDetails@8" (ByVal szFullTho As String, ByVal iBuffLen As Integer) As Integer
Declare Function GBGetToolsVersion Lib "gbhltv32.dll" Alias "_GBGetToolsVersion@8" (ByVal szToolsVer As String, ByVal iBuffLen As Integer) As Integer
'
' HLT Set functions
' -----------------
Declare Function GBSetAddressCase Lib "gbhltv32.dll" Alias "_GBSetAddressCase@4" (ByVal szOption As String) As Integer
Declare Function GBSetRangeIndicator Lib "gbhltv32.dll" Alias "_GBSetRangeIndicator@4" (ByVal szOption As String) As Integer
Declare Function GBSetBuildingRange Lib "gbhltv32.dll" Alias "_GBSetBuildingRange@4" (ByVal szOption As String) As Integer
Declare Function GBSetDataFileType Lib "gbhltv32.dll" Alias "_GBSetDataFileType@4" (ByVal szOption As String) As Integer
Declare Function GBSetDllLevel Lib "gbhltv32.dll" Alias "_GBSetDllLevel@4" (ByVal szOption As String) As Integer
Declare Function GBSetInAddressFmt Lib "gbhltv32.dll" Alias "_GBSetInAddressFmt@4" (ByVal szOption As String) As Integer
Declare Function GBSetInDelimiter Lib "gbhltv32.dll" Alias "_GBSetInDelimiter@4" (ByVal szOption As String) As Integer
Declare Function GBSetInFmtStyle Lib "gbhltv32.dll" Alias "_GBSetInFmtStyle@4" (ByVal szOption As String) As Integer
Declare Function GBSetListSorted Lib "gbhltv32.dll" Alias "_GBSetListSorted@4" (ByVal szOption As String) As Integer
Declare Function GBSetMaxElements Lib "gbhltv32.dll" Alias "_GBSetMaxElements@4" (ByVal szOption As String) As Integer
Declare Function GBSetNoStreetInd Lib "gbhltv32.dll" Alias "_GBSetNoStreetInd@4" (ByVal szOption As String) As Integer
Declare Function GBSetOutDelimiter Lib "gbhltv32.dll" Alias "_GBSetOutDelimiter@4" (ByVal szOption As String) As Integer
Declare Function GBSetOutFmtStyle Lib "gbhltv32.dll" Alias "_GBSetOutFmtStyle@4" (ByVal szOption As String) As Integer
Declare Function GBSetOutTerminator Lib "gbhltv32.dll" Alias "_GBSetOutTerminator@4" (ByVal szOption As String) As Integer
Declare Function GBSetOutAddressFmt Lib "gbhltv32.dll" Alias "_GBSetOutAddressFmt@4" (ByVal szOption As String) As Integer
Declare Function GBSetPostcodeFmt Lib "gbhltv32.dll" Alias "_GBSetPostcodeFmt@4" (ByVal szOption As String) As Integer
Declare Function GBSetReturnCounty Lib "gbhltv32.dll" Alias "_GBSetReturnCounty@4" (ByVal szOption As String) As Integer
Declare Function GBSetUnwantedElements Lib "gbhltv32.dll" Alias "_GBSetUnwantedElements@4" (ByVal szOption As String) As Integer


Global Const User_buf_size = 1024             ' The size of the user buffer
Global Const Error_buf_size = 200             ' The size of the error buffer
Global szUser_buf As String * User_buf_size   ' User buffer passed to HLT functions
Global szError_buf As String * Error_buf_size ' Buffer passed to GBGetError
Global szCurrent_tho As String                ' The current thorofare selection
Global szCurrent_bui As String                ' The current building selection
Global NL As String * 2                       ' String to hold carriage return / line feed -
                                              ' The Value is assigned in RA.Form.Load

Global bMore As Integer                       ' Boolean - True = more data to come
                                              '           False = no more data

