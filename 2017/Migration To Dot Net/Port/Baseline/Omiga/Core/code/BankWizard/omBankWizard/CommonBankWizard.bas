Attribute VB_Name = "Common"
Option Explicit
' Long Integer variable 'Result' flag definitions for
' UKCheckDetails
Global Const BWNOBBD As Long = 128        '0x00000080
Global Const BWNOMEM As Long = 8192       '0x00002000
Global Const BWNOMOD As Long = 16384      '0x00004000
Global Const BWNOSUB As Long = 32768      '0x00008000
Global Const BWNOOFS As Long = 65536      '0x00010000
Global Const BWNOLIC As Long = 131072     '0x00020000
Global Const BWEXLIC As Long = 262144     '0x00040000
Global Const BWNOUNQ As Long = 524288     '0x00080000
Global Const BW30DAY As Long = 1048576    '0x00100000
Global Const BWNOTRN As Long = 2097152    '0x00200000
Global Const BWSCNIR As Long = 1          '0x00000001
Global Const BWSCNBR As Long = 2          '0x00000002
Global Const BWMODFL As Long = 4          '0x00000004
Global Const BWBYFOR As Long = 8          '0x00000008
Global Const BWNOINI As Long = 16         '0x00000010
Global Const BWBADAN As Long = 32         '0x00000020
Global Const BWBADSC As Long = 64         '0x00000040
Global Const BWTXDIS As Long = 256        '0x00000100
Global Const BWXSCRB As Long = 512        '0x00000200
Global Const BWDBDIS As Long = 1024       '0x00000400
Global Const BWROLLN As Long = 2048       '0x00000800
Global Const BWCRDIS As Long = 4096       '0x00001000
Global Const BW2MONT As Long = 4194304    '0x00400000
Global Const BWCLOSE As Long = 8388608    '0x00800000
Global Const BWCLOSR As Long = 16777216   '0x01000000
Global Const BW2MONR As Long = 33554432   '0x02000000
Global Const BWCUDIS As Long = 67108864   '0x04000000
Global Const BWPRDIS As Long = 134217728  '0x08000000
Global Const BWBSDIS As Long = 268435456  '0x10000000
Global Const BWDVDIS As Long = 536870912  '0x20000000
Global Const BWAUDIS As Long = 1073741824 '0x40000000
' Long Integer variable 'Result' flag definitions for
' UKUpdateTables
Global Const BWNOSET As Long = 1          '0x00000001
Global Const BWINCOM As Long = 2          '0x00000002
Global Const BWNOTAB As Long = 8          '0x00000004
' Long Integer variable 'Status' flag definitions for
' UKGetStatus
Global Const LONG_TAB  As Long = 1
Global Const SHORT_TAB As Long = 2
' Input flag definitions
Global Const MIXEDCASE       As Long = 1
Global Const BRCHECK         As Long = 2
Global Const BRDETAILS       As Long = 4
Declare Function UKCheckDetails Lib "bw2uk" Alias "_WUKCheckDetails@8" _
    (UKDetails As UKDetailsType, Index As Long) As Long
Declare Function UKGetNextSubBranch Lib "bw2uk" Alias "_WUKGetNextSubBranch@8" _
    (ByVal Branch As String, Index As Long) As Long
Declare Function UKGetGeneral Lib "bw2uk" Alias "_WUKGetGeneral@8" _
    (UKGeneral As UKGeneralType, Index As Long) As Long
Declare Function UKGetBACS Lib "bw2uk" Alias "_WUKGetBACS@8" _
    (UKBacs As UKBacsType, Index As Long) As Long
Declare Function UKGetCHAPS Lib "bw2uk" Alias "_WUKGetCHAPS@8" _
    (UKChaps As UKChapsType, Index As Long) As Long
Declare Function UKGetCCCC Lib "bw2uk" Alias "_WUKGetCCCC@8" _
    (UKCccc As UKCcccType, Index As Long) As Long
Declare Function UKGetPrint Lib "bw2uk" Alias "_WUKGetPrint@8" _
    (UKPrint As UKPrintType, Index As Long) As Long
Declare Function UKGetStatus Lib "bw2uk" Alias "_WUKGetStatus@0" () As Long
Declare Function UKGetVersion Lib "bw2uk" Alias "_WUKGetVersion@4" _
    (ByVal Version As String) As Long
Declare Function UKUpdateTables Lib "bw2uk" Alias "_WUKUpdateTables@8" _
    (TableSet As Long, Result As Long) As Long
Declare Function UKSearch Lib "bw2uk" Alias "_WUKSearch@28" _
    (ByVal String1 As String, ByVal String2 As String, ByVal String3 As String, _
    SortCodeArray As Long, SuffixArray As Integer, _
    ByVal NoReturn As Long, ByVal StartPos As Long) As Long
' Type structure for UKCheckDetails function call
Type UKDetailsType
    SortCode           As String
    SortCodeLen        As Long
    AccountNo          As String
    AccountNoLen       As Long
    AccCode            As String
    AccCodeLen         As Long
    Result             As Long
    InFlags            As Long
    SubBranchSuffix    As String
    SubBranchSuffixLen As Long
    BranchTitle        As String
    BranchTitleLen     As Long
    ShortName          As String
    ShortNameLen       As Long
    DeletedDate        As String
    DeletedDateLen     As Long
    BACSStatus         As String
    BACSStatusLen      As Long
    DateClosedBACS     As String
    DateClosedBACSLen  As Long
    RedirFlag          As String
    RedirFlagLen       As Long
    RedirSC            As String
    RedirSCLen         As Long
    DR                 As String
    DRLen              As Long
    CR                 As String
    CRLen              As Long
    CU                 As String
    CULen              As Long
    PR                 As String
    PRLen              As Long
    BS                 As String
    BSLen              As Long
    DV                 As String
    DVLen              As Long
    AU                 As String
    AULen              As Long
    SterlingStatus     As String
    SterlingStatusLen  As Long
    EuroStatus         As String
    EuroStatusLen      As Long
    BranchTypeInd      As String
    BranchTypeIndLen   As Long
    BranchNamePlace    As String
    BranchNamePlaceLen As Long
    Address1           As String
    Address1Len        As Long
    Address2           As String
    Address2Len        As Long
    Address3           As String
    Address3Len        As Long
    Address4           As String
    Address4Len        As Long
    Town               As String
    TownLen            As Long
    County             As String
    CountyLen          As Long
    PostCodeMajor      As String
    PostCodeMajorLen   As Long
    PostCodeMinor      As String
    PostCodeMinorLen   As Long
    TelAreaCode        As String
    TelAreaCodeLen     As Long
    TelNumber          As String
    TelNumberLen       As Long
    FaxAreaCode        As String
    FaxAreaCodeLen     As Long
    FaxNumber          As String
    FaxNumberLen       As Long
End Type
' Type structure for UKGetGeneral function call
Type UKGeneralType
    SortCode As String
    SortCodeLen As Long
    BICBank As String
    BICBankLen As Long
    BICBranch As String
    BICBranchLen As Long
    SubBranchSuffix As String
    SubBranchSuffixLen As Long
    BranchTitle As String
    BranchTitleLen As Long
    ShortName As String
    ShortNameLen As Long
    FullName1 As String
    FullName1Len As Long
    FullName2 As String
    FullName2Len As Long
    BankCodeOB As String
    BankCodeOBLen As Long
    NCBCountryCode As String
    NCBCountryCodeLen As Long
    Supervisory As String
    SupervisoryLen As Long
    DelDate As String
    DelDateLen As Long
End Type
' Type structure for UKGetBACS function call
Type UKBacsType
    BACSStatus As String
    BACSStatusLen As Long
    BACSDateLastChange As String
    BACSDateLastChangeLen As Long
    DateClosedBACS As String
    DateClosedBACSLen As Long
    RedirFlag As String
    RedirFlagLen As Long
    RedirSC As String
    RedirSCLen As Long
    BACSSettlement As String
    BACSSettlementLen As Long
    SettSection As String
    SettSectionLen As Long
    SettSubSection As String
    SettSubSectionLen As Long
    HandlingBank As String
    HandlingBankLen As Long
    HandlingBankStream As String
    HandlingBankStreamLen As Long
    AccNumFlag As String
    AccNumFlagLen As Long
    DDIVoucherFlag As String
    DDIVoucherFlagLen As Long
    DR As String
    DRLen As Long
    CR As String
    CRLen As Long
    CU As String
    CULen As Long
    PR As String
    PRLen As Long
    BS As String
    BSLen As Long
    DV As String
    DVLen As Long
    AU As String
    AULen As Long
End Type
' Type structure for UKGetCHAPS function call
Type UKChapsType
    SterlingStatus As String
    SterlingStatusLen As Long
    SterlingDateLastChange As String
    SterlingDateLastChangeLen As Long
    SterlingDateClosed As String
    SterlingDateClosedLen As Long
    SterlingSettlement As String
    SterlingSettlementLen As Long
    EuroStatus As String
    EuroStatusLen As Long
    EuroDateLastChange As String
    EuroDateLastChangeLen As Long
    EuroDateClosed As String
    EuroDateClosedLen As Long
    EuroBICBank As String
    EuroBICBankLen As Long
    EuroBICBranch As String
    EuroBICBranchLen As Long
    EuroSettlement As String
    EuroSettlementLen As Long
    EuroRetIndicator As String
    EuroRetIndicatorLen As Long
End Type
' Type structure for UKGetCCCC function call
Type UKCcccType
    CCCCStatus As String
    CCCCStatusLen As Long
    CCCCDateLastChange As String
    CCCCDateLastChangeLen As Long
    CCCCDateClosed As String
    CCCCDateClosedLen As Long
    CCCCSettlement As String
    CCCCSettlementLen As Long
    DASortCode As String
    DASortCodeLen As Long
    CCCCRetIndicator As String
    CCCCRetIndicatorLen As Long
End Type
' Type structure for UKGetPrint function call
Type UKPrintType
    BranchTypeInd As String
    BranchTypeIndLen As Long
    SortCodeMB As String
    SortCodeMBLen As Long
    MajLocName As String
    MajLocNameLen As Long
    MinLocName As String
    MinLocNameLen As Long
    BranchNamePlace As String
    BranchNamePlaceLen As Long
    SecondEntryInd As String
    SecondEntryIndLen As Long
    BranchNameSec As String
    BranchNameSecLen As Long
    BranchTitle1 As String
    BranchTitle1Len As Long
    BranchTitle2 As String
    BranchTitle2Len As Long
    BranchTitle3 As String
    BranchTitle3Len As Long
    Address1 As String
    Address1Len As Long
    Address2 As String
    Address2Len As Long
    Address3 As String
    Address3Len As Long
    Address4 As String
    Address4Len As Long
    Town As String
    TownLen As Long
    County As String
    CountyLen As Long
    PostCodeMajor As String
    PostCodeMajorLen As Long
    PostCodeMinor As String
    PostCodeMinorLen As Long
    TelAreaCode As String
    TelAreaCodeLen As Long
    TelNumber As String
    TelNumberLen As Long
    FaxAreaCode As String
    FaxAreaCodeLen As Long
    FaxNumber As String
    FaxNumberLen As Long
End Type
