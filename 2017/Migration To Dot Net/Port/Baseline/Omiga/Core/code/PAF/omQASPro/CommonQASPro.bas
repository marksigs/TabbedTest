Attribute VB_Name = "QASProStandardModule"
Option Explicit
'**** Visual Basic (32 bit) header file for QASRVEA.DLL v3.08(28) ****
Global Const qaerr_FATAL = -1000
Global Const qaerr_NOMEMORY = -1001
Global Const qaerr_INITOOLARGE = -1005
Global Const qaerr_ININOEXTEND = -1006
Global Const qaerr_FILEOPEN = -1010
Global Const qaerr_FILEEXIST = -1011
Global Const qaerr_FILEREAD = -1012
Global Const qaerr_FILEWRITE = -1013
Global Const qaerr_FILEDELETE = -1014
Global Const qaerr_FILEACCESS = -1016
Global Const qaerr_FILEVERSION = -1017
Global Const qaerr_FILEHANDLE = -1018
Global Const qaerr_FILECREATE = -1019
Global Const qaerr_FILERENAME = -1020
Global Const qaerr_FILEEXPIRED = -1021
Global Const qaerr_FILENOTDEMO = -1022
Global Const qaerr_READFAIL = -1025
Global Const qaerr_WRITEFAIL = -1026
Global Const qaerr_BADDRIVE = -1027
Global Const qaerr_BADDIR = -1028
Global Const qaerr_DIRCREATE = -1029
Global Const qaerr_BADOPTION = -1030
Global Const qaerr_BADINIFILE = -1031
Global Const qaerr_BADLOGFILE = -1032
Global Const qaerr_BADMEMORY = -1033
Global Const qaerr_BADHOTKEY = -1034
Global Const qaerr_HOTKEYUSED = -1035
Global Const qaerr_BADRESOURCE = -1036
Global Const qaerr_BADDATADIR = -1037
Global Const qaerr_BADTEMPDIR = -1038
Global Const qaerr_NOTDEFINED = -1040
Global Const qaerr_DUPLICATE = -1041
Global Const qaerr_BADACTION = -1042
Global Const qaerr_CCFAILURE = -1050
Global Const qaerr_CCBADCODE = -1051
Global Const qaerr_CCACCESS = -1052
Global Const qaerr_CCNODONGLE = -1053
Global Const qaerr_CCINSTALL = -1060
Global Const qaerr_CCEXPIRED = -1061
Global Const qaerr_CCDATETIME = -1062
Global Const qaerr_CCUSERLIMIT = -1063
Global Const qaerr_CCACTIVATE = -1064
Global Const qaerr_CCBADDRIVE = -1065
Global Const qaerr_UNAUTHORISED = -1070
Global Const qaerr_NOTHREAD = -1080
Global Const qaerr_NOTASK = -1090
Declare Sub QAInitialise Lib "QASRVEA.DLL" (ByVal vi1 As Long)
Declare Sub QAErrorMessage Lib "QASRVEA.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long)
Declare Function QAErrorLevel Lib "QASRVEA.DLL" (ByVal vi1 As Long) As Long
Declare Sub QAVersionInfo Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal vi2 As Long)
Declare Function QADataInfo Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal vi2 As Long, ri3 As Long) As Long
Declare Function QASystemInfo Lib "QASRVEA.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
Declare Sub QAUpdateKey Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal vi2 As Long)
Declare Function QAUpdateCode Lib "QASRVEA.DLL" (ByVal vs1 As String) As Long
Declare Function QALicenseInfo Lib "QASRVEA.DLL" (ri1 As Long, ri2 As Long, ri3 As Long) As Long
Declare Function QAAuthorise Lib "QASRVEA.DLL" (ByVal vs1 As String, ByVal vl2 As Long) As Long
Global Const qasrvid_NEAREST = 1
Global Const qasrvid_DATAPLUS = 2
Global Const qasrvid_PRO = 4
Global Const qasrvid_NAMELOOKUP = 8
Global Const qasrvid_NAMESRCH = 16
Global Const qaerr_SRVBADPARAMS = -9939
Global Const qaerr_SRVBADSRCHKEY = -9938
Global Const qaerr_SRVINVPCODELEN = -9937
Global Const qaerr_SRVTOODEEP = -9936
Global Const qaerr_SRVDPNOTALL = -9935
Global Const qaerr_SRVNOTSTARTED = -9934
Global Const qaerr_SRVALREADYSTARTED = -9933
Global Const qaerr_SRVCANTSTART = -9932
Global Const qaerr_SRVDIDNTSTART = -9931
Global Const qaerr_SRVNAMESFAILED = -9930
Global Const qaerr_SRVNAMESCORRUPT = -9929
Global Const qaerr_SRVNAMESVERSION = -9928
Global Const qaerr_SRVMODULENOTREQ = -9927
Global Const qaerr_SRVNOMODULES = -9926
Global Const qaerr_SRVNOMATCH = -9925
Global Const qaerr_SRVOUTOFPHASE = -9924
Global Const qaerr_SRVBUFFERTOOSMALL = -9923
' IVW Constant Required.
Global Const qaerr_SRVNOADDRESSMATCH = -9978
Declare Function QAServer_Startup Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal rs2 As String) As Long
Declare Sub QAServer_Config Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal rs2 As String)
Declare Sub QAServer_Shutdown Lib "QASRVEA.DLL" ()
Declare Function QAServer_Request Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal rs2 As String) As Long
Declare Function QAServer_Search Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal rs2 As String, ri3 As Long) As Long
Declare Sub QAServer_FreeResponse Lib "QASRVEA.DLL" (ByVal rs1 As String)
Declare Function QAServer_ResponseSize Lib "QASRVEA.DLL" () As Long
Declare Function QAServer_Response Lib "QASRVEA.DLL" (ByVal rs1 As String, ByVal vl2 As Long) As Long
'**** End of File ****
