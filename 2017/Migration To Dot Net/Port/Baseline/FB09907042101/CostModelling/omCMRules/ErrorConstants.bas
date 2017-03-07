Attribute VB_Name = "ErrorConstants"
Option Explicit

Public Enum ERRORS
oeUnspecifiedError
    eNOTMIPLEMENTED = vbObjectError + 512
    eXMLMISSINGELEMENT
    eXMLMISSINGATTRIBUTE
    eXMLINVALIDATTTRIBUTE
    eXMLINVALIDATTTRIBUTEVALUE
    eRECORDNOTFOUND
    eUNSPECIFIEDERROR = vbObjectError + 512 + 999
    eFAILEDTOLOADREQUEST = vbObjectError + 512 + 1
End Enum

