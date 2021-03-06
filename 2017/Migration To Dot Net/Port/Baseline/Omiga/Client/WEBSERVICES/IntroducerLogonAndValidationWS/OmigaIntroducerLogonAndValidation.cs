﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.1.4322.2032
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by wsdl, Version=1.1.4322.2032.
// 
namespace Vertex.Fsd.Omiga.Web.Services.IntroducerLogonAndValidationWS {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Web.Services.WebServiceBindingAttribute(Name="OmigaSoapRpcSoap", Namespace="http://IntroducerLogonAndValidation.Omiga.vertex.co.uk")]
    public abstract class OmigaIntroducerLogonAndValidation : System.Web.Services.WebService {
        
        /// <remarks/>
        [System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
        public abstract GETFIRMSRESPONSEType getIntroducerUserFirms([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk")] GETFIRMSREQUESTType REQUEST);
        
        /// <remarks/>
        [System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.AuthenticateFirm.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.AuthenticateFirm.Omiga.vertex.co.uk")]
        public abstract AUTHENTICATEFIRMRESPONSEType authenticateFirm([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.AuthenticateFirm.Omiga.vertex.co.uk")] AUTHENTICATEFIRMREQUESTType REQUEST);
        
        /// <remarks/>
        [System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.ValidateBroker.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
        public abstract VALIDATEBROKERRESPONSEType validateBroker([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.ValidateBroker.Omiga.vertex.co.uk")] VALIDATEBROKERREQUESTType REQUEST);
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class GETFIRMSREQUESTType {
        
        /// <remarks/>
        public INTRODUCERCREDENTIALSType INTRODUCERCREDENTIALS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string UNITID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string MACHINEID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CHANNELID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERAUTHORITYLEVEL;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PASSWORD;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class INTRODUCERCREDENTIALSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string INTRODUCERUSERNAME;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string INTRODUCERUSERPASSWORD;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public short LOGINATTEMPTS;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
    public class VALIDATEBROKERACTIVITYFSAType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYDESCRIPTION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CATEGORY;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
    public class VALIDATEBROKERFIRMPERMISSIONSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ACTIVITYFSA")]
        public VALIDATEBROKERACTIVITYFSAType[] ACTIVITYFSA;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public short FRMPERMISSIONS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool FRMPERMISSIONSSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string EFFECTIVEDATE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
    public class VALIDATEBROKERFIRMType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("FIRMPERMISSIONS")]
        public VALIDATEBROKERFIRMPERMISSIONSType[] FIRMPERMISSIONS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FIRMID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FSAREF;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FIRMNAME;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.ValidateBroker.Omiga.vertex.co.uk")]
    public class VALIDATEBROKERRESPONSEType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ARFIRM")]
        public VALIDATEBROKERFIRMType[] ARFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("PRINCIPALFIRM")]
        public VALIDATEBROKERFIRMType[] PRINCIPALFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGE")]
        public MESSAGEType[] MESSAGE;
        
        /// <remarks/>
        public ERRORType ERROR;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public RESPONSEAttribType TYPE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.Omiga.vertex.co.uk")]
    public class MESSAGEType {
        
        /// <remarks/>
        public string MESSAGETEXT;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGETYPE")]
        public string[] MESSAGETYPE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.Omiga.vertex.co.uk")]
    public class ERRORType {
        
        /// <remarks/>
        public long NUMBER;
        
        /// <remarks/>
        public string SOURCE;
        
        /// <remarks/>
        public string VERSION;
        
        /// <remarks/>
        public string DESCRIPTION;
        
        /// <remarks/>
        public string ID;
        
        /// <remarks/>
        public string REQUEST;
        
        /// <remarks/>
        public string REQUESTDOC;
        
        /// <remarks/>
        public string REQUESTNODE;
        
        /// <remarks/>
        public string SCHEMAMASTERNODE;
        
        /// <remarks/>
        public string SCHEMAREFNODE;
        
        /// <remarks/>
        public SQLCOMMANDType SQLCOMMAND;
        
        /// <remarks/>
        public string COMPONENT_ID;
        
        /// <remarks/>
        public string COMPONENT_RESPONSE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.Omiga.vertex.co.uk")]
    public class SQLCOMMANDType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("SQLPARAM")]
        public SQLPARAMType[] SQLPARAM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string TYPE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string COMMANDTEXT;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.Omiga.vertex.co.uk")]
    public class SQLPARAMType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string NAME;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string TYPE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string DIRECTION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string VALUE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.Omiga.vertex.co.uk")]
    public enum RESPONSEAttribType {
        
        /// <remarks/>
        WARNING,
        
        /// <remarks/>
        SYSERR,
        
        /// <remarks/>
        SUCCESS,
        
        /// <remarks/>
        APPERR,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.ValidateBroker.Omiga.vertex.co.uk")]
    public class BROKERCREDENTIALSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BROKERIDENTIFICATION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string BROKERFSAREF;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.ValidateBroker.Omiga.vertex.co.uk")]
    public class VALIDATEBROKERREQUESTType {
        
        /// <remarks/>
        public BROKERCREDENTIALSType BROKERCREDENTIALS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string UNITID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string MACHINEID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CHANNELID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERAUTHORITYLEVEL;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PASSWORD;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.AuthenticateFirm.Omiga.vertex.co.uk")]
    public class AUTHENTICATEFIRMACTIVITYFSAType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYDESCRIPTION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CATEGORY;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.AuthenticateFirm.Omiga.vertex.co.uk")]
    public class AUTHENTICATEFIRMFIRMPERMISSIONSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ACTIVITYFSA")]
        public AUTHENTICATEFIRMACTIVITYFSAType[] ACTIVITYFSA;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PRINCIPALFIRMID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ARFIRMID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ACTIVITYID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public short FRMPERMISSIONS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool FRMPERMISSIONSSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string EFFECTIVEDATE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.AuthenticateFirm.Omiga.vertex.co.uk")]
    public class AUTHENTICATEFIRMRESPONSEType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("FIRMPERMISSIONS")]
        public AUTHENTICATEFIRMFIRMPERMISSIONSType[] FIRMPERMISSIONS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGE")]
        public MESSAGEType[] MESSAGE;
        
        /// <remarks/>
        public ERRORType ERROR;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public RESPONSEAttribType TYPE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.AuthenticateFirm.Omiga.vertex.co.uk")]
    public class AUTHENTICATEFIRMFIRMDETAILSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string INTRODUCERID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FIRMID;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.AuthenticateFirm.Omiga.vertex.co.uk")]
    public class AUTHENTICATEFIRMREQUESTType {
        
        /// <remarks/>
        public AUTHENTICATEFIRMFIRMDETAILSType ARFIRM;
        
        /// <remarks/>
        public AUTHENTICATEFIRMFIRMDETAILSType PRINCIPALFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string UNITID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string MACHINEID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CHANNELID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERAUTHORITYLEVEL;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PASSWORD;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class GETFIRMSFIRMType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FIRMID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FSAREF;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FIRMNAME;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string UNITID;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class INTRODUCERDETAILSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ARFIRM")]
        public GETFIRMSFIRMType[] ARFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("PRINCIPALFIRM")]
        public GETFIRMSFIRMType[] PRINCIPALFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string INTRODUCERID;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class INTRODUCERUSERDETAILSType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int CHANGEPASSWORDINDICATOR;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool CHANGEPASSWORDINDICATORSpecified;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerUserFirms.Omiga.vertex.co.uk")]
    public class GETFIRMSRESPONSEType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGE")]
        public MESSAGEType[] MESSAGE;
        
        /// <remarks/>
        public ERRORType ERROR;
        
        /// <remarks/>
        public INTRODUCERUSERDETAILSType INTRODUCERUSER;
        
        /// <remarks/>
        public INTRODUCERDETAILSType INTRODUCER;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public RESPONSEAttribType TYPE;
    }
}
