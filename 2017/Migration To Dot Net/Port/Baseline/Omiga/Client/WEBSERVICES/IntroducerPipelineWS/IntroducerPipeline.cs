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
namespace Vertex.Fsd.Omiga.Web.Services.IntroducerPipelineWS {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Web.Services.WebServiceBindingAttribute(Name="OmigaSoapRpcSoap", Namespace="http://GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public abstract class OmigaIntroducerPipeline : System.Web.Services.WebService {
        
        /// <remarks/>
        [System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
        public abstract GETINTRODUCERPIPELINERESPONSEType getIntroducerPipeline([System.Xml.Serialization.XmlElementAttribute(Namespace="http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk")] GETINTRODUCERPIPELINEREQUESTType REQUEST);
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINEREQUESTType {
        
        /// <remarks/>
        public PIPELINESEARCHType PIPELINESEARCH;
        
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class PIPELINESEARCHType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string INTRODUCERID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool INCLUDEPREVIOUS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool INCLUDEPREVIOUSSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool INCLUDECANCELLED;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool INCLUDECANCELLEDSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool INCLUDEFIRM;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool INCLUDEFIRMSpecified;
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
    public class MESSAGEType {
        
        /// <remarks/>
        public string MESSAGETEXT;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGETYPE")]
        public string[] MESSAGETYPE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINEVALUATIONType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int VALUATIONAMOUNT;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool VALUATIONAMOUNTSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int REBUILDCOST;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool REBUILDCOSTSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool VALUATIONACCEPTABLE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool VALUATIONACCEPTABLESpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool VALUATIONREFERRED;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool VALUATIONREFERREDSpecified;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GETINTRODUCERPIPELINEMAINAPPLICANTType))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GETINTRODUCERPIPELINECURRENTOWNERType))]
    public class GETINTRODUCERPIPELINEPERSONType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string FORENAME;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SURNAME;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINEMAINAPPLICANTType : GETINTRODUCERPIPELINEPERSONType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int TITLE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool TITLESpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool UKRESIDENTINDICATOR;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool UKRESIDENTINDICATORSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string POSTCODE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINECURRENTOWNERType : GETINTRODUCERPIPELINEPERSONType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string TRANSFERDATE;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINECASEType {
        
        /// <remarks/>
        public GETINTRODUCERPIPELINEMAINAPPLICANTType MAINAPPLICANT;
        
        /// <remarks/>
        public GETINTRODUCERPIPELINEPERSONType ORIGINALOWNER;
        
        /// <remarks/>
        public GETINTRODUCERPIPELINECURRENTOWNERType CURRENTOWNER;
        
        /// <remarks/>
        public GETINTRODUCERPIPELINEVALUATIONType VALUATION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string APPLICATIONNUMBER;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string REFERENCENUMBER;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string SUBMISSIONDATE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string ESTIMATEDCOMPLETIONDATE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string STAGE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int RULESSTATUS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool RULESSTATUSSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string UNDERWRITERDECISION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int PURCHASEPRICE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool PURCHASEPRICESpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int LOANAMOUNT;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool LOANAMOUNTSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public System.Double DRAWDOWN;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool DRAWDOWNSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int TYPEOFAPPLICATION;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool TYPEOFAPPLICATIONSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public int NUMBEROFAPPLICANTS;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool NUMBEROFAPPLICANTSSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public bool GUARANTOR;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool GUARANTORSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string OWNINGUNIT;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.GetIntroducerPipeline.Omiga.vertex.co.uk")]
    public class GETINTRODUCERPIPELINERESPONSEType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("CASE")]
        public GETINTRODUCERPIPELINECASEType[] CASE;
        
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
}