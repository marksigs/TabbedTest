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
namespace Vertex.FSD.Omiga.Web.Services.RunIncomeCalculationsWS {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Web.Services.WebServiceBindingAttribute(Name="OmigaSoapRpcSoap", Namespace="http://RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public abstract class OmigaRunIncomeCalculations : System.Web.Services.WebService {
        
        /// <remarks/>
        [System.Web.Services.WebMethodAttribute()]
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("RESPONSE", Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
        public abstract RESPONSEType runIncomeCalculations();
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class REQUESTType {
        
        /// <remarks/>
        public INCOMECALCULATIONRequestType INCOMECALCULATION;
        
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
        public string ADMINSYSTEMSTATE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string USERAUTHORITYLEVEL;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string PASSWORD;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class INCOMECALCULATIONRequestType {
        
        /// <remarks/>
        public APPLICATIONType APPLICATION;
        
        /// <remarks/>
        public CUSTOMERLISTRequestType CUSTOMERLIST;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string CALCULATEMAXBORROWING;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string QUICKQUOTEMAXBORROWING;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class APPLICATIONType {
        
        /// <remarks/>
        public string APPLICATIONNUMBER;
        
        /// <remarks/>
        public string APPLICATIONFACTFINDNUMBER;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.IDUK.Omiga.marlboroughstirling.com")]
    public class ERRORType {
        
        /// <remarks/>
        public long NUMBER;
        
        /// <remarks/>
        public string SOURCE;
        
        /// <remarks/>
        public string DESCRIPTION;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.IDUK.Omiga.marlboroughstirling.com")]
    public class MESSAGEType {
        
        /// <remarks/>
        public string MESSAGETEXT;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("MESSAGETYPE")]
        public string[] MESSAGETYPE;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlTextAttribute()]
        public string[] Text;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class MAXIMUMBORROWINGType {
        
        /// <remarks/>
        public string INCOME1;
        
        /// <remarks/>
        public string INCOME2;
        
        /// <remarks/>
        public string INCOMEMULTIPLIERTYPE;
        
        /// <remarks/>
        public string INCOMEMULTIPLE;
        
        /// <remarks/>
        public string MAXIMUMBORROWINGAMOUNT;
        
        /// <remarks/>
        public string AMOUNTREQUESTED;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class CUSTOMERResponseType {
        
        /// <remarks/>
        public string CUSTOMERNUMBER;
        
        /// <remarks/>
        public string CUSTOMERVERSIONNUMBER;
        
        /// <remarks/>
        public string MONTHLYALLOWABLEINCOME;
        
        /// <remarks/>
        public string ANNUALALLOWABLEINCOME;
        
        /// <remarks/>
        public string NETANNUALALLOWABLEINCOME;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class CUSTOMERLISTResponseType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("CUSTOMER")]
        public CUSTOMERResponseType[] CUSTOMER;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class LENDERDETAILSType {
        
        /// <remarks/>
        public CUSTOMERLISTResponseType CUSTOMERLIST;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class LENDERLISTType {
        
        /// <remarks/>
        public LENDERDETAILSType LENDERDETAILS;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class ALLOWABLEINCOMEType {
        
        /// <remarks/>
        public LENDERLISTType LENDERLIST;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class INCOMECALCULATIONResponseType {
        
        /// <remarks/>
        public ALLOWABLEINCOMEType ALLOWABLEINCOME;
        
        /// <remarks/>
        public MAXIMUMBORROWINGType MAXIMUMBORROWING;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Response.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class RESPONSEType {
        
        /// <remarks/>
        public INCOMECALCULATIONResponseType INCOMECALCULATION;
        
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://msgtypes.IDUK.Omiga.marlboroughstirling.com")]
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class CUSTOMERRequestType {
        
        /// <remarks/>
        public string CUSTOMERNUMBER;
        
        /// <remarks/>
        public string CUSTOMERVERSIONNUMBER;
        
        /// <remarks/>
        public string CUSTOMERROLETYPE;
        
        /// <remarks/>
        public string CUSTOMERORDER;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://Request.RunIncomeCalculations.IDUK.Omiga.marlboroughstirling.com")]
    public class CUSTOMERLISTRequestType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("CUSTOMER")]
        public CUSTOMERRequestType[] CUSTOMER;
    }
}