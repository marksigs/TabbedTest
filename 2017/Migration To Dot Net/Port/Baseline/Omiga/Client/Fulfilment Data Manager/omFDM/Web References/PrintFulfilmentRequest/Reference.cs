﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.1.4322.573
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------
// History
// HMA  21/05/2006	MAR1820 Set Keep Alive = False
// AS	06/06/2006	MAR1862 Add configuration of web requests
//------------------------------------------------------------------------------
// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 1.1.4322.573.
// 
namespace Vertex.Fsd.Omiga.omFDM.PrintFulfilmentRequest 
{
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
	using System.Net;
	using System.Reflection;
    
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="DirectGatewaySoapBinding", Namespace="http://interfaces")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericResponseType))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericRequestType))]
    public class DirectGatewaySoapServiceWse : Microsoft.Web.Services2.WebServicesClientProtocol {
        
		// MAR1820
		private static PropertyInfo requestPropertyInfo = null; 

		protected override System.Net.WebRequest GetWebRequest(Uri uri) 
		{ 
			WebRequest request = base.GetWebRequest(uri); 

			if (requestPropertyInfo==null) 
				// Retrieve property info and store it in a static member for optimizing future use 
				requestPropertyInfo = request.GetType().GetProperty("Request"); 

			// Retrieve underlying web request 
			HttpWebRequest webRequest = 
				(HttpWebRequest)requestPropertyInfo.GetValue(request,null); 

			// Setting KeepAlive 
			webRequest.KeepAlive = false; 
			return request; 
		} 

        /// <remarks/>
        public DirectGatewaySoapServiceWse() {
            this.Url = "http://rdgsmarsdgw1/DGWWebServices/services/DirectGatewaySoapEndpoint";
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("", RequestElementName="request", RequestNamespace="http://interfaces", ResponseElementName="response", ResponseNamespace="http://interfaces", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("ProcessPrintFulfilmentResponse", Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")]
        public ProcessPrintFulfilmentResponseType execute([System.Xml.Serialization.XmlElementAttribute(Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")] ProcessPrintFulfilmentRequestType ProcessPrintFulfilmentRequest) {
            object[] results = this.Invoke("execute", new object[] {
                        ProcessPrintFulfilmentRequest});
            return ((ProcessPrintFulfilmentResponseType)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult Beginexecute(ProcessPrintFulfilmentRequestType ProcessPrintFulfilmentRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("execute", new object[] {
                        ProcessPrintFulfilmentRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public ProcessPrintFulfilmentResponseType Endexecute(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((ProcessPrintFulfilmentResponseType)(results[0]));
        }
    }
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="DirectGatewaySoapBinding", Namespace="http://interfaces")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericResponseType))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericRequestType))]
    public class DirectGatewaySoapService : Vertex.Fsd.Omiga.Core.Web.Services.Protocols.SoapHttpClientProtocol {
        
        /// <remarks/>
        public DirectGatewaySoapService() {
            this.Url = "http://rdgsmarsdgw1/DGWWebServices/services/DirectGatewaySoapEndpoint";
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("", RequestElementName="request", RequestNamespace="http://interfaces", ResponseElementName="response", ResponseNamespace="http://interfaces", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("ProcessPrintFulfilmentResponse", Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")]
        public ProcessPrintFulfilmentResponseType execute([System.Xml.Serialization.XmlElementAttribute(Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")] ProcessPrintFulfilmentRequestType ProcessPrintFulfilmentRequest) {
            object[] results = this.Invoke("execute", new object[] {
                        ProcessPrintFulfilmentRequest});
            return ((ProcessPrintFulfilmentResponseType)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult Beginexecute(ProcessPrintFulfilmentRequestType ProcessPrintFulfilmentRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("execute", new object[] {
                        ProcessPrintFulfilmentRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public ProcessPrintFulfilmentResponseType Endexecute(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((ProcessPrintFulfilmentResponseType)(results[0]));
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")]
    public class ProcessPrintFulfilmentRequestType : GenericRequestType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string PackID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string PackTypeID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PackDetailsType PackDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string PrimaryImageReference;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ImageReference", Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string[] ImageReference;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public class PackDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public StructuredAddressDetailsType Address;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PersonDetailsType Person;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PackDetailsTypeMainApplicantIndicator MainApplicantIndicator;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool MainApplicantIndicatorSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ProductAccountNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ProductApplicationNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string DateTimeCreated;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PackDetailsTypePropertyLocation PropertyLocation;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool PropertyLocationSpecified;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(StructuredAddressDetailsTypeExt))]
    public class StructuredAddressDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string FlatNameOrNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string HouseOrBuildingName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string HouseOrBuildingNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Street;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string District;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string TownOrCity;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string County;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string PostCode;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string CountryCode;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public bool Overriden;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool OverridenSpecified;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public StructuredAddressDetailsTypeAddressType AddressType;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool AddressTypeSpecified;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public enum StructuredAddressDetailsTypeAddressType {
        
        /// <remarks/>
        H,
        
        /// <remarks/>
        P,
        
        /// <remarks/>
        M,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://generic.dg.ingdirect.com/0.0.1")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(ProcessPrintFulfilmentResponseType))]
    public abstract class GenericResponseType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ErrorCode;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ErrorMessage;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string CustomerNumber;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://processprintfulfilment.fulfilment.services.dg.ingdirect.com/0.0.1")]
    public class ProcessPrintFulfilmentResponseType : GenericResponseType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string PrintFulfilmentResponse;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public class PersonDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string FirstName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string MiddleName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string LastName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Salutation;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public SpecialNeedType SpecialNeed;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool SpecialNeedSpecified;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public enum SpecialNeedType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("0")]
        Item0,
        
        /// <remarks/>
        L,
        
        /// <remarks/>
        A,
        
        /// <remarks/>
        B,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://generic.dg.ingdirect.com/0.0.1")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(ProcessPrintFulfilmentRequestType))]
    public abstract class GenericRequestType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ClientDevice;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string TellerID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string TellerPwd;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ProxyID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ProxyPwd;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string Operator;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ProductType;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string SessionID;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string CommunicationChannel;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string CommunicationDirection;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, IsNullable=true)]
        public string ServiceName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string CustomerNumber;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public abstract class StructuredAddressDetailsTypeExt : StructuredAddressDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string DateFromMonth;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string DateFromYear;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ResidentialStatus;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string AddressUIId;
        
        /// <remarks/>
        public string Transaction;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public enum PackDetailsTypeMainApplicantIndicator {
        
        /// <remarks/>
        @true,
        
        /// <remarks/>
        @false,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public enum PackDetailsTypePropertyLocation {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("0")]
        Item0,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("1")]
        Item1,
    }
}
