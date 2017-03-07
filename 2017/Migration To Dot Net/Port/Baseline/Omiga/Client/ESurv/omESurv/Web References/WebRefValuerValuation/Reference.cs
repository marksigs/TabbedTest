﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.1.4322.2032
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------
// History
// HMA  21/05/2006  MAR1820  Set Keep Alive = False
// GHun	08/06/2006	MAR1862 Add configuration of web requests
//------------------------------------------------------------------------------
// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 1.1.4322.2032.
// 
namespace Vertex.Fsd.Omiga.omESurv.WebRefValuerValuation 
{
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
	using System.Net;
    
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="DirectGatewaySoapBinding", Namespace="http://interfaces")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericResponseType))]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(GenericRequestType))]
    public class DirectGatewaySoapService : Vertex.Fsd.Omiga.Core.Web.Services.Protocols.SoapHttpClientProtocol {

        /// <remarks/>
        public DirectGatewaySoapService() 
		{
            this.Url = "http://rdgsmarsdgw1/DGWWebServices/services/DirectGatewaySoapEndpoint";
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("", RequestElementName="request", RequestNamespace="http://interfaces", ResponseElementName="response", ResponseNamespace="http://interfaces", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        [return: System.Xml.Serialization.XmlElementAttribute("RealtimeValuationResponse", Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
        public RealtimeValuationResponseType execute([System.Xml.Serialization.XmlElementAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")] RealtimeValuationRequestType RealtimeValuationRequest) {
            object[] results = this.Invoke("execute", new object[] {
                        RealtimeValuationRequest});
            return ((RealtimeValuationResponseType)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult Beginexecute(RealtimeValuationRequestType RealtimeValuationRequest, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("execute", new object[] {
                        RealtimeValuationRequest}, callback, asyncState);
        }
        
        /// <remarks/>
        public RealtimeValuationResponseType Endexecute(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((RealtimeValuationResponseType)(results[0]));
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestType : GenericRequestType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsType RealtimeValuationRequestDetails;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeMessageType MessageType;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string MessageTimeStamp;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string MessageOriginatorReference;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInData FeedInData;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum RealtimeValuationRequestDetailsTypeMessageType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("Survey Instruction")]
        SurveyInstruction,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("Survey Status")]
        SurveyStatus,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("Valuation Report")]
        ValuationReport,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("External Report")]
        ExternalReport,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("Nothing Available")]
        NothingAvailable,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInData {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetails InstructionDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataLenderDetails LenderDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataPropertyDetails PropertyDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataApplicantDetails ApplicantDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataAgentDetails AgentDetails;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataOtherInformation OtherInformation;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetails {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetailsInstructionSource InstructionSource;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType InstTelephone;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType InstFax;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string InstructionSourceDX;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstInstructionType TypeOfInstruction;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstApplicationType Applicationtype;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum RealtimeValuationRequestDetailsTypeFeedInDataInstructionDetailsInstructionSource {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("ING Direct NV")]
        INGDirectNV,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public class PhoneDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string AreaCode;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string LocalNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Extension;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsTypePhoneType PhoneType;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://commontypes.messages.services.dg.ingdirect.com/0.0.1")]
    public enum PhoneDetailsTypePhoneType {
        
        /// <remarks/>
        C,
        
        /// <remarks/>
        R,
        
        /// <remarks/>
        B,
        
        /// <remarks/>
        U,
        
        /// <remarks/>
        F,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationResponseDetailsType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationResponseDetailsTypeResult Result;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string qstReference;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum RealtimeValuationResponseDetailsTypeResult {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("0")]
        Item0,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("1")]
        Item1,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("2")]
        Item2,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("3")]
        Item3,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("4")]
        Item4,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://generic.dg.ingdirect.com/0.0.1")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(RealtimeValuationResponseType))]
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationResponseType : GenericResponseType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public RealtimeValuationResponseDetailsType RealtimeValuationResponseDetails;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataOtherInformation {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Access;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string InstructionNote1;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string InstructionNote2;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string InstructionNote3;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataAgentDetails {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Agent;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType AgentTel;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataApplicantDetails {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstName Applicant1;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstName Applicant2;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public StructuredAddressDetailsType ApplicantAddress;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType ApplicantTelDay;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType ApplicantTelEve;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType ApplicantMobile;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string ApplicantEmail;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class qstName {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Title;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Initials;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string Surname;
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataPropertyDetails {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstPropertyType PropertyType;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="integer")]
        public string NumberOfBedrooms;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstTenure Tenure;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstYN NewProperty;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string PlotNumber;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstIndemnityType BuildingIndemnityType;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="integer")]
        public string PriceOfProperty;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstPriceType PurchasePriceOrEstimatedValue;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public qstName OccupierName;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public StructuredAddressDetailsType PropertyAddress;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType PropertyTelDay;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType PropertyTelEve;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstPropertyType {
        
        /// <remarks/>
        HT,
        
        /// <remarks/>
        HS,
        
        /// <remarks/>
        HD,
        
        /// <remarks/>
        F,
        
        /// <remarks/>
        B,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("")]
        Item,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstTenure {
        
        /// <remarks/>
        Leasehold,
        
        /// <remarks/>
        Freehold,
        
        /// <remarks/>
        Commonhold,
        
        /// <remarks/>
        Feuhold,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("")]
        Item,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstYN {
        
        /// <remarks/>
        Y,
        
        /// <remarks/>
        N,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstIndemnityType {
        
        /// <remarks/>
        N,
        
        /// <remarks/>
        Z,
        
        /// <remarks/>
        P,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("")]
        Item,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstPriceType {
        
        /// <remarks/>
        PP,
        
        /// <remarks/>
        EV,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://realtimevaluation.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public class RealtimeValuationRequestDetailsTypeFeedInDataLenderDetails {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string LenderIfDifferent;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string LenderReference;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string LenderBranch;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType LenderBranchTelephone;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string LenderStaff;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public PhoneDetailsType LenderStaffTelephone;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified, DataType="integer")]
        public string AdvanceAmount;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public string GrossFee;
        
        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute(Form=System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public StructuredAddressDetailsType ReturnAddress;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://generic.dg.ingdirect.com/0.0.1")]
    [System.Xml.Serialization.XmlIncludeAttribute(typeof(RealtimeValuationRequestType))]
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
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstInstructionType {
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("1")]
        Item1,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("2")]
        Item2,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("3")]
        Item3,
        
        /// <remarks/>
        [System.Xml.Serialization.XmlEnumAttribute("4")]
        Item4,
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://questtypes.externalInterface.services.dg.ingdirect.com/0.0.1")]
    public enum qstApplicationType {
        
        /// <remarks/>
        HM,
        
        /// <remarks/>
        FT,
        
        /// <remarks/>
        RM,
    }
}