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
// This source code was auto-generated by Microsoft.VSDesigner, Version 1.1.4322.2032.
// 
namespace Vertex.Fsd.Omiga.omVex.vexWebService {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="ValuationMessageSoap", Namespace="http://tempuri.org/")]
    public class ValuationMessage : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        /// <remarks/>
        public ValuationMessage() {
            this.Url = "https://www.vexcommunication.com/ValExTest/ValuationMessage.asmx";
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SubmitIncomingURRNMessage", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SubmitIncomingURRNMessage(string mESSAGE, int vERSIONID, string sOURCEGUID) {
            object[] results = this.Invoke("SubmitIncomingURRNMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginSubmitIncomingURRNMessage(string mESSAGE, int vERSIONID, string sOURCEGUID, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("SubmitIncomingURRNMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID}, callback, asyncState);
        }
        
        /// <remarks/>
        public string EndSubmitIncomingURRNMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SubmitIncomingStatusMessage", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SubmitIncomingStatusMessage(string mESSAGE, int vERSIONID, string sOURCEGUID) {
            object[] results = this.Invoke("SubmitIncomingStatusMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginSubmitIncomingStatusMessage(string mESSAGE, int vERSIONID, string sOURCEGUID, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("SubmitIncomingStatusMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID}, callback, asyncState);
        }
        
        /// <remarks/>
        public string EndSubmitIncomingStatusMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SubmitIncomingNewHCRCaseMessage", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SubmitIncomingNewHCRCaseMessage(string mESSAGE, int vERSIONID, string sOURCEGUID) {
            object[] results = this.Invoke("SubmitIncomingNewHCRCaseMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginSubmitIncomingNewHCRCaseMessage(string mESSAGE, int vERSIONID, string sOURCEGUID, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("SubmitIncomingNewHCRCaseMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID}, callback, asyncState);
        }
        
        /// <remarks/>
        public string EndSubmitIncomingNewHCRCaseMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SubmitIncomingInstructionMessage", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SubmitIncomingInstructionMessage(string mESSAGE, int vERSIONID, string sOURCEGUID) {
            object[] results = this.Invoke("SubmitIncomingInstructionMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginSubmitIncomingInstructionMessage(string mESSAGE, int vERSIONID, string sOURCEGUID, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("SubmitIncomingInstructionMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        sOURCEGUID}, callback, asyncState);
        }
        
        /// <remarks/>
        public string EndSubmitIncomingInstructionMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/SubmitIncomingMessage", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string SubmitIncomingMessage(string mESSAGE, int vERSIONID, string vERSIONNAME, string sOURCEGUID) {
            object[] results = this.Invoke("SubmitIncomingMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        vERSIONNAME,
                        sOURCEGUID});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginSubmitIncomingMessage(string mESSAGE, int vERSIONID, string vERSIONNAME, string sOURCEGUID, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("SubmitIncomingMessage", new object[] {
                        mESSAGE,
                        vERSIONID,
                        vERSIONNAME,
                        sOURCEGUID}, callback, asyncState);
        }
        
        /// <remarks/>
        public string EndSubmitIncomingMessage(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((string)(results[0]));
        }
    }
}
