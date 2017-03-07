<%@ Control Language="c#" AutoEventWireup="false" Codebehind="NullableYesNo.ascx.cs" Inherits="Epsom.Web.WebUserControls.NullableYesNo" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:RadioButton id="rdoYes" GroupName="grpYesNo" Runat="server" text="Yes"/>
<asp:RadioButton id="rdoNo" GroupName="grpYesNo" Runat="server" text="No"/>
<asp:customvalidator id="validatorYesNo" runat="server" enableclientscript="false"
        onservervalidate="ValidatorRequired"
        errormessage="You must select Yes or No" text="***" />

