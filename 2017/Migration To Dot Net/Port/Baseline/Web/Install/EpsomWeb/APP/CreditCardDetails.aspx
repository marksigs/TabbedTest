<%@ Page language="c#" Codebehind="CreditCardDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.CreditCardDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/FeesSummary.ascx" TagName="FeesSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
<mp:content id="region1" runat="server">

    <script language="javascript">
      function ProcessSubmit()
      {
        var mainLayer = document.getElementById('mainLayer');
        var holdingLayer = document.getElementById('holdingLayer');
        mainLayer.style.display = 'none';
        holdingLayer.style.display = 'block';
        setTimeout('document.getElementById("waitImage").src = "<%=ResolveUrl("~/_gfx/evaluating.gif")%>";', 50);
        scroll(0,0);
      }
    </script>

<epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
</mp:content>
<mp:content id="region2" runat="server">
<cms:helplink id="helplink" class="helplink" runat="server" helpref="2200" />
<h1><cms:cmsBoilerplate cmsref="620" runat="server" /></h1>

<div id="mainLayer">
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
<asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" />
<asp:panel id="pnlFeesSummary" Runat="server">
      <epsom:feessummary id="ctlFeesSummary" runat="server" PaidUpFrontOnly="true" CompactFormat="true"/>
<table class="displaytable" width="100%">
<tr>
  <td class=prompt><b>Total Charge to be debited</b></td>
  <td><asp:label id="lblTotalCharge" runat="server"></asp:label>
  </td>
  <td> </td>
  <td></td>
</tr>
</table>       
</asp:panel>

<asp:panel runat="server" id="pnlCardDetails">
<h2 class=form_heading>Card Details</h2>
<table class=formtable>
<tr>
  <td class=prompt><span class="mandatoryfieldasterisk">*</span> Card holder name</td>
  <td><asp:textbox id="txtCreditCardName" runat="server"></asp:textbox>
    <asp:requiredfieldvalidator runat="Server"
       text="***" errormessage="Please enter card holder name"
       controltovalidate="txtCreditCardName" enableclientscript="false" id="Requiredfieldvalidator1">
    </asp:requiredfieldvalidator> 
  </td>
</tr>
<tr>
  <td class=prompt><span class="mandatoryfieldasterisk">*</span> Card type</td>
  <td><asp:DropDownList id="cmbCreditCardType" Runat="server"></asp:DropDownList>
  <validators:RequiredSelectedItemValidator Runat="server" EnableClientScript="false"
      Text="***" ErrorMessage="Please enter card type" ControlToValidate="cmbCreditCardType">      
    </validators:RequiredSelectedItemValidator>
  </td>
</tr>
<tr>
  <td class=prompt><span class="mandatoryfieldasterisk">*</span> Card number</td>
  <td><asp:TextBox id="txtCreditCardNumber" Runat="server"></asp:TextBox>
    <asp:RequiredFieldValidator Runat="Server"
       text="***" ErrorMessage="Please enter card number"
       ControlToValidate="txtCreditCardNumber" EnableClientScript="false">
    </asp:RequiredFieldValidator> 
  </td>
</tr>
<tr>
  <td class=prompt><span class="mandatoryfieldasterisk">*</span> CVC number</td>
  <td><asp:textbox id="txtISVNumber" runat="server"></asp:textbox>
    <asp:requiredfieldvalidator runat="Server"
       text="***" errormessage="Please enter CVC number"
       controltovalidate="txtISVNumber" enableclientscript="false" id="Requiredfieldvalidator2">
    </asp:requiredfieldvalidator> 
    <asp:customvalidator id="Customvalidator2" runat="server" errormessage="You must supply a 3-digit CVC number"
        enableclientscript="false" onservervalidate="IsvValidator" text="***" />
  </td>
</tr>
<tr>
  <td class=prompt>Issue number  </td>
  <td><asp:textbox id="txtIssueNumber" runat="server"></asp:textbox>
      <asp:customvalidator id="Customvalidator3" runat="server" errormessage="Invalid issue number"
        enableclientscript="false" onservervalidate="IssueNumberValidator" text="***" />
  </td>
</tr>
<tr>
  <td class=prompt>Start date </td>
  <td>
      <asp:dropdownlist id="cmbStartDateMonth" runat="server" />
      <asp:dropdownlist id="cmbStartDateYear" runat="server" />
      <asp:customvalidator id="cmbStartDateValidator" runat="server" errormessage="Invalid start date"
        enableclientscript="false" onservervalidate="StartDateValidator" text="***" />
  </td>
</tr>
<tr>
  <td class=prompt><span class="mandatoryfieldasterisk">* </span>Expiry date</td>
  <td>
      <asp:dropdownlist id="cmbExpiryDateMonth" runat="server" />
      <asp:dropdownlist id="cmbExpiryDateYear" runat="server" />
      <asp:customvalidator id="Customvalidator1" runat="server" errormessage="Invalid expiry date"
        enableclientscript="false" onservervalidate="ExpiryDateValidator" text="***" />
  </td>
</tr>

</table>
</asp:panel>
<asp:panel id="pnl5" Runat="server">
<table class=formtable>
  <tr>
    <td class=prompt>Have you printed the declaration?</td>
    <td>
       <epsom:nullableyesno id="nynDeclaration" runat="server" mustbetrue="true"
       errormessage="You must print the declaration" />
    </td>
  </tr>
  <tr>
    <td class=prompt>Have you printed the Direct Debit Mandate?</td>
    <td>
       <epsom:nullableyesno id="nynDdm" runat="server" mustbetrue="true"
       errormessage="You must print the Direct Debit Mandate"/>
    </td>
  </tr>
</table>
</asp:panel>
<asp:panel id="pnlErrorMessage" runat="server">
<table class=formtable>
  <tr>
    <td>
    <span class="mandatoryfieldasterisk">
    <asp:label id="lblErrorMessage" runat="server"></asp:label>
    </span>
    </td>
  </tr>
</table> 
</asp:panel>
<asp:panel id="pnlAssignMessage" runat="server">
<table class=formtable>
  <tr>
    <td>
    <span class="mandatoryfieldasterisk">
    <asp:label id="lblAssignMessage" runat="server"></asp:label>
    </span>
    </td>
  </tr>
</table> 
</asp:panel>

<p><cms:cmsBoilerplate cmsref="200" runat="server" /></p>
<p>

<epsom:dipnavigationbuttons id="Dipnavigationbuttons1" runat="server"></epsom:dipnavigationbuttons>
</p>
<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

</div>
    
    <div id="holdingLayer" style="display:none">
      <p>
        Please wait while our system evaluates your submission - this may take up to a minute at busy
        times. Please do not navigate away from this page.
      </p>
      <img id="waitImage" src="<%=ResolveUrl("~/_gfx/evaluating.gif")%>" width="386" height="22" alt="Evaluating mortgage decision" />
    </div>


</mp:content>
</mp:contentcontainer>

