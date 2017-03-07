<%@ Page language="c#" Codebehind="SolicitorDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.SolicitorDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
  <mp:content id="region1" runat="server">
    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="3400" />
    <h1><cms:cmsBoilerplate cmsref="634" runat="server" /></h1>
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
    
    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
	<cms:cmsBoilerplate cmsref="340" runat="server" />
    <fieldset>
      <legend>Solicitor finder</legend>
    
      <asp:panel id="pnlFindError" runat="server" Visible="False">
        <p class="error">
          <asp:label id="lblFindMessage" runat="server" /> 
        </p>  
      </asp:panel>
      
      <table class="formtable">
        <tr>
          <td class="prompt">Company name</td>
          <td><asp:textbox id="txtSearchCompanyName" runat="server" CssClass="text"/> </td>
          <td width="20%">You can search using a single letter</td>
        </tr>
        
        <tr>
          <td class="prompt">Town or postcode</td>
          <td><asp:textbox id="txtSearchTown" runat="server" CssClass="text"/> <asp:textbox id="txtSearchPostcode" runat="server" CssClass="text"/></td>
          <td width="20%">or more for the firm name and town</td>
        </tr>
      
        <tr>
          <td class="prompt"></td>
          <td>
            <asp:button id="cmdValidateFirm" runat="server" causesValidation="false" text="Validate Firm" />
          </td>
        </tr>
        
        <tr>
          <td class="prompt"></td>
          <td><asp:ListBox id="lstFoundSolicitors" runat="server" rows="5" style="width: 44%" AutoPostback="True" /></td>
        </tr>
        
      </table>
      
    </fieldset>

    <table class="formtable">
      <tr>
        <td class=prompt><span class="mandatoryfieldasterisk">* </span>Name of firm</td>
        <td>
          <asp:TextBox id=txtFirm Runat="server" CssClass="text"></asp:TextBox>
          <asp:requiredfieldvalidator runat="server" controltovalidate="txtFirm" text="***" errormessage="Please validate firm or enter a firm name" id="Requiredfieldvalidator3" enableclientscript="false" />
        </td>
      </tr>

      <tr>
        <td class=prompt>Name of solicitor</td>
        <td>
        <asp:textbox id="txtFirstName" runat="server" CssClass="text"></asp:textBox>
        <asp:textbox id="txtLastName" runat="server" cssclass="text"></asp:textbox>
        </td>
      </tr>
    </table>
    
    <epsom:Address id="ctlAddress" runat="server"></epsom:Address>
    
    <epsom:telephonenumbers id="ctlTelephoneNumbers" FirstItemMandatory="true" usageComboGroup="ContactTelephoneUsage" runat="server" />
    
    <table class="formtable">
      <tr>
        <td class="prompt">DX number/District</td>
        <td>
          <asp:textbox id="txtDxNumber" runat="server" CssClass="text" maxlength="30"/>
          <asp:textbox id="txtDistrict" runat="server" cssclass="text" />
          <asp:regularexpressionvalidator runat="server" controltovalidate="txtDxNumber" text="***"
          errormessage="DX number must be alphanumeric" enableclientscript="false" validationexpression="[a-zA-Z0-9]*">
          </asp:regularexpressionvalidator>
        </td>
      </tr>
      
      <tr>
        <td class="prompt">Email Address</td>
        <td>
          <asp:textbox id="txtEmailAddress" runat="server" CssClass="text" />
        </td>
      </tr>
      
    </table>
    
    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server"></sstchur:SmartScroller>
      
  </mp:content>
</mp:contentcontainer>

