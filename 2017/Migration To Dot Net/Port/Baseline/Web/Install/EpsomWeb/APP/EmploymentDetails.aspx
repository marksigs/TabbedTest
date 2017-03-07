<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Page language="c#" Codebehind="EmploymentDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.EmploymentDetails1" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/EmploymentDetails.ascx" TagName="EmploymentDetail" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">
    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
  </mp:content>
  
  <mp:content id="region2" runat="server">
 <cms:helplink id="helplink" class="helplink" runat="server" helpref="2230" />
    <h1>Employment details: <asp:label id="lblCandidate" runat="server" text="" /></h1>
    
    
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
    
    <asp:validationsummary id="Validationsummary1" runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" />
    
    <cms:cmsBoilerplate cmsref="230" runat="server" />
    
    <h2 class="form_heading">Employment Summary</h2>
    
    <fieldset>

      <legend>Employment Summary</legend>
          
      <table class="formtable" >
         <tr>
          <td class="prompt"><span class="mandatoryfieldasterisk">*</span> UK Taxpayer
          <td>
            <epsom:nullableyesno id="nynTaxpayer" runat="server" required="true" autopostback="true" Enabled="false"
            errormessage="Specify if you are a UK Taxpayer" >
            </epsom:nullableyesno>
          </td>
        </tr> 
      </table>
      
      <asp:Panel ID="pnlUkTaxPayerDetails" runat="server">
        <table class="formtable">
          <tr>
            <td class="prompt"></td>
            <td>
              <p>Please give details<br /><asp:textbox id="txtUkTaxPayerDetails" runat="server" textmode="multiline"></asp:textbox></p>
            </td>
          </tr>
        </table>
      </asp:Panel>
      
      <table class="formtable">
        <tr>
          <td class="prompt">Tax office &amp; Reference number</td>
          <td><asp:textbox id=txtTaxOffice runat="server"></asp:textbox><asp:textbox id="txtTexReferenceNumber" runat="server"></asp:textbox></td>
        </tr>
        <tr>
          <td class="prompt">National Insurance number</td>
          <td><asp:TextBox id=txtNationalInsuranceNumber runat="server"></asp:TextBox></td>
        </tr>
        <tr>
          <td class="prompt">Do you self assess to the Inland Revenue</td>
          <td>
            <asp:radiobutton id="rdoSelfAssesYes" runat="server" GroupName="grpSelfAssess" Text="Yes"></asp:radiobutton>
            <asp:radiobutton id="rdoSelfAssesNo" runat="server" GroupName="grpSelfAssess" Text="No"></asp:radiobutton>
          </td>
        </tr>
      </table>
    </fieldset>
    
    <asp:repeater id="rptEmployments" runat="server">
      <ItemTemplate>
          <h2 class="form_heading"><asp:Literal ID="lblEmploymentTitle" Runat="server"></asp:Literal></h2>
          <epsom:EmploymentDetail id="ctlEmploymentDetail" runat="server" /> 
      </ItemTemplate>
    </asp:repeater>
    
    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server"></sstchur:SmartScroller>

  </mp:content>

</mp:contentcontainer>

