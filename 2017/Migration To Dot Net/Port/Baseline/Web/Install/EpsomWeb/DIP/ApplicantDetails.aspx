<%@ Page language="c#" Codebehind="ApplicantDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.ApplicantDetails" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Applicant.ascx" TagName="Applicant" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2060" />

    <h1><asp:Label id="lblHeading" runat="server" /></h1>

   
   
   <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
	
	<cms:cmsBoilerplate cmsref="60" runat="server" />
    
    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" CssClass="validationsummary"
      headertext="Input errors:" displaymode="BulletList" />

    <asp:panel id="pnlFormError" runat="server" Visible="False">
      <p class="error">
        <asp:label id="lblFormErrorMessage" runat="server" /> 
      </p>
    </asp:panel>

<epsom:applicant id="ctlApplicant" runat="server" />

    <br /><br />

    <p class="button_orphan">
      <asp:Button id="cmdAddApplicant" runat="server" CssClass="button" text="Add another applicant" />
    </p>

    <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>
