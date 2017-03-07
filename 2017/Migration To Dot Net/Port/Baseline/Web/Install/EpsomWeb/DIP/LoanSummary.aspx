<%@ Page language="c#" Codebehind="LoanSummary.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.ProductSummary" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="40700" />

    <h1><cms:cmsBoilerplate cmsref="619" runat="server" /></h1>
    
    
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />
	
	<cms:cmsBoilerplate cmsref="190" runat="server" />
    
    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" cssclass="validationsummary" showsummary="true" id="Validationsummary1" />

    <h2 class="form_heading">Main loan details</h2>
    
    <epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" />

    <epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server" />

    <h2 class="form_heading">Product components</h2>

    <epsom:productsummary id="ctlProductSummary" runat="server" />

    <br />

    <p class="button_orphan">
  		<asp:Button ID="cmdAddComponent" runat="server" text="Add another component" CssClass="button" /> 
  	  <asp:customvalidator runat="server" id="ComponentTotalValidator" onservervalidate="AreComponentsValid" 
        validateemptytext="True"  text="***" enableclientscript="false" />
    </p>

		<epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

  </mp:content>

</mp:contentcontainer>

