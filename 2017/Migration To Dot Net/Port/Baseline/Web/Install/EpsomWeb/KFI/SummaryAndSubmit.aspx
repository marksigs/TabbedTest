<%@ Page language="c#" Codebehind="SummaryAndSubmit.aspx.cs" AutoEventWireup="false" 
  SmartNavigation="false" Inherits="Epsom.Web.Kfi.SummaryAndSubmit" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdditionalInfoSummary.ascx" TagName="AdditionalInfoSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdviserSummary.ascx" TagName="AdviserSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantAndGuarantorSummary.ascx" TagName="ApplicantAndGuarantorSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/FeesSummary.ascx" TagName="FeesSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/BrokerSummary.ascx" TagName="BrokerSummary" TagPrefix="Epsom" %>
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
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2080" />
     <h1><cms:cmsBoilerplate cmsref="608" runat="server" /></h1>

    <div id="mainLayer">

      <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
    
    <p><cms:cmsBoilerplate cmsref="80" runat="server" /></p>     
    
      <asp:panel runat="server" id="pnlAdviserDetails">
      <h2 class="form_heading">Adviser Details</h2>
      <epsom:advisersummary id="AdviserSummary" runat="server" />
      </asp:panel>

      <h2 class="form_heading">Product Details</h2>

      <epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" />

      <epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server" />
        
      <epsom:productsummary id="ctlProductSummary" runat="server" />
      
      <epsom:FeesSummary id="ctlFeesSummary" runat="server" />
      
      <epsom:ApplicantAndGuarantorSummary id="ctlApplicantAndGuarantorSummary" runat="server" />

      <epsom:brokerSummary id="ctlBrokerSummary" runat="server"></epsom:brokersummary>
      
      <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    </div>
    
    <div id="holdingLayer" style="display:none">

      <p>
		<cms:cmsBoilerplate cmsref="81" runat="server" />
      </p>

      <img id="waitImage" src="<%=ResolveUrl("~/_gfx/evaluating.gif")%>" width="386" height="22" alt="Evaluating mortgage decision" />

    </div>


  </mp:content>

</mp:contentcontainer>
