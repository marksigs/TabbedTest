<%@ Page language="c#" Codebehind="SummaryAndSubmit.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.SummaryAndSubmit" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdditionalInfoSummary.ascx" TagName="AdditionalInfoSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdviserSummary.ascx" TagName="AdviserSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantAndGuarantorSummary.ascx" TagName="ApplicantAndGuarantorSummary" TagPrefix="Epsom" %>
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

      <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

      <p><cms:cmsBoilerplate cmsref="80" runat="server" /></p>

      <h2 class="form_heading">Reference</h2>

      <itemtemplate>
        <fieldset>
          <legend>Reference details</legend>
          <table class="displaytable_small">
            <tr>
              <td class="prompt">Your reference</td>
              <td><asp:label id="lblIntroducerReference" runat="server" /></td>
            </tr>
          </table>
          <table class="displaytable">
            <tr>
              <td style="text-align:right"><asp:button id="cmdEditReference" runat="server" text="Edit"  cssclass="button"/></td>
            </tr>
          </table>
        </fieldset> 
      </itemtemplate> 
      
      <asp:panel runat="server" id="pnlAdviserDetails">

        <h2 class="form_heading">Adviser Details</h2>

        <epsom:advisersummary id="AdviserSummary" runat="server" />

      </asp:panel>

      <h2 class="form_heading">Product Details</h2>

      <epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" />

      <epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server" />
        
      <epsom:productsummary id="ctlProductSummary" runat="server" />
      
      <epsom:ApplicantAndGuarantorSummary id="ctlApplicantAndGuarantorSummary" runat="server" />
      
      <h2 class="form_heading">Commitments</h2>
	<asp:repeater id="rptCommitments" runat="server">
<itemtemplate>
 <fieldset>
   <legend>Commitment</legend>
      <table class="displaytable">
        <tr>
		      <td class="prompt">Applicant</td>
          <td><asp:label id="lblApplicant" runat="server" /></td>
          <td class="prompt"></td>
          <td></td>
	      </tr>
        <tr>
		      <td class="prompt">Commitment type</td>
          <td><asp:label id="lblCommitmentType" runat="server" /></td>
          <td class="prompt">To be paid on or before completion?</td>
          <td><asp:label id="lblPaidOnCompletion" runat="server" /></td>
	      </tr>
        <tr> 
          <td class="prompt"><asp:label ID="lblTitleBalance" runat="server" text="Outstanding balance"></asp:label></td>
          <td><asp:label id="lblOutstandingBalance" runat="server" /></td>
          <td class="prompt">To be paid by applicants business?</td>
          <td><asp:label id="lblBusinessCommitment" runat="server" /></td>
	      </tr>
        <tr>
          <td class="prompt"><asp:label ID="lblTitleMonthlyPayment" runat="server" text="Monthly repayments"></asp:label></td>
          <td><asp:label id="lblMonthlyRepayment" runat="server" /></td>
	      </tr>
      </table>
	</fieldset>
	  </itemtemplate>
	</asp:repeater>
  <table class="displaytable">
     <tr>
       <td style="text-align:right"><asp:button id="cmdEditCommitments" runat="server" text="Edit"  cssclass="button"/></td>
     </tr>
   </table>

      
      <h2 class="form_heading">Additional Information</h2>

      <epsom:additionalinfosummary id="AdditionalInformation" runat="server" />

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
