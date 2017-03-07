<%@ Page language="c#" Codebehind="Response.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.Response" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdditionalInfoSummary.ascx" TagName="AdditionalInfoSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/AdviserSummary.ascx" TagName="AdviserSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantAndGuarantorSummary.ascx" TagName="ApplicantAndGuarantorSummary" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/Documents.ascx" TagName="Documents" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" masterpagefile="../masterpages/masterpage2.ascx" runat="server">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id=DipMenu runat="server"></epsom:dipmenu>

  </mp:content>
  
  <mp:content id="region2" runat="server">

    <cms:helplink id="helplink" class="helplink" runat="server" helpref="2090" />

    <table>
      <tr>
        <td><h1><cms:cmsBoilerplate cmsref="609" runat="server" /></h1></td>
        <td class="displayimage" align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:image   id="imgDBLogo" runat="server" src="../_gfx/db_logo.gif"></asp:image></td>
      </tr>
    </table>

    <p class="dipnavigationbuttons">
	     <asp:button id="cmdDone" runat="server" text="Done" CssClass="button" />
	   </p>

    <h2 class="form_heading">Response</h2>

    <br />

    <p>
      &nbsp;&nbsp;
      <b>Response: &nbsp;<span class="highlight"><asp:label id="lblResponseStatus" runat="server" /></span></b>
    </p>
    
    <asp:panel id="pnlReferText" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="91" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlAcceptText" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="90" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlCascadeText" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="93" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlDeclineText" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="92" runat="server" /></p>
    </asp:panel>

    <asp:label id="lblAssignMsg" ForeColor="Red" text="" visible="True" enabled="True" runat="server" />
    
    <table class="displaytable_small">
      <tr>
        <td class="prompt">Case reference number</td>
        <td><asp:label id="lblCaseReferenceNumber" runat="server" /></td>
      </tr>
      <tr>
        <td class="prompt">Your reference number</td>
        <td><asp:label id="lblYourReferenceNumber" runat="server" /></td>
      </tr>
      <tr>
        <td class="prompt">Submitted</td>
        <td><asp:label id="lblDateSubmitted" runat="server" /></td>
      </tr>
      <asp:Panel id="pnlSchemeRequested" runat="server">
        <tr>
          <td class="prompt">Scheme requested</td>
          <td><asp:label id="lblSchemeRequested" runat="server" /></td>
        </tr>
      </asp:Panel>
      <asp:Panel id="pnlScheme" runat="server">
        <tr>
          <td class="prompt"><asp:label id="lblSchemePrompt" runat="server" /></td>
          <td><asp:label id="lblScheme" runat="server" /></td>
        </tr>
      </asp:Panel>
    </table>
 
    <p class="dipnavigationbuttons">
      <span style="left:0.5em; position:absolute">
        <asp:Button id="cmdPrintDIP" runat="server" autopostback="false" text="Print DIP" visible="True" CssClass="button" />
      </span> 
      <asp:Button id="cmdKfi" runat="server" text="KFI"  CssClass="button"/>
      &nbsp;
      <asp:Button id="cmdApply" runat="server" text="Apply" CssClass="button" />
      &nbsp;
      <asp:Button id="cmdAssign" runat="server" text="Assign"  CssClass="button"/>
      &nbsp;
      <asp:Button id="cmdNewDip" runat="server" text="Further DIP" CssClass="button" />
    </p>  

    <br />

    <asp:panel id="pnlProductCascade" runat="server">
    
      <p><b>Product details</b></p>
    
      <p>
        <cms:cmsBoilerplate cmsref="99" runat="server" />
      </p>

      <table class="formtable" id="ProductCascadeListBox">
        <tr>
          <td>
            <asp:listbox id="lstProductCascade" runat="server" rows="10" style="width: 98%" />
          </td> 
        </tr>
      </table>

    </asp:panel>

    <br />

    <asp:panel id="pnlShoppingList" runat="server">
    
    <h2 class="form_heading" id="hdrTitle" runat="server">What we need from you</h2>
    
    <asp:panel id="pnlShoppingList1" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="94" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlShoppingList2" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="95" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlShoppingList3" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="96" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlShoppingList4" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="97" runat="server" /></p>
    </asp:panel>
     
     <asp:panel id="pnlShoppingList5" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="98" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlShoppingListDefault" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="99" runat="server" /></p>
    </asp:panel>
  </asp:panel>  
     
 <asp:panel id="pnlKYCApplicant1" runat="server">
 <h2 class="form_heading" id="H21" runat="server">Know your customer</h2>
      <epsom:documents id="Documents" documentcategory="KYCPolicies" runat="server" />
     
    <asp:panel id="pnlKYC1Applicant1" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="1000" runat="server" /></p>
    </asp:panel>
   
       <asp:panel id="pnlKYC2Applicant1" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1001" runat="server" /></p>
    </asp:panel>
    
           <asp:panel id="pnlKYC3Applicant1" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1002" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlKYC4Applicant1" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1003" runat="server" /></p>
    </asp:panel>
    
        <asp:panel id="pnlKYC5Applicant1" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1004" runat="server" /></p>
    </asp:panel>
    
        <asp:panel id="pnlKYCDefaultApplicant1" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="1005" runat="server" /></p>
	    </asp:panel>
   </asp:panel>
      
 
 <asp:panel id="pnlKYCApplicant2" runat="server">
     
    <h2 class="form_heading" id="H22" runat="server">Know your customer</h2>
      <epsom:documents id="Documents1" documentcategory="KYCPolicies" runat="server" />
      
    <asp:panel id="pnlKYC1Applicant2" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="1000" runat="server" /></p>
    </asp:panel>
   
       <asp:panel id="pnlKYC2Applicant2" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1001" runat="server" /></p>
    </asp:panel>
        
           <asp:panel id="pnlKYC3Applicant2" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1002" runat="server" /></p>
    </asp:panel>
    
    <asp:panel id="pnlKYC4Applicant2" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1003" runat="server" /></p>
    </asp:panel>
    
        <asp:panel id="pnlKYC5Applicant2" runat="server">
       <p class="forminfo"><cms:cmsBoilerplate cmsref="1004" runat="server" /></p>
    </asp:panel>
    
        <asp:panel id="pnlKYCDefaultApplicant2" runat="server">
      <p class="forminfo"><cms:cmsBoilerplate cmsref="1005" runat="server" /></p>
  </asp:panel>  
    </p>
    
    </asp:panel>

    
	  <asp:panel runat="server" id="pnlAdviserDetails">

      <h2 class="form_heading">Adviser Details</h2>

      <epsom:advisersummary id="ctlAdviserSummary" runat="server" />

    </asp:panel>

    <h2 class="form_heading">Product Details</h2>

    <epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" />

    <epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server" />
        
    <epsom:productsummary id="ctlProductSummary" DisplayOnly="True" runat="server" />
      
    <epsom:applicantandguarantorsummary id="ctlApplicantAndGuarantorSummary" runat="server" />   
      
    <h2 class="form_heading">Additional Information</h2>

    <epsom:additionalinfosummary id="ctlAdditionalInformation" runat="server" />
	  
    <p class="dipnavigationbuttonswithbar">
 	    <asp:button id="cmdDone2" runat="server" text="Done" CssClass="button" command="Done" />
    </p>

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />
				    
  </mp:content>

</mp:contentcontainer>
