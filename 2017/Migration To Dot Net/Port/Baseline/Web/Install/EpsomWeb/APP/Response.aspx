<%@ Page language="c#" Codebehind="Response.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.Response" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="40800" />
    <h1><cms:cmsBoilerplate cmsref="638" runat="server" /></h1>
    
    <p class="dipnavigationbuttons">
	    <asp:button id="cmdDone" runat="server" text="Done" CssClass="button" />
	  </p>
    
    <h2 class="form_heading" >Summary</h2>

    <h3>Submission <asp:label id=lblResponseStatus runat="server" /></h3>
    <asp:label id="lblResponseMessage" runat="server" />

	  <table class="formtable">
	  	<tr>
	  	  <td class="prompt">Case reference number</td>
	  	  <td><asp:label id="lblReferenceNumberText" runat="server" /></td>
	  	</tr>
	  	  <tr>
	  	  <td class="prompt">Submission date</td>
	  	  <td><asp:label id="lblSubmissionDate" runat="server" /></td>
	  	</tr>
	  	  <tr>
	  	  <td class="prompt">Your reference number</td>
	  	  <td><asp:label id="lblYourReferenceNumber" runat="server" /></td>
	  	</tr>

		</Table>
		
		<asp:panel id="pnlPaymentDetails" runat="server">
		  <br>
	    <table class="formtable">
	  	  <tr>
	  	    <td class="prompt">Payment details</td>
	  	    <td><asp:label id="lblName" runat="server" /></td>
	  	  </tr>
	  	    <tr>
	  	    <td class="prompt"></td>
	  	    <td><asp:label id="lblCardType" runat="server" /></td>
	  	  </tr>
	  	  <tr>
	  	    <td class="prompt"></td>
	  	    <td><asp:label id="lblCardNumber" runat="server" /></td>
	  	  </tr>
	  	  <tr>
	  	    <td class="prompt"></td>
	  	    <td><asp:label id="lblAmount" runat="server" /></td>
	  	  </tr>
		  </Table>
		  
		  <p class="button_orphan">
        <asp:button id="cmdPrintableApplication" runat="server" causesvalidation="False" text="Print Application" cssclass="button" />
        <asp:button id="cmdPrintableReceipt" runat="server" causesvalidation="False" text="Print Receipt" cssclass="button" />
      </p>
		
		</asp:panel>
  
  <p class="dipnavigationbuttonswithbar">
 	  <asp:button id="cmdDone2" runat="server" text="Done" CssClass="button" command="Done" />
  </p>
					    
  </mp:content>

</mp:contentcontainer>
