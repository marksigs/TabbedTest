<%@ Page language="c#" Codebehind="LoanSummary.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.ProductSummary" %>
<%@ Register Src="~/WebUserControls/ProductSummary.ascx" TagName="ProductSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoanDetailsSummary.ascx" TagName="LoanDetailsSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubmissionRouteSummary.ascx" TagName="SubmissionRouteSummary" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
	<mp:content id="region1" runat="server">
		<epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
	</mp:content>
	<mp:content id="region2" runat="server">
		<cms:helplink class="helplink" id="helplink" runat="server" helpref="40700"></cms:helplink>
		<H1>
			<cms:cmsBoilerplate id="CmsBoilerplate1" runat="server" cmsref="619"></cms:cmsBoilerplate></H1>
		<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
		<asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" CssClass="validationsummary"
			headertext="Input errors:" displaymode="BulletList"></asp:validationsummary>
		<H2 class="form_heading">Main loan details</H2>
		<epsom:loandetailssummary id="ctlLoanDetailsSummary" runat="server" DisplayOnly="true"></epsom:loandetailssummary>
		<epsom:submissionroutesummary id="ctlSubmissionRouteSummary" runat="server"></epsom:submissionroutesummary>
		<asp:panel id="pnlSelectProducts" runat="server">
			<P><BR>
				<asp:label id="lblSelectProducts" runat="server" text="You will need to select your mortgage products before proceeding"></asp:label></P>
			<asp:button id="cmdSelectProducts" runat="server" text="Clear products and make new selection"
				cssclass="button"></asp:button>
		</asp:panel>
		<asp:panel id="pnlComponents" runat="server">
			<H2 class="form_heading">Product components</H2>
			<epsom:productsummary id="ctlProductSummary" runat="server"></epsom:productsummary>
			<BR>
			<P class="button_orphan">
				<asp:button id="cmdAddComponent" runat="server" text="Add another component" cssclass="button"></asp:button>
				</P>
		</asp:panel>
		<asp:customvalidator id="ComponentTotalValidator" runat="server" text="***" enableclientscript="false"
					validateemptytext="True" onservervalidate="AreComponentsValid"></asp:customvalidator>
		<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server" buttonbar="true"></epsom:dipnavigationbuttons>
	</mp:content>
</mp:contentcontainer>
