<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Page language="c#" Codebehind="EmploymentDetails1.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.EmploymentDetails1" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentPermanentDetails.ascx" TagName="EmploymentPermanentDetails" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentContractDetails.ascx" TagName="EmploymentContractDetails" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentPrevious.ascx" TagName="EmploymentPrevious" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentSelfEmployedDetails.ascx" TagName="EmploymentSelfEmployedDetails" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentLength.ascx" TagName="EmploymentLength" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
  <mp:content id="region1" runat="server">
    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
  </mp:content>
  <mp:content id="region2" runat="server">
    <H1>Employment details : <asp:label id="lblCandidate" runat="server" text=""></asp:label></H1>
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
    <asp:validationsummary id="Validationsummary1" runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true"></asp:validationsummary>
    <H2 class="form_heading">Type of employment</H2>
    <TABLE class="formtable">
 <TR>
   <TD class="prompt"><asp:label id="lblPermanentEmployee" runat="server" text="Selected employment type "></asp:label></TD>
   <TD><asp:dropdownlist disabled   cssclass="text" id="ddlPermanentEmployee" runat="Server" autopostback="true"></asp:dropdownlist></TD>
 </TR>                                              
 </TABLE>

<epsom:employmentpermanentdetails id="ctlPermanentDetails" runat="server"></epsom:employmentpermanentdetails>
<!-- Permanent employment only section -->
<!--<epsom:employmentcontractdetails id="ctlContractDetails" runat="server"></epsom:employmentcontractdetails>-->
<!-- contract only section cont   -->
<epsom:employmentlength id="ctlEmploymentLength"  runat="server"></epsom:employmentlength>
<!--??????? only section -->
<epsom:employmentselfemployeddetails id="ctlSelfEmployed" runat="server"></epsom:employmentselfemployeddetails>
<!-- self employed only section only section -->
<!--<epsom:employmentprevious id="ctlEmploymentPrevious" runat="server"></epsom:employmentprevious>-->

<epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>
<sstchur:SmartScroller id="ctlSmartScroller" runat="server"></sstchur:SmartScroller>
  </mp:content>
</mp:contentcontainer>
</mp:content></mp:contentcontainer>
