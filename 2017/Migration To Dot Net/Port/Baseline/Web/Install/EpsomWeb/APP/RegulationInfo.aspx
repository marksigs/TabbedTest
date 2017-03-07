<%@ Page language="c#" Codebehind="RegulationInfo.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.RegulationInfo" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/MinimumCriteria.ascx" TagName="MinimumCriteria" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id=region1 runat="server">

    <epsom:dipmenu id=DipMenu runat="server"></epsom:dipmenu></mp:content>

  <mp:content id=region2 runat="server">
  
    <sstchur:smartscroller id="ctlSmartScroller" runat="server" />

    <h1>Regulation Information</h1>

    <!--
        <p>
  		    <asp:button id="cmdTestDipAccepted" runat="server" text="Load Dip Accepted Test Data " cssclass="button" /> 
  		    <asp:button id="cmdTestSchemeOnly" runat="server" text="Load Dip Scheme Only Test Data " cssclass="button" /> 
        </p>
      -->

    <epsom:dipnavigationbuttons id=ctlDIPNavigationButtons1 runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" cssclass="validationsummary"
      headertext="Input errors:" displaymode="BulletList">
    </asp:validationsummary>

    <h2>Adviser's declaration</h2>

    <table class=formtable>
      <tr>
        <td style="padding: 0 0.5em 0 1.5em">Is this application for a regulated loan?</td>
        <td>
          <asp:radiobuttonlist id="rblRegulatedLoan" runat="Server" repeatdirection="Horizontal" repeatlayout="Flow">
            <asp:listitem value="True">Yes</asp:listitem>
            <asp:listitem value="False">No</asp:listitem>
          </asp:radiobuttonlist>
        </td>
      </tr> 
      <tr>
        <td style="padding: 0 0.5em 0 1.5em">What is the level of service for this application?</td>
        <td>
          <asp:radiobuttonlist id="rblLevelOfService" runat="Server" repeatdirection="Horizontal" repeatlayout="Flow">
            <asp:listitem value="True">Advised</asp:listitem>
            <asp:listitem value="False">Non advised</asp:listitem>
          </asp:radiobuttonlist>
        </td>
      </tr> 
      <tr>
        <td style="padding: 0 0.5em 0 1.5em" rowspan="2">Will any of the procuration fee/commission paid for the introduction
        of the mortgage be passed on to the customer?</td>
        <td>
          <asp:radiobuttonlist id="rblFeeToCustomer" runat="Server" autopostback="true" repeatdirection="Horizontal" repeatlayout="Flow">
            <asp:listitem value="None">None</asp:listitem>
            <asp:listitem value="Part">Part</asp:listitem>
            <asp:listitem value="All">All</asp:listitem>
          </asp:radiobuttonlist>
        </td>
      </tr> 
      <tr>
        <td>
          <asp:panel runat="server" id="pnlFeeToCustomer">
            If yes, how much? 
            <asp:textbox runat="server" CssClass="text" id="txtFeeToCustomer" columns="10"></asp:textbox> 
          </asp:panel>
        </td>
      </tr>
      <tr>
        <td style="padding: 0 0.5em 0 1.5em">Has the customer received and accepted a Key Facts Illustration?
        </td>
        <td>
          <Epsom:NullableYesNo id="nynCustomerAcceptedKFI" runat="Server" />
        </td>
      </tr> 
    </table>

    <epsom:minimumcriteria id="ctlMinimumCriteria" runat="server"></epsom:minimumcriteria>

    <epsom:dipnavigationbuttons buttonbar="true" id=ctlDIPNavigationButtons2 runat="server"></epsom:dipnavigationbuttons>
    
  </mp:content>

</mp:contentcontainer>
