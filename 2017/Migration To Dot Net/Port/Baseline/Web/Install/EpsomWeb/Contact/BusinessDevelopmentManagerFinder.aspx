<%@ Page language="c#" Codebehind="BusinessDevelopmentManagerFinder.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Contact.BusinessDevelopmentManagerFinder" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/BusinessDevelopmentManagerDisplay.ascx" TagName="BusinessDevelopmentManagerDisplay" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>BDM Finder</h1>
    
      <p>
        To find a BDM close to you, simply enter your postcode:&nbsp;
        <asp:TextBox ID="txtPostcode" runat="server" size="8" CssClass="text" />
        <asp:Button runat="server" commandName="cmdSearchPostcode" text="Search" CssClass="button" id="cmdSearchPostcode"/>
      </p>      
 
     <!--  <asp:Panel id="pnlBDMAccounts" runat="server">

        <p>
          Find your account's BDM:
        </p>

        <asp:repeater id="rptBDMAccounts" runat="server">
          <itemtemplate>
            <p>
              <asp:LinkButton id="btnBDMAccount" CssClass="fancylink" runat="server" />
              <input id="hiddenIndex" type="hidden" runat="server" name="hiddenIndex"/>
            </p>
          </itemtemplate>
        </asp:repeater>
 
      </asp:Panel>-->

    <h2 class="form_heading" id="hBusinessDevelopmentManagers" runat="server">
      <asp:label id="lblBusinessDevelopmentManagers" runat="server" />
    </h2>

    <br />

    <asp:DataList id="dlBusinessDevelopmentManagers" RepeatColumns="1" runat="server" RepeatDirection="vertical" RepeatLayout="table">
      <ItemTemplate>
        <epsom:BusinessDevelopmentManagerDisplay id="ctlBusinessDevelopmentManagerDisplay" runat="server" />
        <br /><br />
      </ItemTemplate>
    </asp:DataList> 
  
  </mp:content>

</mp:contentcontainer>

