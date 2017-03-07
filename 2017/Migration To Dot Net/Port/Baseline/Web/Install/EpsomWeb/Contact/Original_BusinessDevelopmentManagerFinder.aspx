<%@ Page language="c#" Codebehind="BusinessDevelopmentManagerFinder.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Contact.BusinessDevelopmentManagerFinder" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/BusinessDevelopmentManagerDisplay.ascx" TagName="BusinessDevelopmentManagerDisplay" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>BDM Finder</h1>
    
    <div style="float:left; padding: 2em 2em 0 0; font-size: 0.9em; width:15em">

      <p>
        Find a BDM close to you:
      </p>      

      <p style="text-align: right">
        Postcode: <asp:TextBox ID="txtPostcode" runat="server" size="8" CssClass="text" />
        <br />
        <asp:Button runat="server" commandName="cmdSearchPostcode" text="Search" id="cmdSearchPostcode"/>
      </p>

      <br /><br />

      <asp:Panel id="pnlBDMAccounts" runat="server">

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
 
      </asp:Panel>
    
    </div>

    <p>
      Text to display on page entry
    </p>

    <br />

    <p id="pBusinessDevelopmentManagers" runat="server">
      <asp:label id="lblBusinessDevelopmentManagers" runat="server" />
    </p>

    <br />

    <asp:DataList id="dlBusinessDevelopmentManagers" RepeatColumns="1" runat="server" RepeatDirection="vertical" RepeatLayout="table">
      <ItemTemplate>
        <epsom:BusinessDevelopmentManagerDisplay id="ctlBusinessDevelopmentManagerDisplay" runat="server" />
        <br /><br />
      </ItemTemplate>
    </asp:DataList> 
  
  </mp:content>

</mp:contentcontainer>

