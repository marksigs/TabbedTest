<%@ Page language="c#" Codebehind="DefaultErrorPage.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Errors.DefaultErrorPage" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Problem encountered.</h1>
    <asp:panel id="pnlEnforcedLogOff" runat="server">
    <p>
      You have been logged out of the DB Mortgages website because it is undergoing maintenance or experiencing technical problems.
    </p>
    <p>
      We apologise for the inconvenience and recommend you try again later. In the event that it fails again, please contact our Helpdesk.
    </p>		
    </asp:panel>

    <asp:panel id="pnlLogOffOrRetry" runat="server">
    <p>
      The DB Mortgages website is undergoing maintenance or experiencing technical problems.
    </p>
    <p>
      We apologise for the inconvenience and recommend you logout and try again later. 
      In the event that it fails again, please contact our Helpdesk.
    </p>		
    <p class="button_orphan">
      <asp:button id="cmdLogOff" runat="server" cssclass="button" text="Log Out" />
      <asp:button id="cmdRetry" runat="server" cssclass="button" text="Retry" />
    </p>

    </asp:panel>
    
  </mp:content>

</mp:contentcontainer>
