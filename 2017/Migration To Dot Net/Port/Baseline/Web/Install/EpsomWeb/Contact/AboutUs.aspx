<%@ Page language="c#" Codebehind="AboutUs.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.AboutUs" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>About db mortgages</h1>

    <h2 class="form_heading">&nbsp;</h2>

    <p><cms:cmsBoilerplate cmsref="4040" runat="server" /></p>

    <h2 class="form_heading">Approved Packagers: <asp:label id="lblLetterRange" runat="server" CssClass="highlight" /></h2>
     
    <p class="forminfo">
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="A,D" Text="A-D" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="E,K" Text="E-K" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="I,L" Text="I-L" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="M,P" Text="M-P" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="Q,T" Text="Q-T" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="U,Z" Text="U-Z" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="0,9" Text="0-9" />
    </p> 

    <br />

    <p>
    
      <asp:repeater id="rptApprovedPackagers" runat="server">
        <HeaderTemplate>
          <table class="displaytable">
        </HeaderTemplate>
        <ItemTemplate>
          <tr>
            <td><%# ((Epsom.Web.Proxy.ApprovedPackager)Container.DataItem).Name %></td>
            <td><%# ((Epsom.Web.Proxy.ApprovedPackager)Container.DataItem).Phone %></td>
            <td><%# ((Epsom.Web.Proxy.ApprovedPackager)Container.DataItem).Email %></td>
          </tr>
        </ItemTemplate>
        <FooterTemplate>
          </table>
        </FooterTemplate>
      </asp:repeater>

    </p>

    <p class="forminfo">
      <asp:label id="lblNoPackagersFoundMessage" runat="server" Text="No packagers found in this letter range" />
    </p>

    <br />

    <p class="forminfo">
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="A,D" Text="A-D" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="E,K" Text="E-K" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="I,L" Text="I-L" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="M,P" Text="M-P" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="Q,T" Text="Q-T" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="U,Z" Text="U-Z" />
      &nbsp;
      <asp:linkbutton runat="server" OnCommand="DisplayLetterRange" CommandArgument="0,9" Text="0-9" />
    </p> 

  </mp:content>

</mp:contentcontainer>

