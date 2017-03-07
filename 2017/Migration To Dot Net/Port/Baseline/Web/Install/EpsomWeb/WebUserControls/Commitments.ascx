<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Commitments.ascx.cs" Inherits="Epsom.Web.WebUserControls.Commitments" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Commitment.ascx" TagName="commitment" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

        <p>
          <cms:cmsBoilerplate cmsref="160" runat="server" />
        </p>
  <asp:Repeater id="rptCommitments" runat="server" >
    <ItemTemplate>
      <fieldset>
        <legend><asp:Literal ID="lblCommitmentTitle" Runat="server"></asp:Literal></legend>
        <epsom:commitment id="ctlCommitment" runat="server" /> 
        </fieldset> 
    </ItemTemplate>
  </asp:Repeater>
  
  
  <p class="button_orphan">
    <cms:cmsBoilerplate cmsref="161" runat="server" />
  </p>
   
  <p class="button_orphan">
    <asp:Button ID="cmdRemoveCommitment" Runat="server" Text="Remove"></asp:Button>
    <asp:Button ID="cmdAddCommitment" Runat="server" Text="Add commitment"></asp:Button>
  </p>
