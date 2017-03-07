<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LandlordDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.LandlordDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/LandlordDetail.ascx" TagName="LandLordDetail" TagPrefix="Epsom" %>

<asp:Repeater id="rptLandlordDetails" Runat="server">
  <ItemTemplate>
    <h2 class="form_heading">Landlord <asp:Literal ID="lblLandlordIndex" Runat="server" /></h2>
    <epsom:LandlordDetail id="ctlLandlordDetail" runat="server"/>
  </ItemTemplate>
</asp:Repeater>