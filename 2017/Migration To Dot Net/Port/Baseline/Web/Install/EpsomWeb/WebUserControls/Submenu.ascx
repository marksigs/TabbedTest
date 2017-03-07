<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Submenu.ascx.cs" Inherits="Epsom.Web.WebUserControls.Submenu" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>

<asp:Panel id="pnlUnrestrictedList" Runat="server">
  <ul id="submenu">
    <li>
      <asp:hyperlink id="lnkHome_Unrestricted" navigateurl="~/Default.aspx" runat="server" text="Home" />
    </li>
    <li>
      <asp:hyperlink id="lnkCalculators_Unrestricted" navigateurl="~/Calculators/Calculators.aspx" runat="server" text="Calculators" />
    </li>
    <li>
      <asp:hyperlink id="lnkNews_Unrestricted" navigateurl="~/News/News.aspx" runat="server" text="News" />
    </li>
    <li>
      <asp:hyperlink id="lnkDownloads_Unrestricted" navigateurl="~/Downloads/Downloads.aspx" runat="server" text="Downloads" />
    </li>
    <li>
      <asp:hyperlink id="lnkBDMFinder_Unrestricted" navigateurl="~/Contact/BusinessDevelopmentManagerFinder.aspx" runat="server" text="BDM Finder" />
    </li>
  </ul>
</asp:Panel>

<asp:Panel id="pnlRestrictedList" Runat="server">
  <ul id="submenu">
    <li>
      <asp:hyperlink id="lnkHome_Restricted" navigateurl="~/Default.aspx" runat="server" text="Home" />
    </li>
    <li>
      <asp:linkButton id="lnkDIP_Restricted"  runat="server" text="DIP" />
    </li>
    <li>
      <asp:linkButton id="lnkKFI_Restricted"  runat="server" text="KFI" />
    </li>
    <li>
      <asp:hyperlink id="lnkCaseTracker_Restricted" navigateurl="~/CaseTracker/Default.aspx" runat="server" text="Case tracking" />
    </li>
    <li>
      <asp:hyperlink id="lnkCalculators_Restricted" navigateurl="~/Calculators/Calculators.aspx" runat="server" text="Calculators" />
    </li>
    <li>
      <asp:hyperlink id="lnkNews_Restricted" navigateurl="~/News/News.aspx" runat="server" text="News" />
    </li>
    <li>
      <asp:hyperlink id="lnkDownloads_Restricted" navigateurl="~/Downloads/Downloads.aspx" runat="server" text="Downloads" />
    </li>
    <li>
      <asp:hyperlink id="lnkBDMFinder_Restricted" navigateurl="~/Contact/BusinessDevelopmentManagerFinder.aspx" runat="server" text="BDM Finder" />
    </li>
  </ul>
</asp:Panel>
