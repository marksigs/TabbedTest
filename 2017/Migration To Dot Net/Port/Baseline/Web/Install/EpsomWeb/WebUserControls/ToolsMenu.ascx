<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ToolsMenu.ascx.cs" Inherits="Epsom.Web.WebUserControls.ToolsMenu" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%@ Register Src="~/WebUserControls/LogOnDetails.ascx" TagName="LogOnDetails" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/UsefulLinks.ascx" TagName="Links" TagPrefix="Epsom" %>

<asp:Panel id="pnlRestrictedItems" Runat="server">

<!--
  <div class="sidebox">

    <h2>Login details</h2>

    <epsom:LogOnDetails id="ctlLogOnDetails" runat="server"></epsom:LogOnDetails>

  </div>
-->

</asp:Panel>

<div class="sidebox">

  <h2>Calculators</h2>

  <ul class="linklist">
    <li>
      <a href="<%=ResolveUrl("~/Calculators/DisposableIncomeCalculator.aspx")%>" onclick="window.open(this.href, 'popupwindow', 'width=700,height=500,scrollbars,resizable'); return false;" target="_blank">Disposable income calculator</a>
    </li>
    <li>
      <a href="<%=ResolveUrl("~/Calculators/WebCalc.aspx")%>" onclick="window.open(this.href, 'popupwindow', 'width=700,height=500,scrollbars,resizable'); return false;" target="_blank">Flexible Calculator</a>
    </li>
  </ul>
  
</div>

<div class="sidebox">
  
  <epsom:Links id="ctlUsefulLinks" type="Useful links" runat="server" />
  
</div>

<div class="sidebox">

  <h2>Acrobat reader</h2>

  <p>
    Download the latest version of the Acrobat Reader
  </p>
  <p><a  target="_blank"  href="http://www.adobe.com/downloads/" class="fancylink">Acrobat Reader</a></p>


<!--  <p><button  onclick="javascript:window.open('http://www.adobe.com/downloads/');" class="fancylink">Acrobat Reader</button></p>
-->
</div>
