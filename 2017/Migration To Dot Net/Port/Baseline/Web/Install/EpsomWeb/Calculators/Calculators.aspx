<%@ Page language="c#" Codebehind="Calculators.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Calculators" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Calculators</h1>

    <p>
      <a href="DisposableIncomeCalculator.aspx" onclick="window.open(this.href, 'popupwindow', 'width=700,height=500,scrollbars,resizable'); return false;" target="_blank">Disposable Income Calculator</a>
    </p>

    <p>
      <a href="WebCalc.aspx" onclick="window.open(this.href, 'popupwindow', 'width=700,height=500,scrollbars,resizable'); return false;" target="_blank">Flexible Calculator</a>
    </p>


		
  </mp:content>

</mp:contentcontainer>
