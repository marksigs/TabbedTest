<%@ Page language="c#" Codebehind="Evaluation.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.Evaluation" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">

    <h1>Evaluation</h1>

    <p>
      Please wait while our system evaluates your submission - this may take up to a minute at busy
      times. Please do not navigate away from this page.
    </p>    
      
    <asp:image runat="server" imageurl="~/_gfx/evaluating.gif" width="386" height="22" alternatetext="Evaluating mortgage decision" id="imgEvaluating"/>
           
  </mp:content>

</mp:contentcontainer>

