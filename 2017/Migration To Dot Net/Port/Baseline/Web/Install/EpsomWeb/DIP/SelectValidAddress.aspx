<%@ Page language="c#" Codebehind="SelectValidAddress.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Dip.SelectValidAddress" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SelectValidAddress.ascx" TagName="SelectValidAddress" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>


<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <script language="javascript">
      function ProcessSubmit()
      {
        var mainLayer = document.getElementById('mainLayer');
        var holdingLayer = document.getElementById('holdingLayer');
        mainLayer.style.display = 'none';
        holdingLayer.style.display = 'block';
        setTimeout('document.getElementById("waitImage").src = "<%=ResolveUrl("~/_gfx/evaluating.gif")%>";', 50);
        scroll(0,0);
      }
    </script>

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2100" />


    <h1>Select valid address</h1>

    <div id="mainLayer">

      <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

      <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

      <p>
        The address(s) you have provided has/have not been recognised by our credit reference agency.
        Please select the correct address from the options shown below.
      </p>

      <epsom:SelectValidAddress id="ctlSelectValidAddress" runat="server" />
      
      <epsom:dipnavigationbuttons buttonbar="true" id="ctlDIPNavigationButtons2" runat="server" />

    </div>
    
    <div id="holdingLayer" style="display:none">

      <p>
        Please wait while our system re-evaluates your submission - this may take up to a minute at busy
        times. Please do not navigate away from this page.
      </p>

      <img id="waitImage" src="<%=ResolveUrl("~/_gfx/evaluating.gif")%>" width="386" height="22" alt="Evaluating mortgage decision" />

    </div>

  </mp:content>

</mp:contentcontainer>
