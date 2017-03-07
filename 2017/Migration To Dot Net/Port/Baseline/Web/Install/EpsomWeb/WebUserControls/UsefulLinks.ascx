<%@ Control Language="c#" AutoEventWireup="false" Codebehind="UsefulLinks.ascx.cs" Inherits="Epsom.Web.WebUserControls.Links" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<CMS:CmsRepeater id="rprLinks" type="Useful Links"  Runat="server">
  <HeaderTemplate>
    <h2><asp:label ID="heading" Runat="server"></asp:label></h2>

    <script language="javascript">

      var newUrlwindow;

      function openURL(url)
      {
		      newUrlwindow=window.open(url);
		      if (window.focus) { newUrlwindow.focus(); }
      }

      var warningMessage = "Important note: When you select a hyperlink from this page, you may be leaving a Deutsche Bank ('the Bank') controlled area and moving to an external Internet website. The information provided on any such website (and any website accessed through it other than the Bank's websites) has been produced by independent providers and the Bank does not endorse or accept any responsibility for it, nor for any opinions or recommendations expressed on it. The existence of a hyperlink on this page does not constitute a recommendation or other approval by the Bank of such website or its providers.";	
      function openURLwarning(url)
      {
	      if(confirm(warningMessage))
	      {
		      newUrlwindow=window.open(url,'name');
		      if (window.focus) { newUrlwindow.focus(); }
	      }
	      else {}
      }

    </script>

    <ul class="linklist">
  </HeaderTemplate>

  <ItemTemplate>
    <li>
      <asp:hyperlink id="lnkLink" runat="server"></asp:hyperlink>
    </li>
  </ItemTemplate>

  <FooterTemplate>
    </ul>
  </FooterTemplate>

</CMS:CmsRepeater>