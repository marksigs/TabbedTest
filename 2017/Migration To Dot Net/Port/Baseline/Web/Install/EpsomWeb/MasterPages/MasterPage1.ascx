<%@ Control Language="c#" AutoEventWireup="false" Codebehind="MasterPage1.ascx.cs" Inherits="Epsom.Web.MasterPages.MasterPage1" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/MainMenu.ascx" TagName="MainMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/SubMenu.ascx" TagName="SubMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ToolsMenu.ascx" TagName="ToolsMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/LoginAndRegistrationMenu.ascx" TagName="LoginAndRegistrationMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/UsefulLinks.ascx" TagName="Links" TagPrefix="Epsom" %>

<!DOCTYPE html PUBLIC "-//w3c//dtd xhtml 1.0 transitional//en" "http://www.w3.org/tr/xhtml1/dtd/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

  <head>
    <title>DB Mortgages</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <meta http-equiv="distribution" content="global" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/layout.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/elements.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/hacks.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/presentation.css")%>" />
    <link rel="stylesheet" type="text/css" media="print" href="<%=ResolveUrl("~/_css/print.css")%>" />
    <link rel="shortcut icon" type="image/ico" href="<%=ResolveUrl("~/_gfx/favicon.ico")%>" />

    <!--start: flash detection-->
      <script language="javascript" type="text/javascript" src="<%=ResolveUrl("~/_js/flash_detect.js")%>" >
        <!--
        function getFlashVersion() { return null; };
        //-->
      </script>
      <script language="javascript" type="text/javascript" src="<%=ResolveUrl("~/_js/flash_cookie.js")%>">
        <!--
        var flashVersion = 0;
        var dontKnow = true;
        //-->
      </script>
      <noscript>Javascript must be enabled to use this site</noscript>
    <!--end: flash detection-->
    <script language="javascript" type="text/javascript">
      <!--
        var newwindow;
        function pophelp(url)
        {
	        newwindow=window.open(url,'Help','height=350,width=350,left=150,top=250,scrollbars=yes,dependent=1,resizable=no,toolbar=no,status=no,menubar=no,locationbar=no');
	        if (window.focus) { newwindow.focus(); }
        }
      //-->
    </script>
    <noscript>Javascript must be enabled to use this site</noscript>

  </head>

  <body>

    <form runat="server" id="form1">

      <div id="main" class="show-all">

        <p id="mainmenu">
          <epsom:mainmenu id="topmenu" runat="server" />
        </p>

        <script type="text/javascript">
        <!--
        var requiredVersion = 5;
        if (flashVersion >= requiredVersion) {
          document.write('<object width="990" height="220">');
          document.write('<param >');
          document.write('<embed src="<%=ResolveUrl("~/WebUserControls/DisplayFlashMovie.aspx?cmsRef=5014")%>" width="990" height="220" >' );
          document.write('</embed>');
          document.write('</object>');
          //document.write('<object id="masthead" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"');
          //document.write('codebase="https://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0" id="masthead_A" width="990" height="220" />');
          //document.write('<param name="movie" value="<%=ResolveUrl("~/WebUserControls/DisplayImage.aspx?cmsRef=5014")%>" />');
          //document.write('<param name="allowScriptAccess" value="sameDomain" />');
          //document.write('<param name="quality" value="high" />');
          //document.write('<param name="bgcolor" value="#ffffff">');
          //document.write('<embed name="masthead_A"  " quality="high"');
          //document.write('width="990" height="220" allowScriptAccess="sameDomain"');
          //document.write('type="application/x-shockwave-flash"');
          //document.write('pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" />');
          //document.write('</embed>');
          //document.write('</object>');
        }
        else {
          document.write('<img src="<%=ResolveUrl("~/WebUserControls/DisplayImage.aspx?cmsRef=5013")%>" width="990" height="220" alt="dbmortgages">');
        }
        //-->
        </script><noscript>Javascript must be enabled to use this site</noscript><div id="menubar">
          <epsom:submenu id="submenu" runat="server" />
        </div>

        <div id="columns">
          <div class="cols-wrapper">
            <div class="float-wrapper">
              <div id="col2" class="minpageheight">
                <div class="main-content">

                  <mp:region runat="server" id="region1" />

                </div>
              </div>

              <div id="col1" class="sidecol">

                <div class="sidebox2">
                  <epsom:loginandregistrationmenu id="LoginAndRegistrationMenu1" runat="server" />
                </div>

              </div>

            </div>

            <div id="col3" class="sidecol">

              <epsom:toolsmenu id="ToolsMenu1" runat="server" />

            </div>

            <div class="clear" id="em"></div>

          </div>

        </div>

        <div id="footer" class="clear">
          <p class="pagename">
            Copyright © Deutsche Bank AG
          </p>
          <p class="footerlinks">
            <a href="<%=ResolveUrl("~/Help/TermsAndConditions.aspx")%>" onclick="window.open(this.href, '_blank', 'width=1000,height=750,scrollbars,resizable'); return false;" target="_blank">Terms and conditions of use</a>
            &nbsp;|&nbsp;
            <a href="<%=ResolveUrl("~/Help//PrivacyPolicy.aspx")%>" onclick="window.open(this.href, '_blank', 'width=1000,height=750,scrollbars,resizable'); return false;" target="_blank">Privacy policy</a>
          </p>
          <p class="w3clogo">
           <a href="http://www.w3.org" onclick="window.open(this.href, '_blank', 'width=1000,height=750,scrollbars,resizable'); return false;" target="_blank">
              <img src="<%=ResolveUrl("~/_gfx/w3c-wai.gif")%>" width="80" height="15" alt="W3C compliance" border="none">
           </a>
          </p>
          <br />
          <p>db mortgages is a trading name of DB UK Bank Limited.</p>
          <p>Authorised and Regulated by the Financial Services Authority and a member of The London Stock Exchange.</p>
          <p>Registered in England and Wales No. 315841. Registered Office: 23 Great Winchester Street, London, EC2P 2AX.</p>
        </div>

      </div>

    </form>

  </body>

</html>
