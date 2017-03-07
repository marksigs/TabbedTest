<%@ Control Language="c#" AutoEventWireup="false" Codebehind="MasterPage2.ascx.cs" Inherits="Epsom.Web.MasterPages.MasterPage2" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
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
    <noscript>Javascript must be enabled to use this site</noscript>
      <script language="javascript" type="text/javascript" src="<%=ResolveUrl("~/_js/flash_detect.js")%>">
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
    <!--end: flash detection-->
<noscript>Javascript must be enabled to use this site</noscript>
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
        //if (flashVersion >= requiredVersion) {
        //  document.write('<object width="990" height="71" >');
        //  document.write('<param >');
        //  document.write('<embed src="<%=ResolveUrl("~/WebUserControls/DisplayFlashMovie.aspx?cmsRef=5016")%>" width="990" height="71" >' );
        //  document.write('</embed>');
        //  document.write('</object>');
        //
        //}
        //else {
          document.write('<img src="<%=ResolveUrl("~/WebUserControls/DisplayImage.aspx?cmsRef=5015")%>"  width="990" height="71" alt="dbmortgages">');
        //}
        //-->
        </script><noscript>Javascript must be enabled to use this site</noscript><div id="menubar">
          <epsom:submenu id="submenu" runat="server" />
        </div>

        <div id="columns">
          <div class="cols-wrapper">
            <div class="float-wrapper">
              <div id="col2" class="minpageheight">
                <div class="main-content">

                  <mp:region runat="server" id="region2" />

                </div>
              </div>

              <div id="col1" class="sidecol">

                <div class="sidebox2">
                  <mp:region runat="server" id="region1" />
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
              <img src="<%=ResolveUrl("~/_gfx/w3c-wai.gif")%>" width="80" height="15" border="none" alt="W3C compliance">
           </a>
          </p>
        </div>

      </div>

    </form>

  </body>

</html>

