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

    <!--start: flash detection-->
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
          document.write('<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"');
          document.write('codebase="http://active.macromedia.com/flash2/cabs/swflash.cab#version=4,0,0,0" id="masthead_A" width="990" height="220" />');
          document.write('<param name="movie" value="<%=ResolveUrl("~/_gfx/masthead_A.swf")%>" />');
          document.write('<param name="allowScriptAccess" value="sameDomain" />');
          document.write('<param name="quality" value="high" />');
          document.write('<param name="bgcolor" value="#ffffff">');
          document.write('<embed name="masthead_A" src="<%=ResolveUrl("~/_gfx/masthead_A.swf")%>" quality="high"');
          document.write('width="990" height="220" allowScriptAccess="sameDomain"');
          document.write('type="application/x-shockwave-flash"');
          document.write('pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" />');
          document.write('</embed>');
          document.write('</object>');
        }
        else {
          document.write('<img src="<%=ResolveUrl("~/_gfx/masthead_A.jpg")%>" width="990" height="220" alt="dbmortgages">');
        }
        //-->
        </script><div id="menubar">
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
          <p>
            Copyright © Deutsche Bank AG
          </p>
          <p>
            This website and the content contained in it is for use in the United Kingdom by professional 
            mortgage intermediaries only. It is not suitable for use by the general public.
          </p>
          <p>
            If you are not a professional mortgage intermediary, you should not rely on the information
            contained herein but should seek independent advice. If you reproduce any information on this
            website for the purpose of advising private clients, you must ensure that it conforms to the
            advising and selling rules of the Financial Services Authority in the United Kingdom.
          </p>
          <p>
            DB UK Bank Limited
            <br />
            Registered office: 23 Great Winchester Street, London EC2P 2AX. Registered in England and Wales
            with Company No. 315841, authorised and regulated by the Financial Services Authority and a member
            of the London Stock Exchange.
          </p>
        </div>

      </div>

    </form>

  </body>

</html>
