<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Page language="c#" Codebehind="testdb.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.testdb" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
  <HEAD>
    <title>testdb</title>
    <meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" Content="C#">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/presentation.css")%>" />
   

  </HEAD>
  <body MS_POSITIONING="GridLayout">
  
  <div id="header">
  <img class="dblogo" src="<%=ResolveUrl("~/_gfx/db_logo.gif")%>" alt="Deutsche Bank Mortgages" />
  </div>

    <!--h1 class="form_heading"><u><CMS:Literal id="heading" cmsref="70" Runat="server">Main Heading</CMS:Literal></u></h1-->

    <h1><asp:label id="Label1" runat="server" text="Heading Text"></asp:label></h1>

    <br>
    <hr />
    
    <asp:label id="Label2" runat="server" text="Body Text"></asp:label>
    </p>
    <br /><br />
    <p align="center"><a href="javascript:window.parent.close()">close</a></p>
    
    
    <!-- form action="testdb.aspx" method="get">
      <input type="text" name="helpref"> <input type="submit" value="Submit Query">
    </form-->
    
  </body>
</HTML>
