<%@ Page language="c#" Codebehind="Help.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Help" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<HTML>
  <HEAD>
    <title>DBM - Help</title>
    <meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" Content="C#">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/layout.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/elements.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/hacks.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/presentation.css")%>" />
    <link rel="stylesheet" type="text/css" media="print" href="<%=ResolveUrl("~/_css/print.css")%>" />

  </HEAD>
  
  <body style="padding: 3em"> <!--LANCE - someone should add this as a proper styling class in the css files-->
  
    <h1 align="left"><asp:label id="labelHeading" runat="server" text="Sorry !" /></h1>

    <br>
    
    <p align="left"><asp:label id="labelBody" runat="server" text="There is currently no help available for this page." /></p>
    
    <br /><br />
    
    <p align="right"><a href="javascript:window.parent.close()">close</a></p>
    
  
  </body>
</HTML>
