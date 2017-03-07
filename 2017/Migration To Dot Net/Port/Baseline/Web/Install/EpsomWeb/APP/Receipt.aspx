<%@ Page language="c#" Codebehind="Receipt.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.Receipt" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" > 

<html>
  <head>
    <title>Receipt</title>
    <meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" Content="C#">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/layout.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/elements.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/hacks.css")%>" />
    <link rel="stylesheet" type="text/css" media="screen" href="<%=ResolveUrl("~/_css/presentation.css")%>" />
    <link rel="stylesheet" type="text/css" media="print" href="<%=ResolveUrl("~/_css/print.css")%>" />
    <link rel="shortcut icon" type="image/ico" href="<%=ResolveUrl("~/_gfx/favicon.ico")%>" />
  </head>
  <body>
	
    <form id="Form1" method="post" runat="server">

      <table class="FormTable" width="100%">
        <tr>
          <td colspan="2">
            <table class="FormTable">
              <tr>
                <td valign="top">
                  <h2 class="form_heading">Application Payment Receipt</h2>
                  <h3><asp:label id="lblDateAndTime" runat="server" /></h3>
                </td>
                <td align="right"><img id="imgReceiptLogo" src="../_gfx/db_logo.gif" style="clip:rect(auto auto 77px auto)" runat="server" /></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
        <tr>
		  <td colspan="2" align="center">Thank you for your payment.  Please keep this receipt for your records.</td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
        <tr>
          <td class="prompt">Cardholder name</td>
          <td><asp:Label ID="lblCardholderName" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Card type</td>
          <td><asp:Label ID="lblCardType" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Card number</td>
          <td><asp:Label ID="lblCardNumber" Runat="server" /></td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
        <asp:repeater id="repeaterFees" runat="server">
          <itemtemplate>
            <tr>
              <td class="prompt"><asp:Label ID="lblFeeDescription" runat="server" /></td>
              <td><asp:Label ID="lblFeeValue" runat="server" /></td>
            </tr>
          </itemtemplate>
        </asp:repeater>
        <tr>
          <td class="prompt">Total Amount</td>
          <td><asp:Label ID="lblAmount" Runat="server" /></td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
        <tr>
          <td class="prompt">Our reference</td>
          <td><asp:Label ID="lblOurReference" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Your reference</td>
          <td><asp:Label ID="lblYourReference" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Transaction reference</td>
          <td><asp:Label ID="lblTransactionReference" Runat="server" /></td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
        <tr>
		  <td colspan="2" align="center">This payment receipt does not constitute a legally binding mortgage offer and does not oblige db mortgages to provide you with a mortgage.</td>
        </tr>
        <tr>
		  <td colspan="2"><hr /></td>
        </tr>
      </table>
      
    </form>
	
  </body>
</html>
