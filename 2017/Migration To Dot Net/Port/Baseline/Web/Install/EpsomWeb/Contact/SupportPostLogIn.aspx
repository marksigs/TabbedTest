<%@ Page language="c#" Codebehind="SupportPostLogIn.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Contact.SupportPostLogIn" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">
<mp:content id="region1" runat="server">

<asp:panel id="pnlSupport" Runat="server">
<h1>Support</h1>
<asp:validationsummary runat="server" headertext="Input errors:" CssClass="validationsummary" showsummary="true" displaymode="BulletList"></asp:validationsummary>
<table class="formtable">
  <tr>
    <td class="prompt">Type of support</td>
    <td>
      <asp:dropdownlist id="cmbCategory" runat="server" />
    </td>
  </tr>
  <tr>
    <td class="prompt">Sub-category</td>
    <td>
      <asp:dropdownlist id="cmbSubCategory" runat="server" />
    </td>
  </tr>
</table>
<asp:panel id="pnlContent" Runat="server">
<hr>
1.CHECK THE TUTORIAL<br>
<CMS:Literal runat="server" cmsref="130" />
<asp:LinkButton id="cmdTutorial" text="View Tutorial" runat="server"></asp:LinkButton>
<hr>
2.PHONE US<br>
<CMS:Literal runat="server" cmsref="131" />
<hr>
3.SEND US AN EMAIL<br>
</asp:panel>
<table class="formtable">
  <tr>
    <td class="prompt">Case Ref. Number</td>
    <td>
      <asp:TextBox ID="txtCaseRef" Runat="server" CssClass="text" style="width: 16em" />
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Comments</td>
    <td>
      <asp:TextBox ID="txtComments" TextMode="MultiLine" Columns="60" Rows="12" Runat="server" />
        <asp:RequiredFieldValidator Runat="Server"
          ErrorMessage="Please enter your comments" ControlToValidate="txtComments" text="***" />
    </td>
  </tr>
  <tr>
    <td class="prompt">&nbsp;</td>
    <td>
      <asp:Button id="cmdSend" Runat="server" text="Send" CssClass="button" />
    </td>
  </tr>
</table>
</asp:panel>
<asp:panel id="pnlError" Runat="server" Visible="false">
<table class="formtable">
  <tr><td colspan="4">
    There has been an error sending the email.
    Please try again later, or contact by telephone.
  </td></tr>
</table>
</asp:panel>

<asp:panel id="pnlConfirmation" Runat="server" Visible="false">
<h1>Contact us confirmation</h1>
<p>
<asp:Label id="lblForename" Runat="server"> </asp:Label>
</p>
<p>
<CMS:Literal runat="server" cmsref="132" />
</p>
<p>
<b>Category: <asp:Label id="lblCategory" Runat="server"> </asp:Label></b>
</p>
<p>
<b>Sub-category: <asp:Label id="lblSubCategory" Runat="server" ></asp:Label></b>
</p>
<p>
Case ref.: <asp:Label id="lblCaseRef" Runat="server"> </asp:Label>
</p>
<p>
Comments: <asp:Label id="lblComments" Runat="server" ></asp:Label>
</p>

</asp:panel>


  
</mp:content>
</mp:contentcontainer>
