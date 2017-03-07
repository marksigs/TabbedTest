<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ApplicantOtherNameDisplay.ascx.cs" Inherits="Epsom.Web.WebUserControls.ApplicantOtherNameDisplay" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<table class="displaytable_small">

  <tr>
    <td class="prompt">Title</td>
    <td><asp:Label ID="lblPreviousTitle" Runat="server" /></td>
  </tr>
  
  <tr>
    <td class="prompt">First name</td>
    <td><asp:Label ID="lblPreviousFirstName" Runat="server" /></td>
  </tr>
  
  <tr>
    <td class="prompt">Middle name</td>
    <td><asp:Label ID="lblPreviousMiddleName" Runat="server" /></td>
  </tr>
  
  <tr>
    <td class="prompt">Surname</td>
    <td><asp:Label ID="lblPreviousSurname" Runat="server" /></td>
  </tr>
  
  <tr>
    <td class="prompt">Stopped being used</td>
    <td><asp:Label ID="lblDateLastUsed" Runat="server" /></td>
  </tr>

</table>
