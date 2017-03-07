<%@ Control Language="c#" AutoEventWireup="false" Codebehind="UserProfile.ascx.cs" Inherits="Epsom.Web.WebUserControls.UserProfile" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Reference Control="UserProfileFirm.ascx"%>
<%@ Register Src="~/WebUserControls/AddressDisplay.ascx" TagName="AddressDisplay" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumberDisplay.ascx" TagName="TelephoneNumberDisplay" TagPrefix="Epsom" %>

<h2 class="form_heading">Your User Profile Summary</h2>

<p class="forminfo">You can edit your user profile details using the links below:</p>

<fieldset>
  <legend>FSA details</legend>
  <table class="displaytable_small">
  	<asp:PlaceHolder id="placeHolder1" runat="server"  />
  </table>

  <table class="displaytable">
    <tr>
      <td style="text-align:right"><asp:Button id="cmdAddFsa" runat="server" text="Add New Firm"  CssClass="button"/></td>
    </tr>
  </table>
</fieldset>

<fieldset>
  <legend>Personal Details</legend>
  <table class="displaytable_small">
  	<tr>
	  	<td class="prompt">Title:</td>
	  	<td><asp:Label id="lblTitle" runat="server" /></td>
  	</tr>
  	<tr>
	  	<td class="prompt">First name:</td>
		  <td><asp:Label id="lblFirstName" runat="server" /></td>
  	</tr>
	  <tr>
  		<td class="prompt">Last name:</td>
	  	<td><asp:Label id="lblLastName" runat="server" /></td>
  	</tr>
	  <tr>
  		<td class="prompt">Date of Birth:</td>
	  	<td><asp:Label id="lblDateBirth" runat="server" /></td>
  	</tr>
  </table>
  <table class="displaytable">
    <tr>
      <td style="text-align:right"><asp:Button id="cmdEditPersonal" runat="server" text="Edit"  CssClass="button"/></td>
    </tr>
  </table>
</fieldset>

<fieldset>

  <legend>Correspondence address & contact numbers</legend>

  <epsom:addressdisplay id="ctlAddress" runat="server" />

  <br />

  <epsom:telephonenumberdisplay id="ctlTelephoneNumbers" runat="server" ShowPreference="false"/>

  <table class="displaytable">
    <tr>
      <td style="text-align:right"><asp:Button id="cmdEditAddress" runat="server" text="Edit"  CssClass="button"/></td>
    </tr>
  </table>
</fieldset>

<fieldset>

  <legend>Login details</legend>

  <table class="formtable">
  	<tr>
	  	<td class="prompt">Email address:</td>
		  <td><asp:Label id="lblEmailAddress" runat="server" /></td>
  	</tr>
	  <tr>
  		<td class="prompt">Password:</td>
	  	<td>************</td>
  	</tr>
  </table>

  <table class="displaytable">
    <tr>
      <td style="text-align:right"><asp:Button id="cmdEditLogOn" runat="server" text="Edit"  CssClass="button"/></td>
    </tr>
  </table>

</fieldset>

<p class="dipnavigationbuttonswithbar">
  <asp:button id=cmdCancel runat="server" text="Cancel" cssclass="button" causesvalidation="false"/>
</p>
