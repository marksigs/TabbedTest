<%@ Page language="c#" Codebehind="RegistrationConfirmation.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.RegistrationConfirmation" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/AddressDisplay.ascx" TagName="AddressDisplay" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumberDisplay.ascx" TagName="TelephoneNumberDisplay" TagPrefix="Epsom" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Register</h1>

    <h2 class="form_heading">Confirm Details</h2>

    <p class="forminfo">
      Thank you for registering with db mortgages. An email has been sent to the email address you provided
      containing a summary of all the details you have entered.
      If you work for another firm please edit your profile to add that firm
      as you will be prompted to select your firm when you login next.
      To change your firm simply log out and log back in.
    </p>

    <fieldset>

      <legend>FSA details</legend>

      <table class="displaytable_small">
	      <tr>
		      <td class="prompt">Firm FSA ref. number:</td>
		      <td><asp:Label id="lblFsaNumber" runat="server" /></td>
	      </tr>
	      <tr>
		      <td class="prompt">Company name:</td>
		      <td><asp:Label id="lblFsaCompanyName" runat="server" /></td>
	      </tr>
      <tr>
		      <td class="prompt">Trading As:</td>
		      <td><asp:Label id="lblTradingAs" runat="server" /></td>
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
		      <td class="prompt">Date of birth:</td>
		      <td><asp:Label id="lblDateBirth" runat="server" /></td>
	      </tr>
      </table>
      
    </fieldset>

    <fieldset>

      <legend>Correspondence address & contact numbers</legend>

      <epsom:AddressDisplay id="ctlAddress" runat="server" />
      <epsom:telephonenumberdisplay id="ctlTelephoneNumbers" runat="server" />

    </fieldset>

    <fieldset>

      <legend>Login details</legend>

      <table class="displaytable_small">
	      <tr>
		      <td class="prompt">User name <br /> (email address):</td>
  		    <td><asp:Label id="lblEmailAddress" runat="server" /></td>
	      </tr>
  	    <tr>
	  	    <td class="prompt">Password:</td>
		      <td>************</td>
  	    </tr>
      </table>

    </fieldset>

    <p class="dipnavigationbuttonswithbar">
      <asp:Button ID="cmdDone" runat="server" text="Done" CssClass="button" /> 
      &nbsp;
      <asp:Button ID="cmdEditProfile" runat="server" text="Edit Profile" CssClass="button" /> 
    </p>
        
  </mp:content>

</mp:contentcontainer>
