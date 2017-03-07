<%@ Control Language="c#" AutoEventWireup="false" Codebehind="AdviserSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.AdviserSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/ApplicantSummary.ascx" TagName="ApplicantSummary" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/UserProfileFirm.ascx" TagName="UserProfileFirm" TagPrefix="Epsom" %>

<itemtemplate>

  <fieldset>

    <legend>Adviser details</legend>

    <table class="displaytable_small">
      <epsom:userprofilefirm id="ctlAdviserFirm"  runat="server" />
     	<tr>
  		  <td class="prompt">First name/Last name</td>
	  	  <td><asp:label id="lblPersonName" runat="server" /></td>
  	  </tr>
     	<tr>
	  	  <td class="prompt">Address</td>
		    <td><asp:label id="lblAddress" runat="server" /></td>
    	</tr>
    </table>

    <table class="displaytable">
      <tr>
        <td style="text-align:right"><asp:button id="cmdEditAdviser" runat="server" text="Edit"  cssclass="button"/></td>
      </tr>
    </table>

  </fieldset> 

</itemtemplate>
