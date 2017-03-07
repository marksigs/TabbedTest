<%@ Control Language="c#" AutoEventWireup="false" Codebehind="BusinessDevelopmentManagerDisplay.ascx.cs" Inherits="Epsom.Web.WebUserControls.BusinessDevelopmentManagerDisplay" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="AddressDisplay.ascx" TagName="AddressDisplay" TagPrefix="Epsom" %>

<table>

  <tr style="vertical-align: top">
    <td style="width:15em">
      <asp:image id="imagePhoto" valign="top" imageurl="~/_gfx/arrow.gif" runat="server" />
    </td>
    <td>
      <table class="displaytable" style="width: 20em">
        <tr>
          <td colspan="2"><b><asp:Label id="lblName" runat="server" /></b></td>
	      </tr>
        <tr>
		      <td colspan="2"><asp:Label id="lblEmail" runat="server" /></td>
	      </tr>
	      <tr>
		      <td class="prompt">Phone:</td>
		      <td><asp:Label id="lblPhone" runat="server" /></td>
	      </tr>
	      <tr>
		      <td class="prompt">Mobile:</td>
		      <td><asp:Label id="lblMobile" runat="server" /></td>
	      </tr>
	      <tr>
		      <td>&nbsp;</td>
	      </tr>

      </table>
    </td>
  </tr>

</table>
<p class="forminfo">
  <asp:Label id="lblNotes" runat="server" />
</p> 
