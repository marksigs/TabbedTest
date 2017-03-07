<%@ Control Language="c#" AutoEventWireup="false" Codebehind="AddressDisplay.ascx.cs" Inherits="Epsom.Web.WebUserControls.AddressDisplay" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<table class="displaytable">
 	<tr>
		<td class="prompt">Flat name or number</td>
		<td><asp:label id="lblAddressFlatNumber" runat="server" /></td>
    <asp:panel runat="server" id="pnlBritishForcesPostOffice">
  		<td class="prompt">BFPO</td>
	  	<td><asp:label id="lblBritishForcesPostOffice" runat="server" /></td>
    </asp:Panel>
	</tr>
	<tr>
		<td class="prompt">Building name</td>
		<td><asp:label id="lblAddressBuildingName" runat="server" /></td>
    <asp:panel runat="server" id="pnlOccupancy">
  		<td class="prompt">Occupancy</td>
	   	<td><asp:label id="lblNatureOfOccupancy" runat="server" /></td>
    </asp:Panel>
	</tr>
	<tr>
		<td class="prompt">Building number</td>
		<td><asp:label id="lblAddressBuildingNumber" runat="server" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
	</tr>
  <tr>
    <td class="prompt">Street</td>
    <td><asp:label id="lblAddressStreet" runat="server" /></td>
    <asp:panel runat="server" id="pnlResidentFrom">
      <td class="prompt">Resident from</td>
      <td>
        <asp:Label id="lblResidentFromMonthAndYear" runat="server" />
      </td>
    </asp:Panel> 
  </tr>
  <tr>
    <td class="prompt">District</td>
    <td><asp:label id="lblAddressDistrict" runat="server" /></td>
    <asp:panel runat="server" id="pnlResidentTo">
      <td class="prompt">Resident to</td>
      <td>
        <asp:label id="lblResidentToMonthAndYear" runat="server" />
      </td>
    </asp:Panel>
  </tr>
  <tr>
    <td class="prompt">Town</td>
    <td><asp:label id="lblTown" runat="server" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="prompt">County</td>
    <td><asp:label id="lblCounty" runat="server" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="prompt">Post code</td>
    <td><asp:label id="lblPostcode" runat="server" /></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <asp:panel runat="server" id="pnlCountry">
    <tr>
      <td class="prompt">Country</td>
      <td><asp:label id="lblCountry" runat="server" /></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </asp:Panel>
</table> 
