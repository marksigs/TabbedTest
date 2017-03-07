<%@ Control Language="c#" AutoEventWireup="false" Codebehind="ProductSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.ProductSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:repeater id="rptComponents" runat="server">

  <itemtemplate>
    <fieldset>
  
      <legend><asp:Label ID="lblComponentTitle" Runat="server"></asp:Label></legend>

      <table class="displaytable">
        <tr>
          <td class="prompt">Primary purpose of loan</td>
          <td><asp:Label ID="lblPurposeOfLoan" Runat="server" /></td>
          <td class="prompt">Component Amount</td>
          <td><asp:Label ID="lblComponentAmount" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Secondary purpose of loan</td>
          <td><asp:Label ID="lblSubPurposeOfLoan" Runat="server" /></td>
          <td class="prompt">Term</td>
          <td><asp:Label ID="lblTermOfMortgage" Runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt">Repayment Type</td>
          <td><asp:Label ID="lblRepaymentType" Runat="server"></asp:Label></td>
        </tr>
        <tr>
          <td class="prompt">Repayment Vehicle</td>
          <td><asp:label id="lblRepaymentVehicle" runat="server" /></td>
          <td class="prompt">Monthly cost of repayment vehicle</td>
          <td><asp:label id="lblRepaymentVehicleCost" runat="server" /></td>
        </tr>
      </table>
      <hr />


      <table class="displaytable" id="tblButton">
        <tr>
          <td><asp:Label ID="lblDescription" Runat="server" /></td>
          <td style="text-align:right">
            <asp:Button ID="cmdRemoveComponent" commandName="cmdRemoveComponent" runat="server" text="Remove" CssClass="button" />
            &nbsp; 
            <asp:Button ID="cmdEditComponent" commandName="cmdEditComponent" runat="server" text="Edit" CssClass="button" />
          </td>
        </tr>
      </table>

    </fieldset>

  </itemtemplate>

</asp:repeater>

