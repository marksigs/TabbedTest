<%@ Control Language="c#" AutoEventWireup="false" Codebehind="AdditionalInfoSummary.ascx.cs" Inherits="Epsom.Web.WebUserControls.AdditionalInfoSummary" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<itemtemplate>

  <fieldset>

   <legend>Additional information</legend>

   <table class="displaytable_small">
     <tr>
       <td class="prompt">Total number of dependants</td>
       <td><asp:label id="lblDependants" runat="server" /></td>
     </tr>
     <tr>
       <td class="prompt">Additional information</td>
       <td><asp:label id="lblAdditionalInfo" runat="server" /></td>
     </tr>
   </table>

   <table class="displaytable">
     <tr>
       <td style="text-align:right"><asp:button id="cmdEditAdditionalInfo" runat="server" text="Edit"  cssclass="button"/></td>
     </tr>
   </table>

  </fieldset> 

</itemtemplate>
    
    