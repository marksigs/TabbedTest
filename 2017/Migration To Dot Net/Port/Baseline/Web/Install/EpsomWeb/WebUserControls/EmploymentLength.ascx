<%@ Control Language="c#" AutoEventWireup="false" Codebehind="EmploymentLength.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentLength" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>

<table class="formtable">
	<tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Job title</td>
    <td>
      <asp:textbox id="txtJobTitle" runat="server" />
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtJobTitle" text="***" 
        errormessage="Job Title is required"
        id="Requiredfieldvalidator2" enableclientscript="false" />
    </td>
    </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Employment status</td>
    <td>
      <asp:dropdownlist id="cmbEmploymentStatus" runat="server" autopostback="True"/>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Employment Status is required" controltovalidate="cmbEmploymentStatus" id="Requiredselecteditemvalidator2"/> 
    </td>
  </tr>
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Date Employment Started</td>
    <td>
      <asp:textbox id="txtDateEmploymentStarted" runat="server" columns="10" cssclass="text" />
      <validators:datevalidator id="DateStartedValidator" runat="server" controltovalidate="txtDateEmploymentStarted" text="***" errormessage="Invalid date Employment Started"
        enableclientscript="false" />dd/mm/yyyy
      </td>
      <td>You must enter 1 year's worth of employment dates.  
    </td>
  </tr>
  <asp:panel runat="server" id="pnlDateEmploymentFinished">
  <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> to
    </td>
    <td>
      <asp:textbox id="txtDateEmploymentFinished" runat="server" columns="10" cssclass="text" />
      <validators:datevalidator id="DateFinishedValidator" runat="server" controltovalidate="txtDateEmploymentFinished" text="***" errormessage="Invalid date Employment Finished"
        enableclientscript="false" />
    </td>
  </tr>
  </asp:panel>

<asp:panel runat="server" id ="pnlContractLength">
  <tr>
    <td class=prompt><span class="mandatoryfieldasterisk">* </span>Length of contract</td>
    <td colspan="2">
      <asp:DropDownList id=cmbLengthOfContractYears Runat="server" Height="20">
               <asp:ListItem Runat="server" Value="0">0</asp:ListItem>
               <asp:ListItem Runat="server" Value="1">1</asp:ListItem>
               <asp:ListItem Runat="server" Value="2">2</asp:ListItem>
               <asp:ListItem Runat="server" Value="3">3</asp:ListItem>
               <asp:ListItem Runat="server" Value="4">4</asp:ListItem>
               <asp:ListItem Runat="server" Value="5">5</asp:ListItem>
               <asp:ListItem Runat="server" Value="6">6</asp:ListItem>
               <asp:ListItem Runat="server" Value="7">7</asp:ListItem>
               <asp:ListItem Runat="server" Value="8">8</asp:ListItem>
               <asp:ListItem Runat="server" Value="9" >9</asp:ListItem>
               <asp:ListItem Runat="server" Value="10">10</asp:ListItem>
               <asp:ListItem Runat="server" Value="11">11</asp:ListItem>
               <asp:ListItem Runat="server" Value="12">12</asp:ListItem>
               <asp:ListItem Runat="server" Value="13">13</asp:ListItem>
               <asp:ListItem Runat="server" Value="14">14</asp:ListItem>
               <asp:ListItem Runat="server" Value="15">15</asp:ListItem>
               <asp:ListItem Runat="server" Value="16">16</asp:ListItem>
               <asp:ListItem Runat="server" Value="17">17</asp:ListItem>
               <asp:ListItem Runat="server" Value="18">18</asp:ListItem>
               <asp:ListItem Runat="server" Value="19">19</asp:ListItem>
               <asp:ListItem Runat="server" Value="20">20</asp:ListItem>
            </asp:DropDownList>
            years
            <asp:DropDownList id=cmbLengthOfContractMonths Runat="server" Height="12">
               <asp:ListItem Runat="server" Value="0">0</asp:ListItem>
               <asp:ListItem Runat="server" Value="1">1</asp:ListItem>
               <asp:ListItem Runat="server" Value="2">2</asp:ListItem>
               <asp:ListItem Runat="server" Value="3">3</asp:ListItem>
               <asp:ListItem Runat="server" Value="4">4</asp:ListItem>
               <asp:ListItem Runat="server" Value="5">5</asp:ListItem>
               <asp:ListItem Runat="server" Value="6">6</asp:ListItem>
               <asp:ListItem Runat="server" Value="7">7</asp:ListItem>
               <asp:ListItem Runat="server" Value="8">8</asp:ListItem>
               <asp:ListItem Runat="server" Value="9">9</asp:ListItem>
               <asp:ListItem Runat="server" Value="10">10</asp:ListItem>
               <asp:ListItem Runat="server" Value="11">11</asp:ListItem>
            </asp:DropDownList>
            months
    </td>
  </tr>
    <tr>
    <td class="prompt"><span class="mandatoryfieldasterisk">*</span> Renewable?
          <td>
            <epsom:nullableyesno  id="nynRenewable" runat="server" autopostback="true" required="true" errormessage="Specify if your contract is renewable" enableclientscript=false />
            
          </td>
        </tr> 
    </td>
  </tr>
</asp:panel> 

  </table>