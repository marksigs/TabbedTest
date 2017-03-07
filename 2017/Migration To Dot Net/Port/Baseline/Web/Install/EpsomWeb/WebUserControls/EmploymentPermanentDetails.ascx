<%@ Control Language="c#" AutoEventWireup="false" Codebehind="EmploymentPermanentDetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentPermanentDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/EmploymentLength.ascx" TagName="EmploymentLength" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
  <H2 class="form_heading">Current employment details</H2>
  <TABLE class="formtable">
    <TBODY>
      <TR>
        <TD class="prompt"><asp:label id="lblEmployerName" runat="server" text="Name of employer"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD><asp:textbox columns="10" cssclass="text" id="txtEmployerName" runat="server" text="" tooltip="enter the name of your employer"></asp:textbox>   <asp:requiredfieldvalidator id="txtEmployerName_validator" runat="server" text="***" EnableClientScript="false" errormessage="Employer's name cannot be blank" controltovalidate="txtEmployerName"></asp:requiredfieldvalidator><TD>
     
      </TR>   
      <tr>
     </TABLE>
     <epsom:address id="ctlEmployerAddress" runat="server" isresidential="False" unitedkingdomonly="True" shownatureofoccupancy="False" showresidencydates="False" ></epsom:address></tr>
     <epsom:telephonenumbers id="ctlEmployerTelephoneNumbers" runat="server" />

   <TABLE class="formtable">
      <TR>
        <TD class="prompt"><asp:label id="lblEmployerEmail" runat="server" text="Employer's email address"></asp:label></TD>
        <TD><asp:textbox columns="40" cssclass="text" id="txtEmployerEmail" runat="server" text="" tooltip="enter the email address of your employer"></asp:textbox></TD>
      </TR>
      
      <TR>
        <TD class="prompt"><asp:label id="lblNatureOfBusiness" runat="server" text="Nature of business "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD><asp:textbox columns="40" cssclass="text" id="txtNatureOfBusiness" runat="server" text="" tooltip="enter the nature of your employer's business"></asp:textbox></TD>
 
      </TR>
     
     <TR>
        <TD class="prompt"><asp:label id="Label1" runat="server" text="Job Title"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <td><asp:textbox columns="40" cssclass="text" id="txtJobTitle" runat="server" text="" tooltip="enter your job title"></asp:textbox></td>
     </TR>
      
      <TR>
        <TD class="prompt"><asp:label id="lblEmploymentStarted" runat="server" text="Date employment started"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
    <td><asp:textbox columns="40" cssclass="text"  id="txtDateJobStarted" runat="server" Enabled="True" />&nbsp;dd/mm/yyyy&nbsp;&nbsp;&nbsp;You must enter 1 years worth of employment dates</td>
   
     </TR>
     
     <TR>
        <TD class="prompt"><asp:label id="lblEmploymentStatus" runat="server" text="Employment Type"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
    <td>
      <asp:dropdownlist id="ddlEmploymentStatus" runat="server" Enabled="True" />
</td>
     </TR>
      <TR>
        <TD class="prompt">
          <asp:label id="lblProbation" runat="server" text="Are you currently on a probationary period? "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:RadioButtonList id="rblProbationPeriod" runat="Server" tooltip="state whether you are on probationary period or not"
            RepeatLayout="flow" RepeatDirection="horizontal"  AutoPostBack="True" >
            <asp:ListItem value="1" runat="server">Yes</asp:ListItem>
            <asp:ListItem value="2" runat="server">No</asp:ListItem>
          </asp:RadioButtonList></TD>
            
      </TR>
        
      <TR>
        <TD class="prompt"><asp:label id="Label2" runat="server" text="Please give details"></asp:label></TD>
        <TD> <asp:TextBox columns="40" id="txtProbationaryPeriod" runat="server" Height="100px" Width="400px" AutoPostBack="True"></asp:textbox></TD>
      </TR>
      
      <TR>
        <TD class="prompt">
          <asp:label id="lblNotice" runat="server" text="Are you under notice of termination or redundancy? "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:RadioButtonList id="rblNoticePeriod" runat="Server" tooltip="state whether you are working under notice of termination or redundancy"
            RepeatLayout="flow" RepeatDirection="horizontal">
            <asp:ListItem value="1" runat="server">Yes</asp:ListItem>
            <asp:ListItem value="2" runat="server">No</asp:ListItem>
          </asp:RadioButtonList>
        </TD>
      </TR>
      <TR>
        <TD class="prompt">
          <asp:label id="lblEmployeeNumber" runat="server" text="Employee/Personnel number "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:textbox columns="40" cssclass="text" id="txtEmployeeNumber" runat="server" tooltip="enter your employee/personnel number"></asp:textbox></TD>
      </TR>
      <TR>
        <TD class="prompt">
          <asp:label id="lblTaxRef" runat="server" text="Tax district and Reference number "></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:textbox columns="40" cssclass="text" id="txtTaxRef" runat="server" tooltip="enter your tax reference number and tax district"></asp:textbox></TD>
        
      </TR>
      <TR>
        <TD class="prompt">
          <asp:label id="lblNINumber" runat="server" text="National Insurance number"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:textbox columns="40" cssclass="text" id="txtNINumber" runat="server" tooltip="enter your National Insurance number"></asp:textbox></TD>
      </TR>
      <TR>
        <TD class="prompt">
          <asp:label id="lblFamilyOwnedCompany" runat="server" text="Is the company owned by a family member?"></asp:label><span class="mandatoryfieldasterisk">*</span></TD>
        <TD>
          <asp:radiobuttonlist id="rblFamilyOwnedCompany" runat="server" tooltip="state whether the company you work for is owned by a family member or not"
            RepeatLayout="flow" RepeatDirection="horizontal" AutoPostBack="True">
            <asp:ListItem value="1" runat="server">Yes</asp:ListItem>
            <asp:ListItem value="2" runat="server">No</asp:ListItem>
          </asp:radiobuttonlist></TD>
      </TR>
     
      <TR>
        <TD class="prompt"><asp:label id="Label97" runat="server" text="Please give details"></asp:label></TD>
        <TD> <asp:TextBox columns="40" id="txtCompanyOwnership" runat="server" Height="100px" Width="400px" AutoPostBack="True"></asp:textbox></TD>
      </TR>
  </TABLE>
  <!--<tr><epsom:employmentlength id="ctlEmploymentLength" runat="server"></epsom:employmentlength>
</tr>-->
</TD></TR></TBODY></TABLE> 
</asp:label></TD></TR></TBODY></TABLE></asp:panel></BODY></HTML>

