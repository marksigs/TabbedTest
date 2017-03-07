<%@ Control Language="c#" AutoEventWireup="false" Codebehind="Employmentdetails.ascx.cs" Inherits="Epsom.Web.WebUserControls.EmploymentDetails" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register Src="Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>


<fieldset>
  <legend>Current employment type</legend>
  <table  class="formtable">
    <tr>
      <td class="prompt">Employment status</td>
      <td><asp:DropDownList id=cmbEmploymentStatus runat="server" AutoPostBack="True"></asp:DropDownList></td>
    </tr>
  </table>
</fieldset>

<fieldset>
  <legend>Current employment details</legend>
  <table  class="formtable">
    <tr>
      <td  class="prompt"><span class="mandatoryfieldasterisk">* </span>Job title/trading name</td>
      <td>
        <asp:TextBox id=txtJobTitle runat="server"></asp:TextBox>
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtJobTitle" text="***" errormessage="Please enter a Job title/trading name" id="RequiredfieldJobTitle" enableclientscript="false" />
      </td>
    </tr>
    <tr>
      <td  class="prompt"><span class="mandatoryfieldasterisk">* </span>Name of employer</td>
      <td>
        <asp:TextBox id=txtNameOfEmployer runat="server"></asp:TextBox>
        <asp:requiredfieldvalidator runat="server" controltovalidate="txtNameOfEmployer" text="***" errormessage="Please enter name of Employer" id="RequiredfieldNameOfEmployer" enableclientscript="false" />
      </td>
    </tr>
  </table>

  <epsom:address id="ctlEmployerAddress" runat="server"  isresidential="false" />

  <epsom:TelephoneNumbers id="ctlEmployerTelephoneNumbers" FirstItemMandatory="true" usageComboGroup="ContactTelephoneUsage" runat="server" />

  <table  class="formtable">
    <tr>
      <td  class="prompt">Employer's email address</td>
      <td><asp:TextBox id=txtEmployerEmail runat="server"></asp:TextBox></td>
    </tr>
    <tr>
      <td  class="prompt"><span class="mandatoryfieldasterisk">* </span>Nature of business</td>
      <td><asp:DropDownList id="cmbNatureOfBusiness" runat="server"></asp:DropDownList>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Nature of business is required" controltovalidate="cmbNatureOfBusiness" id="Requiredselecteditemvalidator2"/> 
      </td>
    </tr>
  </table>

  <asp:Panel id="pnlDateEmploymentStarted" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class="prompt"><span class="mandatoryfieldasterisk">* </span><asp:Label id="lblDateEmploymentStarted" runat="server"></asp:Label></TD>
    <TD>
      <asp:TextBox id="txtDateEmploymentStarted" runat="server"></asp:TextBox>
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtDateEmploymentStarted" text="***" errormessage="Please enter the date employment started" id="RequiredfieldDateEmploymentStarted" enableclientscript="false" />
      <validators:datevalidator id="DateValidator1" runat="server" controltovalidate="txtDateEmploymentStarted" 
            text="***" errormessage="Invalid date employment started"
          enableclientscript="false" /> dd/mm/yyyy
    </TD>
  </TR>
</TABLE>
  </asp:Panel>

  <asp:Panel id="pnlDateEmploymentFinished" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>Date employment finished</TD>
    <TD><asp:TextBox id=txtDateEmploymentFinished runat="server"></asp:TextBox>
    <validators:datevalidator id="Datevalidator2" runat="server" controltovalidate="txtDateEmploymentFinished" 
            text="***" errormessage="Invalid date employment finished"
          enableclientscript="false" /> dd/mm/yyyy

    </TD>
  </TR>
</TABLE>
  </asp:Panel>

  <asp:Panel id="pnlDateBusinessAcquired" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>What date did you acquire an interest in the business</TD>
    <TD><asp:TextBox id=txtDateBusinessAcquired runat="server"></asp:TextBox>
    <validators:datevalidator id="Datevalidator3" runat="server" controltovalidate="txtDateBusinessAcquired" 
            text="***" errormessage="Invalid date business acquired"
          enableclientscript="false" /> dd/mm/yyyy

    </TD>
  </TR>
</TABLE>
  </asp:Panel>

  <asp:Panel id="pnlEmploymentType" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt><span class="mandatoryfieldasterisk">* </span>Employment type</TD>
    <TD><asp:DropDownList id=cmbEmploymentType runat="server"></asp:DropDownList>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Employment type is required" controltovalidate="cmbEmploymentType" id="Requiredselecteditemvalidator1"/> 
    </TD>
  </TR>
</TABLE>
  </asp:Panel>

  <asp:Panel id="pnlPobationaryPeriod" runat="server">
<TABLE class="formtable">
  <TR>
    <TD class="prompt"><span class="mandatoryfieldasterisk">* </span>Are you currently on a probationary period</TD>
    <TD>
      <epsom:nullableyesno  id="nynOnProbationaryPeriod" runat="server" autopostback="true" required="true" errormessage="Specify if you currently on a probationary period" enableclientscript=false />
    </TD>
  </TR>
</TABLE>
<asp:Panel id=pnlProbationaryPeriodDetails runat="server">
<TABLE class=formtable>
  <TR>
    <TD class="prompt"></TD>
    <TD><asp:TextBox id="txtOnProbationaryPeriodDetails" textmode="multiline" runat="server" ></asp:TextBox></TD></TR>
    </TABLE>
  </asp:Panel>
  </asp:Panel>

  <asp:Panel id="pnlUnderNotice" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt><span class="mandatoryfieldasterisk">* </span>Are you currently under notice of termination or 
      redundancy?</TD>
    <TD>
      <epsom:nullableyesno  id="nynUnderNotice" runat="server" autopostback="true" required="true" errormessage="Specify if you currently under notice" enableclientscript=false />
    </TD>
  </TR>
</TABLE>
  </asp:Panel>

  <asp:Panel id="pnlEmployeeNumber" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>Employee/Personnel number</TD>
    <TD>
      <asp:TextBox id=txtEmployeeNumber runat="server"></asp:TextBox>
    </TD>
  </TR>
</TABLE>
  </asp:Panel>
  
  <asp:Panel id="pnlCompanyRegistrationNumber" runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>Company registration number</TD>
    <TD><asp:TextBox id=txtCompanyRegistrationNumber runat="server"></asp:TextBox></TD></TR></TABLE></asp:Panel>
    
<asp:Panel id=pnlVatNumber runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>VAT number</TD>
    <TD><asp:TextBox id=txtVatNumber runat="server"></asp:TextBox></TD></TR></TABLE></asp:Panel>
<asp:Panel id=pnlPercentageShare runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>What is your % shareholding</TD>
    <TD><asp:TextBox id=txtPercentageShare runat="server"></asp:TextBox></TD></TR></TABLE>
</asp:Panel>
<asp:Panel id=pnlFamilyOwned runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt><span class="mandatoryfieldasterisk">* </span>Is the company owned by a family member?</TD>
    <TD>
      <epsom:nullableyesno  id="nynFamilyOwned" runat="server" autopostback="true" required="true" errormessage="Specify if the company is owned by a family member" enableclientscript=false />
    </TD>
  </TR>
</TABLE>
<asp:Panel id=pnlFamilyOwnedDetails runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt></TD>
    <TD><asp:TextBox id=txtFamilyOwnedDetails runat="server" TextMode="MultiLine"></asp:TextBox></TD>
    </TR>
</TABLE>
</asp:Panel>
</asp:Panel>
<asp:Panel id=pnlCompanyStatus runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt><span class="mandatoryfieldasterisk">* </span>Is the business?</TD>
    <TD>
      <asp:DropDownList id=cmbCompanyStatus runat="server"></asp:DropDownList>
      <validators:requiredselecteditemvalidator runat="server" enableclientscript="false"
          text="***" errormessage="Business Status is required" controltovalidate="cmbCompanyStatus" id="Requiredselecteditemvalidator3"/> 
    </TD>
  </TR></TABLE></asp:Panel></fieldset> 

<FIELDSET><LEGEND>Accountant Details</LEGEND>
<TABLE class=formtable>
  <TR>
    <TD class=prompt>Do you have an accountant for this employment?</TD>
    <TD>
      <asp:radiobutton id=rdoEmploymentAccountantYes runat="server" Text="Yes" GroupName="grpEmploymentAccountant" AutoPostBack="True"></asp:radiobutton>
      <asp:radiobutton id=rdoEmploymentAccountantNo runat="server" Text="No" GroupName="grpEmploymentAccountant" AutoPostBack="True"></asp:radiobutton>
    </TD>
  </TR>
</TABLE>

<asp:Panel id=pnlAccountantDetails runat="server">
<TABLE class=formtable>
  <TR>
    <TD class=prompt>How long has this accountant acted for you?</TD>
    <td>
      <asp:dropdownlist id="cmbAccountantYears" runat="server" height="20">
               <asp:listitem runat="server" value="0" id="Listitem1">0</asp:listitem>
               <asp:listitem runat="server" value="1" id="Listitem2">1</asp:listitem>
               <asp:listitem runat="server" value="2" id="Listitem3">2</asp:listitem>
               <asp:listitem runat="server" value="3" id="Listitem4">3</asp:listitem>
               <asp:listitem runat="server" value="4" id="Listitem5">4</asp:listitem>
               <asp:listitem runat="server" value="5" id="Listitem6">5</asp:listitem>
               <asp:listitem runat="server" value="6" id="Listitem7">6</asp:listitem>
               <asp:listitem runat="server" value="7" id="Listitem8">7</asp:listitem>
               <asp:listitem runat="server" value="8" id="Listitem9">8</asp:listitem>
               <asp:listitem runat="server" value="9" id="Listitem10">9</asp:listitem>
               <asp:listitem runat="server" value="10" id="Listitem11">10</asp:listitem>
               <asp:listitem runat="server" value="11" id="Listitem12">11</asp:listitem>
               <asp:listitem runat="server" value="12" id="Listitem13">12</asp:listitem>
               <asp:listitem runat="server" value="13" id="Listitem14">13</asp:listitem>
               <asp:listitem runat="server" value="14" id="Listitem15">14</asp:listitem>
               <asp:listitem runat="server" value="15" id="Listitem16">15</asp:listitem>
               <asp:listitem runat="server" value="16" id="Listitem17">16</asp:listitem>
               <asp:listitem runat="server" value="17" id="Listitem18">17</asp:listitem>
               <asp:listitem runat="server" value="18" id="Listitem19">18</asp:listitem>
               <asp:listitem runat="server" value="19" id="Listitem20">19</asp:listitem>
               <asp:listitem runat="server" value="20" id="Listitem21">20</asp:listitem>
            </asp:dropdownlist>
            years
            <asp:dropdownlist id="cmbAccountantMonths" runat="server" height="12">
               <asp:listitem runat="server" value="0" id="Listitem22">0</asp:listitem>
               <asp:listitem runat="server" value="1" id="Listitem23">1</asp:listitem>
               <asp:listitem runat="server" value="2" id="Listitem24">2</asp:listitem>
               <asp:listitem runat="server" value="3" id="Listitem25">3</asp:listitem>
               <asp:listitem runat="server" value="4" id="Listitem26">4</asp:listitem>
               <asp:listitem runat="server" value="5" id="Listitem27">5</asp:listitem>
               <asp:listitem runat="server" value="6" id="Listitem28">6</asp:listitem>
               <asp:listitem runat="server" value="7" id="Listitem29">7</asp:listitem>
               <asp:listitem runat="server" value="8" id="Listitem30">8</asp:listitem>
               <asp:listitem runat="server" value="9" id="Listitem31">9</asp:listitem>
               <asp:listitem runat="server" value="10" id="Listitem32">10</asp:listitem>
               <asp:listitem runat="server" value="11" id="Listitem33">11</asp:listitem>
            </asp:dropdownlist>
            months
    </td>

  </TR>
  <TR>
    <TD class=prompt>Name of accountant (first name/last name)</TD>
    <TD><asp:TextBox id=txtAccountantFirstName runat="server"></asp:TextBox><asp:TextBox id=txtAccountantLastName runat="server"></asp:TextBox></TD></TR>
  <TR>
    <TD class=prompt><span class="mandatoryfieldasterisk">* </span>Name of company/firm</TD>
    <TD>
      <asp:TextBox id=txtAccountantFirm runat="server"></asp:TextBox>
      <asp:requiredfieldvalidator runat="server" controltovalidate="txtAccountantFirm" text="***" errormessage="Please enter the accountant company/Firm" id="Requiredfieldvalidator2" enableclientscript="false" />
    </TD>
  </TR>
  </TABLE>
    <epsom:address id=ctlAccountantAddress runat="server" isresidential="false"></epsom:address>
    <epsom:TelephoneNumbers id="ctlAccountantTelephoneNumbers" usageComboGroup="ContactTelephoneUsage" runat="server"></epsom:TelephoneNumbers>
<TABLE class=formtable>
  <TR>
    <TD class=prompt>Accountant email address</TD>
    <TD><asp:TextBox id=txtAccountantEmail runat="server"></asp:TextBox></TD></TR>
  <TR>
    <TD class=prompt>Accountant's qualifications</TD>
    <td><asp:dropdownlist id="cmbAccountantQualifications" runat="server"></asp:dropdownlist></td>
  </TR>
</TABLE>
</asp:Panel>
</FIELDSET> 

<asp:Panel id=pnlContractDetails runat="server">
<FIELDSET><LEGEND>Contract details</LEGEND>
<TABLE class=formtable>
  <tr>
    <td class=prompt><span class="mandatoryfieldasterisk">* </span>Length of contract</td>
    <td >
      <asp:DropDownList id=cmbLengthOfContractYears Runat="server" autopostback="true" Height="20">
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
            <asp:DropDownList id=cmbLengthOfContractMonths Runat="server" autopostback="true" Height="12">
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
 
  <TR>
    <TD class=prompt>Date contract due to finish</TD>
    <TD>
    <asp:TextBox id=txtContractFinishDate runat="server" Enabled="False"></asp:TextBox>
    </TD>
  </TR>
  <TR>
    <TD class=prompt>Is the contract renewable</TD>
    <TD>
      <epsom:nullableyesno  id="nynContractRenewable" runat="server" autopostback="true" required="true" errormessage="Specify if the contract is renewable" enableclientscript=false />
    </TD>
  </TR>
  <TR>
    <TD class=prompt>How many times has the contract been renewed</TD>
    <TD>
      <asp:TextBox id="txtNumberTimesContractRenewed" runat="server"></asp:TextBox>
      <asp:rangevalidator id="Rangevalidator1" runat="server" enableclientscript="false" 
          errormessage="Number of times contract renewed must be numeric" controltovalidate="txtNumberTimesContractRenewed" 
          minimumvalue="0" maximumvalue="9999999999" type="Currency" text="***"></asp:rangevalidator>
    </TD>
  </TR>
</TABLE>
</FIELDSET> 
</asp:Panel>
