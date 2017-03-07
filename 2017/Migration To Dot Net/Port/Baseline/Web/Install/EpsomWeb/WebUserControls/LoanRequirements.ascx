<%@ Control Language="c#" AutoEventWireup="false" Codebehind="LoanRequirements.ascx.cs" Inherits="Epsom.Web.WebUserControls.LoanRequirements" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<%@ Register Src="~/WebUserControls/PropertyDetails.ascx" TagName="PropertyDetails" TagPrefix="Epsom" %>
 
 <cms:cmsBoilerplate cmsref="30" runat="server" />
    <h2 class=form_heading>Loan details</h2>

    <fieldset>
      <legend>Application Details</legend>
      <table class=formtable>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Category</td>
          <td>
            <asp:dropdownlist id=cmbCategory runat="server" CausesValidation="no" autopostback="true"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id=RequiredSelectedItemValidator2 Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Category is required" ControlToValidate="cmbCategory"></validators:RequiredSelectedItemValidator>
          </td>
          <td rowspan="3">
            <cms:cmsBoilerplate cmsref="31" id="BTLText"  runat="server" />
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Type of application</td>
          <td>
            <asp:dropdownlist id=cmbApplicationType runat="server" CausesValidation="no" autopostback="true"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id=Requiredselecteditemvalidator3 Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Type of application is required" ControlToValidate="cmbApplicationType"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Type of buyer</td>
          <td>
            <asp:dropdownlist id=cmbBuyerType runat="server"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id=RequiredSelectedItemValidator1 Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Type of Buyer is required" ControlToValidate="cmbBuyerType"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
      </table>
    </fieldset> 

    <asp:Panel id="pnlProperyDetails" runat="server" visible="True" Enabled="True">
    
      <fieldset>

        <legend>Property Details</legend>

        <table class="formtable">

          <tr>
            <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Property Location</td>
            <td>
              <asp:dropdownlist id="ddlCategory" runat="server" CausesValidation="no" autopostback="false"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator7" Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Property Location must be selected" ControlToValidate="ddlCategory"></validators:RequiredSelectedItemValidator>
            </td>
          </tr>

          <tr>
            <td class="prompt"><span class=mandatoryfieldasterisk>*</span>What Valuation type do you require?</td>
            <td>
              <asp:dropdownlist id="ddlValuationType" runat="server" CausesValidation="no" autopostback="false"></asp:dropdownlist>
              <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator8" Runat="server" EnableClientScript="false"
               Text="***" ErrorMessage="Valuation type must be selected" ControlToValidate="ddlValuationType">
               </validators:RequiredSelectedItemValidator></td>
          </tr>
        
          <asp:Repeater id="rptDepositSource" runat="server">
            <ItemTemplate>
              <tr>
                <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Deposit source type</td>
                <td><asp:dropdownlist  id="ddlDepositSourceType"  runat="server" CausesValidation="no" AutoPostBack="false"  DataMember="HttpSession.CurrentApplication.DepositSources.DepositSourceType"></asp:dropdownlist>
                <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator6" Runat="server" EnableClientScript="false" Text="***" ErrorMessage="A Valid Deposit Source is required" ControlToValidate="ddlDepositSourceType"></validators:RequiredSelectedItemValidator></td>
                <td><span class=mandatoryfieldasterisk>*</span>Amount</td>    
                <td><asp:textbox id="txtAmount"  runat="server" AutoPostBack="false" DataMember="HttpSession.CurrentApplication.DepositSources.Amount"></asp:textbox><asp:rangevalidator id="Rangevalidator111" runat="server" enableclientscript="false" errormessage="A Valid Deposit Amount is required" controltovalidate="txtAmount" MinimumValue="0" MaximumValue="999999999999" Type="Currency" text="***"></asp:rangevalidator></td>
              </tr>
            </ItemTemplate>
          </asp:Repeater>

        </table>
      
        <p class="button_orphan">
          <asp:button id="btnAddDepositSource" runat="server" enabled="True" text="Add another deposit source" CausesValidation=False />
          <asp:button id="btnRemoveDepositSource" runat="server" enabled="True" text="Remove" CausesValidation=False/>
        </p>
      
        <asp:Panel id="pnlRentalIncome" runat="server" Visible="False">
          <table class="formtable">
            <tr>
              <td class="prompt">Anticipated monthly rental income </td>
              <td><asp:textbox id="txtRentalIncome" runat="server"></asp:textbox></td>
            </tr>   
          </table>
        </asp:Panel>

      </fieldset>
    
    </asp:Panel>

    <fieldset>
      <legend>Loan type</legend>
    
      <table class=formtable>
        <tr>
          <td class="prompt"><asp:Label ID="lblSchemeRequiredIndicator"  Runat="server" CssClass="mandatoryfieldasterisk">*</asp:label>Scheme</td>
          <td>
            <asp:dropdownlist id=cmbScheme runat="server" AutoPostBack="True"></asp:dropdownlist><asp:label ID="lblSchemeRequiredText" Runat="server">&nbsp;(leave blank to find the scheme)</asp:label> 
            <validators:RequiredSelectedItemValidator id="RequiredselecteditemvalidatorScheme" Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Scheme is required" ControlToValidate="cmbScheme"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>* </span>Application income status</td>
          <td>
            <asp:dropdownlist id=cmbStatus runat="server" CausesValidation="no" autopostback="true"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator4" Runat="server" EnableClientScript="false" Text="***" ErrorMessage="Status is required" ControlToValidate="cmbStatus"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
        <asp:Panel ID="pnlSelfCertification"  Runat="server" Visible="False">
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>* </span>Self certification reason</td>
          <td>
            <asp:dropdownlist id="cmbSelfCertificationReason" runat="server"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator9" Runat="server" EnableClientScript="false" Text="***"
             ErrorMessage="Self certification reason is required" ControlToValidate="cmbSelfCertificationReason"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
        </asp:Panel>
        <asp:Panel id="pnlFlexiSurpress" runat="server" Visible="True">
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Flexi features</td>
          <td>
            <asp:RadioButtonList id=rblFlexiFeatures runat="Server" repeatLayout="Flow" repeatDirection="Horizontal">
              <asp:ListItem Value="0">Include Non-Flexi </asp:ListItem>
              <asp:ListItem Value="1">Only with Flexi Features</asp:ListItem>
            </asp:RadioButtonList>
            <asp:customvalidator id="Customvalidator1" runat="server" 
              errormessage="Flexi Product Error"
              enableclientscript="false" onservervalidate="FlexiProductCheck" text="***" >
            </asp:customvalidator> 
          </td> 
        </tr>
      
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Product features</td>
          <td>
            <asp:RadioButtonList id=rblProductClassFeatures runat="Server" repeatLayout="Flow" repeatDirection="Horizontal">
               <asp:ListItem Value="0">Standard only</asp:ListItem>
               <asp:ListItem Value="1">Guarantor only</asp:ListItem>
            </asp:RadioButtonList>
          </td>
    
        </tr>
          </asp:Panel>
      </table>
    </fieldset> 

    <fieldset>
      <legend>Loan amount</legend>
      <table class=formtable>
		<!-- ik_EP2_2367 field hidden, validators removed -->
        <tr style="display:none">
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Credit limit</td>
          <td>
            <asp:textbox id="txtCreditLimit" runat="server" CssClass="text"></asp:textbox>
          </td>
          <td rowspan="2"><cms:cmsBoilerplate cmsref="34" runat="server" /></td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Amount requested</td>
          <td>
            <asp:textbox id="txtAmountRequested" runat="server" CssClass="text" ></asp:textbox>
            <asp:RequiredFieldValidator id=RequiredFieldValidator2 Runat="server" 
              EnableClientScript="false" ErrorMessage="Amount requested is required" ControlToValidate="txtAmountRequested" text="***">
            </asp:RequiredFieldValidator>
             <asp:rangevalidator id="Rangevalidator1" runat="server" enableclientscript="false" errormessage="Amount Requested must be at over £25000" controltovalidate="txtAmountRequested" MinimumValue="25001" MaximumValue="9999999999" Type="Currency" text="***"></asp:rangevalidator>
      
            <asp:CompareValidator id=CompareValidator4 Runat="server" EnableClientScript="false" Text="***" 
              ErrorMessage="Amount requested is Valid" ControlToValidate="txtAmountRequested" Operator="DataTypeCheck" Type="Integer">
            </asp:CompareValidator>
            <asp:CustomValidator id="AmountRequestedValidator" runat="server" 
              errormessage="Amount requested error"
              enableclientscript="false" onservervalidate="AmountRequestedCheck" text="***" >
           </asp:customvalidator> 
              </td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Purchase price / estimated value</td>
          <td>
            <asp:textbox id="txtPurchasePrice" runat="server" CssClass="text"></asp:textbox>
            <asp:RequiredFieldValidator id=RequiredFieldValidator3 Runat="server" EnableClientScript="false" ErrorMessage="Purchase price / estimated value amount is required" ControlToValidate="txtPurchasePrice" text="***"></asp:RequiredFieldValidator>
            <asp:CompareValidator id=CompareValidator5 Runat="server" EnableClientScript="false" Text="***" 
              ErrorMessage="Purchase price is not numeric" ControlToValidate="txtPurchasePrice" Operator="DataTypeCheck" Type="Integer">
            </asp:CompareValidator>
           <asp:rangevalidator id="Rangevalidator2" runat="server" enableclientscript="false" errormessage="Amount in Purchase Price too low or too high" controltovalidate="txtPurchasePrice" MinimumValue="25001" MaximumValue="999999999" Type="Currency" text="***"></asp:rangevalidator>
          </td>
        </tr>

        <asp:panel id=pnlRightToBuyDetails runat="server">
          <tr>
            <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Discount price</td>
            <td>
              <asp:textbox id=txtDiscountPrice runat="server" CssClass="text" CausesValidation="no" AutoPostBack="True"></asp:textbox>
              <asp:CompareValidator id=CompareValidator6 Runat="server" EnableClientScript="false" Text="***" 
                ErrorMessage="Discount price is not numeric" ControlToValidate="txtDiscountPrice" Operator="DataTypeCheck" Type="Integer">
              </asp:CompareValidator>
              <asp:CompareValidator id=CompareValidator7 Runat="server" EnableClientScript="false" Text="***" 
                ErrorMessage="Discount price must be less than purchase price" 
                ControlToValidate="txtDiscountPrice" Operator="LessThanEqual" Type="Currency" ControlToCompare="txtPurchasePrice">
              </asp:CompareValidator>
              <asp:RequiredFieldValidator id="RequiredfieldDiscountPrice" Runat="server" EnableClientScript="false" ErrorMessage="Discount Price is required" ControlToValidate="txtDiscountPrice" text="***"></asp:RequiredFieldValidator>
            </td>
          </tr>
          <tr>
            <td class="prompt">Calculated discount</td>
            <td>
              <asp:textbox id="txtCalculatedDiscount" runat="server" enabled="false" CssClass="text" readonly="true"></asp:textbox>
              <asp:button id="CmdCalculateDiscount" runat="server"  text="Calculate" causesvalidation="false" cssclass="button"></asp:button>
            </td>
          </tr>
        </asp:panel>
      </table>
    </fieldset> 

    <h2 class=form_heading>Submission route</h2>
    
    <table class=formtable>
      <tr>
        <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Submission method</td>
        <td>
          <asp:RadioButtonList id=rblPackaged runat="Server" repeatLayout="Flow" repeatDirection="Horizontal" AutoPostBack="True">
            <asp:ListItem Value="0">Packaged</asp:ListItem>
            <asp:ListItem Value="1">Unpackaged</asp:ListItem>
          </asp:RadioButtonList>
          <asp:requiredfieldvalidator runat="server" id="rblPackagedRequired" controltovalidate="rblPackaged" text="***"
            errormessage="Please specify the submission method" enableclientscript=false ></asp:requiredfieldvalidator> 
        </td>
      </tr>
      
      <asp:panel id=pnlPackagerPicker runat="server">
        <tr>
          <td class="prompt">Packager association</td>
          <td>
            <asp:dropdownlist id=cmbPackagerAssociation runat="server" autoPostback="true" CssClass="span2"></asp:dropdownlist>
          </td>
        </tr>
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Packager</td>
          <td>
            <asp:dropdownlist id=cmbPackager runat="server" CssClass="span2"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id="RequiredSelectedItemValidator5" Runat="server" 
              EnableClientScript="false" Text="***" ErrorMessage="Packager is required"
              ControlToValidate="cmbPackager"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
      </asp:panel>
      <asp:panel id="pnlUnpackagedDetails" runat="server" >
        <tr>
          <td class="prompt"><span class=mandatoryfieldasterisk>*</span>Route</td>
          <td>
            <asp:RadioButtonList id=rblUnpackagedRoute runat="Server" repeatLayout="Flow" repeatDirection="Horizontal" AutoPostBack="True">
              <asp:ListItem Value="0">Direct to DB</asp:ListItem>
              <asp:ListItem Value="1">Other/my network</asp:ListItem>
            </asp:RadioButtonList>
          </td>
        </tr>
      </asp:panel>
    </table>
    
    <asp:panel id=pnlMortgageClubPicker runat="server">
      <table class=formtable>
        <tr>
          <td></td>
          <td><asp:label ID="lblSelectMortgageClubMessage" Runat="server" /></td>
        </tr>
      </table>
      
      <asp:panel id="pnlMortgageClubCombo" runat="server"> 
      <table class=formtable>
        <tr>
          <td class="prompt">Choose Mortgage Club</td>
          <td>
            <asp:dropdownlist id=cmbMortgageClub runat="server" CssClass="span2"></asp:dropdownlist>
            <validators:RequiredSelectedItemValidator id="ValidatorMortgageClub" Runat="server" 
              EnableClientScript="false" Text="***" ErrorMessage="Mortgage Club is required"
              ControlToValidate="cmbMortgageClub"></validators:RequiredSelectedItemValidator>
          </td>
        </tr>
      </table>
      </asp:panel>
    </asp:panel>
    
    <asp:panel id=pnlWarning1 runat="server">
     <p><cms:cmsBoilerplate cmsref="32" runat="server" /></p>
    </asp:panel>
    