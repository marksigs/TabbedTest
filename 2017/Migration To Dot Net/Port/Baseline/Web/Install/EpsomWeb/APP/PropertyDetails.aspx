<%@ Page language="c#" Codebehind="PropertyDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.PropertyDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/TelephoneNumbers.ascx" TagName="TelephoneNumbers" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/Address.ascx" TagName="Address" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">
  <mp:content id="region1" runat="server">
    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>
  </mp:content>
  <mp:content id="region2" runat="server">
    <H1>Property details</H1>
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>
    <asp:validationsummary id="Validationsummary1" runat="server" displaymode="BulletList" headertext="Input errors:"
      CssClass="validationsummary" showsummary="true"></asp:validationsummary>
    <asp:panel id="pnlNewPurchaseInformation" runat="server">
      <H1 class="form_heading">New purchase information</H1>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt">Purchase Price
          </TD>
          <TD>
            <asp:textbox id="txtPurchasePrice" runat="server" enabled="false"></asp:textbox></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Is it a private 
            sale?</TD>
          <TD>
            <Epsom:NullableYesNo id="nynPrivateSale" runat="Server" required="true" errormessage="Specify if it is a private sale"></Epsom:NullableYesNo></TD>
        </TR>
       
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Do you have any 
            business connection with, or are you related to the vendor?</TD>
          <TD>
            <Epsom:NullableYesNo id="nynVendorRelated" runat="Server" required="true" errormessage="Specify business connection with vendor"></Epsom:NullableYesNo></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Is the property 
            being purchased at full market value?</TD>
          <TD>
            <Epsom:NullableYesNo id="nynFullMarketValue" runat="Server" required="true" errormessage="Specify if purchased at full market value"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlFullMarketValueMemo" runat="server">
          <TR>
            <TD class="prompt">&nbsp;</TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtFullMarketValueMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel>
        <asp:panel id="pnlEntryDate" runat="server">
          <TR>
            <TD class="prompt"><SPAN class="mandatoryfieldasterisk">*</SPAN>
              If the property is in Scotland, when is the date of entry?</TD>
            <TD colSpan="2">
              <asp:TextBox id="txtEntryDate" CssClass="text" Runat="server"></asp:TextBox>
              <validators:datevalidator id="Datevalidator1" runat="server" errormessage="Invalid date of entry" controltovalidate="txtEntryDate"
                text="***" enableclientscript="false"></validators:datevalidator>
              <validators:dateinfuturevalidator id="Datevalidator2" runat="server" errormessage="Date of entry must be in the future"
                controltovalidate="txtEntryDate" text="***" enableclientscript="false"></validators:dateinfuturevalidator>dd/mm/yyyy
            </TD>
          </TR>
        </asp:panel></TABLE>
    </asp:panel> <!-- EP2_1955 HMA Define Deposit details in repeater -->
    <asp:panel id="pnlDepositDetails" runat="server">
      <H2 class="form_heading">Deposit source and amount</H2>
      <TABLE class="formtable">
        <asp:repeater id="Repeater1" runat="server">
          <itemtemplate>
            <tr>
              <td class="prompt"><span class="mandatoryfieldasterisk">*</span>
                Deposit source type</td>
              <td>
                <asp:dropdownlist id="Dropdownlist1" runat="server" causesvalidation="no" autopostback="false" datamember="HttpSession.CurrentApplication.DepositSources.DepositSourceType"></asp:dropdownlist>
                <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator9" runat="server" enableclientscript="false" text="***"
                  errormessage="A Valid Deposit Source is required" controltovalidate="ddlDepositSourceType"></validators:requiredselecteditemvalidator></td>
              <td><span class="mandatoryfieldasterisk">*</span>
                Amount</td>
            </tr>
          </itemtemplate>
        </asp:repeater>
        <asp:Repeater id="rptDepositSource" runat="server">
          <ItemTemplate>
            <tr>
              <td class="prompt">Deposit Source</td>
              <td>
                <asp:dropdownlist id="ddlDepositSourceType" runat="server" causesvalidation="no" datamember="HttpSession.CurrentApplication.DepositSources.DepositSourceType"
                  enabled="false"></asp:dropdownlist>
              </td>
              <td>
                <asp:textbox id="txtAmount" runat="server" enabled="false" datamember="HttpSession.CurrentApplication.DepositSources.Amount"></asp:textbox>
              </td>
            </tr>
          </ItemTemplate>
        </asp:Repeater>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Are other 
            financial incentives being offered?</TD>
          <TD>
            <Epsom:NullableYesNo id="nynOtherFinancialIncentives" runat="Server" required="true" errormessage="Specify if other incentives being offered"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlOtherFinancialIncentivesMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details
              <BR>
              <asp:TextBox id="txtOtherFinancialIncentivesMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel></TABLE>
    </asp:panel>
    <asp:panel id="pnlRemortgageDetails" runat="server">
      <H2 class="form_heading">Remortgage information</H2>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Estimated value 
            of property</TD>
          <TD colSpan="3">
            <asp:TextBox enabled="false" id="txtRemortgageEstimatedPropertyValue" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator enabled="false" id="RequiredfieldEstimatedValueOfProperty" runat="server" errormessage="Please enter estimated value of property"
              controltovalidate="txtRemortgageEstimatedPropertyValue" text="***" enableclientscript="false"></asp:requiredfieldvalidator>
            <asp:rangevalidator enabled="false" id="Rangevalidator1" runat="server" errormessage="Estimated property value must be numeric"
              controltovalidate="txtRemortgageEstimatedPropertyValue" text="***" enableclientscript="false" minimumvalue="0"
              maximumvalue="9999999999" type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Amount required 
            to repay existing mortgage</TD>
          <TD colSpan="3">
            <asp:TextBox id="txtRemortgageAmountExisting" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator id="Requiredfieldvalidator1" runat="server" errormessage="Please enter amount to pay existing mortgage"
              controltovalidate="txtRemortgageAmountExisting" text="***" enableclientscript="false"></asp:requiredfieldvalidator>
            <asp:rangevalidator id="Rangevalidator2" runat="server" errormessage="Amount to repay existing mortgage must be numeric"
              controltovalidate="txtRemortgageAmountExisting" text="***" enableclientscript="false" minimumvalue="0"
              maximumvalue="9999999999" type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Amount required 
            to repay any 2nd or subsequent charges</TD>
          <TD colSpan="3">
            <asp:TextBox id="txtRemortgageSubsequentChargeAmount" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator id="Requiredfieldvalidator2" runat="server" errormessage="Please enter amount to pay subsequent charges"
              controltovalidate="txtRemortgageSubsequentChargeAmount" text="***" enableclientscript="false"></asp:requiredfieldvalidator>
            <asp:rangevalidator id="Rangevalidator3" runat="server" errormessage="Remortgage subsequent amount must be numeric"
              controltovalidate="txtRemortgageSubsequentChargeAmount" text="***" enableclientscript="false" minimumvalue="0"
              maximumvalue="9999999999" type="Currency"></asp:rangevalidator></TD>
        </TR> <!--    
				<TR>
					<TD class="prompt">Are you selling your existing property?</TD>
					<TD>
						<Epsom:NullableYesNo id="nynRemortgageSellingExistingProperty" runat="Server"></Epsom:NullableYesNo></TD>
				</TR>
				<TR>
					<TD class="prompt">If yes, what is the selling price</TD>
					<TD colSpan="3">
						<asp:TextBox id="txtRemortgageSellingPrice" Runat="server"></asp:TextBox></TD>
				</TR>
				<TR>
					<TD class="prompt">If no, what are your intentions with regard to this property?</TD>
					<TD colSpan="3">
						<asp:TextBox id="txtRemortgageIntentions" Runat="server"></asp:TextBox></TD>
				</TR>
      --></TABLE>
    </asp:panel>
    <H2 class="form_heading">Details of property to be mortgaged</H2>
    <FIELDSET>
      <LEGEND>Address or property</LEGEND>
      <epsom:Address id="ctlPropertyAddress" runat="server" UnitedKingdomOnly="True"></epsom:Address>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Tenure</TD>
          <TD>
            <asp:DropDownList id="ddlTenure" autopostback="true" Runat="server"></asp:DropDownList>
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator2" runat="server" errormessage="Please Select a tenure"
              controltovalidate="ddlTenure" text="***" enableclientscript="false"></validators:requiredselecteditemvalidator></TD>
        </TR>
        <asp:Panel id="pnlRemainingtermOfLease" runat="server">
          <TR>
            <TD class="prompt">If leasehold, remaining term of lease</TD>
            <TD>
              <asp:TextBox id="txtRemainingTermOfLease" CssClass="text" Runat="server"></asp:TextBox></TD>
          </TR>
        </asp:Panel>
        <asp:panel id="pnlRemortgagePreemptionDate" runat="server">
          <TR>
            <TD class="prompt">Right-to-buy Pre-emption date</TD>
            <TD colSpan="3">
              <asp:textbox id="txtRemortgagePreemptionDate" runat="server"></asp:textbox>
              <validators:datevalidator id="Datevalidator3" runat="server" errormessage="Invalid Right-to-buy Pre-emption date date"
                controltovalidate="txtRemortgagePreemptionDate" text="***" enableclientscript="false"></validators:datevalidator>dd/mm/yyyy
            </TD>
          </TR>
        </asp:panel></TABLE>
    </FIELDSET>
    <FIELDSET>
      <LEGEND>Property type</LEGEND>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Year of 
            construction</TD>
          <TD colSpan="2">
            <asp:TextBox id="txtYearOfConstruction" CssClass="text" autopostback="true" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator id="Requiredfieldvalidator3" runat="server" errormessage="Please enter year of construction"
              controltovalidate="txtYearOfConstruction" text="***" enableclientscript="false"></asp:requiredfieldvalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Is the property 
            new build</TD>
          <TD>
            <Epsom:NullableYesNo id="nynNewBuild" runat="Server" required="true" errormessage="Specify if it is new build"></Epsom:NullableYesNo></TD>
        </TR>
        <TR>
          <TD class="prompt">
            <asp:label id="asteriskLessThan10InPlace" runat="server" cssclass="mandatoryfieldasterisk">*</asp:label>If 
            &lt; 10 years old, which of the following is in place?</TD>
          <TD colSpan="2">
            <asp:DropDownList id="ddlLessThan10InPlace" Runat="server"></asp:DropDownList>
            <validators:requiredselecteditemvalidator id="validateLessThan10InPlace" runat="server" errormessage="Please Select a house builder guarantee"
              controltovalidate="ddlLessThan10InPlace" text="***" enableclientscript="false"></validators:requiredselecteditemvalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Type of 
            property?</TD>
          <TD colSpan="2">
            <asp:DropDownList id="ddlTypeOfProperty" autopostback="true" Runat="server"></asp:DropDownList>
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator5" runat="server" errormessage="Please Select type of property"
              controltovalidate="ddlTypeOfProperty" text="***" enableclientscript="false"></validators:requiredselecteditemvalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Detachment type?</TD>
          <TD colSpan="2">
            <asp:DropDownList id="ddlDetachmentType" autopostback="true" Runat="server"></asp:DropDownList>
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator6" runat="server" errormessage="Please Select detachment type"
              controltovalidate="ddlDetachmentType" text="***" enableclientscript="false"></validators:requiredselecteditemvalidator></TD>
        </TR>
      </TABLE>
    </FIELDSET>
    <FIELDSET>
      <LEGEND>If the property is a flat</LEGEND>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt">Number of floors in the block</TD>
          <TD>
            <asp:TextBox id="txtFlatNumberOfFloors" CssClass="text" Runat="server"></asp:TextBox>
            <asp:rangevalidator id="Rangevalidator4" runat="server" errormessage="Number of floors must be numeric"
              controltovalidate="txtFlatNumberOfFloors" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt">Number of flats in the block</TD>
          <TD>
            <asp:TextBox id="txtFlatNumberOfFlats" CssClass="text" Runat="server"></asp:TextBox>
            <asp:rangevalidator id="Rangevalidator5" runat="server" errormessage="Number of flats must be numeric"
              controltovalidate="txtFlatNumberOfFlats" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt">Which floor is the flat on</TD>
          <TD>
            <asp:TextBox id="txtFlatWhichFloor" CssClass="text" Runat="server"></asp:TextBox>
            <asp:rangevalidator id="Rangevalidator6" runat="server" errormessage="Floor number must be numeric"
              controltovalidate="txtFlatWhichFloor" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt">
            <asp:label id="lblFlatOverCommercial" runat="server" cssclass="mandatoryfieldasterisk">*</asp:label>Is 
            the flat above commercial premises</TD>
          <TD>
            <Epsom:NullableYesNo id="nynFlatAboveCommercialPremises" runat="Server" required="true" errormessage="Specify if flat is above a commercial premises"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlFlatAboveCommercialPremisesMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtFlatAboveCommercialPremisesMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel></TABLE>
    </FIELDSET>
    <FIELDSET>
      <LEGEND>Number of rooms</LEGEND>
      <TABLE class="formtable">
        <asp:repeater id="rptNewPropertyRoomTypes" runat="server">
          <itemtemplate>
            <tr>
              <td class="prompt">
                <asp:label id="lblRoomType" runat="server" /></td>
              <td>
                <asp:textbox id="txtNumberOfRooms" runat="server" CssClass="text" />
                <asp:rangevalidator id="Rangevalidator7" runat="server" enableclientscript="false" errormessage="Number of rooms is invalid"
                  controltovalidate="txtNumberOfRooms" minimumvalue="0" maximumvalue="9999999999" type="Currency" text="***"></asp:rangevalidator>
              </td>
            </tr>
          </itemtemplate>
        </asp:repeater></TABLE>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt">Outbuildings</TD>
          <TD>
            <asp:TextBox id="txtNumOutBuildings" CssClass="text" Runat="server"></asp:TextBox>
            <asp:rangevalidator id="Rangevalidator8" runat="server" errormessage="Outbuilding count must be numeric"
              controltovalidate="txtNumOutbuildings" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt">Garages</TD>
          <TD>
            <asp:TextBox id="txtNumGarages" CssClass="text" Runat="server"></asp:TextBox>
            <asp:rangevalidator id="Rangevalidator9" runat="server" errormessage="Garage count must be numeric"
              controltovalidate="txtNumGarages" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR> <!--        <tr>
          <td class="prompt">Other</td>
          <td><asp:TextBox id="txtNumOtherRooms" Runat="server" CssClass="text" /></td>
        </tr> --></TABLE>
    </FIELDSET>
    <FIELDSET>
      <LEGEND>Details</LEGEND>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Building 
            construction</TD>
          <TD colSpan="2">
            <asp:DropDownList id="cmbBuildingConstruction" autopostback="true" Runat="server"></asp:DropDownList>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator3" Runat="server" EnableClientScript="false" Text="***"
              ErrorMessage="Building construction is required" ControlToValidate="cmbBuildingConstruction"></validators:RequiredSelectedItemValidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Roof 
            construction</TD>
          <TD colSpan="2">
            <asp:DropDownList id="cmbRoofConstruction" autopostback="true" Runat="server"></asp:DropDownList>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator1" Runat="server" EnableClientScript="false" Text="***"
              ErrorMessage="Roof construction is required" ControlToValidate="cmbRoofConstruction"></validators:RequiredSelectedItemValidator></TD>
        </TR>
        <asp:panel id="pnlStandardMaterialsMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtStandardMaterialsMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>What land does 
            the property have with it?</TD>
          <TD>
            <asp:DropDownList id="ddlNumAcres" runat="server"></asp:DropDownList>
            <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator7" Runat="server" EnableClientScript="false" Text="***"
              ErrorMessage="Please select what land the property has" ControlToValidate="ddlNumAcres"></validators:RequiredSelectedItemValidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Is the property 
            affected by an agricultural restriction/covenant</TD>
          <TD colSpan="3">
            <Epsom:NullableYesNo id="nynAgriculturalRestriction" runat="Server" required="true" errormessage="Specify if there are agricultural restrictions"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlAgriculturalRestrictionMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtAgriculturalRestrictionMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Will the 
            property be your main residence</TD>
          <TD colSpan="3">
            <Epsom:NullableYesNo id="nynMainResidence" runat="Server" required="true" errormessage="Specify if property will be main residence"></Epsom:NullableYesNo></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Will any part of 
            the property be used for business purposes</TD>
          <TD colSpan="3">
            <Epsom:NullableYesNo id="nynBusinessPurpose" runat="Server" required="true" errormessage="Specify if property will be used for business purpose"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlBusinessPurposeMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtBusinessPurposeMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel>
        <TR>
          <TD class="prompt">If you previously purchased the property from the local 
            authority, was it within the last 3 years</TD>
          <TD>
            <Epsom:NullableYesNo id="nynLocalAuthorityWithinThreeYears" runat="Server"></Epsom:NullableYesNo></TD>
          <TD>If yes, date purchased</TD>
          <TD>
            <asp:TextBox id="txtDatePurchasedFromLocalAuthority" runat="server" CssClass="text"></asp:TextBox></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Will the 
            property be let or used for any other purpose than your main residence</TD>
          <TD colSpan="3">
            <Epsom:NullableYesNo id="nynAnyOtherPurpose" runat="Server" required="true" errormessage="Specify if any other purpose of property"
              autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlAnyOtherPurposeMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtAnyOtherPurposeMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel></TABLE>
    </FIELDSET>
    <FIELDSET>
      <LEGEND>Details of other occupiers</LEGEND>
      <TABLE class="formtable">
        <TR style="TEXT-ALIGN: center">
          <TD>Please give details of all persons, other than the applicants, over the age 
            of 17 who will occupy the property</TD>
        </TR>
      </TABLE>
      <asp:repeater id="rptOccupiers" runat="server">
        <itemtemplate>
          <table class="formtable">
            <tr>
              <td class="prompt"><span class="mandatoryfieldasterisk">* </span>First name</td>
              <td>
                <asp:textbox id="txtOtherOccupierFirstName" runat="server" CssClass="text" />
                <asp:requiredfieldvalidator runat="server" controltovalidate="txtOtherOccupierFirstName" text="***" errormessage="Please enter other occupier first name"
                  id="Requiredfieldvalidator4" enableclientscript="false" />
              </td>
            </tr>
            <tr>
              <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Surname</td>
              <td>
                <asp:textbox id="txtOtherOccupierLastName" runat="server" CssClass="text" />
                <asp:requiredfieldvalidator runat="server" controltovalidate="txtOtherOccupierLastName" text="***" errormessage="Please enter other occupier last name"
                  id="Requiredfieldvalidator5" enableclientscript="false" />
              </td>
            </tr>
            <tr>
              <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Date of birth</td>
              <td>
                <asp:textbox id="txtOtherOccupierDateOfBirth" runat="server" CssClass="text" />
                <validators:datevalidator id="DateStartedValidator" runat="server" controltovalidate="txtOtherOccupierDateOfBirth"
                  text="***" errormessage="Invalid other occupant date of birth" enableclientscript="false" />dd/mm/yyyy
              </td>
            </tr>
            <tr>
              <td class="prompt"><span class="mandatoryfieldasterisk">* </span>Relationship to 
                applicant(s)</td>
              <td>
                <asp:DropDownList id="ddlOtherOccupierRelationship" runat="server" />
                <validators:RequiredSelectedItemValidator id="Requiredselecteditemvalidator8" Runat="server" EnableClientScript="false" Text="***"
                  ErrorMessage="Please select other occuper relationship to applicant" ControlToValidate="ddlOtherOccupierRelationship"></validators:RequiredSelectedItemValidator>
              </td>
            </tr>
          </table>
        </itemtemplate>
      </asp:repeater>
      <P class="button_orphan">
        <asp:Button id="cmdAddOtherOccupier" runat="server" text="Add another occupier"></asp:Button>
        <asp:Button id="cmdRemoveAnOccupier" runat="server" text="Remove an occupier"></asp:Button></P>
    </FIELDSET>
    <H2 class="form_heading">Buy to let details</H2>
    <asp:panel id="pnlBuyToLetDetails" runat="server">
      <TABLE class="formtable">
        <TR>
          <TD class="prompt">If applying on a buy to let product, what is the anticipated 
            monthly rental</TD>
          <TD>
            <asp:textbox id="txtBuyToLetIncome" runat="server" CssClass="text"></asp:textbox>
            <asp:rangevalidator id="Rangevalidator10" runat="server" errormessage="Buy to let income must be numeric"
              controltovalidate="txtBuyToLetIncome" text="***" enableclientscript="false" minimumvalue="0" maximumvalue="9999999999"
              type="Currency"></asp:rangevalidator></TD>
        </TR>
        <TR>
          <TD class="prompt">If buy to let, will you or a relative ocupy the property now 
            or in the future</TD>
          <TD>
            <Epsom:NullableYesNo id="nynWillRelativeOccupy" runat="Server" autopostback="true"></Epsom:NullableYesNo></TD>
        </TR>
        <asp:panel id="pnlWillRelativeOccupyMemo" runat="server">
          <TR>
            <TD class="prompt"></TD>
            <TD>Please give details<BR>
              <asp:TextBox id="txtWillRelativeOccupyMemo" runat="server" TextMode="MultiLine"></asp:TextBox></TD>
          </TR>
        </asp:panel></TABLE>
    </asp:panel>
    <asp:panel id="pnlValuerDetails" runat="server">
      <H2 class="form_heading">Valuer Details</H2>
      <TABLE class="formtable"> <!-- valuation type moved to Loan requirements screen -->
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">*</SPAN>
            Who should we contact to arrange the valuation</TD>
          <TD>
            <asp:dropdownlist id="cmbAccessArrangements" runat="server"></asp:dropdownlist>
            <validators:requiredselecteditemvalidator id="Requiredselecteditemvalidator4" runat="server" errormessage="Contact for valuation is required"
              controltovalidate="cmbAccessArrangements" text="***" enableclientscript="false"></validators:requiredselecteditemvalidator></TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">*</SPAN>
            Original valuation date</TD>
          <TD>
            <asp:textbox id="txtOriginalValuationDate" runat="server" cssclass="text"></asp:textbox>
            <validators:datevalidator id="Datevalidator4" runat="server" errormessage="Invalid original valuation date"
              controltovalidate="txtOriginalValuationDate" text="***" enableclientscript="false"></validators:datevalidator>dd/mm/yyyy
          </TD>
        </TR>
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">*</SPAN>
            Original valuer company name</TD>
          <TD>
            <asp:textbox id="txtOriginalValuer" runat="server" cssclass="text"></asp:textbox>
            <asp:requiredfieldvalidator id="txtOriginalValuer_validator" runat="server" errormessage="Original Valuer is required"
              controltovalidate="txtOriginalValuer" text="***" enableclientscript="false"></asp:requiredfieldvalidator></TD>
        </TR>
      </TABLE>
      <epsom:address id="ctlValuerAddress" runat="server" unitedkingdomonly="True"></epsom:address>
    </asp:panel>
    <H2 class="form_heading">Access to the property</H2>
    <TABLE class="formtable">
      <TR>
        <TD class="prompt">Name of the Selling Agent</TD>
        <TD>
          <asp:textbox id="txtSellingAgentName" runat="server" CssClass="text"></asp:textbox></TD>
      </TR>
      <TR>
         <td><asp:Label id="lblSellingAgentTelephone"  CssClass="prompt">Selling Agent Telephone</asp:Label></td>
          <td><asp:TextBox id="txtSellingAgentTelArea" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
          <td><asp:TextBox id="txtSellingAgentTelNumber" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
          <td><asp:TextBox id="txtSellingAgentTelExtension" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
      </TR>
    </TABLE>
    <FIELDSET>
      <LEGEND>Vendor's details</LEGEND>
      <TABLE class="formtable">
        <TR>
          <TD class="prompt"><SPAN class="mandatoryfieldasterisk">* </SPAN>Vendor's 
            forename and surname</TD>
          <TD>
            <asp:TextBox id="txtAgentFirstname" CssClass="text" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator id="RequiredfieldAgentFirstname" runat="server" errormessage="Please enter the vendor's firstname"
              controltovalidate="txtAgentFirstname" text="***" enableclientscript="false"></asp:requiredfieldvalidator>
            <asp:TextBox id="txtAgentSurname" CssClass="text" Runat="server"></asp:TextBox>
            <asp:requiredfieldvalidator id="RequiredfieldAgentSurname" runat="server" errormessage="Please enter the vendor's surname"
              controltovalidate="txtAgentSurname" text="***" enableclientscript="false"></asp:requiredfieldvalidator></TD>
        </TR>
      </TABLE>
      <epsom:Address id="ctlAgentAddress" runat="server" UnitedKingdomOnly="True"></epsom:Address>
      <TABLE class="formtable">
        <TR>
          <td><asp:Label id="lblVendorLabel"  CssClass="prompt">Vendor Telephone</asp:Label></td>
          <td><asp:TextBox id="txtVendorTelArea" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
          <td><asp:TextBox id="txtVendorTelNumber" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
          <td><asp:TextBox id="txtVendorTelExt" CssClass="text" size="12" Runat="server"></asp:TextBox></td>
        </TR>
      </TABLE>
    </FIELDSET>
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server" savepage="true" saveandclosepage="true"
      buttonbar="true"></epsom:dipnavigationbuttons>
    <sstchur:SmartScroller id="ctlSmartScroller" runat="server"></sstchur:SmartScroller>
  </mp:content>
</mp:contentcontainer>
