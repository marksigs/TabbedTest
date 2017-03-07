<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Page Language="c#" codebehind="WebCalc.aspx.cs" autoeventwireup="false" Inherits="WebCalculator.WebCalculator" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage3.ascx"><mp:content id=region2 runat="server">
<H1>Flexible Calculator</H1>
<TABLE class=formtable>
  <TR>
    <TD class=prompt width=167><asp:label id=lblAnnualIncome runat="server" text="Value / Price of property"></asp:label></TD>
    <TD width=216><asp:textbox id=txtValuePrice runat="server" Width="110px"></asp:textbox><asp:rangevalidator id=rngValidatorValueOfProperty runat="server" ErrorMessage="Price / Value must be numeric and less than £2,000,000." ControlToValidate="txtValuePrice" MaximumValue="99999999" MinimumValue="0" Type="Integer">***</asp:rangevalidator></TD>
    <TD width=19></TD>
    <TD class=prompt width=122><SPAN class=mandatoryfieldasterisk>* </SPAN><asp:label id=lblLoanAmount runat="server">Loan/Mortgage amount Required</asp:label></TD>
    <TD width=180><asp:textbox id=txtLoanAmount runat="server" Width="110px"></asp:textbox><asp:RequiredFieldValidator id=rfValidatorLoanAmount ErrorMessage="Please enter Loan Amount" ControlToValidate="txtLoanAmount" Runat="Server" Display="Dynamic">***</asp:RequiredFieldValidator><asp:rangevalidator id=rngValidatorLoanAmount runat="server" ErrorMessage="Loan Amount must be numeric and less than £2,000,000." ControlToValidate="txtLoanAmount" MaximumValue="99999999" MinimumValue="0" Type="Integer" Display="Dynamic">***</asp:rangevalidator></TD></TR>
  <TR>
    <TD class=prompt width=167><SPAN class=mandatoryfieldasterisk>* </SPAN><asp:label id=Label2 runat="server"> Does
			the Mortgage have a Fixed or discount Period?</asp:label></TD>
    <TD width=216><asp:radiobutton id=rbFixedPeriodYes runat="server" Width="24px" AutoPostBack="true" GroupName="OptFixedPeriod" Text="Yes"></asp:radiobutton><asp:radiobutton id=rbFixedPeriodNo runat="server" Width="24px" AutoPostBack="true" GroupName="OptFixedPeriod" Text="No" Checked="True"></asp:radiobutton></TD>
    <TD class=prompt width=19></TD>
    <TD width=122></SPAN><SPAN class=mandatoryfieldasterisk>* </SPAN><asp:label id=lblduration runat="server">Duration
        / Term in Years</asp:label></TD>
    <TD width=180><asp:dropdownlist id=dpnDuration runat="server" Width="110px" AutoPostBack="True">
							<asp:ListItem Value="1">1</asp:ListItem>
							<asp:ListItem Value="2">2</asp:ListItem>
							<asp:ListItem Value="3">3</asp:ListItem>
							<asp:ListItem Value="4">4</asp:ListItem>
							<asp:ListItem Value="5">5</asp:ListItem>
							<asp:ListItem Value="6">6</asp:ListItem>
							<asp:ListItem Value="7">7</asp:ListItem>
							<asp:ListItem Value="8">8</asp:ListItem>
							<asp:ListItem Value="9">9</asp:ListItem>
							<asp:ListItem Value="10">10</asp:ListItem>
							<asp:ListItem Value="11">11</asp:ListItem>
							<asp:ListItem Value="12">12</asp:ListItem>
							<asp:ListItem Value="13">13</asp:ListItem>
							<asp:ListItem Value="14">14</asp:ListItem>
							<asp:ListItem Value="15">15</asp:ListItem>
							<asp:ListItem Value="16">16</asp:ListItem>
							<asp:ListItem Value="17">17</asp:ListItem>
							<asp:ListItem Value="18">18</asp:ListItem>
							<asp:ListItem Value="19">19</asp:ListItem>
							<asp:ListItem Value="20">20</asp:ListItem>
							<asp:ListItem Value="21">21</asp:ListItem>
							<asp:ListItem Value="22">22</asp:ListItem>
							<asp:ListItem Value="23">23</asp:ListItem>
							<asp:ListItem Value="24">24</asp:ListItem>
							<asp:ListItem Value="25">25</asp:ListItem>
							<asp:ListItem Value="26">26</asp:ListItem>
							<asp:ListItem Value="27">27</asp:ListItem>
							<asp:ListItem Value="28">28</asp:ListItem>
							<asp:ListItem Value="29">29</asp:ListItem>
							<asp:ListItem Value="30">30</asp:ListItem>
							<asp:ListItem Value="31">31</asp:ListItem>
							<asp:ListItem Value="32">32</asp:ListItem>
							<asp:ListItem Value="33">33</asp:ListItem>
							<asp:ListItem Value="34">34</asp:ListItem>
							<asp:ListItem Value="35">35</asp:ListItem>
						</asp:dropdownlist><asp:RequiredFieldValidator id=rfDuration ErrorMessage="Please select a Duration / Term" ControlToValidate="dpnDuration" Runat="server">***</asp:RequiredFieldValidator></TD></TR>
  <TR>
    <TD class=prompt width=167><asp:label class=mandatoryfieldasterisk id=lblMFA1 runat="server">* </asp:label><asp:label id=lblIntFixedPeriod runat="server">Interest Rate for fixed / Discount Period</asp:label></TD>
    <TD width=216><asp:textbox id=txtFixedInterest1 runat="server" Width="43px"></asp:textbox><asp:label class=prompt id=lblPercent1 runat="server">%</asp:label><asp:rangevalidator id=rngValidatorFixedInt1 runat="server" ErrorMessage="Discount Period Interest Rate must be a numeric value." ControlToValidate="txtFixedInterest1" MaximumValue="100" MinimumValue="0" Type="Double" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfFixedInt1 ErrorMessage="Please enter Interest Rate for Fixed Period" ControlToValidate="txtFixedInterest1" Runat="server" Display="Dynamic">***</asp:RequiredFieldValidator></TD>
    <TD class=prompt width=19></TD>
    <TD class=prompt width=122><SPAN><SPAN class=mandatoryfieldasterisk>* 
      </SPAN><asp:label id=lblRepaymentType runat="server">Repayment
        Type</asp:label></SPAN></TD>
    <TD width=180><asp:dropdownlist id=dpnPaymentType runat="server" Width="110px" AutoPostBack="True">
							<asp:ListItem Value="repayment">Repayment</asp:ListItem>
							<asp:ListItem Value="intonly">Interest Only</asp:ListItem>
							<asp:ListItem Value="partandpart">Part &amp; Part</asp:ListItem>
						</asp:dropdownlist><asp:RequiredFieldValidator id=rfRepaymentType ErrorMessage="Please select Repayment Type" ControlToValidate="dpnPaymentType" Runat="server">***</asp:RequiredFieldValidator></TD></TR>
  <TR>
    <TD class=prompt width=167 height=40></SPAN><SPAN 
      class=mandatoryfieldasterisk>* </SPAN><asp:label id=lblInterestRate runat="server">Interest
        Rate</asp:label></TD>
    <TD width=216 height=40><asp:textbox id=txtInterest1 runat="server" Width="43px"></asp:textbox><asp:label class=prompt id=lblPercent2 runat="server">%</asp:label><asp:rangevalidator id=rngValidatorInt1 runat="server" ErrorMessage="Interest Rate must be a numeric value." ControlToValidate="txtInterest1" MaximumValue="100" MinimumValue="0" Type="Double" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfInt1 ErrorMessage="Please enter Interest Rate" ControlToValidate="txtInterest1" Runat="server" Display="Dynamic">***</asp:RequiredFieldValidator></TD>
    <TD width=19 height=40></TD>
    <TD class=prompt width=122 height=40><asp:label class=mandatoryfieldasterisk id=lblMFA3 runat="server">* </asp:label><asp:label id=lblIntOnlyAmt runat="server">If Part & Part Interest only amount</asp:label></SPAN></TD>
    <TD width=180 height=40><asp:textbox id=txtIntOnlyAmt runat="server" Width="110px"></asp:textbox><asp:rangevalidator id=rngValidatorIntOnlyAmount runat="server" ErrorMessage="Interest Only Amount must be less than or equal to the Loan Amount." ControlToValidate="txtIntOnlyAmt" MaximumValue="2000000" MinimumValue="0" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfIntOnlyAmount ErrorMessage="Please enter Part and Part Interest Only Amount" ControlToValidate="txtIntOnlyAmt" Runat="server" Display="Dynamic">***</asp:RequiredFieldValidator></TD></TR>
  <TR>
    <TD class=prompt width=167><asp:label class=mandatoryfieldasterisk id=lblMFA2 runat="server">* </asp:label><asp:label id=lblMonthsFixedPeriod runat="server">How many months does this period last?</asp:label></TD>
    <TD width=216 height=40><asp:textbox id=txtMonthsFixedPeriod runat="server" Width="110px"></asp:textbox><asp:rangevalidator id=rngValidatorMonthsFixedPeriod runat="server" ErrorMessage="Discount Period must be number of months" ControlToValidate="txtMonthsFixedPeriod" MaximumValue="420" MinimumValue="0" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfMonthsFixedPeriod ErrorMessage="Please enter Discount Period Months" ControlToValidate="txtMonthsFixedPeriod" Runat="server" Display="Dynamic">***</asp:RequiredFieldValidator></TD>
    <TD width=19 height=40></TD>
    <TD class=prompt width=122 height=40></TD>
    <TD width=180 height=40></TD></TR>
  <TR>
    <TD class=prompt width=167>&nbsp;</TD>
  <TR>
    <TD class=prompt width=167 height=24><asp:label id=lblMonthlyOverPayment runat="server">Monthly
        overpayment</asp:label></TD>
    <TD width=216 height=24><asp:textbox id=txtMonthlyOverPayment runat="server" Width="110px"></asp:textbox><asp:rangevalidator id=rngValidatorMonthlyOverpayment runat="server" ErrorMessage="Monthly Overpayment must be integer value." ControlToValidate="txtMonthlyOverPayment" MaximumValue="10000" MinimumValue="0" Type="Integer" Display="Dynamic">***</asp:rangevalidator></TD>
    <TD width=19 height=24></TD>
    <TD class=prompt width=122 height=24><asp:label id=lblInitialFlexPayment runat="server">Initial
        flexible monthly payment</asp:label></TD>
    <TD width=180 height=24><asp:textbox id=txtInitialFlexPayment runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD></TR>
  <TR>
    <TD class=prompt width=167><asp:label id=lblLumpSumOverPayment runat="server">Lump
        sum Overpayment</asp:label></TD>
    <TD width=216><asp:textbox id=txtLumpSumOverPayment1 runat="server" Width="64px" AutoPostBack="True"></asp:textbox><asp:label id=lblInYear1 runat="server"> in year </asp:label><asp:textbox id=txtLumpSumOverPayment2 runat="server" Width="23px"></asp:textbox><asp:rangevalidator id=rngValidatorOverPayment1 runat="server" ErrorMessage="Lump Sum Overpayment amount must be integer value." ControlToValidate="txtLumpSumOverPayment1" MaximumValue="9999999" MinimumValue="0" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:rangevalidator id=rngValidatorOverPayment2 runat="server" ErrorMessage="Lump Sum Overpayment year must be integer value and within mortgage term" ControlToValidate="txtLumpSumOverPayment2" MaximumValue="35" MinimumValue="1" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfLSOP2 ErrorMessage="Please enter Lump Sum Overpayment Year" ControlToValidate="txtLumpSumOverPayment2" Runat="server" Display="Dynamic" Enabled="False">***</asp:RequiredFieldValidator></TD>
    <TD width=19></TD>
    <TD class=prompt width=122><asp:label id=lblRevertingTo runat="server">Reverting
        to</asp:label></TD>
    <TD width=180><asp:textbox id=txtRevertingTo runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD></TR>
  <TR>
    <TD class=prompt width=167 height=27><asp:label id=lblPaymentHoliday runat="server">Payment
        holiday in Year</asp:label></TD>
    <TD width=216 height=27><asp:textbox id=txtPaymentHoliday1 runat="server" Width="32px" AutoPostBack="True"></asp:textbox><asp:label id=lblFor runat="server">for</asp:label><asp:textbox id=txtPaymentHoliday2 runat="server" Width="32px"></asp:textbox><asp:label id=lblMonths runat="server">months</asp:label><asp:rangevalidator id=rngValidatorPaymentHoliday1 runat="server" ErrorMessage="Payment Holiday year must be be numeric and within mortgage term" ControlToValidate="txtPaymentHoliday1" MaximumValue="35" MinimumValue="1" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:rangevalidator id=rngValidatorPaymentHoliday2 runat="server" ErrorMessage="Payment Holiday period must be integer value." ControlToValidate="txtPaymentHoliday2" MaximumValue="36" MinimumValue="1" Type="Integer" Display="Dynamic">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfPayHols2 ErrorMessage="Please enter Payment Holiday Months" ControlToValidate="txtPaymentHoliday2" Runat="server" Display="Dynamic" Enabled="False">***</asp:RequiredFieldValidator></TD>
    <TD width=19 height=27></TD>
    <TD class=prompt width=122 height=27><asp:label id=lblOriginalLoanCost runat="server">Original
        Loan Cost</asp:label></TD>
    <TD width=180 height=27><asp:textbox id=txtOriginalLoanCost runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD></TR>
  <TR>
    <TD class=prompt width=167><asp:label id=lblBorrowBack runat="server" Visible="False">Borrow
        Back</asp:label></TD>
    <TD width=216><asp:textbox id=txtBorrowBack1 runat="server" Width="64px" AutoPostBack="True" Visible="False"></asp:textbox><asp:label id=lblInYear2 runat="server" Visible="False"> in year </asp:label><asp:textbox id=txtBorrowBack2 runat="server" Width="23px" Visible="False"></asp:textbox><asp:rangevalidator id=rngValidatorBorrowBack2 runat="server" Width="8px" ErrorMessage="BorrowBack year must be be numeric and within mortgage term" ControlToValidate="txtBorrowBack2" MaximumValue="35" MinimumValue="1" Type="Integer" Display="Dynamic" Visible="False">***</asp:rangevalidator><asp:rangevalidator id=rngValidatorBorrowBack1 runat="server" Width="8px" ErrorMessage="Borrow Back amount must be integer value." ControlToValidate="txtBorrowBack1" MaximumValue="1000000" MinimumValue="0" Type="Integer" Display="Dynamic" Visible="False">***</asp:rangevalidator><asp:RequiredFieldValidator id=rfBorrowBackYear ErrorMessage="Please enter Borrow Back Year" ControlToValidate="txtBorrowBack2" Runat="server" Display="Dynamic" Enabled="False" Visible="False">***</asp:RequiredFieldValidator>&nbsp;</TD>
    <TD width=19></TD>
    <TD class=prompt width=122><asp:label id=lblFinalLoancost runat="server">Final
        loan cost</asp:label></TD>
    <TD width=180><asp:textbox id=txtFinalLoancost runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD></TR>
  <TR>
    <TD class=prompt width=167 height=27><asp:label id=lblRepaidOn runat="server">Mortgage
        repaid on</asp:label></TD>
    <TD width=216 height=27><asp:textbox id=txtRepaidOn runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD>
    <TD class=prompt width=19 height=27>&nbsp;</TD>
    <TD class=prompt width=122 height=27><asp:label id=lblInterestSaved runat="server">Interest
        saved</asp:label></TD>
    <TD width=180 height=27><asp:textbox id=txtInterestSaved runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD></TR>
  <TR>
    <TD class=prompt width=167><asp:label id=lblCurrentMonthlyPayment runat="server">Current
        monthly payment</asp:label></TD>
    <TD width=216><asp:textbox id=txtCurrentMonthlyPayment runat="server" Width="110px" ReadOnly="True"></asp:textbox></TD>
    <TD class=prompt width=19>&nbsp;</TD>
    <TD class=prompt width=122>&nbsp;</TD>
    <TD class=prompt width=180>&nbsp;</TD></TR>
  <TR>
    <TD class=prompt width=167></TD>
    <TD class=prompt width=216></TD>
    <TD class=prompt width=19></TD>
    <TD width=122><asp:button id=btnReCalc runat="server" Width="110px" Text="Re-Calculate"></asp:button></TD>
  <TR>
    <TD class=prompt width=167>&nbsp;</TD>
  <TR>
  <TR>
    <TD class=prompt width=167>&nbsp;</TD>
    <TD colSpan=8><asp:image id=imgGraph runat="server" Width="500px" Visible="False" BorderWidth="1" BorderStyle="Solid" Height="325px"></asp:image></TD></TR></TABLE><asp:validationsummary id=ValidationSummary runat="server" Width="187px" Height="71px" ShowSummary="False" headertext="Input errors:" ShowMessageBox="True"></asp:validationsummary></FORM>
    </mp:content>

</mp:contentcontainer>
