<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Page language="c#" Codebehind="IncomeDetails.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.IncomeDetails" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix = "sstchur" Namespace = "sstchur.web.SmartNav" Assembly = "sstchur.web.SmartNav" %>
<%@ Register Src="~/WebUserControls/NullableYesNo.ascx" TagName="NullableYesNo" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server" />

  </mp:content>

  <mp:content id="region2" runat="server">
<cms:helplink id="helplink" class="helplink" runat="server" helpref="2260" />
    <h1><cms:cmsBoilerplate cmsref="624" runat="server" /><asp:label id="lblHeading" runat="server" /></h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server" />

    <asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />
    
    <cms:cmsBoilerplate cmsref="260" runat="server" />
    
    <asp:panel id="pnlEmployedIncome" runat="server">

      <h2 class="form_heading">Employed applicants</h2>

      <p class="forminfo"><cms:cmsBoilerplate cmsref="261" runat="server" /></p>

      <table class="formtable">

        <asp:repeater id="rptEmployedEarnedIncomeTypes" runat="server">

          <itemtemplate>

            <tr>
              <td class="prompt">
              <span class="mandatoryfieldasterisk"><asp:label id="lblMandatoryEarnedIncomeType" runat="server" text="*"/></span> 
              <asp:label id="lblEmployedEarnedIncomeType" runat="server" />
              </td>
              <td>
                <asp:textbox id="txtEmployedEarnedIncomeAmount" CssClass="text" runat="server" />
                <asp:rangevalidator id="Rangevalidator1" runat="server" enableclientscript="false" 
                  errormessage="Income amount must be numeric" controltovalidate="txtEmployedEarnedIncomeAmount" 
                  MinimumValue="0" MaximumValue="9999999999" Type="Currency" text="***"></asp:rangevalidator>
                <asp:requiredfieldvalidator controltovalidate="txtEmployedEarnedIncomeAmount" runat="server" text="***" 
                  errormessage="Please enter income amount" id="validateMandatoryEmployedEarnedIncomeAmount" enableclientscript="false"/>
              </td>
            </tr>

          </itemtemplate>

        </asp:repeater>

      </table>

      <table class="formtable">
        <tr>
          <td class="prompt"><b><asp:label id="lblEmployedIncomeTotal" runat="server" text="Total guaranteed income" /></b></td>
          <td>
            <asp:textbox id="txtEmployedIncomeTotal" Enabled="false" CssClass="text" runat="server" tooltip="Total guaranteed/regular income" />
            <asp:button id="cmdCalculateEmployedIncomeTotal" runat="server" text="Update" />
          </td>
        </tr>
      </table>

      <table class="formtable">

        <asp:repeater id="rptNonGuaranteedEarnedIncomeTypes" runat="server">

          <itemtemplate>

            <tr>
              <td class="prompt"><asp:label id="lblNonGuaranteedEarnedIncomeType" runat="server" /></td>
              <td>
                <asp:textbox id="txtNonGuaranteedEarnedIncomeAmount" CssClass="text" runat="server" />
                <asp:rangevalidator id="Rangevalidator2" runat="server" enableclientscript="false" errormessage="Income amount must be numeric" controltovalidate="txtNonGuaranteedEarnedIncomeAmount" MinimumValue="0" MaximumValue="9999999999" Type="Currency" text="***"></asp:rangevalidator>
              </td>
            </tr>

          </itemtemplate>

        </asp:repeater>

      </table>
  
    </asp:panel>
    
    <asp:panel id="pnlSelfEmployedIncome" runat="server">

      <h2 class="form_heading">Self-employed applicants</h2>

      <table class="formtable">
        <tr>
          <td class="prompt">
          <span class="mandatoryfieldasterisk">*</span> 
          <asp:Label id="lblSelfEmployedIncome" runat="server" text="All income details derived from self employment" />
          </td>
          <td>
            <asp:textbox id="txtSelfEmployedIncome" CssClass="text" runat="server" tooltip="Self employment income" />
            <asp:rangevalidator id="Rangevalidator3" runat="server" enableclientscript="false" 
               errormessage="Self-employed income amount must be numeric" controltovalidate="txtSelfEmployedIncome" MinimumValue="0" MaximumValue="9999999999" Type="Currency" text="***"></asp:rangevalidator>
             <asp:requiredfieldvalidator controltovalidate="txtSelfEmployedIncome" runat="server" text="***" 
                  errormessage="Please enter self employed income" id="Requiredfieldvalidator1" enableclientscript="false"/>
          </td>
          <td>Including net profit and drawings</td>
        </tr>
      </table>    

    </asp:panel>
    
    <asp:panel id="pnlOtherIncome" runat="server">

      <h2 class="form_heading">Other income</h2>

      <table class="formtable">

        <asp:repeater id="rptOtherEarnedIncomeTypes" runat="server">

          <itemtemplate>

            <tr>
              <td class="prompt"><asp:label id="lblOtherEarnedIncomeType" runat="server" /></td>
              <td>
                <asp:textbox id="txtOtherEarnedIncomeAmount" runat="server" CssClass="text" autopostback="true"/>
                <asp:rangevalidator id="Rangevalidator4" runat="server" enableclientscript="false" errormessage="Income amount must be numeric" controltovalidate="txtOtherEarnedIncomeAmount" MinimumValue="0" MaximumValue="9999999999" Type="Currency" text="***"></asp:rangevalidator>
              </td>
            </tr>

            <asp:panel id="pnlOtherEarnedIncomeDescription" runat="server" visible="false">
              <tr>
                <td class="prompt">&nbsp;</td>
                <td>
                  Please give details
                  <br />
                  <asp:TextBox id="txtOtherEarnedIncomeDescription" TextMode="MultiLine" runat="server" />
                </td>
              </tr>
            </asp:panel>

          </itemtemplate>

        </asp:repeater>

      </table>

      <table class="formtable">
        <tr>
          <td class="prompt"><b><asp:label id="lblOtherIncomeTotal" runat="server" text="Total income" /></b></td>
          <td>
            <asp:textbox id="txtOtherIncomeTotal" Enabled="false" CssClass="text" runat="server" tooltip="Total other income" />
            <asp:button id="cmdCalculateOtherIncomeTotal" runat="server" text="Recalculate" />
          </td>
        </tr>
      </table>

    </asp:panel>

    <br /><br />
    
    <epsom:dipnavigationbuttons buttonbar="true" savepage="true" saveandclosepage="true" id="ctlDIPNavigationButtons2" runat="server" />

    <sstchur:SmartScroller id="ctlSmartScroller" runat="server" />

  </mp:content>

</mp:contentcontainer>

