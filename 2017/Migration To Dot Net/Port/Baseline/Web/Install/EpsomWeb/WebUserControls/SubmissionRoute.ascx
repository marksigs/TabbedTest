<%@ Control Language="c#" AutoEventWireup="false" Codebehind="SubmissionRoute.ascx.cs" Inherits="Epsom.Web.WebUserControls.SubmissionRoute" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"%>

<asp:validationsummary runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true" id="Validationsummary1" />

    <asp:Panel id="pnlSubmissionRoute" runat="server">
    
      <h2 class="form_heading">How do you want to submit business to db mortgages?</h2>
        <table class="formtable">
        
          <tr>
            <td class="prompt"></td>
            <td>
              <asp:RadioButtonList id="rblSubmissionRoute" runat="Server" AutoPostback="true">
                <asp:listItem Value="1">Direct to db mortgages - no third party involved</asp:listItem>
                <asp:ListItem Value="2">Direct to db mortgages - as a network member</asp:ListItem>
                <asp:ListItem Value="3">Direct to db mortgages - under a master broker arrangement</asp:ListItem>
                <asp:ListItem Value="4">Via a packager</asp:ListItem>
              </asp:RadioButtonList>
            </td>
          </tr>
          
          <asp:Panel id="pnlChoosePackager" runat="server">
          <tr>
            <td class="prompt">
              Choose Packager
            </td>
            <td>
              <asp:dropdownlist id="ddlChoosePackager" runat="server" />
            </td>
          </tr>
          </asp:Panel>

          <asp:Panel id="pnlChooseMortgageClub" runat="server">
          <tr>
            <td class="prompt">
              Choose Mortgage Club
            </td>
            <td>
              <asp:dropdownlist id="ddlChooseMortgageClub" runat="server" />
            </td>
          </tr>
          </asp:Panel>
        </table>
    </asp:Panel>
      
    <h2 class="form_heading">Loan type</h2>
      <table class="formtable">
        <tr>
          <td class="prompt"></td>
          <td>
            <asp:RadioButtonList id="rblMultipleComponent" runat="Server" >
              <asp:ListItem Value="0">Single component</asp:listItem>
              <asp:ListItem Value="1">Multi component</asp:ListItem>
            </asp:RadioButtonList>
          </td>
        </tr>
      </table>