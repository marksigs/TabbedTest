<%@ Page language="c#" Codebehind="InsuranceAndInvestment.aspx.cs" AutoEventWireup="false" 
  SmartNavigation="false" Inherits="Epsom.Web.Kfi.InsuranceAndInvestment" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
  <cms:helplink id="helplink" class="helplink" runat="server" helpref="2130" />
    <h1>Insurance Detail</h1>

    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

       <h2 class="form_heading">Insurances</h2>
       <p><CMS:Literal runat="server" cmsref="140" /></p>
       <p>Get a quote now<a href="http://www.insurancequote.co.uk">www.insurancequote.co.uk</a></p>
        <table class="formtable">
          <tr>
            <td class="promptheader">Insurance type</td>
            <td class="promptheader">Yes</td>
            <td class="promptheader">No</td>
            <td class="promptheader">Company</td>
            <td class="promptheader">Monthly payment</td>
          </tr>
          <tr>
            <td class="prompt">Buildings </td>
            <td>
              <asp:radioButton 
                groupName="rblBuildingYN" 
                id="rdoBuildingYes"                
                runat="server">
              </asp:radiobutton>
            </td>
            <td>
              <asp:radiobutton 
                groupName="rblBuildingYN" 
                id="rdoBuildingNo"                
                runat="server">
              </asp:radiobutton>
            </td>
            <td><asp:TextBox id="txtBuildingsCompany" runat="server" /></td>
            <td><asp:TextBox id="txtBuildingsPayment" runat="server" /></td>
          </tr>
          <tr>
            <td class="prompt">Contents</td>
            <td>
              <asp:radiobutton 
                groupName="grpContentsYN" 
                id="rdoContentsYes"                
                runat="server">
              </asp:radiobutton>
            </td>
            <td>
              <asp:radiobutton 
                groupName="grpContentsYN"                
                id="rdoContentsNo" 
                runat="server">
              </asp:radiobutton>
            </td>
            <td><asp:TextBox id="txtContentsCompany" runat="server" /></td>
            <td><asp:TextBox id="txtContentsPayment" runat="server" /></td>
          </tr>
          <tr>
            <td class="prompt">Payment Protection</td>
            <td>
              <asp:radiobutton 
                groupName="grpPaymentProtectionYN" 
                id="rdoPaymentProtectionYes"                
                runat="server">
              </asp:radiobutton>
            </td>
            <td>              
              <asp:radiobutton 
                groupName="grpPaymentProtectionYN" 
                id="rdoPaymentProtectionNo"                
                runat="server">
              </asp:radiobutton>
            </td>
            <td><asp:TextBox id="txtPaymentProtectionCompany" runat="server" /></td>
            <td><asp:TextBox id="txtPaymentProtectionPayment" runat="server" /></td>
          </tr>
        </table>
       
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" runat="server"></epsom:dipnavigationbuttons>

  </mp:content>

</mp:contentcontainer>

