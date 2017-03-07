<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Page language="c#" Codebehind="SubmissionChecklist.aspx.cs" AutoEventWireup="False" Inherits="Epsom.Web.App.SubmissionChecklist" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

<mp:content id=region1 runat="server">
  <script language="javascript">
      function ProcessSubmit()
      {
        var mainLayer = document.getElementById('mainLayer');
        var holdingLayer = document.getElementById('holdingLayer');
        mainLayer.style.display = 'none';
        holdingLayer.style.display = 'block';
        setTimeout('document.getElementById("waitImage").src = "<%=ResolveUrl("~/_gfx/evaluating.gif")%>";', 50);
        scroll(0,0);
      }
    </script>
 
  <epsom:dipmenu id=DipMenu runat="server"></epsom:dipmenu>
 </mp:content>
 
<mp:content id=region2 runat="server">

  <h1>Submission Checklist</h1>
  <div id="mainLayer">
  <epsom:dipnavigationbuttons id=ctlDIPNavigationButtons1 runat="server"></epsom:dipnavigationbuttons>
  
  <asp:validationsummary id=Validationsummary1 runat="server" displaymode="BulletList" headertext="Input errors:" CssClass="validationsummary" showsummary="true"></asp:validationsummary>

  <p>
    <asp:label id="lblMessage" Runat="server"></asp:label>
  </p>

  
  <h2 class=form_heading>Application data completion status</h2>

  <table>
  
  <asp:Panel id="pnlAdvisor" Runat="server">
    <tr>
      <td class="prompt" >
        Advisor Details
      </td>
      <td>
        <asp:Image id=imgAdvisorDetails Runat="server"></asp:Image>
        <asp:LinkButton id=cmdAdvisorDetails Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
  </asp:Panel>
    <tr>
      <td class="prompt">
        Regulation Info
      </td>
      <td>
        <asp:Image id=imgRegulationInfo Runat="server"></asp:Image>
        <asp:LinkButton id=cmdRegulationInfo Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Production Selection
      </td>
      <td>
        <asp:Image id=imgProductSelection Runat="server"></asp:Image>
        <asp:LinkButton id=cmdProductSelection Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Personal Details
      </td>
      <td>
        <asp:Repeater id="rptPersonalDetails" runat="server">
          <itemtemplate>
            <asp:Image id="imgPersonalDetails" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditPersonalDetails" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Employment Details
      </td>
      <td>
        <asp:Repeater id="rptEmploymentDetails" runat="server">
          <itemtemplate>
            <asp:Image id="imgEmploymentDetails" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditEmploymentDetails" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    <tr>
      <td class="prompt">
        Income Details
      <td>
        <asp:Repeater id="rptIncomeDetails" runat="server">
          <itemtemplate>
            <asp:Image id="imgIncomeDetails" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditIncomeDetails" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    <tr>
      <td class="prompt">
        Lender Landlord Details
      </td>
      <td>
        <asp:Repeater id="rptLenderLandlordDetails" runat="server">
          <itemtemplate>
            <asp:Image id="imgLenderLandlordDetails" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditLenderLandlordDetails" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Payment History
      </td>
      <td>
        <asp:Repeater id="rptPaymentHistory" runat="server">
          <itemtemplate>
            <asp:Image id="imgPaymentHistory" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditPaymentHistory" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Commitments
      </td>
      <td>
        <asp:Repeater id="rptCommitments" runat="server">
          <itemtemplate>
            <asp:Image id="imgCommitments" Runat="server"></asp:Image>
            <asp:LinkButton id="cmdEditCommitments" runat="server"></asp:LinkButton>
          </itemtemplate>
        </asp:Repeater>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Property Details
      </td>
      <td>
        <asp:Image id=imgPropertyDetails Runat="server"></asp:Image>
        <asp:LinkButton id=cmdPropertyDetails Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    
    <asp:Panel id="pnlSolicitor" Runat="server">
    <tr>
      <td class="prompt">
        Solicitor Details
      </td>
      <td>
        <asp:Image id=imgSolicitorDetails Runat="server"></asp:Image>
        <asp:LinkButton id=cmdSolicitorDetails Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    </asp:Panel>
    
    <tr>
      <td class="prompt">
        Fees
      </td>
      <td>
        <asp:Image id=imgFee Runat="server"></asp:Image>
        <asp:LinkButton id=cmdFee Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Declaration
      </td>
      <td>
        <asp:Image id=imgDeclaration Runat="server"></asp:Image>
        <asp:LinkButton id=cmdDeclaration Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
    
    <tr>
      <td class="prompt">
        Direct Debit Details
      </td>
      <td>
        <asp:Image id=imgDirectDebitDetails Runat="server"></asp:Image>
        <asp:LinkButton id=cmdDirectDebitDetails Runat="server" text="Edit item"></asp:LinkButton>
      </td>
    </tr>
  </table>
  </div>
  <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server" />
  
  <div id="holdingLayer" style="display:none">
      <p>
        Please wait while our system automatically saves all your application data - this may take up to a minute at busy
        times. Please do not navigate away from this page.
      </p>
      <img id="waitImage" src="<%=ResolveUrl("~/_gfx/evaluating.gif")%>" width="386" height="22" alt="Evaluating mortgage decision" />
    </div>

</mp:content>
</mp:contentcontainer>
