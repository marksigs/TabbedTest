<%@ Page language="c#" Codebehind="Default.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.CaseTracking" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage4.ascx">

  <mp:content id="region2" runat="server">

    <h1>Case Tracking</h1>
  <asp:Label id="lblNoRecordsMsg"  visible="False" runat="server" /> 
    <asp:Panel id="panelOverview" runat="server">

      <p>
        Case ref or applicant name:
        &nbsp; 
        <asp:TextBox id="txtFind" maxlength="50" runat="server" />
        <asp:Button id="buttonFind" text="Find" runat="server" />
        <asp:Button id="buttonClear" text="Clear" runat="server" />
      </p>

      <p style="float:right">
        order by <asp:DropDownList id="ddlOrderBy" width="110" runat="server" /> <asp:Button id="buttonReverseOrder" text="a..z" runat="server" />
      </p>

      <p>
        Include:  <asp:CheckBox id="chkPrevious"  text="previously owned" runat="server" /> 
                  <asp:CheckBox id="chkCancelled" text="cancelled" runat="server"/> 
                  <asp:CheckBox id="chkFirm"      text="all my firms" runat="server" /> 
        <asp:Label id="pagesize" text="8"  visible="False" runat="server" /> 
        <asp:Label id="sortfield" text="date"  visible="False" runat="server" /> 
      </p>

      <p style="padding: 1em 0 1em 0">
        <asp:Label id="labelSearchResult" runat="server" />
        <asp:Label id="labelIntroducer" runat="server" />
      </p>

      <asp:Label id="labelError" runat="server" />  

      <asp:Panel id="pnlCaseList" runat="server">
      
        <table class="casetrackertable">

          <tr>
            <th>Our ref</th>
            <th>Your ref</th>
            <th>1st Applicant</th>
            <th>Current owner</th>
            <th>Original owner</th>
            <th>Applicant's postcode</th>
            <th>Submit date</th>
            <th>Stage</th>
            <th>Action</th>
          </tr> 
              
          <asp:Repeater id="repeaterCases"  runat="server">

            <itemtemplate>
          
              <tr class="item">
                <td><asp:Label id="labelReferenceNumber" runat="server" /></td>
                <td><asp:Label id="labelYourReferenceNumber" runat="server" /></td>
                <td><asp:Label id="labelApplicantName" runat="server" /></td>
                <td><asp:Label id="labelCurrentOwner" runat="server" /></td>
                <td><asp:Label id="labelOriginalOwner" runat="server" /></td>            
                <td><asp:Label id="labelPostcode" runat="server" /></td>
                <td><asp:Label id="labelSubmitDate" runat="server" /></td>
                <td><asp:Label id="labelStage" runat="server" /> <asp:Label id="labelStatus" runat="server" /></td>
                <td nowrap="true">
                  <asp:Button id="buttonAssign" runat="server" />
                  <asp:Button id="buttonDIP" runat="server" />
                  <asp:Button id="buttonKFI" runat="server" />
                  <asp:Button id="buttonAPP" runat="server" />
                  <asp:Button id="buttonView" runat="server" />
                </td>
              </tr>
                
            </itemtemplate>
          
            <alternatingitemtemplate>
          
              <tr class="alternatingitem">
                <td><asp:Label id="labelReferenceNumber" runat="server" /></td>
                <td><asp:Label id="labelYourReferenceNumber" runat="server" /></td>
                <td><asp:Label id="labelApplicantName" runat="server" /></td>
                <td><asp:Label id="labelCurrentOwner" runat="server" /></td>
                <td><asp:Label id="labelOriginalOwner" runat="server" /></td>            
                <td><asp:Label id="labelPostcode" runat="server" /></td>
                <td><asp:Label id="labelSubmitDate" runat="server" /></td>
                <td><asp:Label id="labelStage" runat="server" /> <asp:Label id="labelStatus" runat="server" /></td>
                <td nowrap="true">
                  <asp:Button id="buttonAssign" runat="server" />
                  <asp:Button id="buttonDIP" runat="server" />
                  <asp:Button id="buttonKFI" runat="server" />
                  <asp:Button id="buttonAPP" runat="server" />
                  <asp:Button id="buttonView" runat="server" />
                </td>
              </tr>
                
            </alternatingitemtemplate>

          </asp:repeater>

        </table>

        <br />

        <p class="button_orphan">
          <asp:Button id="buttonPrev" text="< Prev" runat="server" />
          &nbsp;&nbsp;&nbsp;&nbsp;
          Page &nbsp; <asp:DropDownList id="ddlPageNumber" width="40" runat="server" /> of <asp:Label id="labelPageCount" runat="server" />
          &nbsp;&nbsp;&nbsp;&nbsp;
          <asp:Button id="buttonNext" text="Next >" runat="server" />
            <!-- <asp:Label id="labelPageNumber" runat="server" /> -->
        </p>
        
      </asp:Panel>
        
    </asp:Panel>



    
    <asp:Panel id="panelDetail" runat="server">

      <p class="dipnavigationbuttons">
        <asp:Button id="buttonBack2Overview" runat="server" text="Back to case list" />
      </p>
    
      <h2 class="form_heading">Selected case details</h2>

      <p style="padding: 1em 0">
        Our reference: <b><asp:Label id="labelReferenceNumberDetail" runat="server" /></b>
      </p>
    
      <table id="tblCaseList" class="casetrackertable">
      
        <tr>
          <th>Your Reference Number</th>
          <th>Submit Date</th>
          <th>1st Applicant</th>
          <th>Property postcode</th>
          <th>Current Owner</th>
          <th>Stage</th>
          <th>Task</th>
          <th>Status Date</th>
          <th>Status</th>    
        </tr>
        
        <tr valign="top" style="background-color:#EEE">
          <td rowspan="99"><asp:Label id="labelYourReferenceNumberDetail" runat="server" /></td>
          <td rowspan="99"><asp:Label id="labelSubmitDateDetail" runat="server" /></td>
          <td rowspan="99"><asp:Label id="labelNameDetail" runat="server" /></td>
          <td rowspan="99"><asp:Label id="labelPostCodeDetail" runat="server" /></td>
          <td rowspan="99"><asp:Label id="labelCurrentOwnerDetail" runat="server" /></td>
        </tr>
        
        <asp:Repeater id="repeaterTasks" runat="server">
          <itemtemplate>
            <tr>
              <td><asp:Label id="labelStageDetail" runat="server" /></td>
              <td><asp:Label id="labelTaskDetail" runat="server" /></td>
              <td><asp:Label id="labelStatusDateDetail" runat="server" /></td>
              <td><asp:Label id="labelStatusDetail" runat="server" /></td>
            </tr>
          </itemtemplate>
        </asp:Repeater>
      
      </table>
      
    </asp:Panel>

  </mp:content>

</mp:contentcontainer>