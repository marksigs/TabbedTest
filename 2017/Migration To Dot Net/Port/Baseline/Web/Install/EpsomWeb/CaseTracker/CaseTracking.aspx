<%@ Page language="c#" Codebehind="CaseTracking.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.CaseTracking" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">
 
  <asp:Panel id="panelOverview" runat="server">

  <h1>Case Tracking</h1>
  
    <table class="displaytable" cellpadding="5">
      <tr><td colspan="9"><h2>Case Overview</h2></td></tr>
      <tr>
        <td colspan="9">
          Case Ref. or Applicant Name: &nbsp; 
          <asp:TextBox id="txtFind" runat="server" />  &nbsp; 
          <asp:Button  id="buttonFind" text="Find" runat="server" />
          <asp:Button  id="buttonClear" text="Clear" runat="server" /> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          <font color="gray"><asp:Label   id="labelSearchResult" runat="server" /></font>
        </td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td>&nbsp;</td></tr>
      <tr>
        <th>Reference Number</th>
        <th>Applicant Name</th>
        <th>Post Code</th>
        <th>Current Owner</th>
        <th>Submission Date</th>
        <th>Stage</th>
        <th>Status</th>
        
        <th colspan="3" align="center"> Action  </th>
        
      </tr> 
         
      <asp:Repeater id="repeaterCases"  runat="server">
        <itemtemplate>
      
          <tr valign="middle" style="background-color:lightcyan">
            <td><asp:Label id="labelReferenceNumber" runat="server" /></td>
            <td><asp:Label id="labelApplicantName" runat="server" /></td>
            <td><asp:Label id="labelPostcode" runat="server" /></td>
            <td><asp:Label id="labelCurrentOwner" runat="server" /></td>
            <td><asp:Label id="labelSubmissionDate" runat="server" /></td>
            <td><asp:Label id="labelStage" runat="server" /></td>
            <td><asp:Label id="labelStatus" runat="server" /></td>
            
            <td><asp:Button id="buttonDIP" runat="server" /></td>
            <td><asp:Button id="buttonKFI" runat="server" /></td>
            <td><asp:Button id="buttonApply" runat="server" /> </td>
            <td><asp:Button id="buttonView" runat="server" /></td>
           </tr>
           
        </itemtemplate>
      
        <alternatingitemtemplate>
      
          <tr  valign="middle" style="background-color:ivory">
            <td><asp:Label id="labelReferenceNumber" runat="server" /></td>
            <td><asp:Label id="labelApplicantName" runat="server" /></td>
            <td><asp:Label id="labelPostcode" runat="server" /></td>
            <td><asp:Label id="labelCurrentOwner" runat="server" /></td>
            <td><asp:Label id="labelSubmissionDate" runat="server" /></td>
            <td><asp:Label id="labelStage" runat="server" /></td>
            <td><asp:Label id="labelStatus" runat="server" /></td>
            
            <td><asp:Button id="buttonDIP" runat="server" /></td>
            <td><asp:Button id="buttonKFI" runat="server" /></td>
            <td><asp:Button id="buttonApply" runat="server" /> </td>
            <td><asp:Button id="buttonView" runat="server" /></td>
           </tr>
           
        </alternatingitemtemplate>
      
      </asp:repeater>
    
      <tr valign="bottom">
        <td colspan="11" align="center">
          <asp:HyperLink id="linkPrev" runat="server" />&nbsp;&nbsp;
          <asp:Label id="labelPageNumber" runat="server"/>&nbsp;&nbsp;
          <asp:HyperLink id="linkNext" runat="server" />
        </td>
      </tr>
    
    </table>
    
  </asp:Panel>
  
  
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

  
  
  <asp:Panel id="panelDetail" runat="server">
  
  <table class="displaytable" border="0">
  
    <tr><td align="left"><h2>Case Status</h2></td><td align="right" colspan="8"><asp:Button id="buttonBack2Overview" runat="server" text="Hide Status"></asp:Button></td></tr>
  
    <tr>
      <th>Reference Number</th>
      <th>Submission Date</th>
      <th>Name</th>
      <th>Post Code</th>
      <th>Current Owner</th>
      <th>Stage</th>
      <th>Task</th>
      <th>Status Date</th>
      <th>Status</th>    
    </tr>
    
    <tr valign="top" style="background-color:lightcyan">
    
    <td rowspan="99"><asp:Label id=labelReferenceNumberDetail runat="server" /> </td>
    <td rowspan="99"><asp:Label id=labelSubmissionDateDetail runat="server" /></td>
    <td rowspan="99"><asp:Label id=labelNameDetail runat="server" /></td>
    <td rowspan="99"><asp:Label id=labelPostCodeDetail runat="server" /></td>
    <td rowspan="99"><asp:Label id=labelCurrentOwnerDetail runat="server" /></td>
    
    </tr>
    
    
    <asp:Repeater id="repeaterTasks" runat="server">
    <itemtemplate>
    
    <tr>
    <td><asp:Label id="labelStageDetail" runat="server" /></td>
    <td><asp:Label id="labelTaskDetail" runat="server" /></td>
    <td><asp:Label id="labelDueDateDetail" runat="server" /></td>
    <td><asp:Label id="labelStatusDetail" runat="server" /></td>
    </tr>
    
    </itemtemplate>
    </asp:Repeater>
    
  
  </table>
  </asp:Panel>

  </mp:content>

</mp:contentcontainer>