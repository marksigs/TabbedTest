<%@ Page language="c#" Codebehind="Downloads.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Downloads" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <h1>Downloads</h1>

    <p>Download stationery and product guides</p>

    <asp:Panel id="pnlApprovedPackagersText" runat="server" CssClass="forminfo">
      To view the downloads available to approved packagers, please log in. The page will reload and you
      will see the full range of downloads.
    </asp:Panel>
   

  <asp:repeater id="rptDocumentCategories" runat="server" >
    <itemtemplate>
      <asp:panel id="pnlCategoryName" runat="server" cssclass="forminfo">
        <h2 class="form_heading"><asp:label id="lblCategoryName" runat="server" text=""></asp:label></h2>
      </asp:panel>
    
      <asp:datalist
        id="dlCategory"
        repeatcolumns="3"
        runat="server"
        repeatdirection="horizontal"
        cellspacing="3"
        itemstyle-verticalalign="top"
        repeatlayout="table">
        <itemtemplate>
          <table class="document">
            <tr>
              <td colspan="2">
                <b><asp:label id="lblDocument" runat="server" /></b>
              </td>
            </tr>
            <tr>
              <td><img src="<%=ResolveUrl("~/_gfx/pdf_logo.gif")%>" alt=""></td>
              <td>
                <asp:label id="lblDocumentDescription" runat="server" />
              </td>
            </tr>
            <tr>
              <td>&nbsp;</td>
              <td>
                <asp:linkbutton id="cmdDocument" cssclass="fancylink" runat="server" text="Download" commandname="cmdDownloadDocument" />
              </td>
            </tr>
          </table>
          <br />

        </itemtemplate>
      </asp:DataList> 

      <p class="forminfo">
        <asp:label id="lblNotFound" runat="server" text="" />
      </p>

    </itemtemplate>  
  </asp:repeater>



    
  </mp:content>

</mp:contentcontainer>
