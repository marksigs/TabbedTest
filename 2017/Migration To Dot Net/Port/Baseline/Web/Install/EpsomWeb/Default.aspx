<%@ Page language="c#" Codebehind="Default.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.Default" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>
<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="~/masterpages/masterpage1.ascx">

  <mp:content id="region1" runat="server">

    <div class="span3 focus">
      <h1>Announcements</h1>
      <p><cms:cmsBoilerplate cmsref="74011" runat="server" /></p>

    </div>

    <asp:Panel id="pnlUnrestricted" Runat="server">

      <div class="span2">

        <h2><asp:label id="labelHeading" runat="server" /></h2>
        <!--p><asp:label id="labelWelcome" runat="server" /></p-->
        <p><cms:cmsBoilerplate cmsref="4012" runat="server" /></p>
        <!--img src="<%=ResolveUrl("~/_gfx/banner.jpg")%>" width="370" height="100" alt="banner"/-->
         <!--img src="<%=ResolveUrl("~/News/ArticleImage.aspx")%>" width="370" height="100" alt="banner"/-->
        <asp:image id="imageUnrestricted" runat="server" imageurl="~/WebUserControls/DisplayImage.aspx?cmsRef=5012" width="370" height="100" alt="banner">
        </asp:image> 
      </div>

      <div class="span1">

        <h2>News</h2>
        <p><strong>Product Update</strong></p>
        <!--<p><cms:cmsBoilerplate cmsref="4016" runat="server" /></p> -->
          <p><asp:label id="lblProductUpdateUnrestricted" runat="server"></asp:label>
          </p>
          <p><asp:hyperlink id="linkProductUpdateUnrestricted" class="fancylink" navigateurl="~/News/News.aspx" text="Read most recent Product Update" runat="server" />
          </p>
          <p>
          <a href="<%=ResolveUrl("~/News/News.aspx")%>" class="fancylink" title="Go to products homepage">Go to products homepage</a>
          </p>

        <p>
          <strong>New Headline</strong>
          <!--<p>
          <cms:cmsboilerplate cmsref="4017" runat="server" id="Cmsboilerplate1"/>
          </p>-->
          <p><asp:label id="lblNewsUnrestricted" runat="server"></asp:label>
          </p>
          <p><asp:hyperlink id="linkNewsUnrestricted" class="fancylink" navigateurl="~/News/News.aspx" text="Read lead story" runat="server" />
          </p>
          <p>
          <a href="<%=ResolveUrl("~/News/News.aspx")%>" class="fancylink" title="Go to news homepage">Go to news homepage</a>
          </p>
          
        </p>

      </div>

    </asp:Panel>

    <asp:Panel id="pnlRestricted" Runat="server">

      <div class="span1">

        <h2>New Case</h2>
        <p><cms:cmsBoilerplate cmsref="4018"  runat="server"/>
          <a href="<%=ResolveUrl("~/DIP/Default.aspx")%>" class="fancylink" title="Get a decision in principle">Get a DIP</a>
        </p>
        <p><cms:cmsBoilerplate cmsref="4019" runat="server" />
          <a href="<%=ResolveUrl("~/KFI/Default.aspx")%>" class="fancylink" title="Get a Key Facts Illustration">Get a KFI</a>
        </p>

        <h2>Existing Cases</h2>

        <p><cms:cmsBoilerplate cmsref="4010" runat="server" />
          <a href="<%=ResolveUrl("~/CaseTracker/Default.aspx")%>" class="fancylink" title="Case tracking">Case Tracking</a>
        </p>

      </div>

      <div class="span2">

        <h2>News</h2>
        <p>
          <strong>Product update</strong>
          <!--<p><cms:cmsboilerplate cmsref="4016" runat="server" id="Cmsboilerplate2"/></p>-->
          <p><asp:label id="lblProductUpdateRestricted" runat="server"></asp:label>
          </p>
          <p><asp:hyperlink id="linkProductUpdateRestricted" class="fancylink" navigateurl="~/News/News.aspx" text="Read most recent Product Update" runat="server" />
          </p>
          <p>
          <a href="<%=ResolveUrl("~/News/News.aspx")%>" class="fancylink" title="Go to products homepage">Go to products homepage</a>
          </p>
        </p>

        <p>
          <strong>New Headline</strong>
          <br />
          <!--<cms:cmsBoilerplate cmsref="4017" runat="server" />-->
          <p><asp:label id="lblNewsRestricted" runat="server"></asp:label>
          </p>
          <p><asp:hyperlink id="linkNewsRestricted" class="fancylink" navigateurl="~/News/News.aspx" text="Read lead story" runat="server" />
          </p>
          <p><a href="<%=ResolveUrl("~/News/News.aspx")%>" class="fancylink" title="Go to news homepage">Go to news homepage</a></p>
        </p>

        <h2>Insurance</h2>
        <p>
          <cms:cmsBoilerplate cmsref="4011" runat="server" />
<!--          <br/>
          <asp:hyperlink id="insuranceLink" runat="server" CssClass="fancylink">Sterling!</asp:hyperlink>
-->        </p>

      </div>

    </asp:Panel>

  </mp:content>

</mp:contentcontainer>

