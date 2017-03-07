using System;
using System.Collections;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using Epsom.Web.Helpers;

namespace Epsom.Web
{
	/// <summary>
	/// Summary description for Overview.
	/// </summary>
	public class CaseTracking : System.Web.UI.Page
	{
    //------------------------------------ Overview Panel Elements
    private const int PAGESIZE = 6; //LANCE - should set this in html

    protected Panel panelOverview;    
    protected Repeater repeaterCases;

    protected Button buttonFind;
    protected Button buttonClear;

    private Label labelReferenceNumber;
    private Label labelYourReferenceNumber;
    private Label labelApplicantName;
    private Label labelPostcode;
    private Label labelCurrentOwner;
    private Label labelSubmissionDate;
    private Label labelStage;
    private Label labelStatus;

    private Label labelPageNumber;

    private Button buttonDIP;
    private Button buttonKFI;
    private Button buttonApply;
    private Button buttonView;

    private HyperLink linkPrev;
    private HyperLink linkNext;

    //------------------------------------- Detail Panel Elements
    protected Panel panelDetail;
    protected Repeater repeaterTasks;

    protected Button buttonBack2Overview;

    private Label labelReferenceNumberDetail;
    private Label labelYourReferenceNumberDetail;
    private Label labelSubmissionDateDetail;
    private Label labelNameDetail;
    private Label labelPostCodeDetail;
    private Label labelCurrentOwnerDetail;
    private Label labelStageDetail;
    private Label labelTaskDetail;
    private Label labelDueDateDetail;
    private Label labelStatusDetail;

    //------------------------------------------------------------


    private int currentPageNumber;
    private int case2ViewIndex = -1;
    private string searchText = "";
    private string errorMessage = "";
    protected ArrayList Cases =  new ArrayList();

    private int currentFirmIndex;
    private Epsom.Web.Proxy.User user;
    private Epsom.Web.Proxy.Firm firm;

    private void Page_Load(object sender, System.EventArgs e)
    {
      //Epsom.Web.Proxy.User user = new Epsom.Web.Proxy.User();    //Dummy Code - User object will come from session
      //user.UserId = "REN";                                       //Dummy Code
      //user.IntroducerId = "0";                                   //Dummy Code
      //user.Firms = new Epsom.Web.Proxy.Firm[1];                  //Dummy Code
      //user.Firms[0] = new Epsom.Web.Proxy.Firm();                //Dummy Code
      //user.Firms[0].UnitId = "ADMIN";                           //Dummy Code 
      //HttpSession.CurrentFirmIndex = 0;                          //Dummy Code

      currentFirmIndex = HttpSession.CurrentFirmIndex;
      user = HttpSession.CurrentUser;
      firm = user.Firms[currentFirmIndex];

      if (user == null) { Response.Redirect("~/"); }

      if (this.ViewState["AllIntroducerCases"] == null)
      {
        Epsom.Web.Proxy.IntroducerPipelineResponse ipr = Epsom.Web.Helpers.WebServiceHelper.Instance.GetCasesForIntroducer(user,currentFirmIndex);
        if (ipr.ResponseMessages.Length > 0)
        { this.errorMessage = ipr.ResponseMessages[0].Text; }
        
        if (ipr.Cases == null)
        { this.ViewState["AllIntroducerCases"] = new Epsom.Web.Proxy.Case[0]; }
        else
        { this.ViewState["AllIntroducerCases"] = ipr.Cases;}

      }


      if (!IsPostBack)
      {
        currentPageNumber = Request.QueryString["page"] == null ? 1 : Convert.ToInt32(Request.QueryString["page"]);
        searchText = Request.QueryString["search"] == null ? "" : Request.QueryString["search"];

        SetVisibility();
      }


    }


    private void SetVisibility()
    {
      FilterCases();
      DisplayCaseOverview();

      this.panelOverview.Visible = true;

      SetDetailVisibility();
    }

    private void SetDetailVisibility()
    {
      if (this.case2ViewIndex > -1)
      {
        DisplayCaseDetail();
        panelDetail.Visible = true;
      }

      else { panelDetail.Visible = false; }
    }

    private string SearchQuery()
    {
      return this.searchText == String.Empty ? "" :  "&search=" + (this.searchText);
    }



    private void FilterCases()
    {
      Epsom.Web.Proxy.Case[] AllCases = (Epsom.Web.Proxy.Case[]) this.ViewState["AllIntroducerCases"];

//      /////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//      
//            for (int i=0; i < AllCases.Length; i++)   //Dummy Code
//            { 
//              AllCases[i].Id = i;                   //Dummy Code
//              AllCases[i].ApplicantName = "";         //Dummy Code
//              AllCases[i].ReferenceNumber = "";         //Dummy Code
//            } 
//            AllCases[0].ApplicantName = "Bill Twopence";       //Dummy Code
//            AllCases[1].ApplicantName = "bertie Billison";  //Dummy Code
//            AllCases[2].ApplicantName = "Two Pence";        //Dummy Code
//            AllCases[3].ApplicantName = "Penny GILBERT";    //Dummy Code
//            AllCases[4].ApplicantName = "Allington Cooper";     //Dummy Code
//      
//            AllCases[0].ReferenceNumber = "ryebn123he890";    //Dummy Code
//            AllCases[1].ReferenceNumber = "123uebduhe8";    //Dummy Code
//            AllCases[2].ReferenceNumber = "89eyndyjwydeybn1";   //Dummy Code
//            AllCases[3].ReferenceNumber = "bd9eyndyjw";         //Dummy Code
//            AllCases[4].ReferenceNumber = "Biebn1llyjwTwoeybn1pence";     //Dummy Code
//      /////////////////////////////////////////////////////////////////////////////////////////////////////////


      for (int i=0; i < AllCases.Length; i++)
      {
        if (AllCases[i].ReferenceNumber.ToLower().IndexOf(this.searchText) != -1) { Cases.Add(AllCases[i]); }
        else if (AllCases[i].ApplicantName.ToLower().IndexOf(this.searchText) != -1) { Cases.Add(AllCases[i]); }
      }

    }



    protected void RepeaterCases_ItemCommand(object sender, RepeaterCommandEventArgs e)
    {
      this.case2ViewIndex = Convert.ToInt32(e.CommandArgument);
      Epsom.Web.Proxy.Case[] AllCases = (Epsom.Web.Proxy.Case[]) this.ViewState["AllIntroducerCases"];

      if (e.CommandName == "view") { SetDetailVisibility(); }

      else
      {
        string destinationPage = "~/Default.aspx"; // logical default
        Epsom.Web.Proxy.Case case2View = AllCases[this.case2ViewIndex];

        Epsom.Web.Proxy.ApplicationResponse ar = Epsom.Web.Helpers.WebServiceHelper.Instance.GetApplicationData(case2View.ReferenceNumber,HttpSession.CurrentUser,HttpSession.CurrentFirmIndex);
        Epsom.Web.Proxy.Application app = ar.Application;
//        Epsom.Web.Proxy.Application app = Epsom.Web.Helpers.WebServiceHelper.Instance.NewApplication();
//        //above eventually to come from WS 15
//        app.Applicants[0].FirstName = "harald";                                   //Dummy Code
//        app.Applicants[0].LastName = "hardrada" + " " + e.CommandArgument;        //Dummy Code
//        app.Applicants[0].SecondName = " " + e.CommandArgument;                   //Dummy Code

        Epsom.Web.Helpers.HttpSession.CurrentApplication = app;

        if (e.CommandName == "apply") 
        { destinationPage = "~/APP/Default.aspx"; }
        else if (e.CommandName == "kfi")
        { destinationPage = "~/KFI/Default.aspx"; }
        else
        { destinationPage = "~/DIP/Default.aspx"; } 
        
        Response.Redirect(destinationPage);
      }

    }



    private void DisplayCaseOverview()
    {
      int case2ViewIndex = 0;
      Control repeaterRow;
      PagedDataSource pager = new PagedDataSource();

      pager.AllowPaging = true;
      pager.PageSize = PAGESIZE;
      pager.CurrentPageIndex = currentPageNumber - 1;
      pager.DataSource = Cases;

      this.repeaterCases.DataSource = pager;
      this.repeaterCases.DataBind();

      if (this.errorMessage == String.Empty)
        {
          ((TextBox) panelOverview.FindControl("txtFind")).Text = this.searchText;
          ((Label) panelOverview.FindControl("labelSearchResult")).Text = this.Cases.Count.ToString() + " cases retrieved for ";
          ((Label) panelOverview.FindControl("labelIntroducer")).Text = user.FirstName + " " + user.LastName + " of " + firm.CompanyName;
        }
      else
        { ((Label) panelOverview.FindControl("labelError")).Text = this.errorMessage; }

      for (int i=0; i < pager.Count; i++)   // now build rows
      {
        repeaterRow = repeaterCases.Controls[i];
        case2ViewIndex = pager.FirstIndexInPage + i;
        
        DisplayOverviewRow(repeaterRow, case2ViewIndex); 
      }
        
      linkPrev = (HyperLink) panelOverview.FindControl("linkPrev");
      linkNext = (HyperLink) panelOverview.FindControl("linkNext");
      labelPageNumber = (Label) panelOverview.FindControl("labelPageNumber");

      string searchquery = this.searchText == String.Empty ? "" :  "&search=" + (this.searchText);
        
      linkPrev.NavigateUrl = Request.CurrentExecutionFilePath + "?page=" + (currentPageNumber - 1) + SearchQuery();
      linkPrev.Text = "<< PREV";
      linkPrev.Enabled = !pager.IsFirstPage;

      linkNext.NavigateUrl = Request.CurrentExecutionFilePath + "?page=" + (currentPageNumber + 1) + SearchQuery();
      linkNext.Text = "NEXT >>";
      linkNext.Enabled = !pager.IsLastPage;

      labelPageNumber.Text = "Page " + currentPageNumber + " of " + pager.PageCount;

    }

    private void DisplayOverviewRow(Control repeaterRow, int case2ViewIndex)
    {
      labelReferenceNumber = (Label) repeaterRow.FindControl("labelReferenceNumber");
      labelYourReferenceNumber = (Label) repeaterRow.FindControl("labelYourReferenceNumber");
      labelApplicantName = (Label) repeaterRow.FindControl("labelApplicantName");
      labelPostcode = (Label) repeaterRow.FindControl("labelPostcode");
      labelCurrentOwner = (Label) repeaterRow.FindControl("labelCurrentOwner");
      labelSubmissionDate = (Label) repeaterRow.FindControl("labelSubmissionDate");
      labelStage = (Label) repeaterRow.FindControl("labelStage");
      labelStatus = (Label) repeaterRow.FindControl("labelStatus");

      buttonDIP = (Button) repeaterRow.FindControl("buttonDIP");
      buttonKFI = (Button) repeaterRow.FindControl("buttonKFI");
      buttonApply = (Button) repeaterRow.FindControl("buttonApply");
      buttonView = (Button) repeaterRow.FindControl("buttonView");

      Epsom.Web.Proxy.Case case2View = (Epsom.Web.Proxy.Case) Cases[case2ViewIndex];
          
      labelReferenceNumber.Text = case2View.ReferenceNumber;
      labelYourReferenceNumber.Text = case2View.YourReferenceNumber;
      labelApplicantName.Text  = case2View.ApplicantName;
      labelPostcode.Text = case2View.PropertyPostCode;
      labelCurrentOwner.Text = case2View.CurrentOwner;
      labelSubmissionDate.Text = case2View.SubmissionDate.ToShortDateString();
      labelStage.Text = case2View.Stage;
      //labelStatus.Text = case2View.Status;
      Epsom.Web.Proxy.ListItem comboItem = Epsom.Web.Helpers.ComboValues.ComboItem(Epsom.Web.Helpers.ComboValues.ComboGroupType.UnderwritersDecision, case2View.Status);
      labelStatus.Text = comboItem == null ? "" : comboItem.Text;

      buttonDIP.CommandName = "dip";
      buttonKFI.CommandName = "kfi";
      buttonApply.CommandName = "apply";
      buttonView.CommandName = "view";

      buttonDIP.CommandArgument = case2ViewIndex.ToString();
      buttonKFI.CommandArgument = case2ViewIndex.ToString();
      buttonApply.CommandArgument = case2ViewIndex.ToString(); //LANCE - perhaps use ref no. instead
      buttonView.CommandArgument = case2ViewIndex.ToString();  // view button returns index of element it corresponds to

      buttonDIP.Text = "DIP";
      buttonKFI.Text = "KFI";
      buttonApply.Text = "Apply";
      buttonView.Text = "View";
    }


    private void DisplayCaseDetail()
    {
      Epsom.Web.Proxy.Case[] AllCases = (Epsom.Web.Proxy.Case[]) this.ViewState["AllIntroducerCases"];
      Epsom.Web.Proxy.Case case2View = AllCases[this.case2ViewIndex];
      Epsom.Web.Proxy.CaseResponse caseResponse = Epsom.Web.Helpers.WebServiceHelper.Instance.GetCaseTrackingData("1000473"); //will be case2View.ReferenceNumber

      case2View.CaseTasks = caseResponse.Case.CaseTasks;

      labelReferenceNumberDetail = (Label) panelDetail.FindControl("labelReferenceNumberDetail");
      labelYourReferenceNumberDetail = (Label) panelDetail.FindControl("labelYourReferenceNumberDetail");
      labelSubmissionDateDetail = (Label) panelDetail.FindControl("labelSubmissionDateDetail");
      labelNameDetail = (Label) panelDetail.FindControl("labelNameDetail");
      labelPostCodeDetail = (Label) panelDetail.FindControl("labelPostcodeDetail");
      labelCurrentOwnerDetail = (Label) panelDetail.FindControl("labelCurrentOwnerDetail");     
        
      labelReferenceNumberDetail.Text = case2View.ReferenceNumber;
      labelYourReferenceNumberDetail.Text = case2View.YourReferenceNumber;
      labelSubmissionDateDetail.Text = case2View.SubmissionDate.ToShortDateString();
      labelNameDetail.Text = case2View.ApplicantName;
      labelPostCodeDetail.Text = case2View.PropertyPostCode;
      labelCurrentOwnerDetail.Text = case2View.CurrentOwner;

      repeaterTasks.DataSource = case2View.CaseTasks;
      repeaterTasks.DataBind();

      for (int i=0; i < case2View.CaseTasks.Length; i++)
      {
        labelStageDetail = (Label) repeaterTasks.Controls[i].FindControl("labelStageDetail");
        labelTaskDetail = (Label) repeaterTasks.Controls[i].FindControl("labelTaskDetail");
        labelDueDateDetail = (Label) repeaterTasks.Controls[i].FindControl("labelDueDateDetail");
        labelStatusDetail = (Label) repeaterTasks.Controls[i].FindControl("labelStatusDetail");
        
        labelStageDetail.Text = case2View.CaseTasks[i].Stage;
        labelTaskDetail.Text = case2View.CaseTasks[i].Task;
        labelDueDateDetail.Text = case2View.CaseTasks[i].StatusDate.ToShortDateString();
        labelStatusDetail.Text = case2View.CaseTasks[i].Status;
      }

    }



    private void ButtonFind_Click(object sender, System.EventArgs e)
    {
      this.searchText = ((TextBox) panelOverview.FindControl("txtFind")).Text;
      this.currentPageNumber = 1;

      SetVisibility();
    }

    private void ButtonClear_Click(object sender, System.EventArgs e)
    {
      this.searchText = "";
      this.currentPageNumber = 1;
      
      SetVisibility();
    }


    private void buttonBack2Overview_Click(object sender, System.EventArgs e)
    {
      this.case2ViewIndex = -1;
      SetDetailVisibility();
    }






		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);

      this.repeaterCases.ItemCommand += new RepeaterCommandEventHandler(this.RepeaterCases_ItemCommand);
      this.buttonBack2Overview.Click += new System.EventHandler(this.buttonBack2Overview_Click);
      this.buttonFind.Click += new System.EventHandler(this.ButtonFind_Click);
      this.buttonClear.Click += new System.EventHandler(this.ButtonClear_Click);
     
    }
		#endregion
	}
}
