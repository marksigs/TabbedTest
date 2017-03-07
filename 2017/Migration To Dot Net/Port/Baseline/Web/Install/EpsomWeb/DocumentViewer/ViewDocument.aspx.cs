using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;

using System.IO;
using System.Net;
using System.Text;

using Epsom.Web.Helpers;

namespace Epsom.Web.DocumentViewer
{
  /// <summary>
  /// Summary description for ViewDocument.
  /// </summary>
  public class ViewDocument : System.Web.UI.Page
  {
    protected System.Web.UI.WebControls.Literal DisplayHTML;
    protected System.Web.UI.WebControls.Label DebugLabel;
    private string envt = "", DocServletURL = "", ErrServletURL = "", debugFlag = "", documentForm = "";
    private string adobeURL = "", adobeGIF = "";
	

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

    }
    #endregion

    private void Page_Load(object sender, System.EventArgs e)
    {
      String userID = "", userRef = "", errMsg = "";
      userID = Request["userID"];
      userRef = Request["ref"];
      getSettings();

      if (userID == null) userID = "";
      if (userRef == null) userRef = "";

      if ((userID.Trim().Equals("")) || (userRef.Trim().Equals("")))
      {
        // Incorrect parameter names or insufficient parameter supplied
        errMsg = "Insufficient or invalid parameters supplied";
      }
      else
      {
        // Correct parameter names supplied, perform validation

        // This line is very important because posted values can cause + chars to be replaced
        // with spaces which messes up the string to for decryption so this is why space chars
        // are replaced with a + char
        userRef = userRef.Replace(" ", "+");

//        DevtEncryptionWS.encryptionService encObj = new DocumentViewer.DevtEncryptionWS.encryptionService();
//        DevtEncryptionWS.encryptionResponse encResp = null;

//        encResp = encObj.Decrypt(userRef);

        string decryptedRef = WebServiceHelper.Instance.Decrypt(userRef);
        
        if (decryptedRef != null)
        {
//          string decryptedRef = encResp.result;
          string delim = "|";
          string[] refParts = decryptedRef.Split(delim.ToCharArray());
          if (refParts.Length == 6)
          {
            // Split decrypted string into it's component parts
            string decryptedEmailID = "", appNo = "", reportID = "", reportVersion = "", reportSection = "", reportPage = "";

            decryptedEmailID = refParts[0].ToUpper();
            appNo = refParts[1];
            reportID = refParts[2];
            reportVersion = refParts[3];
            reportSection = refParts[4];
            reportPage = refParts[5];

            if (decryptedEmailID.Equals(userID.ToUpper()))
            {
              string documentLocation = appNo + "|" + reportID + "|" + reportVersion + "|" + reportSection + "|" + reportPage;

              StringBuilder sb = new StringBuilder();

              Boolean PDFok = false;
              try 
              {
                // Call the servlet to with P_CHECK parameter to check if the file can be retrieved
                sb.Append(DocServletURL + "?P_LOCATION=" + documentLocation);
                if (envt.Equals("DEVT"))
                  sb.Append("&P_ENVT=DEVT");
                sb.Append("&P_CHECK=Y");

                System.Net.WebResponse resp = null;
                System.Net.WebRequest req = null;

                req = System.Net.HttpWebRequest.Create(sb.ToString());
                resp = req.GetResponse();

                StreamReader sr = new StreamReader(resp.GetResponseStream());
                String servletResponse = sr.ReadToEnd();
                sr.Close();

                if (servletResponse.StartsWith("PASS")) 
                {
                  PDFok = true;
                } 
                else 
                {
                  errMsg = servletResponse.Substring(5);
                  // Replace pipe symbols with line feeds
                  errMsg = errMsg.Replace("|", "<br/>");
                }							

              } 
              catch (Exception ex) 
              {
                PDFok = false;
              }

              if (PDFok) 
              {
                // build the HTML form which will be used to post values to the ViewerGetFileServlet
                sb = new StringBuilder();
                sb.Append("<form id=\"frm\" method=\"post\" target=\"_blank\" action=\"" + DocServletURL + "\">");
                sb.Append("<input type=\"hidden\" name=\"P_LOCATION\" value=\"" + documentLocation + "\">");
                if (envt.Equals("DEVT"))
                  sb.Append("<input type=\"hidden\" name=\"P_ENVT\" value=\"DEVT\">");
                sb.Append("<input type=\"Submit\" value=\"View Document\">");
                sb.Append("</form>");

                documentForm = sb.ToString();
              }
            }
            else
            {
              // Decrypted URL email address does not match the plaintext one
              // so the URL may have been tampered with
              errMsg = "User ID does not match value in encrypted part of URL";
            }

          }
        }
        else
        {
//          errMsg = "Failed to decrypt URL : " + encResp.result + "<br>Attempting to decrypt<hr>" + userRef + "<hr>";
          errMsg = "Failed to decrypt URL : " + decryptedRef + "<br />Attempting to decrypt<hr />" + userRef + "<hr />";
          errMsg += userRef.Replace(" ", "+");
        }
      }

      buildHTML(errMsg);
      if (!errMsg.Equals(""))
        logError(userID, userRef, errMsg);
    }

    private void getSettings()
    {
      // Retrieve servlet URL from config settings in web.config
      envt = System.Configuration.ConfigurationSettings.AppSettings["ENVT"];
      debugFlag = System.Configuration.ConfigurationSettings.AppSettings["DEBUG"];
      adobeURL = System.Configuration.ConfigurationSettings.AppSettings["AdobeReaderURL"];
      adobeGIF = System.Configuration.ConfigurationSettings.AppSettings["AdobeGIF"];

      switch (envt.ToUpper())
      {
        case "PROD":
          DocServletURL = System.Configuration.ConfigurationSettings.AppSettings["DocServletURLprod"];
          ErrServletURL = System.Configuration.ConfigurationSettings.AppSettings["ErrServletURLprod"];
          break;
        default:
          // Default to devt
          envt = "DEVT";
          DocServletURL = System.Configuration.ConfigurationSettings.AppSettings["DocServletURLdevt"];
          ErrServletURL = System.Configuration.ConfigurationSettings.AppSettings["ErrServletURLdevt"];
          break;
      }
      this.DebugLabel.Visible = false;
      if (debugFlag.ToUpper().Equals("ON"))
        this.DebugLabel.Visible = true;
    }


    private void buildHTML(string errMsg)
    {
      if (errMsg.Equals(""))
      {
        StringBuilder sb = new StringBuilder();
        sb.Append("To view the document you will need Adobe Acrobat Reader installed<br /><br />");
        sb.Append("<a href=\"" + adobeURL + "\" target=\"_blank\"><img src=\"" + adobeGIF + "\" border=\"0\"></a>");
        sb.Append(documentForm);
        this.DisplayHTML.Text = sb.ToString();
      }
      else
      {
        // Build page showing error message
        this.DisplayHTML.Text = "<p class=\"error\">" + errMsg + "</p>";
      }
    }

    private void logError(string userID, string userRef, string errMsg)
    {
      // Call the error logging servlet to log the error to the vdrData database on uug-res-ap165
      WebResponse resp = null;
      StreamReader sr = null;
      try
      {
        WebRequest req = null;
        ErrServletURL += "?P_ENVT=" + envt + "&P_EMAIL=" + userID + "&P_URL=" + userRef + "&P_ERRORDESC=" + errMsg;
        req = HttpWebRequest.Create(ErrServletURL);
        resp = req.GetResponse();
        sr = new StreamReader(resp.GetResponseStream());
        string ret = sr.ReadToEnd();
        this.DebugLabel.Text = ret;
      }
      catch (Exception ex)
      {
        DebugLabel.Text = ex.Message;
      }
      finally
      {
        try
        {
          if (sr != null)	sr.Close();					
        }
        catch (Exception ex1) { }

        try
        {
          if (resp != null) resp.Close();
        }
        catch (Exception ex1) { }
      }
    }
  }
}
