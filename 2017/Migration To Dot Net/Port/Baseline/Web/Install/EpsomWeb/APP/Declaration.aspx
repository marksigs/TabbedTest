<%@ Register Src="~/WebUserControls/DIPMenu.ascx" TagName="DIPMenu" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/DIPNavigationButtons.ascx" TagName="DIPNavigationButtons" TagPrefix="Epsom" %>
<%@ Register Src="~/WebUserControls/ApplicantType.ascx" TagName="applicantType" TagPrefix="Epsom" %>
<%@ Register TagPrefix="mp" namespace="Microsoft.Web.Samples.MasterPages" assembly="MasterPages" %>
<%@ Register TagPrefix="validators" NameSpace="Epsom.Web.Validators" Assembly="Epsom.Web"%>
<%@ Page language="c#" Codebehind="Declaration.aspx.cs" AutoEventWireup="false" Inherits="Epsom.Web.App.Declaration" %>
<%@ Register TagPrefix="CMS" namespace="Epsom.Web.Helpers" assembly="Epsom.Web" %>

<mp:contentcontainer id="Contentcontainer1" runat="server" masterpagefile="../masterpages/masterpage2.ascx">

  <mp:content id="region1" runat="server">

    <epsom:dipmenu id="DipMenu" runat="server"></epsom:dipmenu>

  </mp:content>

  <mp:content id="region2" runat="server">
	<cms:helplink id="helplink" class="helplink" runat="server" helpref="2210" />
    <h1><cms:cmsBoilerplate cmsref="621" runat="server" /></h1>
 
    <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons1" runat="server"></epsom:dipnavigationbuttons>

    <asp:validationsummary id="Validationsummary1" runat="server" showsummary="true" CssClass="validationsummary"
      headertext="Input errors:" displaymode="BulletList"></asp:validationsummary>

    <div class="notprinted">
  	  <cms:cmsBoilerplate cmsref="210" runat="server" />
	  </div>
	
    <table class="formtable">
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>
        <div enabled="true" class="applicationdeclaration"> 
          Please note that in this declaration:

          <ul>

          <li>references to &quot;loan&quot;, &quot;mortgage&quot; and &quot;related security&quot; mean the loan, mortgage (or, in Scotland, standard security) or any other security for any loan or mortgage made in connection with this application;

          <li>references to &quot;I&quot;, &quot;we&quot;, &quot;me&quot;, &quot;us&quot;, &quot;my&quot; and &quot;our&quot; mean the person or persons applying for a loan by way of this application and, where there is more than one applicant, the applicants agree that they are jointly and severally liable for their obligation to repay the loan;

          <li>references to &quot;you&quot;, &quot;your&quot;, &quot;yours&quot; and &quot;yourself&quot; mean DB UK Bank Limited and its successors, transferees and assignees and anyone who at any time in the future has the benefit of the loan, mortgage or related security.

          </ul>

           

          <B>By signing below, </B>

          <BR>I/we declare and agree the following:

           

          <ol type="1">

          <li>       I/we are 18 years of age or over.

          <li>       If the loan is to be regulated by the Financial Services Authority, I/we have received a key facts illustration in relation to the mortgage for which I/we am/are applying and I/we am/are aware of any arrangement fees and/or other fees payable in connection with this application or the advance of any loan to us/me by you.

          <li>       The information given in this application form is true, accurate, complete and not misleading and, where the application form, or part of it, has been completed by someone else, I/we have checked their answers thoroughly and confirm that they are true, accurate, complete and not misleading.

          <li>       The information given by me/us in this form and this declaration will form part of the terms of any mortgage I/we enter into with you. If any information given by me/us is incorrect, I/we agree to make good any loss which you may suffer by acting in reliance on such information.

          <li>       I/we will provide to you any extra information which you request.

          <li>       I/we will notify you promptly of any changes to my/our circumstances that may occur before the mortgage is completed.

          <li>       I/we understand that,

           

          <ol type="a">

          <li>       if there is a material change in my/our circumstances or where my/our ability to meet my/our monthly payments has changed, or

          <li>       if you doubt my/our ability to meet my/our monthly payments for any reason, or

          <li>       if you become aware that any information I/we have given you in relation to my/ our circumstances or the property to be mortgaged or anything else which relates to my/our application is incorrect or misleading or has materially changed, or

          <li>       if you have reason to doubt the value of the property to be mortgaged or if any further investigation is required into a matter which may affect the value of the property, or

          <li>       if satisfactory title to the property to be mortgaged cannot be confirmed, you may refuse to proceed with my/our application and the mortgage.

          </ol>

           

          <li>       Unless otherwise stated in my/our application, I/we declare that I/we have punctually and in the required manner made all payments due under any existing or previous mortgage to which I/we have been a party and that no arrears have arisen under any such existing or previous mortgage.

          <li>       I/we undertake to pay any fees or disbursements incurred by you whether or not the mortgage completes unless you agree otherwise. I/we understand that these fees may include legal fees and that where you make me/us a mortgage offer, I/we will be responsible for the costs and disbursements of any solicitor or conveyancer who acts on your behalf in relation to the mortgage offer (whether or not the mortgage completes) or the mortgage, loan and related security.

          <li>       Where relevant under the terms of any mortgage offer you make to me/us, I/we agree that you may add or deduct any fees (including but not limited to any higher lending charge fee, any completion fee or any telegraphic transfer fee) in relation to the loan, mortgage or related security, from any advance you make to me/us.

          <li>       I/we have sufficient income to sustain any payments required under the terms of this mortgage and I/we understand that if I/we fail to maintain payments due, this may result in the forced sale of the property.

          <li>       You are authorised to give any information relating to this application, including supplying a copy of this application, to my/our intermediary/broker or solicitor/conveyancer. My/our intermediary/broker or solicitor may direct any requests for information about this application to you and you are entitled to accept these requests. My/our intermediary/broker may also write to you with any amendments to this application and may disclose to you any information about me/us in relation to this application which would otherwise be confidential.

          <li>       Any solicitor/conveyancer acting for both you and me/us may disclose to you any information or documentation he/she or you consider relevant in your decision to lend and I/we waive any duty of confidentiality or privilege which may otherwise exist in relation to this mortgage transaction. I/we authorise my/our solicitor/conveyancer to send their file relating to this application and any mortgage, loan and related security to you at your request.

          <li>       I/we understand that any introducer or broker that I/we or you use is not authorised to make any representation or give any undertaking on your behalf and that you will not be bound by or liable for any such representation or undertaking.

          <li>       I/we understand that information given by me/us in this application form may be used for the purposes of obtaining a quote from either a buildings and contents insurer or a mortgage payment protection policy provider and I/we consent to the information in this application form being passed on to any such insurer or provider.

          <li>       In the event that I/we do not arrange appropriate insurance protection for the property, I/we agree that on or after completion of the mortgage, you may insure the property at my/our cost against loss or damage by fire and such other risks and in such amounts you consider necessary.

          <li>       You may make all enquiries you feel appropriate (including with, but not limited to, the Inland Revenue, the Council of Mortgage Lenders Possession Register, my/our bank or building society and, any employer, accountant, mortgage lender or landlord of mine/ours in the three year period preceding your enquiry) for deciding whether to proceed with this application and I/we will be responsible for any costs incurred by you in making such enquiries.

          <li>       I/we understand and agree that you may use automated decision making systems (including credit scoring techniques) in order to help you make your decision as to whether to lend to me/us. I/we also understand that following completion of my/our mortgage, you may use automated systems for the purposes of the research you carry out, or that is carried out on your behalf, in connection with your mortgage lending business.

          <li>       If you decide not to progress my/our mortgage application, you will tell me/us in writing and if your decision is based solely on automated decision making systems, you will give me/us an opportunity to appeal in writing.

          <li>       Where this application is made in joint names, we understand that this will create a &quot;financial association&quot; as joint applicants between us and that, as joint applicants, we are authorised to pass on any information to you about each other. We understand that as a result of our joint application credit reference agencies will link together our records and that they will also link all our previous and subsequent names and addresses. As &quot;financial associates&quot;, our files at credit reference agencies will remain linked until such time as any of us successfully file with the credit reference agencies for a disassociation. Our financial association will create a link between us for the purposes of information recorded from searches made by credit reference agencies and that in future our financial dealings may affect each other. I/we also understand that the information held about me/us at such credit reference agencies may already be linked to one or more people connected to me/us (including other members of my/our household) and that these associated records, together with my/our own records, may also be referred to when you consider my/our application and that the current and previous names, addresses and dates of birth of these connected people may also be supplied.

          <li>       I/we understand that in order to prevent or detect fraud you may make searches at fraud prevention agencies and I/we authorise you to make these searches.

          <li>       I/we understand that if we give you any false or inaccurate information and you suspect fraud and or fraud is identified, you may record this at the fraud prevention agencies to prevent fraud and money laundering. Upon request, you will provide me/us with a list of those fraud prevention agencies with whom you record information.

          <li>       I/we understand that in assessing my/our application you may make searches at credit reference agencies and in so doing you will provide current and previous names, addresses and dates of birth of me/each of us. If I/we provide information about others on a joint application, I am/we are certain that I/we have their agreement. I/we authorise you to make such searches about me/us and to use information provided to you by credit reference agencies and fraud prevention agencies in order to help you make decisions (including credit decisions) about me/us in relation to my/our application, any mortgage which you take over my/our property, any other credit or credit related service you or any of your affiliates provide to me/us, as well as any insurance proposal or claim. You may also make these searches about me/us in order to verify my/our identity, for the purposes of fraud prevention, debt recovery, prevention of money laundering or debtor tracing. I/we acknowledge that when credit reference agencies receive a search from you, they will place a search &quot;footprint&quot; on my/our credit file(s) whether or not my/our application proceeds. If the search is for a credit application, the record of that search (but not the name of the organisation that carried it out) may be seen by other  organisations when I/we apply for credit in the future and I/we understand that it is possible that a large number of these searches carried out in a short period of time may impact on my/our ability to obtain credit.

          <li>       I/we understand that credit reference agencies will supply to you, public information such as County Court Judgements (CCJs) and bankruptcies, electoral register information and fraud prevention information on me/us.

          <li>       I/we understand that you may pass on information as to how I/we run my/our mortgage account to credit reference agencies. This may include the disclosure of any failure by me/us to pay in full and on time and that any such outstanding debt will be recorded by the credit reference agencies. Upon request, you will provide me/us with a list of those credit reference agencies used by you.

          <li>       I/we understand that information provided by you, other organisations and fraud prevention agencies about me/us, and my/our financial associate(s) and my/our business (if I/we have one) to credit reference and fraud prevention agencies may be supplied to other organisations and used by them and you to:

           

          <ol type=”a”>

          <li>       Verify my/our identity if I/we or my/our financial associate apply(ies) for other facilities including all types of insurance applications and claims.

          <li>       Assist other organisations to make decisions on credit, credit related services and on motor, household, life and other insurance proposals and insurance claims, about me/us, my/our partner(s), other members of my/our household or my/our business.

          <li>       Manage my/our personal accounts and my/our (or my partner's) business' accounts and my/our insurance policies (if I/we have one/any).

          <li>       Trace my/our whereabouts and recover payment if I/we do not make payments that I/we owe.

          <li>       Conduct checks for the prevention and detection of crime including fraud and/or money laundering.

          <li>       Undertake statistical analysis and system testing.

          </ol>

          <BR><BR>My/our data may also be used for other purposes for which I/we give my/our specific permission or, in very limited circumstances, when required by law or where permitted under the terms of the Data Protection Act 1998.

          <BR><BR>You may contact me/us by post or telephone, unless I/we have asked you not to, or by e-mail if I/we have given you permission to, to inform me/us of products or services that may be of interest to me/us.

          <li>I/we acknowledge that if I/we borrow from you and do not make the payments that I/we owe, you will trace my/our whereabouts and recover payment.

          <li> I/we understand and agree that my/our loan account may be administered by a third party servicer and that information about the mortgage, loan and related security as well as my/our personal details will be provided to this third party servicer for the purposes of the administration of the loan account.

          <li> You may arrange for a surveyor to provide a valuation report which will be used by you in assessing this application. I/we will be responsible for the cost of this valuation report once the valuer has been instructed, whether or not the mortgage proceeds to completion. I/we acknowledge that part of the valuation fee paid by me/us will be used towards the cost of any initial assessment by you of my/our application. If the application is declined or does not proceed before the valuer has been instructed, the valuation fee may be refunded net of any costs. If you provide me/us with a copy of, or extract from, your valuation report I/we understand that you make no representation or warranty (expressed or implied) nor accept any liability or responsibility in respect of its content. I/we understand that your valuation report is not a detailed or structural report or survey about the condition of the property and that the report may fail to reveal serious defects to the property and I/we acknowledge that I/we will not rely on this report for the purposes of my/our decision to purchase the property. I/we recognise that it is strongly advised to obtain a more comprehensive survey as to the condition and value of the property. 

          <li> I/we understand that any additional security insurance arrangements you make in relation to my/our loan, mortgage or its related security, are for your benefit only and that I/we have no right or claim in relation to them.

          <li> I/we will not let the property without your prior consent in writing and I/we will not create any further security over the property prior to or after completion of the mortgage without informing you and obtaining your prior consent in writing;

          <li> I/we understand that the rate of interest and monthly payment in relation to my/our mortgage may be varied from time to time in accordance with the terms and conditions relevant to the mortgage.

          <li> I/we understand that an early repayment charge may be payable in relation to the mortgage if the mortgage is redeemed within a certain period in accordance with the terms and conditions relevant to the mortgage.

          <li> Where I/we repay only the interest on any money you have lent us, I/we understand that it is my/our responsibility to put in place and maintain a savings vehicle to ensure that at the end of the term of my/our mortgage I/we will be able to pay off my/our debt in full. I/we understand that you advise me/us to obtain independent financial advice in relation to this savings vehicle. I/we will regularly check the performance of the savings vehicle and I/we will not allow anything to be done which might result in my/our savings vehicle coming to an end or being cancelled or becoming void or voidable. I/we will pay any payments or premiums due under my/our savings vehicle on time and in full.

          <li> I/we acknowledge that it is my/our responsibility to ensure that appropriate life cover or other means of repayment is in place to repay the mortgage in the event of my/our death.

          <li> I/we understand that telephone calls with me/us may be recorded or monitored for training, security and/or quality purposes.

          </ol>

          <BR><BR>

          <B>Data Protection Act</b>

          <BR>I/we agree that details of this application, any mortgage offer or loan that you may make to me/us, the property secured by the mortgage, and the conduct of my/our account(s) with you (all of which are my/our &quot;personal details&quot;) may be held by you and used by you in making credit decisions about me/us, for credit control purposes, in administering my/our account(s) with you and for marketing or statistical analysis. I/we agree that you may disclose my/our personal details to:

          <ol type=”a”>

          <li>       any licensed credit reference agency where they will be stored and used by you or other lenders in making credit decisions about me/us and other members of my/our household;

          <li>       fraud prevention agencies;

          <li>       if you suspect fraud, to the police and any other relevant law enforcement agency;

          <li>       to an actual or proposed third party guarantor of my/our obligations under the mortgage;

          <li>       any actual or proposed buildings and/or contents insurer and any actual or proposed insurer you wish to use to provide you with additional security who will use them to help decide whether or not to offer cover and whether or not to process claims;

          <li>       any actual or proposed purchaser of my/our mortgage or any one who takes a charge over it and any person involved in its funding or securitisation, and all their advisors;

          <li>       the Council of Mortgage Lenders Possessions Register;

          <li>       the Land Registry, the Registers of Scotland and H.M. Revenue and Customs and other proper bodies, persons or bureaux; and/or

          <li>       to any other person to comply with any legal or regulatory requirement in the United Kingdom or elsewhere.

          </ol>

          <BR><BR>Information passed on to such third parties may be used for future lending decisions, arrears handling or fraud prevention. 

          <BR><BR>I/we understand that you may possess &quot;sensitive information&quot; about me/us, including information about past criminal convictions and health data. This information will be used only for assessing risk or my/our eligibility for a mortgage or insurance cover. I/we consent to such sensitive information being processed by you and your service providers and agents for these purposes.

          <BR><BR>I/we agree that my/our personal details may be disclosed and used within your group of companies (including to and by their respective service providers, agents and actual and proposed successors) to conduct and monitor and analyse your business.

          <BR><BR>I/we agree that, unless I/we have ticked the boxes below, you may send me/us marketing information relating to other financial products and services which you believe may be of interest to me/us by email, post or telephone.

          <BR><BR>I/we understand and agree that my/our personal details may be transferred (including by your service providers and agents) to countries outside the European Economic Area that may not have the same level of data protection legislation as inside the United Kingdom. This means that personal details processed in such countries may have less protection than inside the United Kingdom. The purpose of these transfers may include data processing (including data processing performed by your agents and service providers) and head office reporting.

          <BR><BR>I/we understand that upon request, you will provide me/us with the details of any organisation to which you have passed on any information about us. I/we will pay a fee of £10 for this information.

          <BR><BR>

          <B>Transfer</B>

          <BR>I/we confirm that you may, without limitation, transfer, charge or otherwise dispose of the loan, mortgage or security (and any of your rights under such loan, mortgage or related security) in whole or in part to any transferee (and transferee shall mean any person, company, association, society or other entity, whether incorporated or unincorporated) at any time without our consent.

          <BR><BR>

          I/we understand that where you transfer to any transferee the right to set the interest rate charged under the loan, the interest rate may be set by reference to that person's own interest rate so that, for example, where the interest rate payable before the transfer is your standard variable interest rate, the transferee may change the interest rate following the transfer to the transferee's own standard variable interest rate.

          <BR><BR>I/we understand that on any transfer of the loan, you will enter into an agreement with the transferee that you (or your agent) will continue to conduct arrears cases as the agent of the transferee and that the transferee will agree that its policy on handling any arrears and exercising any discretion in the setting of the interest rate will be identical to your policy. The agreement will apply for a minimum of three months after the transfer out may be terminated early by the transferee if your performance as agent is not satisfactory or if you, or the transferee, suffer financial difficulties. I/we understand and agree that any actual or proposed purchaser of the mortgage may carry out searches at credit reference agencies and that such searches will be entered on the records of the credit reference agencies.

          <BR><BR>

          Following any transfer of the loan, mortgage or related security, I/we agree that you may seek information about the loan, mortgage or related security from the transferee for the purposes of your credit-scoring research and performance data records.

          <BR><BR>

           

          <B>Insurance Notice and Declaration</B>

          <BR><BR>I/we understand that insurers pass information to the Claims and Underwriting Exchange Register, run by Insurance Database Services Limited (IDS Ltd). The aim is to help you check information provided and also to prevent fraudulent claims. When you deal with my/our request for insurance, you may search the register. When I/we tell you about an incident (such as fire, water damage or theft) which may or may not give rise to a claim, you will pass information relating to it to the register. I/we understand that I/we may ask you for more information about this.

          <BR><BR>I/we will show this notice to anyone who has an interest in the property insured under the policy.

          <BR><BR>I/we understand that you will pass information on this form and about any incident I/we give details of to IDS Ltd so that they can make it available to other insurers. I/we also understand that, in response to any searches you make in connection with this application or any incident that I/we have given details of, IDS Ltd may pass you information it has received from other insurers about other incidents involving anyone insured under this policy.

          <BR><BR>

          All applicants to the mortgage are required to sign the following section.

          <BR><BR>I/we have read and understood the information contained in the declaration section of this application form and, by signing and dating this application, I/we give my/our consent to the use of my/our information in this way.

          <BR><BR>

          YOUR HOME MAY BE REPOSSESSED IF YOU DO NOT KEEP UP

          <BR>REPAYMENTS ON YOUR MORTGAGE

          <BR><BR>

          THIS APPLICATION FORM MUST BE SIGNED AND DATED FOR US TO

          <BR>BE ABLE TO PROCEED WITH THE APPLICATION

          <BR><BR>

           

          DB UK Bank Limited accepts no responsibility for any representations made by any employee or agent of DB UK Bank Limited or any

          other person unless these are incorporated in the offer of loan or are subsequently confirmed by DB UK Bank Limited in writing.

          <BR><BR>

          db mortgages is a trading name of DB UK Bank Limited authorised and regulated

          <BR>by the Financial Services Authority and a Member of The London Stock Exchange.

          <BR>Registered in England and Wales No. 315841

          <BR>Registered office: 23 Great Winchester Street, London EC2P 2AX
        </div> 
      </tr>
      <tr class="notprinted">
        <td style="text-align:right">
          <asp:button cssclass="button" align="right" autopostback="false" id="btnPrint" runat="server" text="Print declaration"></asp:button>
        </td>
      </tr>
      <tr class="notprinted">
        <td>
          <asp:checkbox id="chkAgreeTC" runat="server" text="Agree and understand the above" tooltip="check this box to indicate that you have read and understood the declaration"
            textalign="right"></asp:checkbox>
        </td>
      </tr>
      
     <tr class="notprinted">
       <td>
         <asp:checkbox id="chkKeyFactsAccepted" runat="server" text="Has the customer received and accepted a Key Facts Illustration?" tooltip="check this box to indicate that the customer has accepted a Key Facts Illustration"
           textalign="right"></asp:checkbox>
       </td>
     </tr>
     <tr class="notprinted">
       <td>
         <asp:label id="lblMarketingOptIn" runat="server" text="Marketing data protection opt-in" tooltip="check these boxes to indicate that you wish to receive marketing material from DB"
            textalign="right"></asp:label>
       </td>
     </tr>
  </table>

  <div class="notprinted">
    <epsom:applicanttype id="ctlApplicantType" runat="server" required="false"/>
  </div>

  <table class="notdisplayed">
    <tr>
      <td><br /></td>
    </tr>
    <tr>
      <td>Applicant 1 signature</td>
      <td><textarea cols="50" rows="3" style="overflow: hidden"></textarea></td>
    </tr>
    <tr>
      <td><br /></td>
    </tr>
    <tr>
      <td>Applicant 1 name (printed)</td>
      <td><textarea cols="50" rows="3" style="overflow: hidden"></textarea></td>
    </tr>
    <tr>
      <td><br /></td>
    </tr>
    <tr>
      <td>Date</td>
      <td><textarea cols="20" rows="1" style="overflow: hidden"></textarea></td>
    </tr>
  </table>

  <table class="notdisplayed">
    <tr>
      <td><br /><br /></td>
    </tr>
    <tr>
      <td>Applicant 2 signature</td>
      <td><textarea cols="50" rows="3" style="overflow: hidden"></textarea></td>
    </tr>
    <tr>
      <td><br /></td>
    </tr>
    <tr>
      <td>Applicant 2 name (printed)</td>
      <td><textarea cols="50" rows="3" style="overflow: hidden"></textarea></td>
    </tr>
    <tr>
      <td><br /></td>
    </tr>
    <tr>
      <td>Date</td>
      <td><textarea cols="20" rows="1" style="overflow: hidden"></textarea></td>
    </tr>
  </table>

  <epsom:dipnavigationbuttons id="ctlDIPNavigationButtons2" buttonbar="true" savepage="true" saveandclosepage="true" runat="server"></epsom:dipnavigationbuttons>

  </mp:content>

</mp:contentcontainer>
