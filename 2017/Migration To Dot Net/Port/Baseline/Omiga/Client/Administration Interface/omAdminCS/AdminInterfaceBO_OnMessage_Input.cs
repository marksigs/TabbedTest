using System.Xml.Serialization;
using WebRef = Vertex.Fsd.Omiga.omAdminCS.WebRefCreateMortgageAccount;

namespace Vertex.Fsd.Omiga.omAdminCS
{
	/// <remarks/>
	[System.Xml.Serialization.XmlRootAttribute("REQUEST", Namespace="", IsNullable=false)]
	public class AdminInterfaceBO_OnMessage_Input 
	{
    
		/// <remarks/>
		public BATCHAUDIT BATCHAUDIT;
    
		/// <remarks/>
		public PAYMENTRECORD PAYMENTRECORD;
    
		/// <remarks/>
//		public WebRef.CreateMortgageAccountRequestType CreateMortgageAccountRequest;
		[System.Xml.Serialization.XmlAnyElement()]
		public object[] OtherElements;
  
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string OPERATION;
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string USERID;
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string UNITID;

		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string PASSWORD;

		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string OTHERSYSTEMCUSTOMERNUMBER;
	}

	/// <remarks/>
	public class BATCHAUDIT 
	{
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string BATCHRUNNUMBER;
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string BATCHNUMBER;
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string BATCHAUDITGUID;
	}


	/// <remarks/>
	public class PAYMENTRECORD 
	{
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string APPLICATIONNUMBER;
    
		/// <remarks/>
		[System.Xml.Serialization.XmlAttributeAttribute()]
		public string PAYMENTSEQUENCENUMBER;
	}

}
