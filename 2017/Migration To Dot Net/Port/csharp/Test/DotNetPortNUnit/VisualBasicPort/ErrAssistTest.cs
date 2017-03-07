using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort.Assists;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	class AdrianException : ErrAssistException
	{
		public AdrianException()
			: base("Adrian Error")
		{
		}
	}

	[TestFixture]
	public class ErrAssistExTest
	{
		[Test]
		public void TestThrowError()
		{
			XmlNode xmlResponseNode = null;
			XmlDocument xmlOut = new XmlDocument();
			XmlElement xmlResponseElement = xmlOut.CreateElement("RESPONSE");
			xmlOut.AppendChild((XmlNode)xmlResponseElement);

			try
			{
				try
				{
					throw new RecordNotFoundException();	// May raise a record not found error.
				}
				catch (RecordNotFoundException)
				{
					// Ignore error and continue.
				}

				try
				{
					throw new PasswordExpiredException();	// May raise a warning.
				}
				catch (OmigaWarningException exception)
				{
					// Add warning details to the xml response and continue.
					exception.AddWarning(xmlResponseElement);
				}

				//throw new ArgumentException("Argument is null");

				xmlResponseNode = (XmlNode)xmlResponseElement;
			}
			catch (Exception exception)
			{
				// Create new exception. This will:
				// 		Write the error details to the event log if it is a system error.
				//		Set any COM+ transaction to complete (because of the TransactionAction.SetComplete)
				//		Create the xml response node to return from this function (by calling CreateErrorResponseNode).
				xmlResponseNode = new ErrAssistException(exception, TransactionAction.SetComplete).CreateErrorResponseNode();
			}

			string x = xmlResponseNode.OuterXml;

			/*
				ErrAssistException ex = new ErrAssistException(544, "This is an error");
				int x = ex.OmigaErrorNumber;
				bool b = ex.IsWarning();
				ErrAssistException ex2 = new ErrAssistException(8530 + -2147221504 + 512);
				int x2 = ex2.OmigaErrorNumber;
				bool z = ex2.IsApplicationError();
				bool z2 = ex2.IsWarning();
				ErrAssistException.GetOmigaErrorNumber(500);
				ErrAssistException.GetOmigaErrorNumber(500 + -2147221504 + 512);
				 */

				/*
				try
				{
					//throw new ErrAssistException(303, "Fishy");
					throw new Exception("Fishy");
					//throw new ErrAssistException(544, "Fishy");
					//throw new AdrianException();
				}
				catch (Exception ex)
				{
					XmlDocument xmlDocument = new XmlDocument();
					XmlElement xmlResponse = xmlDocument.CreateElement("RESPONSE");

					ErrAssistException exception = new ErrAssistException(ex);
					if (exception.IsWarning())
					{
						xmlDocument.AppendChild(xmlResponse);
						exception.AddWarning(xmlResponse);
					}
					if (exception.IsSystemError())
					{
						exception.LogError();
					}
					string response = exception.CreateErrorResponse();
				}
			}
			catch (ErrAssistException)
			{
			}
	*/
		}

		[Test]
		public void TestGetMessageText()
		{
			/*//string messageText = ErrAssistException.GetMessageText(OMIGAERROR.RecordNotFound);
			ErrAssistException.ExpandMessageText("This is a test");
			ErrAssistException.ExpandMessageText("This is a test", "one");
			ErrAssistException.ExpandMessageText("This is a %s test", "one");
			ErrAssistException.ExpandMessageText("This is a %s %s test", "one", "two");
			ErrAssistException.ExpandMessageText("This is a %s %s test", "one");
			ErrAssistException.ExpandMessageText("This is a %s %s test", "one", null);
			//Assert.AreEqual(messageText, "Record not found");
			 */
		}
	}
}
