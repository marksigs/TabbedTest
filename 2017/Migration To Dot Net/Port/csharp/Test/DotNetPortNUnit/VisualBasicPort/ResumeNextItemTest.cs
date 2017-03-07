using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using NUnit.Framework;
using Vertex.Fsd.Omiga.VisualBasicPort;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Tests
{
	[TestFixture]
	public class ResumeNextItemTest
	{
		[Test]
		public void TestResumeNextItem()
		{
			/*
			try
			{
				ThrowSuppressedExceptions();
				Assert.Fail(); // Should not get to this line.
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.RecordNotFound);
			}

			XmlDocument xmlDocument = new XmlDocument();
			xmlDocument.AppendChild(xmlDocument.CreateElement("RESPONSE"));
			using (new ResumeNext(OMIGAERROR.RecordNotFound, xmlDocument.DocumentElement))
			{
				ThrowSuppressedExceptions();
			}
			Assert.AreEqual(xmlDocument.DocumentElement.GetAttribute("TYPE"), "WARNING");

			try
			{
				ThrowSuppressedExceptions();
				Assert.Fail(); // Should not get to this line.
			}
			catch (ErrAssistException exception)
			{
				Assert.AreEqual(exception.OmigaError, OMIGAERROR.RecordNotFound);
			}
			 */
		}

		private void ThrowSuppressedExceptions()
		{
			/*
			using (new ResumeNext(OMIGAERROR.NotImplemented))
			{
				new ErrAssistException(OMIGAERROR.NotImplemented).Throw();
				new ErrAssistException(OMIGAERROR.RecordNotFound).Throw();
				new ErrAssistException(118).Throw();

				try
				{
					new ErrAssistException(OMIGAERROR.ObjectNotCreatable).Throw();
					Assert.Fail();
				}
				catch (ErrAssistException exception)
				{
					Assert.AreEqual(exception.OmigaError, OMIGAERROR.ObjectNotCreatable);
				}
			}
			 */
		}
	}
}
