using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class DocumentText : Document
	{
		public DocumentText()
		{
		}

		public DocumentText(string fileName)
			: base(fileName)
		{
		}

		protected override void WriteToFile(FileStream fileStream, byte[] bytes)
		{
			using (StreamWriter textWriter = new StreamWriter(fileStream))
			{
				textWriter.Write(GetEncoding(bytes).GetString(bytes));
			}
		}

		protected override byte[] ReadFromFile(FileStream fileStream)
		{
			byte[] bytes = null;

			using (StreamReader textReader = new StreamReader(fileStream))
			{
				bytes = Encoding.Unicode.GetBytes(textReader.ReadToEnd());
			}

			return bytes;
		}

		protected virtual Encoding GetEncoding(byte[] bytes)
		{
			return Encoding.Unicode;
		}
	}
}
