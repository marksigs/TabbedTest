using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class DocumentBinary : Document
	{
		public DocumentBinary()
		{
		}

		public DocumentBinary(string fileName)
			: base(fileName)
		{
		}

		protected override void WriteToFile(FileStream fileStream, byte[] bytes)
		{
			using (BinaryWriter binaryWriter = new BinaryWriter(fileStream))
			{
				binaryWriter.Write(bytes);
			}
		}

		protected override byte[] ReadFromFile(FileStream fileStream)
		{
			byte[] bytes = null;

			using (BinaryReader binaryReader = new BinaryReader(fileStream))
			{
				bytes = binaryReader.ReadBytes((int)fileStream.Length);
			}

			return bytes;
		}
	}
}
