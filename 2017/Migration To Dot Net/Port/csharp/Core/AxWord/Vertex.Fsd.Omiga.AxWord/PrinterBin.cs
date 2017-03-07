using System;
using System.Collections.Generic;
using System.Text;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class PrinterBin
	{
		private int _number;
		public int Number
		{
			get { return _number; }
			set { _number = value; }
		}

		private string _name;
		public string Name
		{
			get { return _name; }
			set { _name = value; }
		}

		public PrinterBin()
			: this("Default", MicrosoftWordConstants.wdPrinterDefaultBin)
		{
		}

		public PrinterBin(int number) : this(null, number)
		{
		}

		public PrinterBin(string name, int number)
		{
			_name = name;
			_number = number;
		}

		public override string ToString()
		{
			return _name;
		}
	}
}
