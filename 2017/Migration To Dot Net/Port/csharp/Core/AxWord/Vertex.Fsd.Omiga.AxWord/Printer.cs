using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing.Printing;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class Printer
	{
		private string _name;
		public string Name
		{
			get { return _name; }
			set { _name = value; }
		}

		private string _port;
		public string Port
		{
			get { return _port; }
			set { _port = value; }
		}

		private bool _isDefault;
		public bool IsDefault
		{
			get { return _isDefault; }
			set { _isDefault = value; }
		}

		private PrinterBin _defaultBin;
		public PrinterBin DefaultBin
		{
			get { return _defaultBin; }
			set { _defaultBin = value; }
		}

		private PrinterBinCollection _bins;
		public PrinterBinCollection Bins
		{
			get { return _bins; }
			set { _bins = value; }
		}

		public Printer()
			: this(null, null)
		{
		}

		public Printer(string name)
			: this(name, null)
		{
		}

		public Printer(string name, string port)
		{
			_name = name;
			_port = port;
			_bins = new PrinterBinCollection(this);
			if (_bins.Count > 0)
			{
				// First bin in list is always the default.
				_defaultBin = _bins[0];
			}
		}

		public override string ToString()
		{
			return _name;
		}
	}
}

