using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;

namespace Vertex.Fsd.Omiga.AxWord
{
	public class PrinterBinCollection : List<PrinterBin>
	{
		public PrinterBinCollection(Printer printer)
		{
			Initialize(printer);
		}

		private void Initialize(Printer printer)
		{
			if (printer == null)
			{
				throw new ArgumentNullException("printer");
			}

			if (!string.IsNullOrEmpty(printer.Name))
			{
				StringCollection binNames;
				List<int> binNumbers;
				NativeMethods.GetPrinterBins(printer.Name, printer.Port, out binNames, out binNumbers);

				for (int i = 0; i < binNames.Count && i < binNumbers.Count; i++)
				{
					Add(new PrinterBin(binNames[i], binNumbers[i]));
				}
			}
		}
	}
}
