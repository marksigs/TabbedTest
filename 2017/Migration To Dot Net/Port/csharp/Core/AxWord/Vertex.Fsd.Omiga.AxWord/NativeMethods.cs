using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace Vertex.Fsd.Omiga.AxWord
{
	/// <summary>
	/// Summary description for NativeMethods.
	/// </summary>
	public static class NativeMethods
	{
		#region FindExecutable
		[DllImport("shell32.dll")]
		private static extern IntPtr FindExecutable(string fileName, string directory, [Out] StringBuilder result);

		public static string FindExecutable(string fileName)
		{
			string executableFileName = null;
			string directory = "";
			StringBuilder result = new StringBuilder(255);
			if (FindExecutable(fileName, directory, result).ToInt32() > 32)
			{
				executableFileName = result.ToString();
			}

			return executableFileName;
		}
		#endregion FindExecutable

		#region pdfPrint
		[DllImport("WPPdfSDK01.dll")]
		private static extern IntPtr pdfPrint(string fileName, string password, string licenceName, string licenceKey, int licenceCode, string options);

		public static int PdfPrint(string fileName, string password, string licenceName, string licenceKey, int licenceCode, string options)
		{
			return pdfPrint(fileName, password, licenceName, licenceKey, licenceCode, options).ToInt32();
		}
		#endregion pdfPrint

		#region DeviceCapabilities : short
		public enum DeviceCapabilitiesIndices
		{
			DC_FIELDS = 1,
			DC_PAPERS = 2,
			DC_PAPERSIZE = 3,
			DC_MINEXTENT = 4,
			DC_MAXEXTENT = 5,
			DC_BINS = 6,
			DC_DUPLEX = 7,
			DC_SIZE = 8,
			DC_EXTRA = 9,
			DC_VERSION = 10,
			DC_DRIVER = 11,
			DC_BINNAMES = 12,
			DC_ENUMRESOLUTIONS = 13,
			DC_FILEDEPENDENCIES = 14,
			DC_TRUETYPE = 15,
			DC_PAPERNAMES = 16,
			DC_ORIENTATION = 17,
			DC_COPIES = 18,

			//#if(WINVER >= 0x0400)
			DC_BINADJUST = 19,
			DC_EMF_COMPLIANT = 20,
			DC_DATATYPE_PRODUCED = 21,
			DC_COLLATE = 22,
			DC_MANUFACTURER = 23,
			DC_MODEL = 24,
			//#endif // WINVER >= 0x0400

			//#if(WINVER >= 0x0500)
			DC_PERSONALITY = 25,
			DC_PRINTRATE = 26,
			DC_PRINTRATEUNIT = 27,
			DC_PRINTERMEM = 28,
			DC_MEDIAREADY = 29,
			DC_STAPLE = 30,
			DC_PRINTRATEPPM = 31,
			DC_COLORDEVICE = 32,
			DC_NUP = 33,
			DC_MEDIATYPENAMES = 34,
			DC_MEDIATYPES = 35,
			//#endif // WINVER >= 0x0500
		}

		[DllImport("winspool.drv", SetLastError = true)]
		private static extern Int32 DeviceCapabilities(
			string device, 
			string port,
			DeviceCapabilitiesIndices capability,
			IntPtr outputBuffer,
			IntPtr deviceMode);

		public static void GetPrinterBins(string deviceName, string port, out StringCollection binNames, out List<int> binNumbers)
		{
			binNames = new StringCollection();
			binNumbers = new List<int>();

			binNames.Add("Default");
			binNumbers.Add(MicrosoftWordConstants.wdPrinterDefaultBin);

			// Bins
			int binNumberCount = DeviceCapabilities(deviceName, port, DeviceCapabilitiesIndices.DC_BINS, (IntPtr)null, (IntPtr)null);
			IntPtr addressPointer = Marshal.AllocHGlobal((int)binNumberCount * 2);
			try
			{
				binNumberCount = DeviceCapabilities(deviceName, port, DeviceCapabilitiesIndices.DC_BINS, addressPointer, (IntPtr)null);
				if (binNumberCount < 0)
				{
					throw new Win32Exception(Marshal.GetLastWin32Error(), "Error getting bin numbers for " + deviceName);
				}
				int offset = addressPointer.ToInt32();
				for (int i = 0; i < binNumberCount; i++)
				{
					try
					{
						binNumbers.Add(Marshal.ReadIntPtr(new IntPtr(offset + i * 2)).ToInt32());
					}
					catch (AccessViolationException)
					{
						// This exception occurs intermittently for some unknown reason; 
						// use default bin instead.
						binNumbers.Add(MicrosoftWordConstants.wdPrinterDefaultBin);
					}
				}
			}
			finally
			{
				Marshal.FreeHGlobal(addressPointer);
			}

			// BinNames
			int binNameCount = DeviceCapabilities(deviceName, port, DeviceCapabilitiesIndices.DC_BINNAMES, (IntPtr)null, (IntPtr)null);
			addressPointer = Marshal.AllocHGlobal((int)binNameCount * 24);
			try
			{
				binNameCount = DeviceCapabilities(deviceName, port, DeviceCapabilitiesIndices.DC_BINNAMES, addressPointer, (IntPtr)null);
				if (binNameCount < 0)
				{
					throw new Win32Exception(Marshal.GetLastWin32Error(), "Error getting bin names for " + deviceName);
				}

				int offset = addressPointer.ToInt32();
				for (int i = 0; i < binNameCount && i < binNumberCount; i++)
				{
					binNames.Add(Marshal.PtrToStringAnsi(new IntPtr(offset + i * 24)));
				}
			}
			finally
			{
				Marshal.FreeHGlobal(addressPointer);
			}
		}

		#endregion DeviceCapabilities

	}
}
