using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DbaGen
{
	/// <summary>
	/// From MSDN: The default value is for a DLL case address is 0x11000000 
	/// (285,212,672). If you neglect to change this value, your component will 
	/// conflict with every other in-process component compiled using the default. 
	/// Staying well away from this address is recommended. Choose a base address 
	/// between 16 megabytes (16,777,216 or 0x1000000) and two gigabytes 
	/// (2,147,483,648 or 0x80000000). The base address must be a multiple of 64K. 
	/// The memory used by your component begins at the initial base address and 
	/// is the size of the compiled file, rounded up to the next multiple of 64K.
	/// </summary>
	public struct DllBaseAddress
	{
		private const uint _baseMinimum = 0x1000000;	// 16 MB
		private const uint _baseMaximum = 0x80000000;	// 2 GB
		private const uint _baseMultiple = 0x10000;		// 64 K

		private string _root;
		private uint _address;

		public DllBaseAddress(uint address)
		{
			_root = null;
			_address = address;
		}

		public DllBaseAddress(string root, uint address)
		{
			_root = root;
			_address = address;
		}

		public static DllBaseAddress NewDllBaseAddress()
		{
			return NewDllBaseAddress(null);
		}

		public static DllBaseAddress NewDllBaseAddress(string root)
		{
			uint address = 0;
			uint result = 0;

			if (string.IsNullOrEmpty(root))
			{
				// Randomly generate base address.
				Random random = new Random();
				result = (uint)random.Next((int)_baseMinimum, Int32.MaxValue);
			}
			else
			{
				// Base address derived from root.
				result = _baseMinimum;
				foreach (char character in root)
				{
					result += character * _baseMultiple;
				}
			}

			if (result < _baseMinimum || result > _baseMaximum)
			{
				throw new InvalidOperationException("Generated DLL base address is < " + _baseMinimum.ToString() + " or > " + _baseMaximum.ToString() + ".");
			}
			else
			{
				address = (uint)(Math.Ceiling((double)result / _baseMultiple) * _baseMultiple);
			}

			return new DllBaseAddress(root, address);
		}

		public override string ToString()
		{
			return ToString("x");
		}

		public string ToString(string format)
		{
			string text = null;

			if (string.Compare(format, "H") == 0)
			{
				text = "&H" + _address.ToString("X");
			}
			else if (format == "x")
			{
				text = "0x" + _address.ToString("x");
			}
			else if (format == "X")
			{
				text = "0X" + _address.ToString("X");
			}
			else 
			{
				text = _address.ToString();
			}

			return text;
		}
	}
}
