/*
--------------------------------------------------------------------------------------------
Workfile:			GuidAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Guid support.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from guidAssist.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Helper routines for Omiga guids.
	/// </summary>
	/// <remarks>
	/// Omiga generally represents a guid as a 32 character string, where each successive byte in the guid 
	/// is represented by the next two characters in the string. This is the same format SQL Query Analyser 
	/// uses when displaying a guid that is stored in a BINARY(16) column.
	/// </remarks>
	public static class GuidAssist
	{
		/// <summary>
		/// Creates a new Omiga guid string.
		/// </summary>
		/// <returns>The new Omiga guid.</returns>
		public static string CreateGUID() 
		{
			return ToString(ToByteArray(Guid.NewGuid()));
		}

		/// <summary>
		/// Converts an array of bytes into an Omiga guid string.
		/// </summary>
		/// <param name="bytes">The bytes to convert.</param>
		/// <returns>An Omiga guid string.</returns>
		public static string ToString(byte[] bytes)
		{
			StringBuilder stringBuilder = new StringBuilder();

			foreach (byte b in bytes)
			{
				stringBuilder.Append(b.ToString("X2"));
			}

			return stringBuilder.ToString();
		}

		/// <summary>
		/// Converts a native .Net guid into a byte array
		/// </summary>
		/// <param name="guid">The .Net guid.</param>
		/// <returns>A byte array containing the re-ordered bytes in the guid.</returns>
		/// <remarks>
		/// The bytes in the array are re-ordered so that if the array is converted into an Omiga guid 
		/// string (e.g., using <see cref="ToString"/>, then the bytes will be in the correct order, i.e., 
		/// the same order used by SQL Server for the BINARY(16) type.
		/// </remarks>
		public static byte[] ToByteArray(Guid guid)
		{
			byte[] bytes = guid.ToByteArray();

			SwapBytes(ref bytes, 0, 3);
			SwapBytes(ref bytes, 1, 2);
			SwapBytes(ref bytes, 4, 5);
			SwapBytes(ref bytes, 6, 7);

			return bytes;
		}

		private static void SwapBytes(ref byte[] bytes, int index1, int index2)
		{
			if (index1 > bytes.Length - 1)
			{
				throw new ArgumentException("Index out of bounds.", "index1");
			}
			else if (index2 > bytes.Length - 1)
			{
				throw new ArgumentException("Index out of bounds.", "index2");
			}

			byte byte1 = bytes[index1];
			byte byte2 = bytes[index2];
			bytes[index1] = byte2;
			bytes[index2] = byte1;
		}
	}
}
