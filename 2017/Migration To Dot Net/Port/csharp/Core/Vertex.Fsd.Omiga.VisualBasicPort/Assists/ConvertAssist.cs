/*
--------------------------------------------------------------------------------------------
Workfile:			ConvertAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Conversion routines.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from ConvertAssist.bas and ConvertAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Text;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Common conversion routines.
	/// </summary>
	public static class ConvertAssist
	{
		/// <summary>
		/// Converts a string to an array of ASCII bytes.
		/// </summary>
		/// <param name="bytes">The existing array of bytes.</param>
		/// <param name="text">The string to be converted</param>
		/// <exception cref="ErrAssistException">The length of the string exceeds the length of the byte array - 1.</exception>
		public static void StringToByteArray(ref byte[] bytes, string text) 
		{
			// HasValue string plus null terminator will fit into the byte array
			if (text.Length > bytes.Length - 1)
			{
				throw new InvalidParameterException();
			}

			bytes = Encoding.ASCII.GetBytes(text);
		}

		/// <summary>
		/// Converts a byte array to a string.
		/// </summary>
		/// <param name="bytes">The byte array.</param>
		/// <returns>The converted byte array.</returns>
		public static string ByteArrayToString(byte[] bytes) 
		{
			string text = String.Empty;
			int index = 0;
			while (index < bytes.Length && bytes[index] != Convert.ToByte(0))
			{
				text = text + (char)(bytes[index]);
				index++;
			} 
			return text;
		}

		/// <summary>
		/// Safe conversion of a string expression to a DateTime object.
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or Date.Now if expression is either null or the empty string.</returns>
		public static DateTime CSafeDate(string expression) 
		{
			return expression != null && expression.Length > 0 ? Convert.ToDateTime(expression) : DateTime.Now;
		}

		/// <summary>
		/// Safe conversion of a string expression to a 32 bit integer.
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or 0 if expression is either null or the empty string.</returns>
		public static int CSafeLng(string expression) 
		{
			return expression != null && expression.Length > 0 ? Convert.ToInt32(expression) : 0;
		}

		/// <summary>
		/// Safe conversion of a string expression to a double.
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or 0 if expression is either null or the empty string.</returns>
		public static double CSafeDbl(string expression) 
		{
			return expression != null && expression.Length > 0 ? Convert.ToDouble(expression) : 0;
		}

		/// <summary>
		/// Safe conversion of a string expression to an ASCII byte.
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or 0 if expression is either null or the empty string.</returns>
		public static byte CSafeByte(string expression) 
		{
			return expression != null && expression.Length > 0 ? Encoding.ASCII.GetBytes(Convert.ToString(expression))[0] : (byte)0;
		}

		/// <summary>
		/// Safe conversion of a string expression to a 16 bit integer (short).
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or 0 if expression is either null or the empty string.</returns>
		public static short CSafeInt(string expression) 
		{
			return expression != null && expression.Length > 0 ? Convert.ToInt16(expression) : (short)0;
		}

		/// <summary>
		/// Safe conversion of a string expression to a bool.
		/// </summary>
		/// <param name="expression">The expression to convert.</param>
		/// <returns>The converted expression or false if expression is either null or the empty string.</returns>
		public static bool CSafeBool(string expression) 
		{
			return expression != null && expression.Length > 0 ? Convert.ToBoolean(expression) : false;
		}
	}
}
