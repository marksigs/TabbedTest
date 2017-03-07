/*
--------------------------------------------------------------------------------------------
Workfile:			EncryptAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Encryption and decryption routines. Based on code from Encrypt.dll.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
RF		27/10/1999	Created
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		09/05/2007	First .Net version. Ported from EncryptAssist.bas.
--------------------------------------------------------------------------------------------
*/
using System;
using Microsoft.VisualBasic;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// Encryption and decryption routines.
	/// </summary>
	public static class EncryptAssist
	{
		private const int cint_MIN = 10;
		private const int cint_MAX = 40;

		/// <summary>
		/// Encrypts a string.
		/// </summary>
		/// <param name="decryptedText">The string to encrypt.</param>
		/// <returns>The encrypted string.</returns>
		/// <remarks>
		/// The encryption uses a symmetric key algorithm.
		/// </remarks>
		public static string Encrypt(string decryptedText) 
		{
			string encryptedText = String.Empty;
			// initialise the random-number generator
			VBMath.Rnd((float)-1.0);
			VBMath.Randomize(1.0);
			for (int charIndex = 0; charIndex < decryptedText.Length; charIndex++)
			{
				short increment = Convert.ToInt16(Conversion.Int(VBMath.Rnd() * cint_MAX) + cint_MIN);
				int char1 = Strings.Asc(Convert.ToString(decryptedText).Substring(charIndex - 1, 1));
				short char2 = (short)(char1 + increment);
				if (char2 >= Strings.Asc("~"))
				{
					// new character exceeds ascii 127 ('~'), thus deduct increment
					// precede new character with "~" in output string so that Decrypt
					// knows that the increment should be added and not deducted
					encryptedText = encryptedText + "~";
					char2 = (short)(char1 - increment);
				}
				encryptedText = encryptedText + (char)(char2);
			}

			return encryptedText;
		}
	}
}
