/*
--------------------------------------------------------------------------------------------
Workfile:			GeneralAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		General helper object.
--------------------------------------------------------------------------------------------
History:
Prog	Date		Description
PSC		16/08/1999	Created
RF		05/10/1999	Added IsDigits, HasDuplicatedChars
RF		27/10/1999	Fix to HasDuplicatedChars
DP		31/08/2000	Added debugging functionality.
CL		01/12/2000	CORE000005 Correct duplicate charecters
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		14/05/2007	First .Net version. Ported from GeneralAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using System.Reflection;
using Vertex.Fsd.Omiga.VisualBasicPort.Core;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Assists
{
	/// <summary>
	/// General helper methods.
	/// </summary>
	public class GeneralAssist
	{
		private AppInstance _appInstance;

		/// <summary>
		/// Creates a GeneralAssist object for the calling assembly.
		/// </summary>
		public GeneralAssist()
		{
			_appInstance = new AppInstance(Assembly.GetCallingAssembly());
		}

		/// <summary>
		/// Creates a GeneralAssist object for a specified assembly.
		/// </summary>
		/// <param name="assembly">The assembly.</param>
		public GeneralAssist(Assembly assembly)
		{
			_appInstance = new AppInstance(assembly);
		}

		/// <summary>
		/// Get the version information for the assembly to which this GeneralAssist object refers, in the format "XX.YY.ZZ".
		/// </summary>
		/// <returns>The component version information.</returns>
		public string GetVersion() 
		{
			return _appInstance.Major + "." + _appInstance.Minor + "." + _appInstance.Revision;
		}

		/// <summary>
		/// Gets the comments information for the assembly to which this GeneralAssist object refers.
		/// </summary>
		/// <returns></returns>
		public string GetComments() 
		{
			return _appInstance.Comments;
		}

		/// <summary>
		/// Indicates whether all the characters in a string are alphabetic.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if all characters in the string are alphabetic; otherwise, false.</returns>
		public static bool IsAlpha(string text) 
		{
			bool isAlpha = text.Length > 0;
			for (int index = 0; index < text.Length && isAlpha; index++)
			{
				isAlpha = Char.IsLetter(text, index);
			}
			return isAlpha;
		}

		/// <summary>
		/// Indicates whether a character is categorized as an uppercase letter.
		/// </summary>
		/// <param name="character">A character.</param>
		/// <returns>true if <paramref name="character"/> is an uppercase letter; otherwise, false.</returns>
		public static bool IsUpperAlphaChar(char character) 
		{
			return Char.IsUpper(character);
		}

		/// <summary>
		/// Indicates whether a character is categorized as an lowercase letter.
		/// </summary>
		/// <param name="character">A character.</param>
		/// <returns>true if <paramref name="character"/> is an lowercase letter; otherwise, false.</returns>
		public static bool IsLowerAlphaChar(char character) 
		{
			return Char.IsLower(character);
		}

		/// <summary>
		/// Indicates whether all the characters in a string are numeric.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if all characters in the string are numeric; otherwise, false.</returns>
		public static bool IsDigits(string text) 
		{
			bool isDigits = text.Length > 0;
			for (int index = 0; index < text.Length && isDigits; index++)
			{
				isDigits = Char.IsDigit(text, index);
			}
			return isDigits;
		}

		/// <summary>
		/// Indicates whether any character in a string is uppercase.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if any character in the string is uppercase; otherwise, false.</returns>
		public static bool ContainsUpperAlpha(string text) 
		{
			bool contains = false;
			for (int index = 0; index < text.Length && !contains; index++)
			{
				contains = Char.IsUpper(text, index);
			}
			return contains;
		}

		/// <summary>
		/// Indicates whether any character in a string is lowercase.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if any character in the string is lowercase; otherwise, false.</returns>
		public static bool ContainsLowerAlpha(string text)
		{
			bool contains = false;
			for (int index = 0; index < text.Length && !contains; index++)
			{
				contains = Char.IsLower(text, index);
			}
			return contains;
		}

		/// <summary>
		/// Indicates whether any character in a string is numeric.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if any character in the string is numeric; otherwise, false.</returns>
		public static bool ContainsDigits(string text)
		{
			bool contains = false;
			for (int index = 0; index < text.Length && !contains; index++)
			{
				contains = Char.IsDigit(text, index);
			}
			return contains;
		}

		/// <summary>
		/// Indicates whether any character in a string is special, i.e. non-alpha and non-numeric.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <returns>true if any character in the string is special; otherwise, false.</returns>
		public static bool ContainsSpecialChars(string text)
		{
			bool contains = false;
			for (int index = 0; index < text.Length && !contains; index++)
			{
				contains = !Char.IsLetter(text, index) && !Char.IsDigit(text, index);
			}
			return contains;
		}

		/// <summary>
		/// Checks that all the alphas or digits in a string are the same.
		/// </summary>
		/// <param name="text">The string to check.</param>
		/// <param name="caseSensitive">If true then checking is case sensitive.</param>
		/// <returns>true if there are duplicated characters; otherwise, false.</returns>
		/// <remarks>
		/// If >1 digit, check if digits are all the same. If >1 alpha, check if alphas are all the 
		/// same. E.g., the following strings return true: "AAAa123", "111AAA111", "Q2q345", "Fred11". 
		/// The following strings return false: "ABCDEFG", "1234567".
		/// </remarks>
		public static bool HasDuplicatedChars(string text, bool caseSensitive) 
		{
			bool hasDuplicatedAlphas = false;
			bool hasDuplicatedDigits = false;
			int alphaCount = 0;
			int numberCount = 0;
			char lastAlphaChar = '\0';
			char lastDigitChar = '\0';
			for (int index = 0; index < text.Length && !hasDuplicatedAlphas && !hasDuplicatedDigits; index++)
			{
				char character = text[index];
				if (Char.IsLetter(character))
				{
					alphaCount++;
					if (!caseSensitive)
					{
						character = Char.ToUpper(character);
					}
					if (alphaCount > 1 && character == lastAlphaChar)
					{
						hasDuplicatedAlphas = true;
					}
					lastAlphaChar = character;
				}
				else if (Char.IsDigit(character))
				{
					numberCount++;
					if (numberCount > 1 && character == lastDigitChar)
					{
						hasDuplicatedDigits = true;
					}
					lastDigitChar = character;
				}
			}

			return (hasDuplicatedAlphas && alphaCount > 1) || (hasDuplicatedDigits && numberCount > 1);
		}

		/// <summary>
		/// Converts a string to mixed case.
		/// </summary>
		/// <param name="text">The string to convert to mixed case.</param>
		/// <returns>The converted string.</returns>
		/// <remarks>
		/// The first character of each word will be returned in uppercase. 
		/// The remainder of each word will be returned in lower case.
		/// </remarks>
		public static string ConvertToMixedCase(string text) 
		{
			string mixedCaseText = String.Empty;
			
			// Make everthing consistent.
			text = text.ToLower();

			int spaceIndex = text.IndexOf(' ');
			int startIndex = 0;
			int endIndex = spaceIndex == -1 ? text.Length : spaceIndex;

			// While have spaces.
			while (spaceIndex > -1)
			{
				// Capitalise the first char after the space
				mixedCaseText += text.Substring(startIndex, 1).ToUpper();
				startIndex = startIndex++;
				// everything else to the next space remains lowercase
				mixedCaseText += text.Substring(startIndex + 1, endIndex - startIndex);
				startIndex = spaceIndex + 1;
				// Any more spaces to deal with?
				spaceIndex = text.IndexOf(' ', startIndex);
				endIndex = spaceIndex == -1 ? text.Length : spaceIndex;
			}
			// Make sure we actually have something
			if (endIndex > startIndex)
			{
				mixedCaseText += text.Substring(startIndex, 1).ToUpper();
				startIndex++;
				mixedCaseText += text.Substring(startIndex, endIndex - startIndex);
			}
			return mixedCaseText;
		}
	}
}
