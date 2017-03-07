/*
--------------------------------------------------------------------------------------------
Workfile:			RegAssist.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		Registry access helper object.
--------------------------------------------------------------------------------------------
.Net Specific History:

Prog	Date		Description
AS		21/06/2007	First .Net version. Ported from RegAssist.cls.
--------------------------------------------------------------------------------------------
*/
using System;
using Microsoft.Win32;

// Missing XML comment for publicly visible type or member.
#pragma warning disable 1591

namespace Vertex.Fsd.Omiga.VisualBasicPort.Obsolete
{
	[Obsolete("Use the Microsoft.Win32.RegistryKey type instead")]
    public static class RegAssist
    {
		[Obsolete("Use Microsoft.Win32.RegistryKey.SetValue instead")]
		public static void SetValueEx(RegistryKey registryKey, string valueName, object value)
		{
			registryKey.SetValue(valueName, value);
		}

		[Obsolete("Use Microsoft.Win32.RegistryKey.SetValue instead")]
		public static void SetValueEx(RegistryKey registryKey, string valueName, RegistryValueKind registryValueKind, object value)
		{
			registryKey.SetValue(valueName, value, registryValueKind);
		}

		[Obsolete("Use Microsoft.Win32.RegistryKey.OpenSubKey and Microsoft.Win32.RegistryKey.GetValue instead")]
		public static object QueryValue(RegistryKey parentRegistryKey, string keyName, string valueName)
		{
			object value = null;

			using (RegistryKey registryKey = parentRegistryKey.OpenSubKey(keyName))
			{
				QueryValueEx(registryKey, valueName, out value);
			}

			return value;
		}

		[Obsolete("Use Microsoft.Win32.RegistryKey.GetValue instead")]
		public static void QueryValueEx(RegistryKey registryKey, string valueName, out object value) 
        {
			value = registryKey.GetValue(valueName);
        }

		[Obsolete("Use Microsoft.Win32.RegistryKey.CreateSubKey and Microsoft.Win32.RegistryKey.SetValue instead")]
		public static void SetKeyValue(RegistryKey parentRegistryKey, string keyName, string valueName, object value)
		{
			using (RegistryKey registryKey = parentRegistryKey.CreateSubKey(keyName))
			{
				SetValueEx(registryKey, valueName, value);
			}
		}

		[Obsolete("Use Microsoft.Win32.RegistryKey.CreateSubKey and Microsoft.Win32.RegistryKey.SetValue instead")]
		public static void SetKeyValue(RegistryKey parentRegistryKey, string keyName, string valueName, object value, RegistryValueKind registryValueKind) 
        {
			using (RegistryKey registryKey = parentRegistryKey.CreateSubKey(keyName))
			{
				SetValueEx(registryKey, valueName, registryValueKind, value);
			}
        }

		[Obsolete("Use Microsoft.Win32.RegistryKey.CreateSubKey instead")]
		public static void CreateNewKey(string keyName, RegistryKey parentRegistryKey) 
        {
			using (RegistryKey registryKey = parentRegistryKey.CreateSubKey(keyName))
			{
			}
        }

		[Obsolete("Use Microsoft.Win32.RegistryKey.DeleteSubKey instead")]
		public static void DeleteKey(RegistryKey registryKey, string keyName) 
        {
			registryKey.DeleteSubKey(keyName);
        }
    }
}
