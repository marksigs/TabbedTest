/*
--------------------------------------------------------------------------------------------
Workfile:			GlobalProperty.cs
Copyright:			Copyright © 2007 Marlborough Stirling
Description:		GlobalProperty object.
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
AS		09/05/2007	First version.
--------------------------------------------------------------------------------------------
*/
using System;
using System.EnterpriseServices;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Vertex.Fsd.Omiga.VisualBasicPort.Core
{
	/// <summary>
	/// The GlobalProperty class defines methods and properties that are accessible to the whole application.
	/// </summary>
	public static class GlobalProperty
	{
		private static string _databaseConnectionString;
		private static int _databaseRetries = 1;
		private static string _databaseProvider;

		private static SharedPropertyGroupManager _sharedPropertyGroupManager;
		private static SharedPropertyGroup _sharedPropertyGroup;
		private static bool _useSharedProperties = true;

		private static SharedProperty GetSharedProperty(string property)
		{
			SharedProperty sharedProperty = null;
			const string sharedPropertyGroupName = "OmigaVisualBasicPort";

			try
			{
				if (_sharedPropertyGroupManager == null)
				{
					_sharedPropertyGroupManager = new SharedPropertyGroupManager();
				}
		
				if (_sharedPropertyGroupManager != null)
				{
					bool exists = false;
					if (_sharedPropertyGroup == null)
					{
						PropertyLockMode propertyLockMode = PropertyLockMode.SetGet;
						PropertyReleaseMode propertyReleaseMode = PropertyReleaseMode.Process;
						_sharedPropertyGroup = _sharedPropertyGroupManager.CreatePropertyGroup(sharedPropertyGroupName, ref propertyLockMode, ref propertyReleaseMode, out exists);
					}
					if (_sharedPropertyGroup != null)
					{
						sharedProperty = _sharedPropertyGroup.CreateProperty(property, out exists);
					}
				}
			}
			catch (Exception e)
			{
				throw new InvalidOperationException("Unable to create SharedPropertyGroup: " + sharedPropertyGroupName + ".", e);
			}

			return sharedProperty;
		}

		/// <summary>
		/// The GetSharedPropertyValue method retrieves the value for a shared property.
		/// </summary>
		/// <param name="property">The shared property whose value should be retrieved.</param>
		/// <returns>The value of the shared property.</returns>
		/// <exception cref="InvalidOperationException">An <see cref="Exception"/> was thrown.</exception>
		/// <remarks>
		/// All shared properties are stored in a shared property group called "OmigaVisualBasicPort", using 
		/// the <see cref="System.EnterpriseServices.SharedPropertyGroupManager"/> type. There is some debate as 
		/// to whether the SharedPropertyGroupManager should be used at all because of its relatively
		/// poor locking performance, however it has always been used in Omiga 4 Visual Basic 6 code, 
		/// and therefore continues to be used here.
		/// </remarks>
		public static object GetSharedPropertyValue(string property)
		{
			object propertyValue = null;

			if (_useSharedProperties)
			{
				try
				{
					SharedProperty sharedProperty = GetSharedProperty(property);
					if (sharedProperty != null)
					{
						propertyValue = sharedProperty.Value;
					}
				}
				catch (Exception e)
				{
					throw new InvalidOperationException("Unable to get SharedProperty.Value: " + property + ".", e);
				}
			}
				
			return propertyValue;
		}

		/// <summary>
		/// The SetSharedPropertyValue method sets a value for a shared property.
		/// </summary>
		/// <param name="property">The shared property whose value should be set.</param>
		/// <param name="propertyValue">The value to which the shared property should be set.</param>
		/// <exception cref="InvalidOperationException">An <see cref="Exception"/> was thrown.</exception>
		/// <remarks>
		/// All shared properties are stored in a shared property group called "OmigaVisualBasicPort", using 
		/// the <see cref="System.EnterpriseServices.SharedPropertyGroupManager"/> type. There is some debate as 
		/// to whether the SharedPropertyGroupManager should be used at all because of its relatively
		/// poor locking performance, however it has always been used in Omiga 4 Visual Basic 6 code, 
		/// and therefore continues to be used here.
		/// </remarks>
		public static void SetSharedPropertyValue(string property, object propertyValue)
		{
			if (_useSharedProperties)
			{
				try
				{
					SharedProperty sharedProperty = GetSharedProperty(property);
					if (sharedProperty != null)
					{
						sharedProperty.Value = propertyValue;
					}
				}
				catch (Exception e)
				{
					throw new InvalidOperationException("Unable to set SharedProperty.Value: " + property + ".", e);
				}
			}
		}		

		/// <summary>
		/// The UseSharedProperties property determines whether shared properties should be used.
		/// </summary>
		/// <remarks>
		/// If true, then shared properties will be used; defaults to <b>true</b>.
		/// </remarks>
		public static bool UseSharedProperties
		{
			get { return _useSharedProperties; }
			set { _useSharedProperties = value; }
		}

		/// <summary>
		/// The DatabaseConnectionString property holds the application's connection string for 
		/// connecting to the database.
		/// </summary>
		/// <exception cref="InvalidOperationException">An <see cref="Exception"/> was thrown.</exception>
		/// <remarks>
		/// Historically an Omiga 4 application hold its database connection string in the registry, 
		/// whereas for .Net applications Microsoft recommend that the connection string is held in the
		/// app.config or web.config file. The DatabaseConnectionString property hides this 
		/// implementation detail from the caller, and takes care of retrieving the connection string
		/// from the registry (or the .config file). Currently only registry connection strings are 
		/// supported.
		/// <para>
		/// The database connection string is held in the registry key 
		/// HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\DATABASE CONNECTION. It can either be defined by a single 
		/// entry, ConnectionString (in standard ADO .Net connection string format), 
		/// or split up into the entries Server, Database Name, User Id, and Password.
		/// </para>
		/// </remarks>
		public static string DatabaseConnectionString
		{
			get
			{
				if (_databaseConnectionString == null || _databaseConnectionString == "0")
				{
					object propertyValue = GetSharedPropertyValue("DatabaseConnectionString");

					if (propertyValue != null)
					{
						_databaseConnectionString = propertyValue.ToString();
					}

					if (_databaseConnectionString == null || _databaseConnectionString == "0")
					{
						// Get Omiga 4 style connection string from the registry.
						try
						{
							RegistryKey hklm = Registry.LocalMachine;
							RegistryKey hkSoftware = hklm.OpenSubKey("Software");
							RegistryKey hkOmiga4 = hkSoftware.OpenSubKey("Omiga4");
							RegistryKey hkDatabaseConnection = hkOmiga4.OpenSubKey("DATABASE CONNECTION");

							object registryConnectionString = hkDatabaseConnection.GetValue("ConnectionString");
							string connectionString = (registryConnectionString != null) ? registryConnectionString.ToString() : String.Empty;

							if (connectionString.Length == 0)
							{
								object registryUserId = hkDatabaseConnection.GetValue("User ID");
								string userId = (registryUserId != null) ? registryUserId.ToString() : String.Empty;

								// Standard Security
								if (userId.Length > 0)
								{
									_databaseConnectionString =
										"Server=" + hkDatabaseConnection.GetValue("Server") + ";" +
										"Database=" + hkDatabaseConnection.GetValue("Database Name") + ";" +
										"User ID=" + userId + ";" +
										"Password=" + hkDatabaseConnection.GetValue("Password") + ";";
								}
								else // Integrated Security
								{
									_databaseConnectionString =
										"Server=" + hkDatabaseConnection.GetValue("Server") + ";" +
										"Database=" + hkDatabaseConnection.GetValue("Database Name") + ";" +
										"Integrated Security=SSPI;Persist Security Info=False";
								}
							}

							_databaseProvider = "SQLOLEDB";

							object registryRetries = hkDatabaseConnection.GetValue("Retries");
							if (registryRetries != null)
							{
								_databaseRetries = Convert.ToInt32(registryRetries.ToString());
							}
						}
						catch (Exception e)
						{
							throw new InvalidOperationException("Unable to read the database connection string from the registry.", e);
						}
					}
				}

				return _databaseConnectionString;
			}
			set
			{
				_databaseConnectionString = value;
				SetSharedPropertyValue("DatabaseConnectionString", _databaseConnectionString);
			}
		}

		/// <summary>
		/// The database provider as defined by the registry key 
		/// HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\DATABASE CONNECTION\Provider.
		/// </summary>
		public static string DatabaseProvider
		{
			get { return _databaseProvider; }
		}

		/// <summary>
		/// The number of times to retry a database operation, as defined by the registry key
		/// HKEY_LOCAL_MACHINE\SOFTWARE\OMIGA4\DATABASE CONNECTION\Retries.
		/// </summary>
		public static int DatabaseRetries
		{
			get { return _databaseRetries; }
		}
	}
}
