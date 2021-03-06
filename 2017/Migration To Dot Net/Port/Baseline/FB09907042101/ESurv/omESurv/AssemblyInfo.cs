/*
--------------------------------------------------------------------------------------------
Workfile:			AssemblyInfo.cs
Copyright:			Copyright � 2005 Vertex Financial Services

Description:		
--------------------------------------------------------------------------------------------
History:

Prog	Date		Description
JD		28/10/2005	MAR342/MAR366 Continued work. Get config correct
PSC		10/11/2005	MAR510 Change logging configuration
PSC		11/11/2005	MAR510 Use individual logging file
PSC		03/01/2006	MAR961 Remove log4net
--------------------------------------------------------------------------------------------
*/

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.EnterpriseServices;

//
// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
//
[assembly: AssemblyTitle("EPSOM2 Omiga 4")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Vertex Financial Services")]
[assembly: AssemblyProduct("omESurv")]
[assembly: AssemblyCopyright("� 2007")]
[assembly: AssemblyTrademark("� 2007 Vertex Financial Services. Build: FB.099.07.04.21.01  [21 Apr 2007]")]
[assembly: AssemblyCulture("")]		

//
// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Revision and Build Numbers 
// by using the '*' as shown below:

[assembly: AssemblyVersion("1.0.*")]

//
// In order to sign your assembly you must specify a key to use. Refer to the 
// Microsoft .NET Framework documentation for more information on assembly signing.
//
// Use the attributes below to control which key is used for signing. 
//
// Notes: 
//   (*) If no key is specified, the assembly is not signed.
//   (*) KeyName refers to a key that has been installed in the Crypto Service
//       Provider (CSP) on your machine. KeyFile refers to a file which contains
//       a key.
//   (*) If the KeyFile and the KeyName values are both specified, the 
//       following processing occurs:
//       (1) If the KeyName can be found in the CSP, that key is used.
//       (2) If the KeyName does not exist and the KeyFile does exist, the key 
//           in the KeyFile is installed into the CSP and used.
//   (*) In order to create a KeyFile, you can use the sn.exe (Strong Name) utility.
//       When specifying the KeyFile, the location of the KeyFile should be
//       relative to the project output directory which is
//       %Project Directory%\obj\<configuration>. For example, if your KeyFile is
//       located in the project directory, you would specify the AssemblyKeyFile 
//       attribute as [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Delay Signing is an advanced option - see the Microsoft .NET Framework
//       documentation for more information on this.
//

[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile(@"C:\Build Epsom2\2 Int Code\Keys\Mars.snk")]
[assembly: AssemblyKeyName("")]
[assembly: ComVisible(false)]
//MAR812 GHun
[assembly: ApplicationName("OmigaDotNet")]
[assembly: ApplicationActivation(ActivationOption.Library)]
[assembly: ApplicationAccessControl(false)]
//MAR812 End
