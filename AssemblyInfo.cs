using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("mca64Inventor AddIn Template")]
[assembly: AssemblyDescription("AddIn template for Autodesk Inventor.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("YourCompany")]
[assembly: AssemblyProduct("mca64Inventor AddIn")]
[assembly: AssemblyCopyright("Copyright © 2024")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components. If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(true)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("D0A4FCF1-B4D0-4824-894A-7FA419AF92F8")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
[assembly: AssemblyVersion("1.0.0.*")]

// To sign your assembly, you must specify a key to use. Refer to the Microsoft .NET Framework documentation for more information on assembly signing.
// Use the attributes below to control which key is used for signing.
// Notes:
//   (*) If no key is specified, the assembly is not signed.
//   (*) KeyName refers to a key installed in the Crypto Service Provider (CSP) on your machine.
//   (*) KeyFile refers to a file containing a key.
//   (*) If the KeyFile and KeyName values are specified, the following occurs:
//       (1) If the KeyName can be found in the CSP, that key is used.
//       (2) If the KeyName does not exist and the KeyFile does exist, the key in the KeyFile is installed into the CSP and used.
//   (*) To create a KeyFile, use the sn.exe (Strong Name) utility.
//       When specifying the KeyFile, the location of the file should be relative to the project output directory, 
//       e.g., %Project Directory%\obj\<configuration>. If the KeyFile is in your project directory, 
//       use [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
//   (*) Delay Signing is an advanced option - see the .NET Framework documentation for more information.
[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile("")]
[assembly: AssemblyKeyName("")]