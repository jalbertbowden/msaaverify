' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("MsaaVerify")> 
<Assembly: AssemblyDescription("Microsoft Active Accessibility Verification Tool")> 
<Assembly: AssemblyCompany("Microsoft Corporation")> 
<Assembly: AssemblyProduct("")> 
<Assembly: AssemblyCopyright("")> 
<Assembly: AssemblyTrademark("")> 
<Assembly: CLSCompliant(True)> 

<Assembly: Security.Permissions.SecurityPermission(Security.Permissions.SecurityAction.RequestMinimum)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("F43F15E3-CE96-4498-B975-AEE8C08115D5")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:

<Assembly: AssemblyVersion("1.1.*")> 
<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")> 
<Assembly: ComVisibleAttribute(False)> 