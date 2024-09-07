// DirectOutput.dll bootstrap loader
//
// This demonstrates how a program that's written as a DirectOutput.dll
// client can be located in an arbitrary install folder of its own, and
// still load DirectOutput.dll from the native DOF install folder,
// without requiring a separate copy of DirectOutput.dll.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using System.IO;

namespace DofDllImport
{
	internal class Program
	{
		// This is the internal program main entrypoint.  If you're refactoring
		// an existing program to use this assembly bootstrap scheme, simply
		// rename your existing Main() to InternalMain(), and add the new Main()
		// below.
		static void InternalMain(string[] args)
		{
			Console.WriteLine("Creating DirectOutput.Pinball object");

			var pinball = new DirectOutput.Pinball();
			pinball.Setup();

			Console.WriteLine("Success!");
		}

		// C# program entrypoint.  This is a small bootstrapper function that
		// sets up a custom handler for assembly loading, and invokes the
		// internal main, which contains the real body of the program.
		static void Main(string[] args)
		{
			// Set up an assembly resolver event handler.  This will be invoked
			// when .NET tries to load an assembly from a DLL; it allows us to
			// intercept the load and apply our own special .dll search rules.
			AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolve;

			// Now call the internal main entrypoint.  It's critical that the
			// actual Main() program entrypoint function DOES NOT reference any
			// classes from the DirectOutput assembly, because that would trigger
			// the .NET loader to try to load the referenced assembly before
			// invoking Main(), and thus before we have a chance to set up our
			// assembly load event callback.  The Main() function must thus be
			// restricted to just setting up the event callback, and then calling
			// the separate internal entrypoint function.  This call is where
			// we're going to trigger the DirectOutput.dll load.
			InternalMain(args);
		}

		// Assembly resolver event callback.  .NET calls this when the program
		// tries to load an assembly, giving us a chance to search for the .DLL
		// in custom locations.
		static Assembly AssemblyResolve(object source, ResolveEventArgs args)
		{
			// Check the assembly name to see if it's one we recognize.  args.Name
			// starts with the assembly name, then a comma and the version and
			// language suffixes.  We only care about the assembly name.
			string dllName;
			if (args.Name.StartsWith("DirectOutput,"))
			{
				// load from DirectOutput.dll
				dllName = "DirectOutput.dll";
			}
			else
			{
				// defer to the system loader for other assemblies
				return null;
			}

			// no assembly found yet
			Assembly assembly = null;

			// Try resolving the DOF COM object code base, which should give us
			// a path to the DOF binaries folder.
			var dofComObjectType = Type.GetTypeFromCLSID(new Guid("A23BFDBC-9A8A-46C0-8672-60F23D54FFB6"));
			if (dofComObjectType != null)
			{
				string codeBase = dofComObjectType.Assembly.CodeBase;
				if (codeBase != null && codeBase.StartsWith("file:///"))
				{
					string codeBasePath = Path.GetDirectoryName(codeBase.Substring(8));
					if (TryLoadAssembly(codeBasePath, dllName, out assembly))
						return assembly;
				}
			}

			// DOF COM object lookup failed.  Perhaps we have our own program config
			// file where the user can tell us the location.
			string dofBasePath = @"d:\per\temp\dof222";
			if (dofBasePath != null)
			{
				// With an explicitly configured DOF base path, we should start by
				// looking in the architecture-appropriate subfolder.
				string cpuArch = (sizeof(long) == 8) ? "x64" : "x86";
				string dofPath = Path.Combine(dofBasePath, cpuArch);
				if (TryLoadAssembly(dofPath, dllName, out assembly))
					return assembly;

				// We could also look in the DOF install folder, in case it's the old
				// flat install setup.
				if (TryLoadAssembly(dofBasePath, dllName, out assembly))
					return assembly;
			}

			// Fall back on the current assembly path.  Note: omit this if you
			// explicitly want to bypass a .dll found in the current folder, which
			// might be the case if you're refactoring an existing program, where
			// users might have previously installed a .dll here per older setup
			// instructions.
			var localPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
			if (TryLoadAssembly(localPath, dllName, out assembly))
				return assembly;

			// not found
			return null;
		}

		static bool TryLoadAssembly(string dir, string name, out Assembly assembly)
		{
			// presume failure
			assembly = null;

			// build the full path, and try loading
			var fullPath = Path.Combine(dir, name);
			return (new FileInfo(fullPath).Exists && (assembly = Assembly.LoadFile(fullPath)) != null);
		}

	}
}
