# DOF DLL Import Demonstration

This is a small demonstration program showing how a C# program can
link to the DirectOutput DLL, with the host program residing in its
own separate folder structure.

Normally, C# requires referenced assemblies to come from one of the
system locations or from the same folder containing the main program
assembly.  This creates installation hassles for programs like DOFLinx
that depend on DirectOutput.dll, but need to be installed in their own
separate program folder structure.  There are two straightforward, but
less-than-ideal solutions that have been used in the past:

1. Install the program in the DirectOutput folder.  This is obviously
bad because it intermingles files from two programs in one folder,
which can lead to user confusion later when updating one or the other
program.

2. Copy DirectOutput.dll (and any other DOF dependencies) to the other
program's install folder.  This is somewhat less obviously bad than
the first option, and in fact, it's not entirely uncommon for programs
to include private copies of dependency DLLs, specifically to avoid
the "DLL hell" of incompatible version conflicts that often happen
when multiple programs try to share a common central copy of a given
DLL.  But it *is* problematic for DOF, because DOF has a bunch of
shared resources (configuration file and asset files) that it depends
upon to run properly, and in most cases these should always come from
the central DOF install folder.  The DOF DLLs use their own DLL file
locations as the starting point to look for asset files, so making a
copy of a DOF DLL in another program's directory will prevent DOF from
finding the shared config/asset files.

This demonstration program provides a solution that's easy to use in
new or existing C# programs.  It can be easily retrofitted into an
existing C# program with minimal changes.  The principle of operation
is to use the assembly-loading hooks that .NET provides, to override
.NET's normal DLL search scheme for the DirectOutput DLL in
particular, loading DirectOutput from the central DOF install folder.
This eliminates the need for the C# client to make its own copy of the
DOF DLL, while still allowing the client program to be installed in
its own folder structure, with no connection to the DOF install
folder. 

The demonstration program shows two ways of finding DOF.  The first is
to look up the registered DOF COM object.  If found, the object's .NET
"code base" path is used, since this gives us the location of the
globally registered DLL implementing the COM object.  This is the most
automatic way, requiring no manual user configuration steps (other
than setting up DOF in the normal way, which is obviously required
before the client program can use DOF anyway) and no extra work in the
calling program.  The second way is to use an explicit path specified
by the client program, such as from the program's own configuration
file.  This is less automatic than the DOF COM object lookup, but it
gives the client program full control over where the DLL is loaded
from.


## Using in your program

To incorporate this into an existing program:

1. Rename your existing Program.Main() function to InternalMain().

2. Add all of the code from the demo program *other than* its
InternalMain() function, which is just there as an example (and for
testing that the demo program actually does what it purports to).

## Visual Basic Version

The DofDllImportVB subfolder has the equivalent code in Visual Basic.
The structure is essentially identical.
