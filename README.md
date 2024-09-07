# DOF DLL Import Demonstration

This is a small demonstration program showing how a C# program can
link to the DirectOutput DLL, with the host program residing in its
own separate folder structure.

Normally, .NET will only load the DLL files for referenced assemblies 
from certain pre-determined locations, primarily the directory where
the main .NET program is running and a set of Windows system folders
where the standard .NET DLLs reside.

This creates installation hassles for programs like DOFLinx that
depend on DirectOutput.dll, but which need to be installed in their
own separate folder tree.  .NET by itself has no way of knowing to
look for DirectOutput.dll in the separate DirectOutput folder tree,
so programs that want to link to it are stuck with one of two
less-than-ideal installation strategies:

1. Install your program directly in the DirectOutput folder, so that
.NET can find DirectOutput.dll when it looks in the program folder.
This is obviously bad for most use cases because it intermingles files
from two different applications in one folder, which can lead to user
confusion later when updating one or the other program.

2. Copy DirectOutput.dll (and any other DOF dependencies) to your
program's install folder.  This isn't as *obviously* bad as the first
option, and in fact, it's not entirely uncommon for programs to
include private copies of dependency DLLs simply to avoid any version
mismatches that result from an external, shared copy of a DLL getting
updated by other programs.  But it *is* problematic for DOF, because
DOF has a bunch of shared resources (configuration file and asset
files) that it depends upon to run properly, and in most cases these
should always come from the central DOF install folder.  The DOF DLLs
use their own DLL file locations as the starting point to look for
these asset and config files, so making a copy of a DOF DLL in another
program's directory will prevent DOF from finding the shared
config/asset files.

This demonstration program provides a solution that's easy to use in
new or existing C# or Visual Basic .NET programs.  It can be easily
retrofitted into an existing program with no changes to the program's
existing logic - it only requires copying and pasting in a little bit
of new code.  

The principle of operation is to use the .NET's assembly-loading
hooks, to override .NET's normal DLL search scheme for the
DirectOutput DLL in particular, loading DirectOutput from the central
DOF install folder.  This eliminates the need for the C# client to
make its own copy of the DOF DLL, while still allowing the client
program to be installed in its own folder structure, with no
connection to the DOF install folder. 

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

3. Find all of the copies of **DirectOutput.dll** in the
Solution Explorer project tree, under Project Name > References.
Select each, go to the Properties pane, and set **Copy Local** to **False**.

4. Find all copies of DirectOutput.dll in your project's Output Path
folder, where your .EXE files are generated.  This is the folder set
in your project properties under **Build > Output Path**, and is
usually called something like **bin\Debug** or **bin\x64\Debug**,
or **bin\Release** or **bin\x64\Release** if you're building in 
release mode.  **Delete** all copies of DirectOutput.dll in these
folders.  If you don't, .NET will just load these and bypass the
custom file search that finds the main Direct Output shared copy
that we're explicitly trying to load.  The custom resolver is only
invoked when .NET **can't** find the file on its own using its
standard search algorithm.


## Visual Basic Version

The DofDllImportVB subfolder has the equivalent code in Visual Basic.
The structure is essentially identical.

## Reference copy of DirectOutput.dll

The project folder tree contains a reference copy of DirectOutput.dll,
in the `References/` folder.   This is used only at **compile time**,
to import the DirectOutput type and function bindings into the project.

This file is **not** copied to the `bin/` folder during the build, and
you should **not** copy it there manually.  It's only there for use at
compile-time.  It's important **not** to include it in the binaries
(.exe) folder, since the whole point of this exercise is to force .NET
to invoke the custom loader code to find the external version of the
DLL in the separate DirectOutput install folder.  If there's a copy of
the DLL the local folder, .NET will just load that, bypassing the
custom event handler and thus failing to accomplish the goal of
loading the shared DirectOutput copy of the file.

Since the DirectOutput.dll copy is only there to import the types
into the main program, you can replace it with any other version of
the library that you wish to use instead for its type imports.
