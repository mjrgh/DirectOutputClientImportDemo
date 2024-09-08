# DOF DLL Import Demonstration

This is a small demonstration program showing how a C# or Visual Basic
.NET program can link to the shared DirectOutput DLL in the main DOF
install folder, with the host program residing in its own separate
application folder structure, with no relation to the DOF install
folder location.

The challenge with installing a DOF client program in its own
application folder is that .NET normally only looks for referenced DLL
files in the same folder where the application is running (along with
some pre-defined system folders, where the main .NET assemblies
themselves reside).  If the client program resides in its own separate
application folder, .NET won't be able to find DirectOutput.dll when
the program starts, so the program will immediately fail with an
error.  Traditionally, programs faced with this problem have had to
choose between two less-than-ideal installation compromises:

1.  Instead of installing your program in its own application folder,
install your program directly in the DirectOutput folder, so that .NET
can find DirectOutput.dll when your program runs.  This solves the DLL
search problem, but it's obviously bad for most use cases, because it
intermingles files from multiple applications in the Direct Output
folder, which can lead to user confusion later when updating one or
the other program.  If your program consists of *only* an .EXE file,
the commingling problem isn't terrible, and you can ask users to think
of your program as a DOF extension to justify its placement in the
main DOF folder.  But if your program has auxiliary files like
configuration files or asset files, this approach is probably
untenable because of the maintenance confusion it can create for
users.

2. Use your own separate program install folder, and **copy**
DirectOutput.dll (and any other DOF dependencies) to your install
folder.  This isn't as *obviously* bad as the first option, and in
fact, it's not entirely uncommon practice on Windows for programs to
include private copies of dependency DLLs merely to avoid the version
conflicts that can arise from using shared copies of DLLs.  But it
*is* problematic for DOF, because DOF has a bunch of shared resources
(configuration file and asset files) that it depends upon to run
properly, and in most cases these should always come from the central
DOF install folder.  The DOF DLLs use their own DLL file locations as
the starting point to look for these asset and config files, so making
a copy of a DOF DLL in another program's directory will prevent DOF
from finding the shared config/asset files.

This demonstration program provides a third solution, which allows
the client program to be installed in its own folder tree, independent
from the Direct Output install folder, **without** needing its own
copy of DirectOutput.dll.  Instead, the client program loads the shared "main"
copy of DirectOutput.dll from the Direct Output install folder, while
still residing in its own independent folder structure.  This solves
the DLL loading problem without any of the compromises of the two
traditional solutions above.

The technique used in this demo can be easily retrofitted into an
existing C# or Visual Basic .NET program with no changes to the
program's existing logic.  It only requires copying and pasting in a
little bit of new code into the existing program.

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

## Adding this to an existing program

To incorporate this into an existing program:

1. Rename your existing Program.Main() function to InternalMain().

2. Add all of the code from the demo program *other than* its
InternalMain() function, which is just there as an example (and for
testing that the demo program actually does what it purports to).

3. Find all of the copies of **DirectOutput.dll** in the
Solution Explorer project tree, under Project Name > References.
For each one, open its properties (right click -> Properties),
and set **Copy Local** to **False**.

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

## Using for a new program

For a new program, you can simply use the C# or VB project as the
initial program skeleton.  Delete the contents of the InternalMain()
function and write your own startup code there instead.

Note that the projects were generated for .NET Framework 4.8, using
the Console App template.  If you wish to target a different .NET
version, you can usually get Visual Studio to retarget it for you via
the project properties under Application > Target Framework.  If you
prefer to start with a different application template entirely,
generate your starter app as you normally would, and follow the
instructions above for adding the demo code to an existing
application.


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
