' DirectOutput.dll bootstrap loader - VB version
'
' This demonstrates how a program that's written as a DirectOutput.dll
' client can be located in an arbitrary install folder of its own, And
' still load DirectOutput.dll from the native DOF install folder,
' without requiring a separate copy of DirectOutput.dll.


Imports System.IO
Imports System.Reflection
Imports System.Reflection.Assembly

Module Module1

	Sub InternalMain(ByVal cmdArgs() As String)
		Console.WriteLine("Creating DirectOutput.Pinball object")

		Dim pinball = New DirectOutput.Pinball()
		pinball.Setup()

		Console.WriteLine("Success!")
	End Sub

	Sub Main(ByVal cmdArgs() As String)

		' Set up an assembly resolver event handler.  This will be invoked
		' when .NET tries to load an assembly from a DLL; it allows us to
		' intercept the load And apply our own special .dll search rules.
		AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf AssemblyResolve

		' Now call the internal main entrypoint.  It's critical that the
		' actual Main() program entrypoint function DOES Not reference any
		' classes from the DirectOutput assembly, because that would trigger
		' the .NET loader to try to load the referenced assembly before
		' invoking Main(), and thus before we have a chance to set up our
		' assembly load event callback.  The Main() function must thus be
		' restricted to just setting up the event callback, and then calling
		' the separate internal entrypoint function.  This call is where
		' we're going to trigger the DirectOutput.dll load.
		InternalMain(cmdArgs)

	End Sub

	Private Function AssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly

		' Check the assembly name to see if it's one we recognize
		Dim dllName As String
		If args.Name.StartsWith("DirectOutput,") Then
			' load from DirectOutput.dll
			dllName = "DirectOutput.dll"
		Else
			' defer to the system loader for other assemblies
			Return Nothing
		End If

		' no assembly found yet
		Dim result As Assembly = Nothing

		' Try resolving the DOF COM object code base, which should give us
		' a path to the DOF binaries folder.
		Dim dofComObjectType = Type.GetTypeFromCLSID(New Guid("A23BFDBC-9A8A-46C0-8672-60F23D54FFB6"))
		If dofComObjectType IsNot Nothing Then
			Dim codeBase As String = dofComObjectType.Assembly.CodeBase
			If (codeBase IsNot Nothing And codeBase.StartsWith("file:///")) Then

				Dim codeBasePath = Path.GetDirectoryName(codeBase.Substring(8))
				If TryLoadAssembly(codeBasePath, dllName, result) Then
					Return result
				End If
			End If
		End If

		' DOF COM object lookup failed.  Perhaps we have our own program config
		' file where the user can tell us the location.
		Dim dofBasePath As String = "d:\per\temp\dof222"
		If dofBasePath IsNot Nothing Then

			' With an explicitly configured DOF base path, we should start by
			' looking in the architecture-appropriate subfolder.
			Dim cpuArch As String
			Dim testInt As Long
			If Len(testInt) = 8 Then cpuArch = "x64" Else cpuArch = "x86"
			Dim dofPath As String = Path.Combine(dofBasePath, cpuArch)
			If TryLoadAssembly(dofPath, dllName, result) Then
				Return result
			End If
		End If

		' We could also look in the DOF install folder, in case it's the old
		' flat install setup.
		If TryLoadAssembly(dofBasePath, dllName, result) Then
			Return result
		End If

		' Fall back on the current assembly path.  Note omit this if you
		' explicitly want to bypass a .dll found in the current folder, which
		' might be the case if you're refactoring an existing program, where
		' users might have previously installed a .dll here per older setup
		' instructions.
		Dim localPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
		If TryLoadAssembly(localPath, dllName, result) Then
			Return result
		End If

		' Not found
		Return Nothing

	End Function


	Private Function TryLoadAssembly(dir As String, name As String, ByRef result As Assembly) As Boolean

		' presume failure
		result = Nothing

		' build the full path, and try loading
		Dim fullPath = Path.Combine(dir, name)
		Dim fullPathInfo = New FileInfo(fullPath)
		If Not fullPathInfo.Exists Then Return False
		result = Assembly.LoadFile(fullPath)
		Return result IsNot Nothing
	End Function

End Module
