Option Explicit

Const DOTNET_RELEASE_KEY = "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\Release"
Const DOTNET_MIN_RELEASE = 461808
Const DOTNET_INSTALLER = "NDP472-KB4054531-Web.exe"

Dim shell: Set shell = CreateObject("WScript.Shell")
Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
Dim appli: Set appli = CreateObject("Shell.Application")
Dim release: release = 0

On Error Resume Next
release: release = shell.RegRead(DOTNET_RELEASE_KEY)
On Error Goto 0

If release < DOTNET_MIN_RELEASE Then
    shell.CurrentDirectory = fso.getParentFolderName(WScript.ScriptFullName)
    appli.ShellExecute DOTNET_INSTALLER, "", "", "runas"
End If
