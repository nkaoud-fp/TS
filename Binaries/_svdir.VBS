
'	Copyright (c) Microsoft Corporation
'
'	_svdir.vbs
'
'	utility to manage VDirs for Soap Toolkit 3
'	for usage information  call it with the HELP option 
'
'
Option Explicit
'On Error Resume Next

Dim objArgs

const TargetServer 	= "IIS://localhost"
const DefaultSiteName	= "w3svc/1/Root"
const  ScriptMap	= ",1,GET,POST"
const  ScriptMapIIS4	= ",1,PUT,DELETE,TRACE"

'error codes returned from ScriptMap
'script returns 0 in the success case
const NO_EXECUTION 	= 1		' no changes performed
const PARAM_MISSING 	= 2		' missing parameter
const CSCRIPT_REQUIRED	= 3		' script was started from wscript.exe
const UNKNOWN_PARAM 	= 4		' unknown command line parameter
const ISAPI_REG 	= 5		' soap isapi not correctly registered
const ILLEGAL_PARAM 	= 6		' illegal command line parameter
const VDIR_ERROR	= 7		' failure during vdir creation
const IIS_VERSION	= 8		' unable to determine IIS version information
const SETMAP_ERROR	= 9		' Failure setting the scriptmap
const GETOBJ_ERROR	= 9		' Failure sto acquire an IIS object


call DetectExeType

Set objArgs = WScript.Arguments
if objArgs.Count < 1 then
	call DisplaySimpleHelp
	WScript.Quit(PARAM_MISSING)
end if

if ucase(objArgs(0)) = "HELP" then
	call DisplayHelp
	WScript.Quit(NO_EXECUTION)
end if

If ucase(objArgs(0)) = "UPDATE" Then
	call VDirUpdate
	WScript.Quit(0)
end if

If ucase(objArgs(0)) = "CREATE" Then
	call VDirCreate
	WScript.Quit(0)
end if


'we can quit with an error message
call DisplayError("Unknown Paramter ("+objArgs(0)+")", UNKNOWN_PARAM, 0)





function VDirUpdate
	Dim args
	Dim ServiceNumber
	Dim VDirName
	Dim SiteName
	Dim ISAPILocation
	Dim  i

	On Error Resume Next

	VDirCommand = ""
	ServiceNumber = ""

	Set args = WScript.Arguments
	if (args.Count > 4) or (args.Count < 2) then
		call DisplayError("Illegal number of parameters.", ILLEGAL_PARAM, 0)
	end if

	VDirName = args(1)

	for i = 2 to args.Count -1
		if	ucase(left(args(i),3)) = "-S:" then
		WScript.Echo args(i)
			if (len(ServiceNumber) > 0) then
				call DisplayError("Illegal service number ("+args(i)+")", ILLEGAL_PARAM, 0)
			end if
			ServiceNumber = Right(args(i), Len(args(i)) - 3)
		end if
		if	ucase(left(args(i),3)) = "-W:" then
			if (len(SiteName) > 0) then
				call DisplayError("Illegal website ("+args(i)+")", ILLEGAL_PARAM, 0)
			end if
			SiteName = Right(args(i), Len(args(i)) - 3)
		end if
	next

	if len(ServiceNumber) = 0 then
		ServiceNumber="1"
	end if
	if len(SiteName) = 0 then
		SiteName="Root"
	end if

	VDirName = CreateVDirName(TargetServer, ServiceNumber, SiteName, VDirName)
	ISAPILocation = DetectIsapiLocation()

	call addMapping( VDirName, ISAPILocation)
end function


function VDirCreate
	Dim args
	Dim ServiceNumber
	Dim VDirName
	Dim SiteName
	Dim VDirPath
	DIM ISAPILocation
	Dim i
	'On Error Resume Next

	VDirCreate= ""

	Set args = WScript.Arguments
	if (args.Count > 5) or (args.Count < 3) then
		call DisplayError("Illegal number of parameters.", ILLEGAL_PARAM, 0)
	end if

	VDirName = args(1)
	VDirPath = args(2)

	for i = 3 to args.Count -1
		if	ucase(left(args(i),3)) = "-S:" then
		WScript.Echo args(i)
			if (len(ServiceNumber) > 0) then
				call DisplayError("Illegal service number ("+args(i)+")", ILLEGAL_PARAM, 0)
			end if
			ServiceNumber = Right(args(i), Len(args(i)) - 3)
		end if
		if	ucase(left(args(i),3)) = "-W:" then
			if (len(SiteName) > 0) then
				call DisplayError("Illegal website ("+args(i)+")", ILLEGAL_PARAM, 0)
			end if
			SiteName = Right(args(i), Len(args(i)) - 3)
		end if
	next

	if len(ServiceNumber) = 0 then
		ServiceNumber="1"
	end if
	if len(SiteName) = 0 then
		SiteName="Root"
	end if

	call createVDir( TargetServer, ServiceNumber, SiteName, VDirName, VDirPath)

	ISAPILocation = DetectIsapiLocation()
	VDirName = CreateVDirName(TargetServer, ServiceNumber, SiteName, VDirName)

	call addMapping( VDirName, ISAPILocation)
end function


Function createVDir(Server, ServiceNumber, SiteName, VDirName, VDirPath)
	'On Error Resume Next

	createVDir = TRUE

	Dim IIsObjectPath
	Dim IIsObject
	Dim vroot
	Dim vDir

	IIsObjectPath = Server + "/W3SVC/" + ServiceNumber
	Set IIsObject = GetObject(IIsObjectPath)
	call CHK_ERR("Error trying to get IIS Object: " + IIsObjectPath, GETOBJ_ERROR, Err.Number)

	set vRoot = IIsObject.GetObject("IIsWebVirtualDir", "Root")
	call CHK_ERR ("Error trying to access IISWebVirtualDirectory 'Root' on: " + IIsObjectPath, GETOBJ_ERROR, Err.Number)

	Set vDir = vRoot.Create("IIsWebVirtualDir", VDirName)
	call CHK_ERR("Unable to create " + vRoot.ADsPath + "/" + VDirName , GETOBJ_ERROR, Err.Number)

	vDir.AccessRead = False
	vDir.AccessExecute = True
	vDir.AspBufferingOn= True
	vDir.AccessScript = True
	vDir.AspAllowSessionState = False
	vDir.ContentIndexed = False
	vDir.Put "Path", VDirPath
	call CHK_ERR("Can't set properties for " + vRoot.ADsPath + "/" + VDirName, VDIR_ERROR, Err.Number)

	if IsIIS4() then
		Const INPROC_IIS4 = True
		Const OUTPROC_IIS4 = False
	
		vDir.AppCreate INPROC_IIS4
	else
		Const INPROC = 0
		Const OUTPROC = 1
 		Const POOLED = 2

		vdir.AppCreate2 INPROC
	end if
	call CHK_ERR("Can't create web-application definition in " + vRoot.ADsPath + "/" + VDirName, VDIR_ERROR, Err.Number)

	vDir.AppFriendlyName = VDirName
	vDir.SetInfo
	call CHK_ERR("Unable to save configuration for " + vRoot.ADsPath + "/" + VDirName, VDIR_ERROR, Err.Number)

End Function




function CreateVDirName(server, number, site, vdir)
	Dim temp
	On Error Resume Next

	temp = server + "/w3svc/" +number + "/" + site + "/" + vdir
	CreateVDirName = temp
end function


Function addMapping(IIsObjectPath, isapiname)
	On Error Resume Next

	Dim IIsObject
	dim i
	dim list
	dim mapString
	dim doneIt

	addMapping = ""
	doneIt = false

	Set IIsObject = GetObject(IIsObjectPath)
	call CHK_ERR("Error trying to get the path of the application: " + IIsObjectPath, GETOBJ_ERROR, Err.Number)

	list = IISObject.ScriptMaps
	call CHK_ERR("Error trying to query scriptmap in: " + IIsObjectPath, GETOBJ_ERROR, Err.Number)

	if IsIIS4() then
		mapString = ".wsdl," + isapiname + ScriptMapIIS4
	else
		mapString = ".wsdl," + isapiname + ScriptMap
	end if

	for i = 0 to UBound(list)
		if (left(list(i), 5) = ".wsdl") then
			list(i) = mapString
			doneIt = true
		end if
	next

	if not doneIt then
		i = ubound(list) + 1
		redim preserve list(i)
		list(i)= mapString
	end if

	IIsObject.Put "scriptmaps", (list)
	IIsObject.SetInfo

	call CHK_ERR("Error trying to set the scriptmap to " + mapString, SETMAP_ERROR, Err.Number)

	WScript.Echo "Registered virtual DIR: "
	WScript.Echo "  " + IIsObjectPath
	WScript.Echo " with addition scriptmap entry:"
	WScript.Echo "  " + mapString
	WScript.Echo

	WScript.Quit(0)
End Function


 '	This can detect the type of exe the script is running under and warns the
' 	user of the popups.
Sub DetectExeType()
	Dim ScriptHost
	On Error Resume Next

	ScriptHost = WScript.FullName
	ScriptHost = Right(ScriptHost, Len(ScriptHost) - InStrRev(ScriptHost, "\"))
	If (UCase(ScriptHost) = "WSCRIPT.EXE") Then
		Dim msg

		msg = 				"This script is executed using WScript. It requires CScript to function correctly."
		msg = AddLine(msg,	"Please run the script by executing")
		msg = AddLine(msg,	"       SOAPVDIR.CMD <cmd> <param>")

		WScript.Echo msg

		WScript.Quit (CSCRIPT_REQUIRED)
	End If
End Sub

function IsIIS4
	Dim ShellObject
	dim mver
	On Error Resume Next

	IsIIS4 = FALSE
	Set ShellObject = WScript.CreateObject("WScript.Shell")

	mver = ShellObject.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\w3svc\Parameters\MajorVersion")

	if (Err <> 0) or (len(mver) < 1) then
		call DisplayError("Can't determine IIS Version", IIS_VERSION, Err.Number)
	End If
	if CInt(mver) < 5 then
		IsIIS4 = TRUE
	end if
end function

function DetectIsapiLocation
	Dim ShellObject
	Dim location
	On Error Resume Next

	Set ShellObject = WScript.CreateObject("WScript.Shell")

	location = ShellObject.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSSOAP\30\SOAPISAP\isapi")
	if (Err <> 0) or (len(location) < 1) then
		call DisplayError("Soap Toolkit 3 Isapi is not correctly registered.", ISAPI_REG, Err.Number)
	End If

	DetectIsapiLocation  = location
end function

Function CommonHelp
	WScript.echo
	WScript.echo	"Usage:"
	WScript.echo	"           SOAPVDIR.CMD <cmd> [<value>*]"
	WScript.echo
	WScript.echo	"Description:"
	WScript.echo	"  Create VDir on localhost to use Soap Toolkit 3 ISAPI or"
	WScript.echo	"  Update VDir on localhost to use Soap Toolkit 3 ISAPI instead of"
	WScript.echo	"  the Soap Toolkit 2 ISAPI"
	WScript.echo
	CommonHelp=""
end function

Function DisplaySimpleHelp
	call CommonHelp()
	WScript.echo	"Supported Commands:"
	WScript.echo	"  HELP, UPDATE, CREATE"
	WScript.echo
	WScript.echo	"Sample"
	WScript.echo	"  soapvdir.cmd UPDATE soapsample"
	WScript.echo	"  soapvdir.cmd CREATE soapsample c:\inetpub\wwwroot\soap"
	WScript.echo
	WScript.echo	"For Extended Help type:"
	WScript.echo	"  soapvdir.cmd HELP"
	WScript.echo
	DisplaySimpleHelp=0
end function


Function DisplayHelp
	call CommonHelp()
	WScript.echo	"Standard Command:"
	WScript.echo	"  Displays this help message"
	WScript.echo	"    soapvdir.cmd HELP "
	WScript.echo
	WScript.echo	"  Updates the VDIR 'name'"
	WScript.echo	"   soapvdir.cmd UPDATE name [-s:server] [-w:site]"
	WScript.echo
	WScript.echo	"  Create a new Vdir"
	WScript.echo	"   soapvdir.cmd CREATE name path [-s:server] [-w:site]"
	WScript.echo
	WScript.echo	"UPDATE Command:"
	WScript.echo	"  'name' is the name of the virtual directory to update."
	WScript.echo	"  By default it is expected that this virtual directory is located"
	WScript.echo	"  on the root (w3svc/1/root)."
	WScript.echo	"  The default server (1) can be changed using the '-s:' option."
	WScript.echo	"  The default site (root) can be changed using the '-w:' option."
	WScript.echo
	WScript.echo	"CREATE Command:"
	WScript.echo	"  'name' is the name of the virtual directory to create."
	WScript.echo	"  'path' is the location on disk."
	WScript.echo	"  By default the virtual directory is created below the"
	WScript.echo	"  root (w3svc/1/root)."
	WScript.echo	"  The default server (1) can be changed using the '-s:' option."
	WScript.echo	"  The default site (root) can be changed using the '-w:' option."
	WScript.echo
	WScript.echo	"Samples:"
	WScript.echo	"  soapvdir.cmd UPDATE soapsample"
	Wscript.echo	"    is identical to"
	WScript.echo	"  soapvdir.cmd UPDATE soapsample -s:1 -w:root"
	WScript.echo
	WScript.echo	"  soapvdir.cmd CREATE soapsample c:\soap"
	Wscript.echo	"    is identical to"
	WScript.echo	"  soapvdir.cmd CREATE soapsample c:\soap -s:1 -w:root"
	WScript.echo
	DisplayHelp=0
end function


function CHK_ERR(msg, code, errcode)
	on error resume next
	CHK_ERR=FALSE

	if errcode <> 0 then
		call DisplayError(msg, code, errcode)
	end if
	CHK_ERR = TRUE
end function

function DisplayError(msg, code, errcode)
	DisplayError = ""
	on error resume next

	WScript.Echo
	if errcode <> 0 then
		WScript.Echo "Error (0x" + hex(errcode) + "):"
	else
		WScript.Echo "Error:"
	end if
	WScript.Echo "  " + msg
	WScript.Echo
	WScript.Quit(UNKNOWN_PARAM)
end function

' concatenates two string with a return in the middle
' this is usefull for output in msg-boxes
function AddLine(msg, line)
	on error resume next
	Addline = msg + chr(13) + line
end function
