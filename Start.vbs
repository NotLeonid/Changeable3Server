Set s=WScript.CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
script=fso.GetAbsolutePathName(WScript.ScriptFullName)
startup=fso.BuildPath(s.SpecialFolders("Startup"),"Changeable3Server.lnk")
Set shortcut=s.CreateShortcut(startup)
shortcut.TargetPath=script
shortcut.Save
s.Run "cmd /c server.bat", 0, True
