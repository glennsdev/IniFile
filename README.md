# IniFile
Read Write INI File VB.NET

Microsoft Visual Studio has provide a build-in application configuration file in XML format. But sometimes for any reason we might need an alternative solution. Back to old days, you might consider to use INI file.

Latest release: https://github.com/glennsdev/IniFile/releases/latest

## Ver 1.0

- This DLL libary made with pure VB.NET
- No dependency. NOT using kernel32.dll and no GetPrivateProfileStringA() calls
- No need [section] feature and no plan to supporting it 

## Add to project

How to add reference of IniFile.dll on Microsoft Visual Studio

- From top menu choose **Project** then **Properties**
- Open **References** tab
- Click **Add...** button
- Click **Browse...** button 
- Select **IniFile.dll**

## Basic usage

At first we might need to imports IniFile library

```vb.net
Imports GLib.IniFile
```

Load INI file:
```vb.net
Dim oIniFile As New GLib.IniFile

'-- If file not found it will be created automatically

Dim lLoad As Boolean = oIniFile.LoadFile("TEST.INI")
```

Read two items value:
```vb.net
Dim str1 As String = oIniFile.Items("str1")
Dim str2 As String = oIniFile("str2")
```

Change two items value:
```vb.net
oIniFile.Items("str1") = "ABC"
oIniFile("str2") = "DEF"
```

Save INI file:
```vb.net
Dim lSave As Boolean = oIniFile.SaveFile()
```
