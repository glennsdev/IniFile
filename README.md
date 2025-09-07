# Read Write INI File VB.NET

Microsoft Visual Studio provides a built-in application configuration file in XML format. However, for simpler use cases, an INI file might be a more straightforward solution.

Download latest release (compiled DLL): https://github.com/glennsdev/IniFile/releases/download/1.2/IniFile.dll.zip

## Ver 1.2

- This DLL libary made with VB.NET
- No dependency. NOT using kernel32.dll and no GetPrivateProfileStringA() calls
- No need [section] feature and no future plan to supporting it
- Supporting remark ";" sign

## Add to project

How to add reference of IniFile.dll on Microsoft Visual Studio

- From top menu choose **Project** then **Properties**
- Open **References** tab
- Click **Add...** button
- Click **Browse...** button 
- Select **IniFile.dll**

## Basic usage

At first we might need to imports IniFile library:
```vb.net
Imports GLib.IniFile
```

Load INI file:
```vb.net
Dim oIniFile As New GLib.IniFile

Dim lLoad As Boolean 

'-- Load TEST.INI file on current EXE directory
lLoad = oIniFile.LoadFile("TEST.INI")

If (lLoad = False) Then
  '-- File not found
End If

'-- Load TEST.INI file on "D:\myapp" directory
lLoad = oIniFile.LoadFile("D:\myapp\SAMPLE.INI")

If (lLoad = False) Then
  '-- File not found
End If
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
Dim lSave As Boolean 

'-- Save using the file name passed on LoadFile()
lSave = oIniFile.SaveFile()

If (lSave) Then
  '-- Success
End If

'-- Save to specific directory and file name
lSave = oIniFile.SaveFile("D:\sample\SETTINGS.INI")

If (lSave) Then
  '-- Success
End If
```
