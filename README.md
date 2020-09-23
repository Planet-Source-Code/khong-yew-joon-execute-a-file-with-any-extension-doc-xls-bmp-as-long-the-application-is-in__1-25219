<div align="center">

## Execute a file with any extension \.doc,\.xls,\.bmp as long the application is install in windows


</div>

### Description

Pass the file name and the function wil check from windows what is the application exe to run the file.Exp you want to run a abc.doc document from your program,u need to know msword.exe path and then you run the shell(applicate exe abc.doc,1) to execute the abc.doc

This work with any extension as long as it register to windows. etc .xls,.vbp,.doc
 
### More Info
 
strname = "Full path of your file"

exp C:\Myfolder\Abc.doc or C:\abc.xls or c:\abc.bmp

Execute a file with any extension

for more access www.efastclick.com


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[KHONG YEW JOON](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/khong-yew-joon.md)
**Level**          |Advanced
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/khong-yew-joon-execute-a-file-with-any-extension-doc-xls-bmp-as-long-the-application-is-in__1-25219/archive/master.zip)

### API Declarations

```
Private Declare Function FindExecutable _
 Lib "shell32.dll" Alias "FindExecutableA" _
 (ByVal lpFile As String, _
 ByVal lpDirectory As String, _
 ByVal lpResult As String) As Long
```


### Source Code

```
Public Function runapp(strname As String, appname As String) As Long
Dim strResult As String
Dim lngResult As Long
Dim i, s_msg
 s_msg = MsgBox("Launch " & appname & " ?", vbYesNo, appname)
 If s_msg = vbYes Then
 strResult = String(255, 0)
 lngResult = FindExecutable(strname, vbNullString, strResult)
 strResult = Trim(Replace(strResult, "/dde", "", 1))
'run the file and not an .exe file
 i = Shell(Trim(Replace(strResult, vbNullChar, "", 1)) & " " & strname, 1)
 End If
End Function
```

