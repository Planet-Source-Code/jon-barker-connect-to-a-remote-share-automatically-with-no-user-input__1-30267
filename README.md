<div align="center">

## Connect to a remote share AUTOMATICALLY, with NO user input\.


</div>

### Description

Ive seen a lot of source code that can bring up the window asking the user to connect to a remote share, but none of which would work without user input. This API call (WNetAddConnection) is very easy and simple to use, altho i cant guarrentee it will work on 98/ME.... tHe_cLeanER
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Beginner
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-connect-to-a-remote-share-automatically-with-no-user-input__1-30267/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Private Sub Form_Load()
Dim r As Long
r = WNetAddConnection("\\dedicated\xpserver", vbNullString, "x:")
If r <> 0 Then
 MsgBox "ERROR: " & Err.Description
End If
End Sub
```

