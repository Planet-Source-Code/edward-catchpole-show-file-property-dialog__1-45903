<div align="center">

## Show File Property Dialog


</div>

### Description

This function will display the file property dialog for any file you specify - it is the same one that Windows Explorer shows when you right click and goto 'Properties'.

The original article can be found at: http://www.mvps.org/vbnet/code/shell/propertypage.htm
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Edward Catchpole](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/edward-catchpole.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/edward-catchpole-show-file-property-dialog__1-45903/archive/master.zip)





### Source Code

```
Private Type SHELLEXECUTEINFO
  cbSize    As Long
  fMask     As Long
  hwnd     As Long
  lpVerb    As String
  lpFile    As String
  lpParameters As String
  lpDirectory  As String
  nShow     As Long
  hInstApp   As Long
  lpIDList   As Long   'Optional
  lpClass    As String  'Optional
  hkeyClass   As Long   'Optional
  dwHotKey   As Long   'Optional
  hIcon     As Long   'Optional
  hProcess   As Long   'Optional
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEx Lib "shell32" _
  Alias "ShellExecuteExA" _
 (SEI As SHELLEXECUTEINFO) As Long
Private Sub Form_Load()
  Command1.Caption = "Show Properties"
 'assure string points to a valid file
 'on your system
  Text1.Text = "c:\windows\notepad.exe"
End Sub
Private Sub Command1_Click()
 'show the properties dialog, passing the filename
 'and the owner of the dialog
  Call ShowProperties(Text1.Text, Me.hwnd)
End Sub
Private Sub Command2_Click()
  Unload Me
End Sub
Private Sub ShowProperties(sFilename As String, hWndOwner As Long)
 'open a file properties property page for
 'specified file if return value
  Dim SEI As SHELLEXECUTEINFO
 'Fill in the SHELLEXECUTEINFO structure
  With SEI
   .cbSize = Len(SEI)
   .fMask = SEE_MASK_NOCLOSEPROCESS Or _
        SEE_MASK_INVOKEIDLIST Or _
        SEE_MASK_FLAG_NO_UI
   .hwnd = hWndOwner
   .lpVerb = "properties"
   .lpFile = sFilename
   .lpParameters = vbNullChar
   .lpDirectory = vbNullChar
   .nShow = 0
   .hInstApp = 0
   .lpIDList = 0
  End With
 'call the API to display the property sheet
  Call ShellExecuteEX(SEI)
End Sub
```

