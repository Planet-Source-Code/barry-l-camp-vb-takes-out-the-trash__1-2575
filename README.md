<div align="center">

## VB Takes Out The Trash


</div>

### Description

Empties the Recycle Bin, regardless of what drive/folder assigned to.
 
### More Info
 
No form required. (Sample uses a standard module instead of a form).

Empties the Recycle Bin. If workstation has more than one drive, *all* Recycled folders (e.g. C:\Recycled) on *all* drives are emptied.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Barry L\. Camp](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/barry-l-camp.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/barry-l-camp-vb-takes-out-the-trash__1-2575/archive/master.zip)

### API Declarations

SHEmptyRecycleBin (Shell32)


### Source Code

```
' ShellTrash Demo
' by Barry L. Camp (blcamp@yahoo.com)
Option Explicit ' The Author's preference.
Const SHERB_NOCONFIRMATION = &H1& ' No dialog confirming the deletion of the objects will be displayed.
Const SHERB_NOPROGRESSUI = &H2& ' No dialog indicating the progress will be displayed.
Const SHERB_NOSOUND = &H4& ' No sound will be played when the operation is complete.
Private Declare Function SHEmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBinA" _
 (ByVal hWnd As Long, ByVal lpBuffer As String, ByVal dwFlags As Long) As Long
Sub Main()
 Dim rc As Long
 Dim nFlags As Long
 ' Suppresses all UI elements, for "quiet" operation.
 nFlags = SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND
 rc = SHEmptyRecycleBin(0&, vbNullString, nFlags)
End Sub
```

