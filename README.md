<div align="center">

## Copy a Folder


</div>

### Description

Copy a Folder in another Folder. If the Folder not exists then this program will create the Folder.
 
### More Info
 
1 Form, 1 Commandbutton


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Manuel W\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manuel-w.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/manuel-w-copy-a-folder__1-7439/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
On Error Resume Next
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set fld = fso.createfolder("c:\windowscopy")
' For Example:
path1$ = "c:\win98\config"
path2$ = "c:\windowscopy\"
If fso.folderexists(path1$) Then
If Not fso.folderexists("c:\windowscopy") Then
'Generate Path
Set fld = fso.createfolder("c:\windowscopy")
End If
'Copy now
fso.copyfolder path1$, path2$, True
'On Error:
Else
MsgBox "Verzeichnis konnte nicht kopiert werden!"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set fso = Nothing
End Sub
```

