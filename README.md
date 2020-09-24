<div align="center">

## Recursive Get ALL files from a Folder and its Subfolders and its Subfolder Subfolders etc\.


</div>

### Description

Recursive Get ALL files from a Folder and its Subfolders and its Subfolder Subfolders etc.
 
### More Info
 
First go to Project->References and include

"Microsoft Scripting Runtime"

Insert a Textbox called textbox1 with its multiline property set to true. Also a Command Button called Command1 .


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alexander Triantafyllou](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alexander-triantafyllou.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alexander-triantafyllou-recursive-get-all-files-from-a-folder-and-its-subfolders-and-its-s__1-44895/archive/master.zip)





### Source Code

```
public filetext as String
private sub command1_click()
Dim fso As New FileSystemObject
myfoldertext="C:\folder\"
call get_all_directory_files(fso.getfolder(myfoldertext))
text1.text=filetext
set fso=nothing
end sub
Public Sub get_all_directory_files(ByVal tfolder As folder)
Dim objfile As file
Dim objfolder As folder
Dim fso As New FileSystemObject
If tfolder <> "" Then
For Each objfile In tfolder.Files
'do the stuff we want with the files
filetext=filetext+objfile+ vbNewLine
Next
For Each objfolder In tfolder.SubFolders
Call get_all_directory_files(objfolder)
Next
Set fso = Nothing
End If
End Sub
```

