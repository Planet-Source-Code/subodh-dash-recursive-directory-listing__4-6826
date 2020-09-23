<div align="center">

## Recursive Directory Listing


</div>

### Description

This will list all the sub directories under a directory in a recursive manner. You may wish to remove the "msgbox" from the code and do something useful.
 
### More Info
 
A directory name

As the process looks through all sub directories, it is a time consuming process. For testing purpose, use a directory which has few sub directories inside it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Subodh Dash](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/subodh-dash.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/subodh-dash-recursive-directory-listing__4-6826/archive/master.zip)

### API Declarations

Free to use /Free to distribute


### Source Code

```
Option Explicit
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
dim rd
rd = RecursiveDir ("c:\sheridan")
function RecursiveDir (path)
	dim folderpath, fol, FolderName
	Set folderpath = fso.getfolder(path)
	Set fol = folderpath.SubFolders
	For Each Foldername In fol
		msgbox Foldername
		RecursiveDir = FolderName
		RecursiveDir FolderName
	Next
end function
```

