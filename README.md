<div align="center">

## RmDir


</div>

### Description

This Procedure Deletes all Files in Directory as well as all Sub Directories and Files
 
### More Info
 
vFile = Directory to Delete


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert A\. Charest Jr\.](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-a-charest-jr.md)
**Level**          |Unknown
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-a-charest-jr-rmdir__1-842/archive/master.zip)





### Source Code

```
'###########################################
'# Removes an Entire Directory Structure #
'# ------------------------------------- #
'# Created By : Robert A. Charest Jr.   #
'# E-mail   : charest@friendlybeaver.com #
'###########################################
Public Sub RmTree(ByVal vDir As Variant)
  Dim vFile As Variant
  ' Check if "\" was placed at end
  ' If So, Remove it
  If Right(vDir, 1) = "\" Then
    vDir = Left(vDir, Len(vDir) - 1)
  End If
  ' Check if Directory is Valid
  ' If Not, Exit Sub
  vFile = Dir(vDir, vbDirectory)
  If vFile = "" Then
    Exit Sub
  End If
  ' Search For First File
  vFile = Dir(vDir & "\", vbDirectory)
  ' Loop Until All Files and Directories
  ' Have been Deleted
  Do Until vFile = ""
    If vFile = "." Or vFile = ".." Then
      vFile = Dir
    ElseIf (GetAttr(vDir & "\" & vFile) And _
      vbDirectory) = vbDirectory Then
      RmTree vDir & "\" & vFile
      vFile = Dir(vDir & "\", vbDirectory)
    Else
      Kill vDir & "\" & vFile
      vFile = Dir
    End If
  Loop
  ' Remove Top Most Directory
  RmDir vDir
End Sub
```

