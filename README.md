<div align="center">

## Check if File Exists \- Including Hidden Files


</div>

### Description

Simple function to check if file exists. Detects Normal & Hidden files. Improvement of code from Greg G., and incorporating a suggestion by Larry Rebich.
 
### More Info
 
A Valid Pathname as String.

Boolean TRUE if file exists at path specified

Simply returns FALSE if an invalid path is encountered


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Redding \- Blue Knot Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-redding-blue-knot-software.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-redding-blue-knot-software-check-if-file-exists-including-hidden-files__1-2045/archive/master.zip)





### Source Code

```
Public Function FileExists(strFile as String) As String
 On Error Resume Next 'Doesn't raise error - FileExists will be false
      'if error occurs
 'a valid path would be someting like "C:\Windows\win.ini" - Full path
 'must be specified
 FileExists = Dir(strFile, vbHidden) <> ""
End Function
```

