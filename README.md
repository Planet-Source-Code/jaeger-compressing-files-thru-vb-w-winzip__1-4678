<div align="center">

## Compressing Files thru VB \(w/ WinZip\)


</div>

### Description

this codes compresses any files of your choice to a designated target folder using Winzipr..
 
### More Info
 
the source folder and files, the target folder with designated zip name (ie. myfile.zip)

you have winzip

you will receive an error message if you don't have WinZip


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jaeger](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jaeger.md)
**Level**          |Unknown
**User Rating**    |3.0 (18 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jaeger-compressing-files-thru-vb-w-winzip__1-4678/archive/master.zip)





### Source Code

```
Dim retval As Double
retval = Shell("C:\program files\winzip\WINzip32 -a c:\TargetFolder\AssignedNameFile.zip c:\SourceFolder\SourceFile(s)", 6)
note: i used shell function here.. if anyone has a better idea than of this, pls tell me..
```

