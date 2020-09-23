<div align="center">

## VBLanguageMacro


</div>

### Description

Check and/or Execute UserDefined Visual Basic Code with just four lines of Code and one API function.
 
### More Info
 
When running the App in any other version of VB than VB6, change the API Declaration "vba6.dll" to "vba5.dll" etc... the number should be the version you're using.


<span>             |<span>
---                |---
**Submitted On**   |2004-04-21 10:29:02
**By**             |[PJK](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pjk.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[VBLanguage1736044212004\.zip](https://github.com/Planet-Source-Code/pjk-vblanguagemacro__1-53307/archive/master.zip)

### API Declarations

```
Undocumented VB Function
Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, _
 ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long
```





