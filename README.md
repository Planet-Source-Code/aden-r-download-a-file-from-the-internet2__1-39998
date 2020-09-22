<div align="center">

## Download A File From The Internet2


</div>

### Description

This File Allows You To Download Files From The Internet.

The Source Code Is Fairly Easy To Use Too!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aden R](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aden-r.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aden-r-download-a-file-from-the-internet2__1-39998/archive/master.zip)





### Source Code

```
'Insert This In A Space In The Form
'Not Between Any Other Pieces Of Code Though
Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, _
  ByVal szFileName As String, _
  ByVal dwReserved As Long, _
  ByVal lpfnCB As Long) As Long
  Public Function DownloadFile(URL As String, _
LocalFilename As String) As Boolean
Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function
ret = DownloadFile("http://www.URL Of The File", _
"c:\Which Dir To Save The File To.jpg, exe etc.")
```

