<div align="center">

## Voice Clock  v1\.0


</div>

### Description

This is a complete application that will speak the time every hour on the hour. It is a good example of using a resource file and of using API calls to play .wav files

If you download and install voice clock, please take the time to rate it. Also, I would appreciate any comments that will help me to understand what is good about it and what it lacks.

Thank you

madwhit@netzero.net
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-02-11 12:06:26
**By**             |[madwhit](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/madwhit.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD33552112000\.zip](https://github.com/Planet-Source-Code/madwhit-voice-clock-v1-0__1-5877/archive/master.zip)

### API Declarations

```
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
```





