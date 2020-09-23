<div align="center">

## A Compare Time Function \(like you can compare dates in VB\)


</div>

### Description

This will allow you to compare times. I noticed that there is a 'Date' type in VB, but no 'Time' type. So if you want to compare Dates you are fine, but for Time comparisons you are a bit stuffed. This is very simple, and will allow you to convert times into numbers so that you can make easy comparisons with them.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Proxy Avoidance](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/proxy-avoidance.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/proxy-avoidance-a-compare-time-function-like-you-can-compare-dates-in-vb__1-63363/archive/master.zip)





### Source Code

```
' This is the sort of code that makes you think 'Why didnt I think of that?!?
'
' EG:
' IF TimeNo("21:55:32") < TimeNo("20:40:12") Then
' msgbox "WHOOO!"
' end if
'
' The code is also cross-compatible with different time formats...
'
' IF TimeNo("21:55:32") < TimeNo("8:40PM") Then
' msgbox "WHOOO!"
' end if
Public Function TimeNo(Time As String) As Long
TimeNo = CLng(Replace(Format(Time, "hhnnss"), ":", ""))
End Function
```

