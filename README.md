<div align="center">

## Advanced Begin & End Date Calculations Simplified


</div>

### Description

Allows calculation of Begin or End dates based upon the RANGE (Week, Month, Year), the DATE to use as the source or comparison date and PREV or CURRENT range. Examples:

'BeginDateCalc("W","P",#11/15/2000#) returns: 11/5/00 as the first day or the PREVIOUS WEEK is Sunday the 5th. You could easily modify the code to allow the last day of the week to be any day you wish.

'BeginDateCalc("M","P",#11/15/2000#) = 10/1/00

'BeginDateCalc("M","C",#11/15/2000#) = 11/1/00

'BeginDateCalc("Wm","C",#11/15/2000#) = 11/1/00 'Wm is used to tell us Week but Month limited. 'Notice the same with "W" instead of "Wm" would result in 10/29/00
 
### More Info
 
Range, Calculation, Date

'Public Domain: This code may be used and distributed freely as long as header remains unchanged.

'The person(s) supplying this code can not be held liable for use, misuse or damage caused by the use of this code.

'Written by Chad M. Kovac

'CEO, Tech Knowledgey, Inc.

'GlobalReplaceCode@TechKnowledgeyinc.com

'http://www.TechKnowledgeyInc.com

'10/04/00 MS Access 97/2000

Caculated Date


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chad M\. Kovac](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chad-m-kovac.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chad-m-kovac-advanced-begin-end-date-calculations-simplified__1-12866/archive/master.zip)

### API Declarations

Created for use in VBA with MS Access.


### Source Code

```
Function EndDateCalc(Range As String, Prev_or_Current As String, Optional FDate As Date) As Date
On Error GoTo Errored
GoTo Main
Errored:
Call Errored_Out(Err.Source, Err.Number, Err.Description, False)
Main:
If FDate <= #1/1/1900# Then FDate = Now()
On Error Resume Next
Reselect:
Select Case Prev_or_Current
Case "P"
 Select Case Range
 Case "D"
 EndDateCalc = DateValue(Format(FDate, "mm/dd/yyyy"))
 Case "W"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w")), "mm/dd/yyyy"))
 Case "Wm"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w")), "mm/dd/yyyy"))
 If Format(EndDateCalc, "yyyymm") > Format(FDate, "yyyymm") Then
 Range = "M"
 GoTo Reselect
 End If
 Case "M"
 Err.Clear
 EndDateCalc = DateValue(Format(FDate - (Val(Format(FDate, "dd"))), "mm/31/yyyy"))
 If Err.Number > 0 Then
 Err.Clear
 EndDateCalc = DateValue(Format(FDate - (Val(Format(FDate, "dd"))), "mm/30/yyyy"))
 If Err.Number > 0 Then
 Err.Clear
 EndDateCalc = DateValue(Format(FDate - (Val(Format(FDate, "dd"))), "mm/29/yyyy"))
 If Err.Number > 0 Then
 Err.Clear
 EndDateCalc = DateValue(Format(FDate - (Val(Format(FDate, "dd"))), "mm/28/yyyy"))
 If Err.Number > 0 Then EndDateCalc = #1/1/90#
 End If
 End If
 End If
End Select
Case "C"
 Select Case Range
 Case "D"
 EndDateCalc = DateValue(Format(FDate, "mm/dd/yyyy"))
 Case "W"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 7), "mm/dd/yyyy"))
 Case "Wm"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 7), "mm/dd/yyyy"))
 If Format(EndDateCalc, "yyyymm") > Format(FDate, "yyyymm") Then
 Range = "M"
 GoTo Reselect
 End If
 Case "M"
 Err.Clear
 EndDateCalc = DateValue(Format(FDate, "mm/31/yyyy"))
 If Err.Number > 0 Then EndDateCalc = DateValue(Format(FDate, "mm/30/yyyy"))
 End Select
Case "N"
 Select Case Range
 Case "D"
 EndDateCalc = DateValue(Format(FDate + 1, "mm/dd/yyyy"))
 Case "W"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 7), "mm/dd/yyyy")) + 7
 Case "Wm"
 EndDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 7), "mm/dd/yyyy")) + 7
 If Format(EndDateCalc, "yyyymm") > Format(FDate, "yyyymm") Then
 Range = "M"
 GoTo Reselect
 End If
 Case "M"
 Err.Clear
 EndDateCalc = DateValue(Month(FDate) + 1 & "/31/" & Format(FDate, "yyyy"))
 If Err.Number > 0 Then
 Err.Clear
 EndDateCalc = DateValue(Month(FDate) + 1 & "/30/" & Format(FDate, "yyyy"))
 If Err.Number > 0 Then
 Err.Clear
 EndDateCalc = DateValue(Month(FDate) + 1 & "/29/" & Format(FDate, "yyyy"))
 If Err.Number > 0 Then EndDateCalc = DateValue(Month(FDate) + 1 & "/28/" & Format(FDate, "yyyy"))
 End If
 End If
 End Select
End Select
End Function
Function BeginDateCalc(Range As String, Prev_or_Current As String, Optional FDate As Date) As Date
'Public Domain: This code may be used and distributed freely as long as header remains unchanged. _
'The person(s) supplying this code can not be held liable for use, misuse or damage caused by the use of this code.
'
'Allows calculation of Begin or End dates based upon the RANGE (Week, Month, Year), the DATE to use as the source or comparison date and PREV or CURRENT range. Examples:
'BeginDateCalc("W","P",#11/15/2000#) returns: 11/5/00 as the first day or the PREVIOUS WEEK is Sunday the 5th. You could easily modify the code to allow the last day of the week to be any day you wish.
'BeginDateCalc("M","P",#11/15/2000#) = 10/1/00
'BeginDateCalc("M","C",#11/15/2000#) = 11/1/00
'BeginDateCalc("Wm","C",#11/15/2000#) = 11/1/00 ' Wm is used to tell us Week but Month limited. Notice the same with "W" instead of "Wm" would result in 10/29/00
'
' Written by Chad M. Kovac
' CEO, Tech Knowledgey, Inc.
' GlobalReplaceCode@TechKnowledgeyinc.com
' http://www.TechKnowledgeyInc.com
' 10/04/00 MS Access 97/2000
On Error GoTo Errored
GoTo Main
Errored:
Call Errored_Out(Err.Source, Err.Number, Err.Description, False)
Main:
If FDate <= #1/1/1900# Then FDate = Now()
On Error Resume Next
Select Case Prev_or_Current
Case "P"
 Select Case Range
 Case "D"
 If Format(FDate, "w") = 2 Then
 BeginDateCalc = DateValue(Format(FDate - 3, "mm/dd/yyyy"))
 Else
 BeginDateCalc = DateValue(Format(FDate - 1, "mm/dd/yyyy"))
 End If
 Case "W"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") + 6), "mm/dd/yyyy"))
 Case "M"
 BeginDateCalc = DateValue(Format(FDate - (Val(Format(FDate, "dd"))), "mm/01/yyyy"))
 Case "Wm"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") + 6), "mm/dd/yyyy"))
 If Format(BeginDateCalc, "yyyymm") < Format(FDate, "yyyymm") Then _
 BeginDateCalc = Format(FDate, "mm/01/yyyy")
 End Select
Case "C"
 Select Case Range
 Case "D"
 BeginDateCalc = DateValue(Format(FDate, "mm/dd/yyyy"))
 Case "W"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 1), "mm/dd/yyyy"))
 Case "M"
 BeginDateCalc = DateValue(Format(FDate, "mm/01/yyyy"))
 Case "Wm"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 1), "mm/dd/yyyy"))
 If Format(BeginDateCalc, "yyyymm") < Format(FDate, "yyyymm") Then _
 BeginDateCalc = Format(FDate, "mm/01/yyyy")
 End Select
Case "N"
 Select Case Range
 Case "D"
 BeginDateCalc = DateValue(Format(FDate + 1, "mm/dd/yyyy"))
 Case "W"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 1), "mm/dd/yyyy")) + 7
 Case "M"
 BeginDateCalc = DateValue(Month(FDate) + 1 & "/01/" & Format(FDate, "yyyy"))
 Case "Wm"
 BeginDateCalc = DateValue(Format(FDate - (Format(FDate, "w") - 1), "mm/dd/yyyy"))
 If Format(BeginDateCalc, "yyyymm") < Format(FDate, "yyyymm") Then _
 BeginDateCalc = Format(FDate, "mm/01/yyyy")
 End Select
End Select
End Function
```

