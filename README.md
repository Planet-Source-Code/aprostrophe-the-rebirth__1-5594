<div align="center">

## Aprostrophe 'The Rebirth'


</div>

### Description

Have you ever try so send a SQL String to a database that has apostrophes ? If YES you will get a run time ERROR Here is your solution....A function that formats the variable before sending it to the database
 
### More Info
 
Ziltch

Take a string, looks for Aprostrophes or Quotation marks appearing more than twice between commas, if so it will double them up.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Intermediate
**User Rating**    |4.1 (168 globes from 41 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aprostrophe-the-rebirth__1-5594/archive/master.zip)





### Source Code

```
Public Function Apos2(strSQL As String) As String
 Dim F As Long, N As Long, Q As Long
 Dim O As String, P As String, A As String
 Q = -1
 For F = 1 To Len(strSQL)
  P = Mid(strSQL, F, 1)
  If P = "'" Or P = """" Then
   If Q > 0 Then
    O = O + "'" + A
    A = ""
   End If
   Q = Q + 1
  ElseIf P = "," Then
   O = O & A
   Q = -1
   A = ""
  End If
  If Q <= 0 Then
   O = O & P
  Else
   A = A & P
  End If
 Next
 Apos2 = O & A
End Function
24 Jan 00
Some Alterations,
and some documentation,
Though F stays in the loop, for sentimental reasons
Public Function Apos3(strSQL As String) As String
'F is the current position in the original string
'lCountOfApos Counts the occurrences of apostrophes and quotes
'lCharaterAtPositionF equals the Character at position F
'If lCharaterAtPositionF is equal to a apostrophes or quote Then
'If lCountOfApos grater than zero
'Then add a additional apostrophe to sOutput along with sBuffer
'sBuffer is a Buffer that is used to store characters after the Second
'occurrence of a apostrophes or quote whilst not encountering a Comma, Quote or apostrophe
'Clear as mud
  Dim F As Long, lCountOfApos As Long
  Dim sOutput As String, lCharaterAtPositionF As String, sBuffer As String
  lCountOfApos = -1
  For F = 1 To Len(strSQL)
    lCharaterAtPositionF = Mid(strSQL, F, 1)
    If lCharaterAtPositionF = "'" Or lCharaterAtPositionF = """" Then
      If lCountOfApos > 0 Then
        sOutput = sOutput + "'" + sBuffer
        sBuffer = ""
      End If
      lCountOfApos = lCountOfApos + 1
    End If
    If lCountOfApos <= 0 Then
      sOutput = sOutput & lCharaterAtPositionF
    Else
      sBuffer = sBuffer & lCharaterAtPositionF
      If lCharaterAtPositionF = "," Or Right(sBuffer, 5) = " AND " Or Right(sBuffer, 4) = " OR " Then
        sOutput = sOutput & sBuffer
        lCountOfApos = -1
        sBuffer = ""
      End If
    End If
  Next
  Apos3 = sOutput & sBuffer
End Function
```

