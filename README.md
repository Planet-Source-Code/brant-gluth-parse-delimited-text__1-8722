<div align="center">

## Parse Delimited Text


</div>

### Description

The code will take a passed string and a delimiter and parse the string into a variant array for reading. Works very well with csv files, etc.
 
### More Info
 
strSource = line ofn text from a text file (ex:/ "00-40-4200",20000216,0.00,15.00)

spits it back at you as a variant array which can be iterated with a for..next loop


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brant Gluth](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brant-gluth.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brant-gluth-parse-delimited-text__1-8722/archive/master.zip)





### Source Code

```
Public Function ParseDelimitedText(strSource As String, strDelimiter As String) As Variant()
  'Comm:
  'Will take the passed string and parse it out to an array which can then be itereated through
  'with a for ..next loop bounded by lbound(ParseDelimitedText) and ubound(ParseDelimitedText)
  'quote delimited doesn't really work with this, but as you'd need top pass the string loaded with
  'chr$(34)'s anyway I guess it doesn't matter.
  'enh: 06/07/2000 switched delimiter from comma to anything BUT quotes
  'decl:
  Dim intTest As Integer
  Dim intStart As String, intEnd As String
  Dim varHold() As Variant
  'Code:
  intStart = 1
  ReDim varHold(0)
  Do While InStr(intStart, strSource, strDelimiter) <> 0 Or intStart < Len(strSource)
    If intStart <> 1 Then ReDim Preserve varHold(UBound(varHold) + 1)
    intEnd = InStr(intStart, strSource, strDelimiter)
    If intEnd = 0 Then intEnd = Len(strSource)
    'increase the array to hold the new value
    varHold(UBound(varHold)) = CVar(Mid$(strSource, intStart, intEnd - intStart))
    intStart = intEnd + 1 'slap the end as the new start position
  Loop
  'Assign:
  ParseDelimiter = varHold
  'for debugging to the immediate window
    For intTest = LBound(varHold) To UBound(varHold)
        Debug.Print "#" & intTest & ": " & varHold(intTest)
    Next
End Function
```

