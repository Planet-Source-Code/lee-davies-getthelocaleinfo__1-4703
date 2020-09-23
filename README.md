<div align="center">

## GetTheLocaleInfo


</div>

### Description

Gets you Locale information from your machine. It's well freaky because it knows what country you come from!!!!!!
 
### More Info
 
Just use type - Call GetTheLocaleInfo()


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lee Davies](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lee-davies.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lee-davies-getthelocaleinfo__1-4703/archive/master.zip)

### API Declarations

```
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
```


### Source Code

```
Public Const LOCALE_USER_DEFAULT = &H400
Public Const LOCALE_IDATE = &H21      ' short date format ordering
Public Const LOCALE_SLANGUAGE = &H2     ' localized name of language
Public Const LOCALE_SCOUNTRY = &H6     ' localized name of country
Public Const LOCALE_SCURRENCY = &H14    ' local monetary symbol
Public Const LOCALE_ILDATE = &H22      ' long date format ordering
Sub GetTheLocaleInfo()
  Dim strBuffer As String * 100
  Dim lngReturn As Long
  Dim strResult As String
  Dim msg As String
  lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IDATE, strBuffer, 99)
  strResult = LPSTRToVBString(strBuffer)
  Select Case strResult
    Case "0":
      msg = "mm/dd/yy"
    Case "1":
      msg = "dd/mm/yy"
    Case "2":
      msg = "yy/mm/dd"
    Case Else:
      msg = "#Error#"
  End Select
  Debug.Print "You are using the " & msg & " short date format"
  lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ILDATE, strBuffer, 99)
  strResult = LPSTRToVBString(strBuffer)
  Select Case strResult
    Case "0":
      msg = "mm/dd/yyyy"
    Case "1":
      msg = "dd/mm/yyyy"
    Case "2":
      msg = "yyyy/mm/dd"
    Case Else:
      msg = "#Error#"
  End Select
  Debug.Print "You are using the " & msg & " Long date format"
  lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLANGUAGE, strBuffer, 99)
  strResult = LPSTRToVBString(strBuffer)
  Debug.Print "You are using " & strResult & " language"
  lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCOUNTRY, strBuffer, 99)
  strResult = LPSTRToVBString(strBuffer)
  Debug.Print "You live in " & strResult & "!"
  lngReturn = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, strBuffer, 99)
  strResult = LPSTRToVBString(strBuffer)
  Debug.Print "You use " & strResult & " as your currency!"
End Sub
Public Function LPSTRToVBString(ByVal s As String) As String
  Dim nullpos As Integer
  nullpos = InStr(s, Chr(0))
  If nullpos > 0 Then
    LPSTRToVBString = Left(s, nullpos - 1)
  Else
    LPSTRToVBString = ""
  End If
End Function
```

