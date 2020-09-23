<div align="center">

## Get the Text from DOC,RTF,HTML


</div>

### Description

With this Code, you can get the plaintext, from a DOC, RTF or HTML File
 
### More Info
 
The HTML Routine is not Perfect


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Holger Kaslowitz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/holger-kaslowitz.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB\.NET
**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__10-2.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/holger-kaslowitz-get-the-text-from-doc-rtf-html__10-272/archive/master.zip)





### Source Code

```
Public Class File2String
  Private WithEvents RichText As New Windows.Forms.RichTextBox()
  Public Function FromRTF(ByVal File As String) As String
    RichText.LoadFile(File)
    Return Replace(RichText.Text, Chr(10), vbCrLf)
  End Function
  Public Function FromDOC(ByVal File As String) As String
    Dim TempString As String
    TempString = OpenTexFile(File)
    Dim LastPos As Integer = InStrRev(TempString, vbCrLf & vbCrLf)
    Dim FirstPos As Integer = InStrRev(TempString, "Ù", LastPos) + 1
    TempString = Mid(TempString, FirstPos, LastPos - FirstPos)
    TempString = Replace(TempString, "F" & Chr(9), "")
    TempString = Replace(TempString, "e'", "")
    TempString = Mid(TempString, InStrRev(TempString, Chr(1)) + 1)
    Return TempString
  End Function
  Public Function FromHTML(ByVal File As String) As String
    RichText.LoadFile(File, RichTextBoxStreamType.PlainText)
    Dim Lastpos As Integer = 1
    Dim LastPos2 As Integer = 1
    Dim TempText As String
    Do While Lastpos < Len(RichText.Text) - 5
      Lastpos = InStr(Lastpos + 1, RichText.Text, ">")
      LastPos2 = InStr(Lastpos, RichText.Text, "<")
      If LastPos2 <> 0 Then
        TempText = TempText & Mid(RichText.Text, Lastpos, LastPos2 - Lastpos)
      End If
    Loop
    TempText = Replace(TempText, "  ", "")
    TempText = Replace(TempText, vbCrLf, "")
    TempText = Replace(TempText, ">", "")
    TempText = Replace(TempText, Chr(10), " ")
    TempText = Replace(TempText, Chr(9), "")
    Return Trim(TempText)
  End Function
  Public Function OpenTexFile(ByVal Fil As String) As String
    Dim Text As String
    Dim Textfile As System.IO.FileStream = System.IO.File.OpenRead(Fil)
    Dim i As Long
    Dim TempBytes(Textfile.Length) As Byte
    Textfile.Read(TempBytes, 0, Textfile.Length)
    Textfile.Close()
    For i = 0 To TempBytes.Length - 1
      If TempBytes(i) = 0 Then
      ElseIf TempBytes(i) = 13 Then
        Text = Text & vbCrLf
      Else
        Text = Text & Chr(TempBytes(i))
      End If
    Next
    Return Text
  End Function
End Class
```

