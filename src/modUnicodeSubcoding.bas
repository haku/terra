Attribute VB_Name = "modUnicodeSubcoding"
'after failing to pass unicode strings via API, i have resorted to
'encoding them as numbers in the string.  this trick will only work
'file names as it assumes the "|" chr is not used for anything else.
'unicode chr's are represnted as numbers between a "|" and a ";".

Option Explicit

Public Function us_Encode(sSource As String) As String
If Not IsUnicode(sSource) Then
    us_Encode = sSource
    Exit Function
End If

Dim i As Long, sbOutput As New cStringBuilder, sChr As String

For i = 1 To Len(sSource)
    sChr = Mid$(sSource, i, 1)
    
    If IsUnicode(sChr) Then
        sbOutput.Append "|" & Hex$(AscW(sChr)) & ";"
    Else
        sbOutput.Append sChr
    End If
Next i

us_Encode = sbOutput.ToString
End Function

Public Function us_Decode(sSource As String) As String
If InStr(1, sSource, "|") <= 0 Then
    us_Decode = sSource
    Exit Function
End If

Dim i As Long, j As Long, sbOutput As New cStringBuilder, sChr As String, sChrCode As String

For i = 1 To Len(sSource)
    sChr = Mid$(sSource, i, 1)
    
    If sChr = "|" Then
        j = InStr(i, sSource, ";")
        sChrCode = Mid$(sSource, i + 1, j - i - 1)
        sChr = ChrW$(Val("&h" & sChrCode & "&"))
        
        i = j
    End If
    
    sbOutput.Append sChr
Next i

us_Decode = sbOutput.ToString
End Function
