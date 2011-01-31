Attribute VB_Name = "modStr"
'this crazy idea to try and pass unicode via a VB6 API really does not work.
'
'Option Explicit
'
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
'Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
'
'Public Function StringFromAddr(ByVal address As Long, ByVal length As Long, _
'    Optional ByVal isUnicode As Boolean) As String
'
'    ' determine the length, if necessary
'    If length < 0 Then
'        If isUnicode Then
'            length = lstrlenW(address)
'        Else
'            length = lstrlenA(address)
'        End If
'    End If
'
'    ' copy the characters
'    StringFromAddr = Space$(length)
'    If isUnicode Then
'        CopyMemory ByVal StrPtr(StringFromAddr), ByVal address, length * 2
'    Else
'        CopyMemory ByVal StringFromAddr, ByVal address, length
'    End If
'End Function
'
