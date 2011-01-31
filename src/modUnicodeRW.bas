Attribute VB_Name = "modUnicodeRW"
'i think this module is from vbaccelerator.com

Option Explicit
Option Compare Binary
Option Base 0

' API declarations.
'Private Const OFS_MAXPATHNAME As Long = 128
'Private Const OF_WRITE       As Long = &H1
'Private Const OF_READ         As Long = &H0
'Private Const OF_CREATE       As Long = &H1000

Private Const ForReading      As Long = 1

Public Enum ForWriteEnum
   ForWriting = 2
   ForAppending = 8
End Enum

Public Enum TristateEnum
   TristateTrue = -1        'Opens the file as Unicode
   TristateFalse = 0        'Opens the file as ASCII
   TristateUseDefault = -2  'Use default system setting
End Enum

'Private Type OVERLAPPED
'   Internal             As Long
'   InternalHigh         As Long
'   offset               As Long
'   OffsetHigh           As Long
'   hEvent               As Long
'End Type
'
'Private Type OFSTRUCT
'   cBytes               As Byte
'   fFixedDisk           As Byte
'   nErrCode             As Integer
'   Reserved1            As Integer
'   Reserved2            As Integer
'   szPathName           As String * OFS_MAXPATHNAME
'End Type

'Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'Public Function AppPath() As String
'   AppPath = App.Path
'   If Right$(AppPath, 1) <> "\" Then
'      AppPath = AppPath & "\"
'   End If
'End Function

Public Function UnicodeFile_Read_FSO( _
   ByVal sFileName As String, _
   Optional ByVal TriState As TristateEnum = TristateTrue) As String
   
   Dim objFSO           As Object
   Dim objStream        As Object

   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If (Not objFSO Is Nothing) Then
      Set objStream = objFSO.OpenTextFile( _
         sFileName, ForReading, False, TriState)
      If (Not objStream Is Nothing) Then
         With objStream
            UnicodeFile_Read_FSO = .ReadAll
            .Close
         End With
         Set objStream = Nothing
      End If
      Set objFSO = Nothing
   End If
End Function

Public Sub UnicodeFile_Write_FSO( _
   ByVal sFileName As String, _
   ByVal sText As String, _
   Optional ByVal ForWrite As ForWriteEnum = ForWriting, _
   Optional ByVal TriState As TristateEnum = TristateTrue)

   Dim objFSO           As Object
   Dim objStream        As Object

   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If (Not objFSO Is Nothing) Then
      Set objStream = objFSO.OpenTextFile( _
         sFileName, ForWrite, True, TriState)
         
      If (Not objStream Is Nothing) Then
         With objStream
            .Write sText
            .Close
         End With
         Set objStream = Nothing
      End If
      Set objFSO = Nothing
   End If
End Sub

'Public Function UnicodeFile_Read_VB(ByVal sFileName As String, _
'    Optional ByVal bRemoveBOM As Boolean) As String
'
'Dim FF As Long, b() As Byte, s As String
'Const uBOM As String = "ÿþ"
'
'On Error Resume Next
'
'FF = FreeFile
'Open sFileName For Binary Access Read As FF
'    ReDim b(LOF(FF))
'    Get FF, , b
'Close FF
's = b
'
'If bRemoveBOM Then
'   If InStr(s, uBOM) = 1 Then
'      s = Replace$(s, uBOM, "")
'   End If
'End If
'
'UnicodeFile_Read_VB = s
'End Function

'Public Sub UnicodeFile_Write_VB(ByVal sFileName As String, _
'   ByVal sText As String, _
'   Optional ByVal bInsertBOM As Boolean)
'
'Dim FF As Long, b() As Byte
'
'On Error Resume Next
'Kill sFileName
'On Error GoTo 0
'
'FF = FreeFile
'Open sFileName For Binary Access Write As #FF
'    If bInsertBOM Then
'       ReDim b(1)
'       b(0) = &HFF
'       b(1) = &HFE
'       Put #FF, , b
'       Erase b
'    End If
'    b = sText
'    Put #FF, , b
'Close #FF
'End Sub

'Public Function UnicodeFile_Read_API(ByVal sFileName As String) As String
'   Dim lpFileInfo       As OFSTRUCT
'   Dim lpOverlapped     As OVERLAPPED
'   Dim szPathName       As String * OFS_MAXPATHNAME
'   Dim hFile            As Long
'   Dim sText            As String
'   Dim lLength          As Long
'   Dim lLengthRet       As Long
'
'   szPathName = sFileName
'
'   With lpFileInfo
'      .cBytes = Len(lpFileInfo)
'      .fFixedDisk = 1
'      .szPathName = szPathName
'   End With
'
'   lLength = FileLen(sFileName)
'   sText = String(lLength, " ")
'
'   hFile = OpenFile(sFileName, lpFileInfo, OF_READ)
'   If (hFile) Then
'      ReadFile hFile, ByVal StrPtr(sText), lLength, lLengthRet, lpOverlapped
'      UnicodeFile_Read_API = MidB(sText, 1, lLength)
'      CloseHandle (hFile)
'   End If
'End Function

'Public Function UnicodeFile_Write_API(ByVal sFileName As String, ByVal sText As String) As Boolean
'   Dim lpFileInfo       As OFSTRUCT
'   Dim lpOverlapped     As OVERLAPPED
'   Dim szPathName       As String * OFS_MAXPATHNAME
'   Dim hFile            As Long
'   Dim lResult          As Long
'   Dim lLengthRet       As Long
'
'   szPathName = sFileName
'
'   With lpFileInfo
'      .cBytes = Len(lpFileInfo)
'      .fFixedDisk = 1
'      .szPathName = szPathName
'   End With
'
'   hFile = OpenFile(sFileName, lpFileInfo, OF_CREATE)
'   If (hFile) Then
'      lResult = WriteFile(hFile, ByVal StrPtr(sText), LenB(sText), lLengthRet, lpOverlapped)
'      UnicodeFile_Write_API = lResult <> 0
'      CloseHandle (hFile)
'   End If
'End Function
