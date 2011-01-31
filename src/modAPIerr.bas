Attribute VB_Name = "modAPIerr"

Option Explicit
Option Compare Binary
Option Base 0

Public Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Public Declare Function GetLastError Lib "kernel32.dll" () As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long

Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Public Const LANG_NEUTRAL As Long = &H0
Public Const ERROR_SUCCESS As Long = 0&
'

Public Function GetErrDes() As String
Dim Buffer As String
Buffer = Space$(200)
FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
GetErrDes = Trim$(Buffer)
End Function
