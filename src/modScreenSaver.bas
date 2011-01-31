Attribute VB_Name = "modScreenSaver"

Option Explicit
Option Compare Binary
Option Base 0

Public Declare Function SystemParametersInfo _
    Lib "User32" _
    Alias "SystemParametersInfoA" _
      (ByVal uiAction As Long, _
       ByVal uiParam As Long, _
       pvParam As Any, _
       ByVal fWInIni As Long) As Boolean

Private Const SPI_GETSCREENSAVEACTIVE As Long = &H10
Private Const SPI_GETSCREENSAVERRUNNING As Long = &H72
Private Const SPI_SETSCREENSAVEACTIVE As Long = 17
'

Public Function ss_GetActive() As Boolean
Dim bActive As Boolean

SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0, bActive, False

If bActive Then
    ss_GetActive = True
Else
    ss_GetActive = False
End If
End Function

Public Function ss_SetActive(bAct As Boolean)
Dim lActiveFlag As Long

lActiveFlag = IIf(bAct, 1, 0)

SystemParametersInfo SPI_SETSCREENSAVEACTIVE, lActiveFlag, 0, 0
End Function
