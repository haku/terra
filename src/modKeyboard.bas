Attribute VB_Name = "modKeyboard"
Option Explicit

Public Const MMkey_Play As Long = 917504
Public Const MMkey_Stop As Long = 851968
Public Const MMkey_Prev_Item As Long = 65536
Public Const MMkey_Next_Item As Long = 131072
Public Const MMkey_Prev_Track As Long = 786432
Public Const MMkey_Next_Track As Long = 720896
Public Const MMkey_Play2 As Long = 271450112
Public Const MMkey_Pause As Long = 271515648
Public Const MMkey_Prev2 As Long = 271712256
Public Const MMkey_Next2 As Long = 271646720

'Public Function kb_IsMmKey(l As Long) As Boolean
'kb_IsMmKey = False
'
'Select Case l
'    Case MMkey_Play, MMkey_Stop, MMkey_Prev_Item, MMkey_Next_Item, _
'        MMkey_Prev_Track, MMkey_Next_Track, _
'        MMkey_Play2, MMkey_Pause, MMkey_Prev2, MMkey_Next2
'
'        kb_IsMmKey = True
'
'End Select
'End Function
