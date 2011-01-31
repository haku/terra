Attribute VB_Name = "modStrings"
Option Explicit

Public Const str_mdb_FileNotMedia = "unwanded file ext."
Public Const str_mdb_FileMis = "file not found."
Public Const str_mdb_FileDupFst = "duplicate (first)."
Public Const str_mdb_FileDupSub = "duplicate (subsiquent)."
Public Const str_mdb_FileMisDupExi = "duplicate (on hdd)."
Public Const str_mdb_FileMisDupMis = "duplicate (file not found.)"
'

Function GettrrFileActDes(i As trrFileAction) As String
Select Case i
    Case trrDoNothing:   GettrrFileActDes = "do nothing."
    Case trrDelLibRef:   GettrrFileActDes = "del lib ref."
    Case trrMoveFile:    GettrrFileActDes = "move file."
    Case trrMarkMissing: GettrrFileActDes = "mark as missing."
End Select
End Function
