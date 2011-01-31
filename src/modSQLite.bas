Attribute VB_Name = "modSQLite"
'sqlite notes:
'select sfile, cast(lendcnt as real) / cast(lstartcnt as real) from tbl_mediafiles order by lendcnt desc limit 10;
'select sfile, cast(lendcnt as real) / cast(lstartcnt as real) AS p from tbl_mediafiles where p=>0.7 order by p desc;

Option Explicit
Option Compare Binary
Option Base 0

Public Declare Sub sqlite3_open Lib "SQLite3VB.dll" (ByVal FileName As String, ByRef handle As Long)
Public Declare Sub sqlite3_close Lib "SQLite3VB.dll" (ByVal DB_Handle As Long)
Public Declare Function sqlite3_last_insert_rowid Lib "SQLite3VB.dll" (ByVal DB_Handle As Long) As Long
Public Declare Function sqlite3_changes Lib "SQLite3VB.dll" (ByVal DB_Handle As Long) As Long

Public Declare Function sqlite_get_table Lib "SQLite3VB.dll" ( _
    ByVal DB_Handle As Long, _
    ByVal SQLString As String, _
    ByRef ErrStr As String _
    ) As Variant()

Public Declare Function sqlite_libversion Lib "SQLite3VB.dll" () As String ' Now returns a BSTR

'// This function returns the number of rows from the last sql statement. Use this to ensure you have a valid array
Public Declare Function number_of_rows_from_last_call Lib "SQLite3VB.dll" () As Long

Public Const file_sqlite_db As String = "mediadb.db3"
'Public Const tbl_mediafiles_colUbnd As Long = 2 'don't think this does anything anymore.

'this query returns all info on all items in db.
Public Const mdb_DefaultSqlQuery As String = _
    "SELECT sfile, dadded, lstartcnt, lendcnt, dlastplay, " & _
    "lmd5, lduration, benabled, bmissing FROM tbl_mediafiles " & _
    "ORDER BY sfile COLLATE NOCASE ASC;"

Public sqlite_Handle
'

Public Sub mdb_connect()
On Error GoTo mdb_connect_err

Dim f As String, bNew As Boolean

f = App.Path & "\" & file_sqlite_db
If Dir(f) = "" Then bNew = True Else bNew = False

sqlite3_open f, sqlite_Handle

If bNew = True Then
    mdb_CreateNew
Else
    mdb_CheckField "lduration", "INT(6)"
    mdb_CheckField "benabled", "INT(1)"
    mdb_CheckField "bmissing", "INT(1)"
End If

Exit Sub
mdb_connect_err:
Main_Err "mdb_connect_err."
err.Clear
End Sub

Public Sub mdb_disconnect()
On Error GoTo mdb_disconnect_err

sqlite3_close sqlite_Handle

mdb_disconnect_err:
err.Clear
End Sub

Function mdb_Query(sSql As String) As Variant
On Error GoTo errHandle

Dim sErr As String

mdb_Query = sqlite_get_table(sqlite_Handle, sSql, sErr)

Exit Function
errHandle:
Debug.Print "mdb_Query errror: " & err.Description
err.Clear
End Function

Public Function mdb_Function(sSql As String) As Variant
On Error GoTo mdb_Function_err
Dim vRet As Variant, vI As Variant, x As Long
mdb_Function = ""

vRet = mdb_Query(sSql)
x = 0
For Each vI In vRet
    If x = 1 Then
        mdb_Function = vI
        Exit For
    End If
    x = x + 1
Next vI

mdb_Function_err:
err.Clear
End Function

Sub mdb_CreateNew()
Const sSQL_new As String = _
    "create table tbl_mediafiles(" & _
    "sfile VARCHAR(10000) not null collate nocase primary key," & _
    "dadded DATETIME," & _
    "lstartcnt INT(6)," & _
    "lendcnt INT(6)," & _
    "dlastplay DATETIME," & _
    "lmd5 BIGINT," & _
    "lduration INT(6)," & _
    "benabled INT(1)," & _
    "bmissing INT(1));"

mdb_Query sSQL_new
End Sub

Sub mdb_CheckField(sField As String, sType As String)
On Error GoTo mdb_CheckField_err

'"select lduration from tbl_mediafiles limit 1;"
'"alter table tbl_mediafiles add column lduration INT(6);"
Dim bColPres As Boolean, sSQL_check As String, sSQL_add As String

sSQL_check = "select " & sField & " from tbl_mediafiles limit 1;"
sSQL_add = "alter table tbl_mediafiles add column " & sField & " " & sType & ";"

Dim vRet As Variant
vRet = mdb_Query(sSQL_check)
bColPres = (number_of_rows_from_last_call > 0)

If Not bColPres Then
    Debug.Print sField & " not present, adding..."
    mdb_Function sSQL_add
Else
'    Debug.Print sField & " present."
End If

Exit Sub
mdb_CheckField_err:
Main_Err "mdb_CheckField"
err.Clear
End Sub

Public Function mdb_AddFile(ByVal sFile As String, Optional bClearHashOnRefind As Boolean = False) As Boolean
On Error GoTo mdb_AddFile_err

Dim sSql As String, n As Long, e As String, f As String

mdb_AddFile = False

e = "encoding file name."
f = Replace(sFile, "'", "''")
f = us_Encode(f)

e = "running query."
sSql = "select * from tbl_mediafiles where sfile='" & f & "' COLLATE NOCASE;"
mdb_Query sSql
n = number_of_rows_from_last_call
If n > 0 Then
    'Debug.Print "file " & f & " already in db " & n & " times."
    
    e = "checking for 'missing' flag."
    If mdb_IsMissing(sFile) Then
        AddToLog frmMain.txtLog, "file re-found: '" & sFile & "'."
        
        e = "setting 'missing' flag to false."
        mdb_SetMissing sFile, False
        
        If bClearHashOnRefind Then
            e = "clearing crc."
            mdb_SetHash sFile, 0
        End If
    End If
    
    Exit Function
End If

e = "adding new item"
sSql = "insert into tbl_mediafiles (sfile,dadded,lstartcnt,lendcnt,lduration,benabled) VALUES ('" & _
    f & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',0,0,0,1);"

'If IsUnicode(sSQL) Then Debug.Print "unicode sql: " & sSQL
'Debug.Print sSQL
mdb_Query sSql
mdb_AddFile = True

Exit Function
mdb_AddFile_err:
Main_Err "mdb_AddFile/" & e
err.Clear
End Function

Public Sub mdb_RemFile(ByVal f As String)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "delete from tbl_mediafiles where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Sub mdb_SetHash(ByVal f As String, lHash As Long)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set lmd5=" & _
    lHash & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Function mdb_IncPlyCnt(ByVal f As String, Optional lN As Long = 1, Optional bEnd As Boolean = False)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

If bEnd Then
    sSql = "update tbl_mediafiles set lendcnt=lendcnt+" & Trim$(Str$(lN)) & _
        " where sfile='" & f & "';"
Else
    sSql = "update tbl_mediafiles set lstartcnt=lstartcnt+" & Trim$(Str$(lN)) & _
        ", dlastplay='" & Format(Now, "yyyy-mm-dd hh:mm:ss") & _
        "' where sfile='" & f & "';"
End If

mdb_Query sSql
End Function

Public Sub mdb_SetPlaybackCnt(ByVal f As String, lSt As Long, lNd As Long)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set lstartcnt=" & Trim$(Str$(lSt)) & _
    ", lendcnt=" & Trim$(Str$(lNd)) & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Sub mdb_SetDuration(ByVal f As String, lDuration As Long)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set lduration=" & Trim$(Str$(lDuration)) & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Sub mdb_SetEnabled(ByVal f As String, bEnabled As Boolean)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set benabled=" & IIf(bEnabled, "1", "0") & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Function mdb_IsMissing(ByVal f As String) As Boolean
Dim sSql As String, vRet As Variant, vI As Variant, b As Long

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "SELECT bmissing FROM tbl_mediafiles WHERE sfile='" & f & "' COLLATE NOCASE;"
vRet = mdb_Query(sSql)
For Each vI In vRet
    If Mid$(vI, 1, 8) <> "bmissing" Then
        b = vI
        Exit For
    End If
Next vI

mdb_IsMissing = IIf(b = 1, True, False)
End Function

Public Sub mdb_SetMissing(ByVal f As String, bMissing As Boolean)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set bmissing=" & IIf(bMissing, "1", "0") & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Sub mdb_SetdAdded(ByVal f As String, dAdded As Date)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set dadded='" & _
    Format(dAdded, "yyyy-mm-dd hh:mm:ss") & "'" & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Sub mdb_SetdLastPlayed(ByVal f As String, dLastPlayed As Date)
Dim sSql As String

f = Replace(f, "'", "''")
f = us_Encode(f)

sSql = "update tbl_mediafiles set dlastplay='" & _
    Format(dLastPlayed, "yyyy-mm-dd hh:mm:ss") & "'" & _
    " where sfile='" & f & "';"

mdb_Query sSql
End Sub

Public Function mdb_SearchToPL(sTxt As String, plRet As typPl, Optional sEsc As String = "\") As Boolean
Dim sSql As String, mdbInter As typMdb

sSql = "SELECT sfile, dadded, lstartcnt, lendcnt, dlastplay, lmd5, lduration, benabled, bmissing " & _
    "FROM tbl_mediafiles " & _
    "WHERE sfile LIKE '%" & us_Encode(sTxt) & "%' ESCAPE '" & sEsc & "' " & _
    "AND (bmissing=0 OR bmissing is NULL) AND (benabled=1 OR benabled is NULL) " & _
    "ORDER BY sfile COLLATE NOCASE ASC;"

'Debug.Print sSql

mdb_QueryToMdb sSql, mdbInter
pl_MakeFromMdb mdbInter, plRet
End Function

'assumes sSQL query inclures the following fields, in this order:
'sfile, dadded, lstartcnt, lendcnt, dlastplay, lmd5, lduration, benabled, bmissing.
Public Function mdb_QueryToMdb(sSql As String, mdbOut As typMdb) As Boolean
Dim vRet As Variant, vI As Variant, _
    i As Long, x As Long, _
    lQueryRowCnt As Long

Dim lTimer As Long, lQueryTime As Long, lProcTime As Long

mdb_QueryToMdb = False

lTimer = GetTickCount

'Debug.Print sSQL
vRet = mdb_Query(sSql)
lQueryRowCnt = number_of_rows_from_last_call 'does not include extra one for header, so no -1

lQueryTime = GetTickCount - lTimer
lTimer = GetTickCount

If lQueryRowCnt < 1 Then
    mdbOut.lCnt = 0
    ReDim mdbOut.Items(0)
    mdbOut.lIndex = -1
    
    GoTo NoItems
End If

ReDim mdbOut.Items(0 To lQueryRowCnt - 1)
mdbOut.lCnt = lQueryRowCnt

i = 0
x = 0
For Each vI In vRet
    If i > 0 Then
        If Not IsEmpty(vI) Then
            Select Case x
                Case 0: mdbOut.Items(i - 1).sFile = us_Decode((vI))
                Case 1: mdbOut.Items(i - 1).dAdded = vI
                Case 2: mdbOut.Items(i - 1).lStartCnt = vI
                Case 3: mdbOut.Items(i - 1).lEndCnt = vI
                Case 4: mdbOut.Items(i - 1).dLastPlay = vI
                Case 5: mdbOut.Items(i - 1).lMD5 = vI
                Case 6: mdbOut.Items(i - 1).lDuration = vI
                Case 7: mdbOut.Items(i - 1).bEnabled = IIf(vI = "0", False, True)
                Case 8: mdbOut.Items(i - 1).bMissing = IIf(vI = "0", False, True)
            End Select
        Else
            Select Case x
                Case 7: mdbOut.Items(i - 1).bEnabled = True
                Case 8: mdbOut.Items(i - 1).bMissing = False
            End Select
        End If
    End If
    
    i = i + 1
    If i > lQueryRowCnt Then
        i = 0
        x = x + 1
        If x > lQueryRowCnt - 1 Then Exit For
    End If
    
Next vI

lProcTime = GetTickCount - lTimer
Debug.Print "query;  run time=" & lQueryTime & "  proc time=" & lProcTime

NoItems:
mdb_QueryToMdb = True
End Function
