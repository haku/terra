Attribute VB_Name = "modPL"
'functions relating to the manipulation of typPl objects.
'mostly these have been moved from frmMain.
'i will move the rest as it as and when i can.
Option Explicit

'##########################################################
'## setting up new lists ##################################
'##########################################################

Public Sub pl_SetAsNew(List As typPl, Optional n As String = "list")
ReDim List.Items(0)
List.lCnt = 0
List.lIndex = -1
List.sName = n
List.lCurrent = -1
List.bEnab = True
List.lScroll = 0
List.bArc = False
pl_RecountTotalDuration List
End Sub

Public Sub pl_RecountTotalDuration(List As typPl)
Dim i As Long

List.lTotalDuration = 0
List.bTotalDurationComplete = True

If List.lCnt < 1 Then Exit Sub

For i = 0 To List.lCnt - 1
    List.lTotalDuration = List.lTotalDuration + List.Items(i).lDuration
    If List.Items(i).lDuration <= 0 Then List.bTotalDurationComplete = False
Next i
End Sub

'##########################################################
'## adding items to lists #################################
'##########################################################

Public Sub pl_AddItem(List As typPl, sFile As String, _
    Optional lMD5 As Long = 0, _
    Optional lSrartCnt As Long = 0, Optional lEndCnt As Long = 0, _
    Optional lDuration As Long = 0, Optional dLastPlay As Date = Empty, _
    Optional lSource0 As Long = -1, Optional lSource1 As Long = -1)

If List.lCnt = 0 Then
    ReDim List.Items(0 To 0)
    List.lCnt = 1
Else
    ReDim Preserve List.Items(LBound(List.Items) To UBound(List.Items) + 1)
    List.lCnt = UBound(List.Items) + 1
End If

List.Items(List.lCnt - 1).sFile = sFile
List.Items(List.lCnt - 1).lMD5 = lMD5
List.Items(List.lCnt - 1).bSel = False
List.Items(List.lCnt - 1).lStartCnt = lSrartCnt
List.Items(List.lCnt - 1).lEndCnt = lEndCnt
List.Items(List.lCnt - 1).lDuration = lDuration
List.Items(List.lCnt - 1).dLastPlay = dLastPlay

List.Items(List.lCnt - 1).lSource(0) = lSource0
List.Items(List.lCnt - 1).lSource(1) = lSource1

pl_RecountTotalDuration List
End Sub

Public Sub pl_AddDir(List As typPl, sFld As String)
On Error GoTo pl_AddDir_err

Dim i As Long, cFlds As Collection, cFiles As Collection

Set cFlds = New Collection
BuildDirCollection sFld, cFlds

Set cFiles = New Collection
For i = 1 To cFlds.count
    AddFilesToCollection cFlds(i), cFiles
Next i

If cFiles.count < 1 Then Exit Sub

For i = 1 To cFiles.count
    If IsMediaFile(cFiles(i)) Then
        pl_AddItem List, cFiles(i)
    End If
Next i

Exit Sub
pl_AddDir_err:
Main_Err "pl_AddDir_err."
err.Clear
End Sub

'##########################################################
'## GUI stuff relating to lists ###########################
'##########################################################

Public Sub pl_SelAll(List As typPl)
Dim i As Long
If List.lCnt < 1 Then Exit Sub
For i = 0 To List.lCnt - 1
    List.Items(i).bSel = True
Next i
End Sub

Public Sub pl_DebuildScroll(vsbSB As VScrollBar)
vsbSB.Enabled = False
vsbSB.Min = 0
vsbSB.Max = 0
End Sub

Public Sub pl_RebuildScroll(List As typPl, lH As Long, vsbSB As VScrollBar)
Dim lTotalH As Long

'todo: currently using item height from media lib.
lTotalH = List.lCnt * mdbl_ItmH

If lTotalH > lH Then
    vsbSB.Min = 0
    vsbSB.Max = List.lCnt - (lH / mdbl_ItmH) + 1
    vsbSB.LargeChange = lH / mdbl_ItmH
    vsbSB.SmallChange = 1
    If vsbSB.Enabled <> True Then vsbSB.Enabled = True
Else
    If vsbSB.Enabled <> False Then vsbSB.Enabled = False
    If vsbSB.Value <> 0 Then vsbSB.Value = 0
End If
End Sub

Public Sub pl_SetScroll(vsbSB As VScrollBar, ByVal v As Long)
If vsbSB.Enabled Then
    If v >= vsbSB.Max Then
        v = vsbSB.Max
    ElseIf v <= vsbSB.Min Then
        v = vsbSB.Min
    End If
    vsbSB.Value = v
End If
End Sub

'##########################################################
'## searching lists #######################################
'##########################################################

Public Function pl_FindIndexFromFile(List As typPl, sFile As String) As Long
Dim i As Long, b As Boolean, f As String

If List.lCnt < 1 Then Exit Function

f = LCase$(sFile)
b = False

For i = 0 To List.lCnt - 1
    If LCase$(List.Items(i).sFile) = f Then
        b = True
        Exit For
    End If
Next i

If b Then
    pl_FindIndexFromFile = i
Else
    pl_FindIndexFromFile = -1
End If
End Function

'##########################################################
'## manipulating lists ####################################
'##########################################################

Public Sub pl_SwapItems(List As typPl, lFirst As Long, lSecond As Long, _
    Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1)

Dim iFirst As typPlItem, iSecond As typPlItem

iFirst = List.Items(lFirst)
iSecond = List.Items(lSecond)

List.Items(lFirst) = iSecond
List.Items(lSecond) = iFirst

If List.lCurrent = lFirst Then
    List.lCurrent = lSecond
ElseIf List.lCurrent = lSecond Then
    List.lCurrent = lFirst
End If

If bUpdateCue Then
    frmMain.cue_ItemMightHaveChanged lListI, lFirst, lSecond
    
'    If cue_ItemInCue(lListI, lFirst) Then
'        cue_ItemChanged lListI, lFirst, lSecond
'    End If
'    If cue_ItemInCue(lListI, lSecond) Then
'        cue_ItemChanged lListI, lSecond, lFirst
'    End If
End If
End Sub

Public Sub pl_Reverse(List As typPl, _
    Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1)

Dim i As Long

For i = 0 To (List.lCnt - 1) / 2
    pl_SwapItems List, i, (List.lCnt - 1) - i, bUpdateCue, lListI
Next i
End Sub

Public Sub pl_Sort(List As typPl, lSortBy As Long, Optional l_First As Long = -1, Optional l_Last As Long = -1, _
    Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1)

'lSortBy:
'   0=path
'   1=playcount - started
'   2=playcount - finished
'   3=dLastPlay

Dim l_Low As Long, l_Middle As Long, l_High As Long, v_Test As Variant

If l_First = -1 Then l_First = 0
If l_Last = -1 Then l_Last = List.lCnt - 1

If l_First < l_Last Then
    l_Middle = (l_First + l_Last) / 2
    l_Low = l_First
    l_High = l_Last
    
    Select Case lSortBy
        Case 0: v_Test = LCase$(List.Items(l_Middle).sFile)
        Case 1: v_Test = List.Items(l_Middle).lStartCnt
        Case 2: v_Test = List.Items(l_Middle).lEndCnt
        Case 3: v_Test = List.Items(l_Middle).dLastPlay
    End Select
    
    Do
        Select Case lSortBy
            Case 0 'path
                While LCase$(List.Items(l_Low).sFile) < v_Test
                    l_Low = l_Low + 1
                Wend
                While LCase$(List.Items(l_High).sFile) > v_Test
                    l_High = l_High - 1
                Wend
            
            Case 1 'start count
                While List.Items(l_Low).lStartCnt < v_Test
                    l_Low = l_Low + 1
                Wend
                While List.Items(l_High).lStartCnt > v_Test
                    l_High = l_High - 1
                Wend
            
            Case 2 'end count
                While List.Items(l_Low).lEndCnt < v_Test
                    l_Low = l_Low + 1
                Wend
                While List.Items(l_High).lEndCnt > v_Test
                    l_High = l_High - 1
                Wend
            
            Case 3 'last played date
                While List.Items(l_Low).dLastPlay < v_Test
                    l_Low = l_Low + 1
                Wend
                While List.Items(l_High).dLastPlay > v_Test
                    l_High = l_High - 1
                Wend
            
        End Select
        
        If (l_Low <= l_High) Then
            pl_SwapItems List, l_Low, l_High, bUpdateCue, lListI
            l_Low = l_Low + 1
            l_High = l_High - 1
        End If
    Loop While (l_Low <= l_High)
    
    If l_First < l_High Then pl_Sort List, lSortBy, l_First, l_High, bUpdateCue, lListI
    If l_Low < l_Last Then pl_Sort List, lSortBy, l_Low, l_Last, bUpdateCue, lListI
End If
End Sub

Public Sub pl_MoveItems(List As typPl, iDir As Integer, Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1)
Dim i As Long, a As Long, b As Long, s As Long

If List.lCnt < 1 Then Exit Sub

If iDir = 0 Then
    a = 0
    b = UBound(List.Items)
    s = 1
Else
    a = UBound(List.Items)
    b = 0
    s = -1
End If

For i = a To b Step s
    If List.Items(i).bSel = False Then GoTo NextI
    
    If (i <= 0 And iDir = 0) _
        Or (i >= UBound(List.Items) And iDir = 1) _
        Then GoTo NextI
    
    pl_SwapItems List, i, i + IIf(iDir = 0, -1, 1), bUpdateCue, lListI
    
NextI:
Next i
End Sub

Public Sub pl_RemoveItem(List As typPl, lIndex As Long, _
    Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1)

Dim i As Long

If List.lCnt < 1 Then Exit Sub

If List.lCurrent = lIndex Then
    'note: when active item is deleted, this makes it it go back to the start of the list.
    List.lCurrent = -1
End If

If lIndex < UBound(List.Items) Then
    For i = lIndex To UBound(List.Items) - 1
        pl_SwapItems List, i, i + 1, bUpdateCue, lListI
    Next i
End If

If bUpdateCue Then
    frmMain.cue_ItemMightHaveChanged lListI, UBound(List.Items), -1
'    If cue_ItemInCue(lListI, UBound(List.Items)) Then
'        cue_ItemChanged lListI, UBound(List.Items), -1
'    End If
End If

If UBound(List.Items) > 0 Then
    ReDim Preserve List.Items(UBound(List.Items) - 1)
    List.lCnt = UBound(List.Items) + 1
    If List.lIndex > UBound(List.Items) Then List.lIndex = -1
    pl_RecountTotalDuration List
Else
    pl_SetAsNew List, List.sName
End If
End Sub

'##########################################################
'## user input ############################################
'##########################################################

Public Sub pl_MouseDown(List As typPl, vsB As VScrollBar, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim cI As Long, i As Long, d As Long

If Button <> 1 Then Exit Sub
If List.lCnt < 1 Then Exit Sub
If Y < 0 Then Exit Sub

cI = vsB.Value + Fix(Y / mdbl_ItmH)

If cI > List.lCnt - 1 Then
    cI = -1
    For i = 0 To List.lCnt - 1
        List.Items(i).bSel = False
    Next i
Else
    If Shift = 0 Then
        For i = 0 To List.lCnt - 1
            List.Items(i).bSel = False
        Next i
        List.Items(cI).bSel = Not List.Items(cI).bSel
    ElseIf (Shift And vbCtrlMask) Then
        List.Items(cI).bSel = Not List.Items(cI).bSel
    ElseIf (Shift And vbShiftMask) And List.lIndex >= 0 Then
        d = IIf(List.lIndex <= cI, 1, -1)
        For i = List.lIndex + d To cI Step d
            List.Items(i).bSel = Not List.Items(i).bSel
        Next i
    End If
End If

If cI >= 0 Then List.lIndex = cI
End Sub

Public Function pl_KeyDown(List As typPl, vsB As VScrollBar, lHeight As Long, _
    KeyCode As Integer, Shift As Integer, _
    Optional bUpdateCue As Boolean = False, Optional lListI As Long = -1) As Boolean

Dim i As Long

Select Case KeyCode
    Case vbKeyPageUp, vbKeyPageDown
        If (Shift And vbCtrlMask) Then
            pl_MoveItems List, IIf(KeyCode = vbKeyPageUp, 0, 1), bUpdateCue, lListI
            pl_KeyDown = True
        Else
            pl_SetScroll vsB, vsB.Value + _
                (vsB.LargeChange * IIf(KeyCode = vbKeyPageUp, -1, 1))
        End If
    
    Case vbKeyUp, vbKeyDown
        i = List.lIndex + IIf(KeyCode = vbKeyUp, -1, 1)
        If i < 0 Then
            i = 0
        ElseIf i > List.lCnt - 1 Then
            i = List.lCnt - 1
        End If
        
        If (Shift And vbShiftMask) Then
            'List.Items(List.lIndex).bSel = Not List.Items(List.lIndex).bSel
            List.Items(i).bSel = List.Items(List.lIndex).bSel
        End If
        
        List.lIndex = i
        
        If i < vsB.Value Then
            pl_SetScroll vsB, i
        ElseIf i > vsB.Value + Fix(lHeight / mdbl_ItmH) - 1 Then
            pl_SetScroll vsB, i - Fix(lHeight / mdbl_ItmH) + 1
        End If
        
        pl_KeyDown = True
    
    Case vbKeySpace
        If (Shift And vbShiftMask) Then
            
        ElseIf (Shift And vbCtrlMask) Then
            
        Else
            For i = 0 To List.lCnt - 1
                List.Items(i).bSel = False
            Next i
        End If
        If List.lIndex >= 0 Then
            List.Items(List.lIndex).bSel = _
                Not List.Items(List.lIndex).bSel
            pl_KeyDown = True
        End If
    
End Select
End Function

'##########################################################
'## converting lists ######################################
'##########################################################

Public Function pl_MakeFromMdb(mdbIn As typMdb, plOut As typPl) As Boolean
Dim i As Long

pl_MakeFromMdb = False

pl_SetAsNew plOut

If mdbIn.lCnt < 1 Then GoTo LeaveSub

For i = 0 To mdbIn.lCnt - 1
    pl_AddItem plOut, _
        mdbIn.Items(i).sFile, _
        mdbIn.Items(i).lMD5, _
        mdbIn.Items(i).lStartCnt, _
        mdbIn.Items(i).lEndCnt, _
        mdbIn.Items(i).lDuration, _
        mdbIn.Items(i).dLastPlay
    
Next i

LeaveSub:
pl_MakeFromMdb = True
End Function

'##########################################################
'## drawing ###############################################
'##########################################################

Public Sub pl_Blank(lHdc As Long, lW As Long, lH As Long, _
    Optional sEmptyText As String = "no list selected.")

On Error GoTo pl_Blank_err

Dim rcMain As RECT

rcMain.Left = 0
rcMain.Top = 0
rcMain.right = lW
rcMain.bottom = lH

'cls the buffer
FillRect lHdc, rcMain, gdi_Main_Brush(0)

'set drawing param
SelectObject lHdc, gdi_Main_hFontNormal
SetTextColor lHdc, GetSysColor(COLOR_BTNTEXT)
SetBkMode lHdc, TRANSPARENT

'draw text
DrawText lHdc, sEmptyText, Len(sEmptyText), rcMain, DT_CENTER

Exit Sub
pl_Blank_err:
Main_Err "pl_Blank_err."
err.Clear
End Sub

'~~ notes for pl_Draw ~~
'   sColums: 0 for hidden, 1 for visible.  in this order:
'   duration,hash code,date last played,play counts

Public Sub pl_Draw( _
    List As typPl, lFrom As Long, _
    lHdc As Long, lW As Long, lH As Long, _
    Optional bFocused As Boolean = False, _
    Optional sEmptyText As String = "this playlist is empty.", _
    Optional sColums As String = "0000", _
    Optional bHideExt As Boolean = False, _
    Optional lTrunk As Long = 0, _
    Optional bTrunkReverse As Boolean = False, _
    Optional bShowIndex As Boolean = True)

On Error GoTo pl_Draw_err

Dim rcMain As RECT, rc As RECT, rc2 As RECT
Dim s As String, i As Long, x As Long, b As Boolean
Dim lColL(0 To 4) As Long, lFlg As Long

rcMain.Left = 0
rcMain.Top = 0
rcMain.right = lW
rcMain.bottom = lH

'cls the buffer
FillRect lHdc, rcMain, gdi_Main_Brush(0)

'set drawing param
SelectObject lHdc, gdi_Main_hFontNormal
SetTextColor lHdc, GetSysColor(COLOR_BTNTEXT)
SetBkMode lHdc, TRANSPARENT

'is there anything to draw?====================================================
If List.lCnt < 1 Then
    s = sEmptyText
    DrawText lHdc, s, Len(s), rcMain, DT_CENTER
    GoTo NoItems
End If

'check the position of the focus box===========================================
'moved to "got focus" event
'If List.lIndex < lFrom Then
'    List.lIndex = lFrom
'ElseIf List.lIndex > lFrom + Fix(lH / mdbl_ItmH) - 1 Then
'    List.lIndex = lFrom + Fix(lH / mdbl_ItmH) - 1
'End If

'calculate columns=============================================================
SelectObject lHdc, gdi_Main_hFontNormal

'duration
If Mid(sColums, 1, 1) = "1" Then
    s = "000:00"
    DrawText lHdc, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(4) = lW - rc.right * dListCellPaddingH
Else
    lColL(4) = lW
End If

'hash code
If Mid(sColums, 2, 1) = "1" Then
    s = "DDDDDDDD"
    DrawText lHdc, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(3) = lColL(4) - rc.right * dListCellPaddingH
Else
    lColL(3) = lColL(4)
End If

'date last played
If Mid(sColums, 3, 1) = "1" Then
    s = Format(Now, sDateFormatString)
    DrawText lHdc, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(2) = lColL(3) - rc.right * dListCellPaddingH
Else
    lColL(2) = lColL(3)
End If

'play counts
If Mid(sColums, 4, 1) = "1" Then
    s = "000/000"
    DrawText lHdc, s, Len(s), rc, dt_left + DT_CALCRECT
    lColL(1) = lColL(2) - rc.right * dListCellPaddingH
Else
    lColL(1) = lColL(2)
End If

lColL(0) = 1 'leave a 1 px gap at the left side.

'draw the items================================================================
'item area rect; these parts don't change.
rc.Left = 0
rc.right = rcMain.right

For i = lFrom To List.lCnt - 1
    'item area
    rc.Top = (i - lFrom) * mdbl_ItmH
    rc.bottom = rc.Top + mdbl_ItmH
    
    'do we need to draw the selectiion bg?  also change forecolour.
    If List.Items(i).bSel Then
        FillRect lHdc, rc, gdi_Main_Brush(1)
        SetTextColor lHdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
    Else
        SetTextColor lHdc, GetSysColor(COLOR_BTNTEXT)
    End If
    
    'is this item playing?
    'If pb_GetPlayState > 0 And cMedia.FileName = List.Items(i).sFile Then
    'note: the "is playing" bit has been removed delibratly.  afterall, the play state is easily seen from the toolbar area.
    
    If List.lCurrent = i Then
        SelectObject lHdc, gdi_Main_hFontSel
    Else
        SelectObject lHdc, gdi_Main_hFontNormal
    End If
    
    'draw the item's text
    rc2 = rc
    For x = 0 To UBound(lColL)
        Select Case x
            Case 1: b = IIf(Mid(sColums, 4, 1) = "1", True, False) 'playcount
            Case 2: b = IIf(Mid(sColums, 3, 1) = "1", True, False) 'date last played
            Case 3: b = IIf(Mid(sColums, 2, 1) = "1", True, False) 'hash
            Case 4: b = IIf(Mid(sColums, 1, 1) = "1", True, False) 'duration
            Case Else: b = True
        End Select
        If Not b Then GoTo NextX
        
        rc2.Left = lColL(x)
        If x = UBound(lColL) Then
            rc2.right = rc.right
        Else
            rc2.right = lColL(x + 1)
        End If
        
        Select Case x
            Case 0 'file name
                s = IIf(bShowIndex, i, "") & " "
                
                If List.Items(i).lSource(0) = 0 Then
                    s = s & "lib / "
                ElseIf List.Items(i).lSource(0) > 0 Then
                    s = s & mdb_PL(List.Items(i).lSource(0) - 1).sName & " / "
                End If
                
                s = s & frmMain.gen_GetShowName(List.Items(i).sFile)
                lFlg = dt_left
            
            Case 1 'play count
                If List.Items(i).lStartCnt > 0 Then
                    s = List.Items(i).lStartCnt & "/" & List.Items(i).lEndCnt
                    lFlg = DT_CENTER
                Else
                    s = ""
                End If
            
            Case 2 'd last play
                s = IIf(List.Items(i).dLastPlay < 1, "", Format(List.Items(i).dLastPlay, sDateFormatString))
                'If Len(s) > 0 Then Debug.Print s
                lFlg = dt_left
            
            Case 3 'hash code
                If List.Items(i).lMD5 <> 0 Then
                    s = Hex(List.Items(i).lMD5)
                    lFlg = DT_CENTER
                Else
                    s = ""
                End If
            
            Case 4 'l duration
                If List.Items(i).lDuration <> 0 Then
                    s = ConvertSecToMin(List.Items(i).lDuration)
                    lFlg = DT_RIGHT
                Else
                    s = ""
                End If
            
        End Select
        
        If Len(s) > 0 Then
            DrawText lHdc, s, Len(s), rc2, DT_NOPREFIX + DT_VCENTER + lFlg
        End If
NextX:
    Next x
    
    'draw list index box?
    If List.lIndex = i And bFocused Then
        FrameRect lHdc, rc, gdi_Main_Brush(2)
    End If
    
    'at end of display area?
    If rc.bottom > lH Then Exit For
    
Next i

NoItems:
Exit Sub
pl_Draw_err:
Main_Err "pl_Draw_err."
err.Clear
End Sub

'##########################################################
'## read / write ##########################################
'##########################################################

Public Sub pl_Read(List As typPl, f As String)
On Error GoTo pl_Read_err

Dim i As Long, a As String, b As String, c, _
    arr() As String, arrLine() As String, arrDate() As String, _
    sF As String, lH As Long, lStartCnt As Long, lEndCnt As Long, _
    lDuration As Long, dLastPlay As Date

pl_SetAsNew List
List.sFilePath = f

a = UnicodeFile_Read_FSO(f)
arr = Split(a, vbNewLine)

For i = 0 To UBound(arr)
    If Mid$(arr(i), 1, 1) <> "#" Then GoTo NextI
    
    a = Mid$(arr(i), 1, InStr(1, arr(i), "="))
    b = Mid$(arr(i), Len(a) + 1)
    
    Select Case a
        Case "#listname="
            List.sName = b
        
        Case "#listcount="
            'not a lot to do here.
        
        Case "#listenab="
            List.bEnab = IIf(Mid$(b, 1, 1) = "0", False, True)
        
        Case "#listscroll="
            List.lScroll = Val(b)
        
        Case "#listarc="
            List.bArc = IIf(Mid$(b, 1, 1) = "0", False, True)
        
        Case "#file="
            arrLine = Split(b, "|")
            
            sF = arrLine(0)
            
            If UBound(arrLine) >= 1 Then
                lH = Val(arrLine(1))
            Else
                lH = 0
            End If
            
            If UBound(arrLine) >= 2 Then
                lStartCnt = Val(arrLine(2))
            Else
                lStartCnt = 0
            End If
            
            If UBound(arrLine) >= 3 Then
                lEndCnt = Val(arrLine(3))
            Else
                lEndCnt = 0
            End If
            
            If UBound(arrLine) >= 4 Then
                lDuration = Val(arrLine(4))
            Else
                lDuration = 0
            End If
            
            If UBound(arrLine) >= 5 Then
                If Len(arrLine(5)) > 0 Then
                    arrDate = Split(arrLine(5), "-", , vbTextCompare)
                    c = arrDate(2) & "/" & arrDate(1) & "/" & arrDate(0) & " " & _
                        arrDate(3) & ":" & arrDate(4) & ":" & arrDate(5)
                    dLastPlay = c
                Else
                    dLastPlay = Empty
                End If
            Else
                dLastPlay = Empty
            End If
            
            pl_AddItem List, sF, lH, lStartCnt, lEndCnt, lDuration, dLastPlay
        
    End Select
    
NextI:
Next i

Exit Sub
pl_Read_err:
Main_Err "pl_Read_err."
err.Clear
End Sub

Public Sub pl_Write(List As typPl, f As String)
On Error GoTo pl_Write_err

Dim a As New cStringBuilder, i As Long

a.Append "#listname=" & List.sName & vbNewLine & _
    "#listcount=" & List.lCnt & vbNewLine & _
    "#listenab=" & IIf(List.bEnab, "1", "0") & vbNewLine & _
    "#listscroll=" & List.lScroll & vbNewLine & _
    "#listarc=" & IIf(List.bArc, "1", "0") & vbNewLine

For i = 0 To List.lCnt - 1
    a.Append "#file=" & List.Items(i).sFile & "|" & _
        List.Items(i).lMD5 & "|" & _
        List.Items(i).lStartCnt & "|" & _
        List.Items(i).lEndCnt & "|" & _
        List.Items(i).lDuration & "|" & _
        IIf(List.Items(i).dLastPlay < 1, "", Format(List.Items(i).dLastPlay, "yyyy-mm-dd-hh-mm-ss")) & vbNewLine & _
        List.Items(i).sFile & vbNewLine
Next i

UnicodeFile_Write_FSO f, a.ToString

Exit Sub
pl_Write_err:
Main_Err "pl_Write_err."
err.Clear
End Sub

'##########################################################
'##########################################################
'##########################################################
