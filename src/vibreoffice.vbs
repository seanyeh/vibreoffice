' vibreoffice - Vi Mode for LibreOffice/OpenOffice
'
' The MIT License (MIT)
'
' Copyright (c) 2014 Sean Yeh
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

Option Explicit

' --------
' Globals
' --------
global VIBREOFFICE_STARTED as boolean ' Defaults to False
global VIBREOFFICE_ENABLED as boolean ' Defaults to False

global oXKeyHandler as object

' Global State
global MODE as string
global VIEW_CURSOR as object
global MULTIPLIER as integer

' -----------
' Singletons
' -----------
Function getCursor
    getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    dim oTextCursor
    oTextCursor = getCursor().getText.createTextCursorByRange(getCursor())
    ' oTextCursor.gotoRange(oTextCursor.getStart(), False)

    getTextCursor = oTextCursor
End Function

' -----------------
' Helper Functions
' -----------------
Sub restoreStatus 'restore original statusbar
    dim oLayout
    oLayout = thisComponent.getCurrentController.getFrame.LayoutManager
    oLayout.destroyElement("private:resource/statusbar/statusbar")
    oLayout.createElement("private:resource/statusbar/statusbar")
End Sub

Sub setRawStatus(rawText)
    thisComponent.Currentcontroller.StatusIndicator.Start(rawText, 0)
End Sub

Sub setStatus(statusText)
    setRawStatus(MODE & " | " & statusText & " | special: " & getSpecial() & " | " & "modifier: " & getMovementModifier())
End Sub

Sub setMode(modeName)
    MODE = modeName
    setRawStatus(modeName)
End Sub

Function gotoMode(sMode)
    Select Case sMode
        Case "NORMAL":
            setMode("NORMAL")
            setMovementModifier("")
        Case "INSERT":
            setMode("INSERT")
        Case "VISUAL":
            setMode("VISUAL")

            dim oTextCursor
            oTextCursor = getTextCursor()
            ' Deselect TextCursor
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            ' Show TextCursor selection
            thisComponent.getCurrentController.Select(oTextCursor)
    End Select
End Function

Sub cursorReset(oTextCursor)
    oTextCursor.gotoRange(oTextCursor.getStart(), False)
    oTextCursor.goRight(1, False)
    oTextCursor.goLeft(1, True)
    thisComponent.getCurrentController.Select(oTextCursor)
End Sub

Function samePos(oPos1, oPos2)
    samePos = oPos1.X() = oPos2.X() And oPos1.Y() = oPos2.Y()
End FUnction

Function genString(sChar, iLen)
    dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function


' -----------------------------------
' Special Mode (for chained commands)
' -----------------------------------
global SPECIAL_MODE As string
global SPECIAL_COUNT As integer

Sub setSpecial(specialName)
    SPECIAL_MODE = specialName

    If specialName = "" Then
        SPECIAL_COUNT = 0
    Else
        SPECIAL_COUNT = 2
    End If
End Sub

Function getSpecial()
    getSpecial = SPECIAL_MODE
End Function

Sub delaySpecialReset()
    SPECIAL_COUNT = SPECIAL_COUNT + 1
End Sub

Sub resetSpecial(Optional bForce)
    If IsMissing(bForce) Then bForce = False

    SPECIAL_COUNT = SPECIAL_COUNT - 1
    If SPECIAL_COUNT <= 0 Or bForce Then
        setSpecial("")
    End If
End Sub


' -----------------
' Movement Modifier
' -----------------
'f,i,a
global MOVEMENT_MODIFIER As string

Sub setMovementModifier(modifierName)
    MOVEMENT_MODIFIER = modifierName
End Sub

Function getMovementModifier()
    getMovementModifier = MOVEMENT_MODIFIER
End Function


' --------------------
' Multiplier functions
' --------------------
Sub _setMultiplier(n as integer)
    MULTIPLIER = n
End Sub

Sub resetMultiplier()
    _setMultiplier(0)
End Sub

Sub addToMultiplier(n as integer)
    dim sMultiplierStr as String
    dim iMultiplierInt as integer

    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 then
        sMultiplierStr = CStr(MULTIPLIER) & CStr(n)
        _setMultiplier(CInt(sMultiplierStr))
    End If
End Sub

' Should only be used if you need the raw value
Function getRawMultiplier()
    getRawMultiplier = MULTIPLIER
End Function

' Same as getRawMultiplier, but defaults to 1 if it is unset (0)
Function getMultiplier()
    If MULTIPLIER = 0 Then
        getMultiplier = 1
    Else
        getMultiplier = MULTIPLIER
    End If
End Function


' -------------
' Key Handling
' -------------
Sub sStartXKeyHandler
    sStopXKeyHandler()

    oXKeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
    thisComponent.CurrentController.AddKeyHandler(oXKeyHandler)
End Sub

Sub sStopXKeyHandler
    thisComponent.CurrentController.removeKeyHandler(oXKeyHandler)
End Sub

Sub XKeyHandler_Disposing(oEvent)
End Sub


' --------------------
' Main Key Processing
' --------------------
function KeyHandler_KeyPressed(oEvent) as boolean
    ' Exit if plugin is not enabled
    If IsMissing(VIBREOFFICE_ENABLED) Or Not VIBREOFFICE_ENABLED Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

    dim bConsumeInput, bIsMultiplier, bIsModified, bIsSpecial
    bConsumeInput = True ' Block all inputs by default
    bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    bIsSpecial = getSpecial() <> ""


    ' --------------------------
    ' Process global shortcuts, exit if matched (like ESC)
    If ProcessGlobalKey(oEvent) Then
        ' Pass

    ' If INSERT mode, allow all inputs
    ElseIf MODE = "INSERT" Then
        bConsumeInput = False

    ' If Change Mode
    ' ElseIf MODE = "NORMAL" And Not bIsSpecial And getMovementModifier() = "" And ProcessModeKey(oEvent) Then
    ElseIf ProcessModeKey(oEvent) Then
        ' Pass

    ' Replace Key
    ElseIf getSpecial() = "r" And Not bIsModified Then
        dim iLen
        iLen = Len(getCursor().getString())
        getCursor().setString(genString(oEvent.KeyChar, iLen))

    ' Multiplier Key
    ElseIf ProcessNumberKey(oEvent) Then
        bIsMultiplier = True
        delaySpecialReset()

    ' Normal Key
    ElseIf ProcessNormalKey(oEvent.KeyChar, oEvent.Modifiers) Then
        ' Pass

    ' If is modified but doesn't match a normal command, allow input
    '   (Useful for built-in shortcuts like Ctrl+s, Ctrl+w)
    ElseIf bIsModified Then
        bConsumeInput = False

    ' Movement modifier here?
    ElseIf ProcessMovementModifierKey(oEvent.KeyChar) Then
        delaySpecialReset()

    ' If bIsSpecial but nothing matched, return to normal mode
    ElseIf bIsSpecial Then
        gotoMode("NORMAL")

    ' Allow non-letter keys if unmatched
    ElseIf asc(oEvent.KeyChar) = 0 Then
        bConsumeInput = False
    End If
    ' --------------------------

    ' Reset Special
    resetSpecial()

    ' Reset multiplier if last input was not number and not in special mode
    If not bIsMultiplier and getSpecial() = "" and getMovementModifier() = "" Then
        resetMultiplier()
    End If
    setStatus(getMultiplier())

    KeyHandler_KeyPressed = bConsumeInput
End Function

Function KeyHandler_KeyReleased(oEvent) As boolean
    dim oTextCursor

    ' Show terminal-like cursor
    oTextCursor = getTextCursor()
    If oEvent.Modifiers = 2 Or oEvent.Modifiers = 8 And oEvent.KeyChar = "c" Then
        ' Allow Ctrl+c for Copy, so don't change cursor
        ' Pass
    ElseIf MODE = "NORMAL" Then
        cursorReset(oTextCursor)
    ElseIf MODE = "INSERT" Then
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        thisComponent.getCurrentController.Select(oTextCursor)
    End If

    KeyHandler_KeyReleased = (MODE = "NORMAL") 'cancel KeyReleased
End Function


' ----------------
' Processing Keys
' ----------------
Function ProcessGlobalKey(oEvent)
    dim bMatched, bIsControl
    bMatched = True
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)

    ' PRESSED ESCAPE (or ctrl+[)
    if oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Then
        ' Move cursor back if was in INSERT (but stay on same line)
        If MODE <> "NORMAL" And Not getCursor().isAtStartOfLine() Then
            getCursor().goLeft(1, False)
        End If

        resetSpecial(True)
        gotoMode("NORMAL")
    Else
        bMatched = False
    End If
    ProcessGlobalKey = bMatched
End Function


Function ProcessNumberKey(oEvent)
    dim c
    c = CStr(oEvent.KeyChar)

    If c >= "0" and c <= "9" Then
        addToMultiplier(CInt(c))
        ProcessNumberKey = True
    Else
        ProcessNumberKey = False
    End If
End Function


Function ProcessModeKey(oEvent)
    dim bIsModified
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    ' Don't change modes in these circumstances
    If MODE <> "NORMAL" Or bIsModified Or getSpecial <> "" Or getMovementModifier() <> "" Then
        ProcessModeKey = False
        Exit Function
    End If

    ' Mode matching
    dim bMatched
    bMatched = True
    Select Case oEvent.KeyChar
        ' Insert modes
        Case "i", "a", "I", "A", "o", "O":
            If oEvent.KeyChar = "a" Then getCursor().goRight(1, False)
            If oEvent.KeyChar = "I" Then ProcessMovementKey("^")
            If oEvent.KeyChar = "A" Then ProcessMovementKey("$")

            If oEvent.KeyChar = "o" Then
                ProcessMovementKey("$")
                ProcessMovementKey("l")
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    getCursor().setString(chr(13) & chr(13))
                    ProcessMovementKey("l")
                End If
            End If

            If oEvent.KeyChar = "O" Then
                ProcessMovementKey("^")
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    ProcessMovementKey("h")
                    getCursor().setString(chr(13))
                    ProcessMovementKey("l")
                End If
            End If

            gotoMode("INSERT")
        Case "v":
            gotoMode("VISUAL")
        Case Else:
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(keyChar, modifiers)
    dim i, bMatched, bIsVisual, iIterations, bIsControl
    bIsControl = (modifiers = 2) or (modifiers = 8)

    bIsVisual = (MODE = "VISUAL") ' is this hardcoding bad? what about visual block?

    ' ----------------------
    ' 1. Check Movement Key
    ' ----------------------
    iIterations = getMultiplier()
    bMatched = False
    For i = 1 To iIterations
        dim bMatchedMovement

        ' Movement Key
        bMatchedMovement = ProcessMovementKey(KeyChar, bIsVisual, modifiers)
        bMatched = bMatched or bMatchedMovement


        ' If Special: d/c + movement
        If bMatched And (getSpecial() = "d" Or getSpecial() = "c") Then
            getTextCursor().setString("")
        End If
    Next i

    ' Reset Movement Modifier
    setMovementModifier("")

    ' Exit already if movement key was matched
    If bMatched Then
        ' If Special: d/c : change mode
        If getSpecial() = "d" Then gotoMode("NORMAL")
        If getSpecial() = "c" Then gotoMode("INSERT")

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 2. Undo/Redo
    ' --------------------
    If keyChar = "u" Or (bIsControl And keyChar = "r") Then
        For i = 1 To iIterations
            Undo(keyChar = "u")
        Next i

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 3. Check Special/Delete Key
    ' --------------------

    ' There are no special/delete keys with modifier keys, so exit early
    If modifiers > 1 Then
        ProcessNormalKey = False
        Exit Function
    End If

    ' Only 'x' or Special (dd, cc) can be done more than once
    If keyChar <> "x" and getSpecial() = "" Then
        iIterations = 1
    End If
    For i = 1 To iIterations
        dim bMatchedSpecial

        ' Special/Delete Key
        bMatchedSpecial = ProcessSpecialKey(keyChar)

        bMatched = bMatched or bMatchedSpecial
    Next i


    ProcessNormalKey = bMatched
End Function


' Function for both undo and redo
Sub Undo(bUndo)
    On Error Goto ErrorHandler

    If bUndo Then
        thisComponent.getUndoManager().undo()
    Else
        thisComponent.getUndoManager().redo()
    End If
    Exit Sub

    ' Ignore errors from no more undos/redos in stack
ErrorHandler:
    Resume Next
End Sub


Function ProcessSpecialKey(keyChar)
    dim oTextCursor, bMatched, bIsSpecial
    bMatched = True
    bIsSpecial = getSpecial() <> ""


    If keyChar = "d" Or keyChar = "c" Or keyChar = "s" Then
        ' Special Cases: 'dd' and 'cc'
        If bIsSpecial Then
            dim bIsSpecialCase
            bIsSpecialCase = (keyChar = "d" And getSpecial() = "d") Or (keyChar = "c" And getSpecial() = "c")

            If bIsSpecialCase Then
                ProcessMovementKey("^", False)
                ProcessMovementKey("j", True)

                oTextCursor = getTextCursor()
                thisComponent.getCurrentController.Select(oTextCursor)
                oTextCursor.setString("")
            Else
                bMatched = False
            End If

            ' Go to INSERT mode after 'cc', otherwise NORMAL
            If bIsSpecialCase And keyChar = "c" Then
                gotoMode("INSERT")
            Else
                gotoMode("NORMAL")
            End If


        ' visual mode: delete selection
        ElseIf MODE = "VISUAL" Then
            oTextCursor = getTextCursor()
            thisComponent.getCurrentController.Select(oTextCursor)
            oTextCursor.setString("")

            If keyChar = "c" Or keyChar = "s" Then gotoMode("INSERT")
            If keyChar = "d" Then gotoMode("NORMAL")


        ' Enter Special mode: 'd' or 'c', ('s' => 'cl')
        ElseIf MODE = "NORMAL" Then

            ' 's' => 'cl'
            If keyChar = "s" Then
                setSpecial("c")
                gotoMode("VISUAL")
                ProcessNormalKey("l", 0)
            Else
                setSpecial(keyChar)
                gotoMode("VISUAL")
            End If
        End If

    ' If is 'r' for replace
    ElseIf keyChar = "r" Then
        setSpecial("r")

    ' Otherwise, ignore if bIsSpecial
    ElseIf bIsSpecial Then
        bMatched = False


    ElseIf keyChar = "x" Then
        oTextCursor = getTextCursor()
        thisComponent.getCurrentController.Select(oTextCursor)
        oTextCursor.setString("")

        ' Reset Cursor
        cursorReset(oTextCursor)

        ' Goto NORMAL mode (in the case of VISUAL mode)
        gotoMode("NORMAL")

    ElseIf keyChar = "D" Or keyChar = "C" Then
        If MODE = "VISUAL" Then
            ProcessMovementKey("^", False)
            ProcessMovementKey("$", True)
            ProcessMovementKey("l", True)
        Else
            ' Deselect
            oTextCursor = getTextCursor()
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            thisComponent.getCurrentController.Select(oTextCursor)
            ProcessMovementKey("$", True)
        End If

        getTextCursor().setString("")

        If keyChar = "D" Then
            gotoMode("NORMAL")
        ElseIf keyChar = "C" Then
            gotoMode("INSERT")
        End IF

    ' S only valid in NORMAL mode
    ElseIf keyChar = "S" And MODE = "NORMAL" Then
        ProcessMovementKey("^", False)
        ProcessMovementKey("$", True)
        getTextCursor().setString("")
        gotoMode("INSERT")

    Else
        bMatched = False
    End If

    ProcessSpecialKey = bMatched
End Function


Function ProcessMovementModifierKey(keyChar)
    dim bMatched

    bMatched = True
    Select Case keyChar
        Case "f", "t", "F", "T", "i", "a":
            setMovementModifier(keyChar)
        Case Else:
            bMatched = False
    End Select

    ProcessMovementModifierKey = bMatched
End Function


Function ProcessSearchKey(oTextCursor, searchType, keyChar, bExpand)
    '-----------
    ' Searching
    '-----------
    dim bMatched, oSearchDesc, oFoundRange, bIsBackwards, oStartRange
    bMatched = True
    bIsBackwards = (searchType = "F" Or searchType = "T")

    If Not bIsBackwards Then
        ' VISUAL mode will goRight AFTER the selection
        If MODE <> "VISUAL" Then
            ' Start searching from next character
            oTextCursor.goRight(1, bExpand)
        End If

        oStartRange = oTextCursor.getEnd()
        ' Go back one
        oTextCursor.goLeft(1, bExpand)
    Else
        oStartRange = oTextCursor.getStart()
    End If

    oSearchDesc = thisComponent.createSearchDescriptor()
    oSearchDesc.setSearchString(keyChar)
    oSearchDesc.SearchCaseSensitive = True
    oSearchDesc.SearchBackwards = bIsBackwards

    oFoundRange = thisComponent.findNext( oStartRange, oSearchDesc )

    If not IsNull(oFoundRange) Then
        dim oText, foundPos, curPos, bSearching
        oText = oTextCursor.getText()
        foundPos = oFoundRange.getStart()

        ' Unfortunately, we must go go to this "found" position one character at
        ' a time because I have yet to find a way to consistently move the
        ' Start range of the text cursor and leave the End range intact.
        If bIsBackwards Then
            curPos = oTextCursor.getEnd()
        Else
            curPos = oTextCursor.getStart()
        End If
        do until oText.compareRegionStarts(foundPos, curPos) = 0
            If bIsBackwards Then
                bSearching = oTextCursor.goLeft(1, bExpand)
                curPos = oTextCursor.getStart()
            Else
                bSearching = oTextCursor.goRight(1, bExpand)
                curPos = oTextCursor.getEnd()
            End If

            ' Prevent infinite if unable to find, but shouldn't ever happen (?)
            If Not bSearching Then
                bMatched = False
                Exit Do
            End If
        Loop

        If searchType = "t" Then
            oTextCursor.goLeft(1, bExpand)
        ElseIf searchType = "T" Then
            oTextCursor.goRight(1, bExpand)
        End If

    Else
        bMatched = False
    End If

    ' If matched, then we want to select PAST the character
    ' Else, this will counteract some weirdness. hack either way
    If Not bIsBackwards And MODE = "VISUAL" Then
        oTextCursor.goRight(1, bExpand)
    End If

    ProcessSearchKey = bMatched

End Function


Function ProcessInnerKey(oTextCursor, movementModifier, keyChar, bExpand)
    dim bMatched, searchType1, searchType2, search1, search2

    ' Setting searchType
    If movementModifier = "i" Then
        searchType1 = "T" : searchType2 = "t"
    ElseIf movementModifier = "a" Then
        searchType1 = "F" : searchType2 = "f"
    Else ' Shouldn't happen
        ProcessInnerKey = False
        Exit Function
    End If

    Select Case keyChar
        Case "(", ")", "{", "}", "[", "]", "<", ">", "t", "'", """":
            Select Case keyChar
                Case "(", ")":
                    search1 = "(" : search2 = ")"
                Case "{", "}":
                    search1 = "{" : search2 = "}"
                Case "[", "]":
                    search1 = "[" : search2 = "}"
                Case "<", ">":
                    search1 = "<" : search2 = ">"
                Case "t":
                    search1 = ">" : search2 = "<"
                Case "'":
                    search1 = "'" : search2 = "'"
                Case """":
                    ' Matches "smart" quotes, which is default in libreoffice
                    search1 = "“" : search2 = "”"
            End Select

            dim bMatched1, bMatched2
            bMatched1 = ProcessSearchKey(oTextCursor, searchType1, search1, False)
            bMatched2 = ProcessSearchKey(oTextCursor, searchType2, search2, True)

            bMatched = (bMatched1 And bMatched2)

        Case Else:
            bMatched = False

    End Select

    ProcessInnerKey = bMatched
End Function


' -----------------------
' Main Movement Function
' -----------------------
'   Default: bExpand = False, keyModifiers = 0
Function ProcessMovementKey(keyChar, Optional bExpand, Optional keyModifiers)
    dim oTextCursor, bSetCursor, bMatched
    oTextCursor = getTextCursor()
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        dim bIsControl
        bIsControl = (keyModifiers = 2) or (keyModifiers = 8)

        ' Ctrl+d and Ctrl+u
        If bIsControl and keyChar = "d" Then
            getCursor().ScreenDown(bExpand)
        ElseIf bIsControl and keyChar = "u" Then
            getCursor().ScreenUp(bExpand)
        Else
            bMatched = False
        End If

        ProcessMovementKey = bMatched
        Exit Function
    End If

    ' Set global cursor to oTextCursor's new position if moved
    bSetCursor = True


    ' ------------------
    ' Movement matching
    ' ------------------

    ' ---------------------------------
    ' Special Case: Modified movements
    If getMovementModifier() <> "" Then
        Select Case getMovementModifier()
            ' f,F,t,T searching
            Case "f", "t", "F", "T":
                bMatched  = ProcessSearchKey(oTextCursor, getMovementModifier(), keyChar, bExpand)
            Case "i", "a":
                bMatched = ProcessInnerKey(oTextCursor, getMovementModifier(), keyChar, bExpand)

            Case Else:
                bSetCursor = False
                bMatched = False
        End Select

        If Not bMatched Then
            bSetCursor = False
        End If
    ' ---------------------------------

    ElseIf keyChar = "l" Then
        oTextCursor.goRight(1, bExpand)

    ElseIf keyChar = "h" Then
        oTextCursor.goLeft(1, bExpand)

    ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
    ' So this is a weird hack
    ElseIf keyChar = "k" Then
        'oTextCursor.goUp(1, False)
        getCursor().goUp(1, bExpand)
        bSetCursor = False

    ElseIf keyChar = "j" Then
        'oTextCursor.goDown(1, False)
        getCursor().goDown(1, bExpand)
        bSetCursor = False
    ' ----------

    ElseIf keyChar = "^" Then
        getCursor().gotoStartOfLine(bExpand)
        bSetCursor = False
    ElseIf keyChar = "$" Then
        dim oldPos, newPos
        oldPos = getCursor().getPosition()
        getCursor().gotoEndOfLine(bExpand)
        newPos = getCursor().getPosition()

        ' If the result is at the start of the line, then it must have
        ' jumped down a line; goLeft to return to the previous line.
        '   Except for: Empty lines (check for oldPos = newPos)
        If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
            getCursor().goLeft(1, bExpand)
        End If

        ' maybe eventually cursorGoto... should return True/False for bsetCursor
        bSetCursor = False

    ElseIf keyChar = "w" or keyChar = "W" Then
        oTextCursor.gotoNextWord(bExpand)
    ElseIf keyChar = "b" or keyChar = "B" Then
        oTextCursor.gotoPreviousWord(bExpand)
    ElseIf keyChar = "e" Then
        oTextCursor.gotoEndOfWord(bExpand)

    ElseIf keyChar = ")" Then
        oTextCursor.gotoNextSentence(bExpand)
    ElseIf keyChar = "(" Then
        oTextCursor.gotoPreviousSentence(bExpand)
    ElseIf keyChar = "}" Then
        oTextCursor.gotoNextParagraph(bExpand)
    ElseIf keyChar = "{" Then
        oTextCursor.gotoPreviousParagraph(bExpand)

    Else
        bSetCursor = False
        bMatched = False
    End If

    ' If oTextCursor was moved, set global cursor to its position
    If bSetCursor Then
        getCursor().gotoRange(oTextCursor.getStart(), False)
    End If

    ' If oTextCursor was moved and is in VISUAL mode, update selection
    if bSetCursor and bExpand then
        thisComponent.getCurrentController.Select(oTextCursor)
    end if

    ProcessMovementKey = bMatched
End Function


Sub initVibreoffice
    dim oTextCursor
    ' Initializing
    VIBREOFFICE_STARTED = True
    VIEW_CURSOR = thisComponent.getCurrentController.getViewCursor

    resetMultiplier()
    gotoMode("NORMAL")

    ' Show terminal cursor
    oTextCursor = getTextCursor()
    cursorReset(oTextCursor)

    sStartXKeyHandler()
End Sub


Sub Main
    If Not VIBREOFFICE_STARTED Then
        initVibreoffice()
    End If

    ' Toggle enable/disable
    VIBREOFFICE_ENABLED = Not VIBREOFFICE_ENABLED

    ' Restore statusbar
    If Not VIBREOFFICE_ENABLED Then restoreStatus()
End Sub
