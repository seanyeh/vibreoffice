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
global VISUAL_BASE as object ' Position of line that is first selected when 
                             ' VISUAL_LINE mode is entered

' -----------
' Singletons
' -----------
Function getCursor
    getCursor = VIEW_CURSOR
End Function

Function getTextCursor
    dim oTextCursor
    On Error Goto ErrorHandler
    oTextCursor = getCursor().getText.createTextCursorByRange(getCursor())

    getTextCursor = oTextCursor
    Exit Function

ErrorHandler:
    ' Text Cursor does not work in some instances, such as in Annotations
    getTextCursor = Nothing
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

' Selects the current line and makes it the Visual base line for use with 
' VISUAL_LINE mode.
Function formatVisualBase()
    dim oTextCursor
    oTextCursor = getTextCursor()
    VISUAL_BASE = getCursor().getPosition()

    ' Select the current line by moving cursor to start of the bellow line and 
    ' then back to the start of the current line.
    getCursor().gotoEndOfLine(False)
    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
        If getCursor().goRight(1, False) Then
            getCursor().goLeft(1, True)
        End If
    End If
    getCursor().gotoStartOfLine(True)
End Function

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
        Case "VISUAL_LINE":
            setMode("VISUAL_LINE")
            ' Select the current line and set it as the Visual base line
            formatVisualBase()
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
End Function

Function genString(sChar, iLen)
    dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Yanks selection to system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    dim dispatcher As Object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection()
    dim oTextCursor, dispatcher As Object

    ' Deselect if in NORMAL mode to avoid overwriting the character underneath
    ' the cursor
    If MODE = "NORMAL" Then
        oTextCursor = getTextCursor()
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        thisComponent.getCurrentController.Select(oTextCursor)
    End If

    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(ThisComponent.CurrentController.Frame, ".uno:Paste", "", 0, Array())
End Sub


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
    dim oTextCursor

    ' Exit if plugin is not enabled
    If IsMissing(VIBREOFFICE_ENABLED) Or Not VIBREOFFICE_ENABLED Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

    ' Exit if TextCursor does not work (as in Annotations)
    oTextCursor = getTextCursor()
    If oTextCursor Is Nothing Then
        KeyHandler_KeyPressed = False
        Exit Function
    End If

    dim bConsumeInput, bIsMultiplier, bIsModified, bIsControl, bIsSpecial
    bConsumeInput = True ' Block all inputs by default
    bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)
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

	' Normal Key (must be before MultiplierKey, so that 0 is seen as startOfLine)
    ElseIf ProcessNormalKey(oEvent.KeyChar, oEvent.Modifiers) Then

    ' Multiplier Key
    ElseIf ProcessNumberKey(oEvent) Then
        bIsMultiplier = True
        delaySpecialReset()

            ' Pass

    ' If is modified but doesn't match a normal command, allow input
    '   (Useful for built-in shortcuts like Ctrl+a, Ctrl+s, Ctrl+w)
    ElseIf bIsModified Then
        ' Ctrl+a (select all) sets mode to VISUAL
        If bIsControl And oEvent.KeyChar = "a" Then
            gotoMode("VISUAL")
        End If
        bConsumeInput = False

    ' Movement modifier here?
    ElseIf ProcessMovementModifierKey(oEvent.KeyChar) Then
        delaySpecialReset()

    ' If standard movement key (in VISUAL mode) like arrow keys, home, end
    ElseIf (MODE = "VISUAL" Or MODE = "VISUAL_LINE") And ProcessStandardMovementKey(oEvent) Then
        ' Pass

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
    If oTextCursor Is Nothing Then
        ' Do nothing
    ElseIf oEvent.Modifiers = 2 Or oEvent.Modifiers = 8 And oEvent.KeyChar = "c" Then
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

    ' keycode can be viewed here: http://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1awt_1_1Key.html
    ' PRESSED ESCAPE (or ctrl+[) (or ctrl+C)
    if oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Or (oEvent.KeyCode = 514 And bIsControl) Then
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


Function ProcessStandardMovementKey(oEvent)
    dim c, bMatched
    c = oEvent.KeyCode

    bMatched = True

    If (MODE <> "VISUAL" And MODE <> "VISUAL_LINE") Then
        bMatched = False
        'Pass
    ElseIf c = 1024 Then
        ProcessMovementKey("j", True)
    ElseIf c = 1025 Then
        ProcessMovementKey("k", True)
    ElseIf c = 1026 Then
        ProcessMovementKey("h", True)
    ElseIf c = 1027 Then
        ProcessMovementKey("l", True)
    ElseIf c = 1028 Then
        ProcessMovementKey("^", True)
    ElseIf c = "0" Then
        ProcessMovementKey("0", True) ' key for zero (0)
    ElseIf c = 1029 Then
        ProcessMovementKey("$", True)
    Else
        bMatched = False
    End If

    ProcessStandardMovementKey = bMatched
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
    dim bMatched, oTextCursor
    bMatched = True
    oTextCursor = getTextCursor()
    Select Case oEvent.KeyChar
        ' Insert modes
        Case "i", "a", "I", "A", "o", "O":
            If oEvent.KeyChar = "a" And NOT oTextCursor.isEndOfParagraph() Then getCursor().goRight(1, False)
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
                ProcessMovementKey("0")
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
        Case "V":
            gotoMode("VISUAL_LINE")
        Case Else:
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(keyChar, modifiers)
    dim i, bMatched, bIsVisual, iIterations, bIsControl
    bIsControl = (modifiers = 2) or (modifiers = 8)

    bIsVisual = (MODE = "VISUAL" Or MODE = "VISUAL_LINE") ' is this hardcoding bad? what about visual block?

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
        If bMatched And (getSpecial() = "d" Or getSpecial() = "c" Or getSpecial() = "y") Then
            yankSelection((getSpecial() <> "y"))
        End If
    Next i

    ' Reset Movement Modifier
    setMovementModifier("")

    ' Exit already if movement key was matched
    If bMatched Then
        ' If Special: d/c : change mode
        If getSpecial() = "d" Or getSpecial() = "y" Then gotoMode("NORMAL")
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
    ' 3. Paste
    '   Note: in vim, paste will result in cursor being over the last character
    '   of the pasted content. Here, the cursor will be the next character
    '   after that. Fix?
    ' --------------------
    If keyChar = "p" or keyChar = "P" Then
        dim oTextCursor
        oTextCursor = getTextCursor()
        ' Move cursor right if "p" to paste after cursor
        If keyChar = "p" And NOT oTextCursor().isEndOfParagraph() Then
            ProcessMovementKey("l", False)
        End If

        For i = 1 To iIterations
            pasteSelection()
        Next i

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 4. Check Special/Delete Key
    ' --------------------

    ' There are no special/delete keys with modifier keys, so exit early
    If modifiers > 1 Then
        ProcessNormalKey = False
        Exit Function
    End If

    ' Only 'x' or Special (dd, cc) can be done more than once
    If keyChar <> "x" And keyChar <> "X" And getSpecial() = "" Then
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
    dim oTextCursor, bMatched, bIsSpecial, bIsDelete
    bMatched = True
    bIsSpecial = getSpecial() <> ""


    If keyChar = "d" Or keyChar = "c" Or keyChar = "s" Or keyChar = "y" Then
        bIsDelete = (keyChar <> "y")

        ' Special Cases: 'dd' and 'cc'
        If bIsSpecial Then
            dim bIsSpecialCase
            bIsSpecialCase = (keyChar = "d" And getSpecial() = "d") Or (keyChar = "c" And getSpecial() = "c")

            If bIsSpecialCase Then
                ProcessMovementKey("0", False)
                ProcessMovementKey("j", True)

                oTextCursor = getTextCursor()
                thisComponent.getCurrentController.Select(oTextCursor)
                yankSelection(bIsDelete)
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
        ElseIf MODE = "VISUAL" Or MODE = "VISUAL_LINE" Then
            oTextCursor = getTextCursor()
            thisComponent.getCurrentController.Select(oTextCursor)

            yankSelection(bIsDelete)

            If keyChar = "c" Or keyChar = "s" Then gotoMode("INSERT")
            If keyChar = "d" Or keyChar = "y" Then gotoMode("NORMAL")


        ' Enter Special mode: 'd', 'c', or 'y' ('s' => 'cl')
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

	' gg to go to beginning of text
	ElseIf keyChar = "g" Then
		If bIsSpecial Then
			If getSpecial() = "g" Then
                ' If cursor is to left of current visual selection then select 
                ' from right end of the selection to the start of file.
                ' If cursor is to right of current visual selection then select 
                ' from left end of the selection to the start of file.
                If MODE = "VISUAL" Then
                    dim oldPos
                    oldPos = getCursor().getPosition()
                    getCursor().gotoRange(getCursor().getStart(), True)
                    If NOT samePos(getCursor().getPosition(), oldPos) Then
                        getCursor().gotoRange(getCursor().getEnd(), False)
                    End If

                ' If in VISUAL_LINE mode and cursor is bellow the Visual base 
                ' line then move it to the Visual base line, reformat the 
                ' Visual base line, and move cursor to start of file.
                ElseIf MODE = "VISUAL_LINE" Then
                    Do Until getCursor().getPosition().Y() <= VISUAL_BASE.Y()
                        getCursor().goUp(1, False)
                    Loop
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        formatVisualBase()
                    End If
                End If

                dim bExpand
                bExpand = MODE = "VISUAL" Or MODE = "VISUAL_LINE"
                getCursor().gotoStart(bExpand)
			End If
		ElseIf MODE = "NORMAL" Or MODE = "VISUAL" Or MODE = "VISUAL_LINE" Then
			setSpecial("g")
		End If
			
		
    ' Otherwise, ignore if bIsSpecial
    ElseIf bIsSpecial Then
        bMatched = False


    ElseIf keyChar = "x" Or keyChar = "X" Then
        oTextCursor = getTextCursor()
        If keyChar = "X" And MODE <> "VISUAL" And MODE <> "VISUAL_LINE" Then
            oTextCursor.collapseToStart()
            oTextCursor.goLeft(1, True)
        End If
        thisComponent.getCurrentController.Select(oTextCursor)
        yankSelection(True)

        ' Reset Cursor
        cursorReset(oTextCursor)

        ' Goto NORMAL mode (in the case of VISUAL mode)
        gotoMode("NORMAL")

    ElseIf keyChar = "D" Or keyChar = "C" Then
        If MODE = "VISUAL" Or MODE = "VISUAL_LINE" Then
            ProcessMovementKey("0", False)
            ProcessMovementKey("$", True)
            ProcessMovementKey("l", True)
        Else
            ' Deselect
            oTextCursor = getTextCursor()
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            thisComponent.getCurrentController.Select(oTextCursor)
            ProcessMovementKey("$", True)
        End If

        yankSelection(True)

        If keyChar = "D" Then
            gotoMode("NORMAL")
        ElseIf keyChar = "C" Then
            gotoMode("INSERT")
        End IF

    ' S only valid in NORMAL mode
    ElseIf keyChar = "S" And MODE = "NORMAL" Then
        ProcessMovementKey("0", False)
        ProcessMovementKey("$", True)
        yankSelection(True)
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
        If MODE <> "VISUAL" And MODE <> "VISUAL_LINE" Then
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
    If Not bIsBackwards And (MODE = "VISUAL" Or MODE = "VISUAL_LINE") Then
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

    ElseIf keyChar = "k" Then
        If MODE = "VISUAL_LINE" Then
            ' This variable represents the line that the user last selected.
            dim lastSelected

            ' If cursor is already on or above the Visual base line.
            If getCursor().getPosition().Y() <= VISUAL_BASE.Y() Then
                lastSelected = getCursor().getPosition().Y()
                ' If on Visual base line then format it for selecting above 
                ' lines.
                If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                    getCursor().gotoEndOfLine(False)
                    ' Make sure that cursor is on the start of the line bellow 
                    ' the Visual base line. This is needed to make sure the 
                    ' new line character will be selected.
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        getCursor().goRight(1, False)
                    End If
                End If

                ' Move cursor to start of the line above last selected line.
                Do Until getCursor().getPosition().Y() < lastSelected
                    If NOT getCursor().goUp(1, bExpand) Then
                        Exit Do
                    End If
                Loop
                getCursor().gotoStartOfLine(bExpand)

            ' If cursor is already bellow the Visual base line.
            ElseIf getCursor().getPosition().Y() > VISUAL_BASE.Y() Then
                ' Cursor will be under the last selected line so it needs to 
                ' be moved up before setting lastSelected.
                getCursor().goUp(1, bExpand)
                lastSelected = getCursor().getPosition().Y()
                ' Move cursor up another line to deselect the last selected
                ' line.
                getCursor().goUp(1, bExpand)

                ' For the case when the last selected line was the line bellow 
                ' the Visual base line, simply reformat the Visual base line.
                If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                    formatVisualBase()

                Else
                    ' Make sure that the current line is fully selected.
                    getCursor().gotoEndOfLine(bExpand)

                    ' Make sure cursor is at the start of the line we 
                    ' deselected. It needs to always be bellow the user's 
                    ' selection when under the Visual base line.
                    If getCursor().getPosition().Y() < lastSelected Then
                        getCursor().goRight(1, bExpand)
                    End If
                End If

            End If

        Else
        ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
        ' So this is a weird hack
            'oTextCursor.goUp(1, False)
            getCursor().goUp(1, bExpand)
        End If
        bSetCursor = False

    ElseIf keyChar = "j" Then
        If MODE = "VISUAL_LINE" Then
            ' If cursor is already on or bellow the Visual base line.
            If getCursor().getPosition().Y() >= VISUAL_BASE.Y() Then
                ' If on Visual base line then format it for selecting bellow 
                ' lines.
                If VISUAL_BASE.Y() = getCursor().getPosition().Y() Then
                    getCursor().gotoStartOfLine(False)
                    getCursor().gotoEndOfLine(bExpand)
                    ' Move cursor to next line if not already there.
                    If getCursor().getPosition().Y() = VISUAL_BASE.Y() Then
                        getCursor().goRight(1, bExpand)
                    End If

                End If

                If getCursor().goDown(1, bExpand) Then
                    getCursor().gotoStartOfLine(bExpand)

                ' If cursor is on last line then select from current position 
                ' to end of line.
                Else
                    getCursor().gotoEndOfLine(bExpand)
                End If

            ' If cursor is above the Visual base line.
            ElseIf getCursor().getPosition().Y() < VISUAL_BASE.Y() Then
                ' Move cursor to start of bellow line.
                getCursor().goDown(1, bExpand)
                getCursor().gotoStartOfLine(bExpand)
            End If

        Else
        ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
        ' So this is a weird hack
            'oTextCursor.goDown(1, False)
            getCursor().goDown(1, bExpand)
        End If
        bSetCursor = False

    ' ----------

    ElseIf keyChar = "0" Then
        getCursor().gotoStartOfLine(bExpand)
        bSetCursor = False
    ElseIf keyChar = "^" Then
        dim oldLine
        oldLine = getCursor().getPosition().Y()

        ' Select all of the current line and put it into a string.
        getCursor().gotoEndOfLine(False)
        If getCursor().getPosition.Y() > oldLine Then
            ' If gotoEndOfLine moved cursor to next line then move it back.
            getCursor().goLeft(1, False)
        End If
        getCursor().gotoStartOfLine(True)
        dim s as String
        s = getCursor().String

        ' Undo any changes made to the view cursor, then move to start of 
        ' line. This way any previous selction made by the user will remain.
        getCursor().gotoRange(oTextCursor, False)
        getCursor().gotoStartOfLine(bExpand)

        ' This integer will be used to determine the position of the first 
        ' character in the line that is not a space or a tab.
        dim i as Integer
        i = 1

        ' Iterate through the characters in the string until a character that 
        ' is not a space or a tab is found.
        Do While i <= Len(s)
            dim c
            c = Mid(s,i,1)
            If c <> " " And c <> Chr(9) Then
                Exit Do
            End If
            i = i + 1
        Loop

        ' Move the cursor to the first non space/tab character.
        getCursor().goRight(i - 1, bExpand)
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

    ElseIf keyChar = "G" Then
        If MODE = "VISUAL_LINE" Then
            ' If cursor is above Visual base line then move cursor down to it. 
            Do Until getCursor().getPosition.Y() >= VISUAL_BASE.Y()
                getCursor().goDown(1, False)
            Loop
            ' If cursor is on Visual base line then move it to start of line.
            If getCursor().getPosition.Y() = VISUAL_BASE.Y() Then
                getCursor().gotoStartOfLine(False)
            End If
        End If
        getCursor().gotoEnd(bExpand)
        bSetCursor = False

    ElseIf keyChar = "w" or keyChar = "W" Then
        ' For the case when the user enters "cw":
        If getSpecial() = "c" Then
            ' If the cursor is on a word then delete from the current position to 
            ' the end of the word.
            ' If the cursor is not on a word then delete from the current position 
            ' to the start of the next word or until the end of the paragraph.

            If NOT oTextCursor.isEndOfParagraph() Then
               ' Move cursor to right in case it is already at start or end of 
               ' word.
               oTextCursor.goRight(1, bExpand)
            End If

            Do Until oTextCursor.isEndOfWord() Or oTextCursor.isStartOfWord() Or oTextCursor.isEndOfParagraph()
                oTextCursor.goRight(1, bExpand)
            Loop

        ' For the case when the user enters "w" or "dw":
        Else
            ' Note: For "w", using gotoNextWord would mean that the cursor 
            ' would not be moved to the next word when it involved moving down 
            ' a line and that line happened to begin with whitespace. It would 
            ' also mean that the cursor would not skip over lines that only 
            ' contain whitespace.

            If NOT (getSpecial() = "d" And oTextCursor.isEndOfParagraph()) Then
                ' Move cursor to right in case cursor is already at the start 
                ' of a word. 
                ' Additionally for "w", move right in case already on an empty 
                ' line.
                oTextCursor.goRight(1, bExpand)
            End If

            ' Stop looping when the cursor reaches the start of a word, an empty 
            ' line, or cannot be moved further (reaches end of file).
            ' Additionally, if "dw" then stop looping if end of paragraph is reached.
            Do Until oTextCursor.isStartOfWord() Or (oTextCursor.isStartOfParagraph() And oTextCursor.isEndOfParagraph())
                ' If "dw" then do not delete past the end of the line
                If getSpecial() = "d" And oTextCursor.isEndOfParagraph() Then
                    Exit Do
                ' If "w" then stop advancing cursor if cursor can no longer 
                ' move right
                ElseIf NOT oTextCursor.goRight(1, bExpand) Then
                    Exit Do
                End If
            Loop
        End If
    ElseIf keyChar = "b" or keyChar = "B" Then
        ' When the user enters "b", "cb", or "db":

        ' Note: The function gotoPreviousWord causes a lot of problems when 
        ' trying to emulate vim behavior. The following method doesn't have to 
        ' account for as many special cases.

        ' "b": Moves the cursor to the start of the previous word or until an empty 
        ' line is reached.

        ' "db": Does same thing as "b" only it deletes everything between the 
        ' orginal cursor position and the new cursor position. The exception to 
        ' this is that if the original cursor position was at the start of a 
        ' paragraph and the new cursor position is on a separate paragraph with 
        ' at least two words then don't delete the new line char to the "left" 
        ' of the original paragraph.

        ' "dc": Does the same as "db" only the new line char described in "db" 
        ' above is never deleted.


        ' This variable is used to tell whether or not we need to make a 
        ' distinction between "b", "cb", and "db".
        dim dc_db as boolean

        ' Move cursor to left in case cursor is already at the start of a word 
        ' or on on an empty line. If cursor can move left and user enterd "dc" 
        ' or "db" and the cursor was originally on the start of a paragraph 
        ' then set dc_db to true and unselect the new line character separating 
        ' the paragraphs. If cursor can't move left then there is no line above 
        ' the current one and no need to make a distinction between "b", "cb", 
        ' and "db".
        dc_db = False
        If oTextCursor.isStartOfParagraph() And oTextCursor.goLeft(1, bExpand) Then
            If getSpecial() = "c" Or getSpecial() = "d" Then
                dc_db = True
                ' If all conditions above are met then unselect the \n char.
                oTextCursor.collapseToStart()
            End If
        End If

        ' Stop looping when the cursor reaches the start of a word, an empty 
        ' line, or cannot be moved further (reaches start of file).
        Do Until oTextCursor.isStartOfWord() Or (oTextCursor.isStartOfParagraph() And oTextCursor.isEndOfParagraph())
            ' Stop moving cursor if cursor can no longer move left
            If NOT oTextCursor.goLeft(1, bExpand) Then
                Exit Do
            End If
        Loop

        If dc_db Then
            ' Make a clone of oTextCursor called oTextCursor2 and use it to 
            ' check if there are at least two words in the "new" paragraph. 
            ' If there are <2 words then the loop will stop when the cursor 
            ' cursor reaches the start of a paragraph. If there >=2 words then 
            ' then the loop will stop when the cursor reaches the end of a word.
            dim oTextCursor2
            oTextCursor2 = getCursor().getText.createTextCursorByRange(oTextCursor)
            Do Until oTextCursor2.isEndOfWord() Or oTextCursor2.isStartOfParagraph()
                oTextCursor2.goLeft(1, bExpand)
            Loop
            ' If there are less than 2 words on the "new" paragraph then set 
            ' oTextCursor to oTextCursor 2. This is because vim's behavior is 
            ' to clear the "new" paragraph under these conditions.
            If oTextCursor2.isStartOfParagraph() Then
                oTextCursor = oTextCursor2
                oTextCursor.gotoRange(oTextCursor.getStart(), bExpand)
                ' If user entered "db" then reselect the \n char from before.
                If getSpecial() = "d" Then
                    oTextCursor.goRight(1, bExpand)
                End If
            End If
        End If
    ElseIf keyChar = "e" Then
        ' When the user enters "e", "ce", or "de":

        ' The function gotoNextWord causes a lot of problems when trying to 
        ' emulate vim behavior. The following method doesn't have to account 
        ' for as many special cases.

        ' Moves the cursor to the end of the next word or end of file if there 
        ' are no more words.


        ' Move cursor to right in case cursor is already at the end of a word.
        oTextCursor.goRight(1, bExpand)

        ' gotoEndOfWord gets stuck sometimes so manually moving the cursor 
        ' right is necessary in these cases.
        Do Until oTextCursor.gotoEndOfWord(bExpand)
            ' If cursor can no longer move right then break loop
            If NOT oTextCursor.goRight(1, bExpand) Then
                Exit Do
            End If
        Loop

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

    If oTextCursor Is Nothing Then
        ' Do nothing
    Else
        cursorReset(oTextCursor)
    End If

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
