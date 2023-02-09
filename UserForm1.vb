Private formRedaction As clsRedaction
Private formDoc As Document

Private Sub CommandButton2_Click()
    ' Starting
    formRedaction.resetRedactedColorCounts
    Application.ScreenUpdating = False
    
    Dim startRedactionPage As Integer
    startRedactionPage = CInt(TB_startRedactionPage.text)
    
    If Me.LB_colorsToRedact.ListCount = 0 Then
        log_text ("please load colors or start macro from redactionMacro")
        GoTo EndRedaction
    End If

    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' XXXXXXXXXXXXXX get user selected colors
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    build_user_color_selection_array
    userColorSelectionArray = formRedaction.getToRedactColorsAsIndex
    redactStoryRangeArray = formRedaction.getRedactStoryRangeAsIntArray
    
    If UBound(userColorSelectionArray) = 0 And userColorSelectionArray(0) = "" Then
        log_text ("No colors selected. Use CTRL and left mouse to select multiple. Exiting...")
        GoTo EndRedaction
    End If
    
    Dim currentHighlightColor As String
    
    ' in stories, check if the color is in the array the user has asked us, if so, replace
    For i = 0 To UBound(redactStoryRangeArray)
        Set currentPosition = formDoc.StoryRanges(redactStoryRangeArray(i))
        reset_search_parameters currentPosition
        With currentPosition.Find
        .Highlight = True
            Do While .Execute(FindText:="", Forward:=True, Format:=True) = True
                ' Start redaction from page 2?
                If currentPosition.Information(wdActiveEndPageNumber) < startRedactionPage Then
                    GoTo skipReplace
                End If
                
                currentHighlightColor = LTrim(Str(currentPosition.HighlightColorIndex))
                If is_in_array(currentHighlightColor, userColorSelectionArray) = True Then
                    ' replace!
                    Dim myRange As Range
                    Set myRange = currentPosition
                    check_and_redact_range myRange
                    currentPosition.Collapse wdCollapseEnd
                ElseIf currentHighlightColor = "9999999" Then
                    If currentPosition.storyType <> wdMainTextStory Then
                        ' save location of multiple highlights
                        multipleHighlightsText = multipleHighlightsText & "Page " & currentPosition.Information(wdActiveEndPageNumber) & ": " & Left(currentPosition.text, 100) & vbCrLf
                        GoTo skipReplace
                    Else
                        ' multiple highlights detected, find begining and end of correct highlight colors
                        go_through_chars_to_redact_multiple_highlights currentPosition
                    End If
                End If

skipReplace:
            Loop
        End With
    Next
    
    If multipleHighlightsText <> "" Then
        log_text ("The following text was highlighted multiple times (a highlighted text with another highlight color inside). This was ignored / not replaced. Please take note and manually clean up. " & vbCrLf & vbCrLf & multipleHighlightsText)
    End If

	' save and finish up    
    Dim fileSuffix As String
	fileSuffix = TB_fileSuffix.text
	
	' also sets the active document / formDoc to the orignal file!
    save_file fileSuffix
    send_finish_log fileSuffix
    
EndRedaction:
    Application.ScreenUpdating = True
    formDoc.Activate
End Sub

Private Sub go_through_chars_to_redact_multiple_highlights(currentRange As Variant)
    Dim replaceStartPos As Integer
    Dim prevHighlightColor As String
    Dim currentHighlightColor As String

    userColorSelectionArray = formRedaction.getToRedactColorsAsIndex

    replaceStartPos = 0
    prevHighlightColor = ""
    Set activeStoryRange = currentRange
    For Each Char In activeStoryRange.Characters
        currentHighlightColor = LTrim(Str(Char.HighlightColorIndex))
        
        ' Char should be replaced
        If is_in_array(currentHighlightColor, userColorSelectionArray) = True Then
            ' no replace start pos, this is the first char of the highlighted text
            If replaceStartPos = 0 Then
                replaceStartPos = Char.start
            ElseIf currentHighlightColor <> prevHighlightColor Then
            ' not the first character, but colors changed to another to be replaced Character
                check_and_redact_range formDoc.Range(start:=replaceStartPos, End:=Char.start)
                ' this is set to zero, because we're skipping all the characters from the replaced string, and will pass by the replaceStartPos = 0 if clause
                replaceStartPos = Char.End - 1
            End If
            prevHighlightColor = currentHighlightColor
        Else
            ' if Char is not highligted AND (but the prev chars where highlighted / there was a replaceStartPos), then replace the string
            If (replaceStartPos <> 0) Then
                check_and_redact_range formDoc.Range(start:=replaceStartPos, End:=Char.start)
                replaceStartPos = 0
            End If
        End If
    Next Char
    
    If (replaceStartPos <> 0) Then
        check_and_redact_range formDoc.Range(start:=replaceStartPos, End:=currentRange.End)
        replaceStartPos = 0
    End If
End Sub

' !!! RECURSIVE FUNCTION !!!
'
' this will go check if in the current range there is a footnote or field reference.
' - if so, it will split the range and call itself / repeat until
' - if there is not footnote or field ref inside the range: call redact_text to really redact the text
Private Function check_and_redact_range(currentRange As Range)

    ' get the current highlight color
    ' we need this for counting
    Dim highlightColor As Integer
    highlightColor = currentRange.Characters(1).HighlightColorIndex
    
    ' perform check since we have the color here anyway
    If (highlightColor < 1) Or (highlightColor > 16) Then
        log_text "Tried to redact string at " & currentRange.start & " with " & Left(currentRange.text, 10) & ", but no or individual color detected. Cannot redact it. Skipping."
        Exit Function
    End If

    ' getting redaction text
    Dim redactionText As String
    redactionText = TB_redactionText.text

    Dim footnoteOrFieldFound As Boolean
    footnoteOrFieldFound = False
    
    For Each Footnote In formDoc.Footnotes
        
        If (currentRange.start <= Footnote.Reference.start) And (Footnote.Reference.End <= currentRange.End) Then
            Dim firstRange As Range
            Dim secondRange As Range
            
            footnoteOrFieldFound = True
            'first range could have another field that comes later, as array is not ordered!
            Set firstRange = formDoc.Range(start:=currentRange.start, End:=Footnote.Reference.start)
            ' perform check since we have the color here anyway
            highlightColor = firstRange.Characters(1).HighlightColorIndex
            If (highlightColor < 1) Or (highlightColor > 16) Then
                log_text "Tried to redact string at " & currentRange.start & " with " & Left(currentRange.text, 10) & ", but no or individual color detected. Cannot redact it. Skipping."
                Exit Function
            End If
            firstRange.text = redactionText
            formRedaction.addRedactedCountByColor LTrim(Str(highlightColor))
            
            ' second Range: define and run check again
            Set secondRange = formDoc.Range(start:=Footnote.Reference.End, End:=currentRange.End)
            check_and_redact_range secondRange
        End If
    Next
    
    For Each field In currentRange.fields
        ' there are fields in this range:
        Set firstRange = formDoc.Range(start:=currentRange.start, End:=field.Range.start)
            check_and_redact_range firstRange
        Set secondRange = formDoc.Range(start:=field.Range.End, End:=currentRange.End)
            check_and_redact_range firstRange
    Next

    If footnoteOrFieldFound = False Then
        currentRange.text = redactionText
        formRedaction.addRedactedCountByColor LTrim(Str(highlightColor))
    End If
End Function

Private Function redact_text(currentRange As Range)
    ' get the current highlight color
    ' we need this for counting
    Dim highlightColor As Integer
    highlightColor = currentRange.Characters(1).HighlightColorIndex
    
    ' perform check since we have the color here anyway
    If (highlightColor < 1) Or (highlightColor > 16) Then
        log_text "Tried to redact string at " & currentRange.start & " with " & Left(currentRange.text, 10) & ", but no or individual color detected. Cannot redact it. Skipping."
        Exit Function
    End If

    ' getting redaction text
    Dim redactionText As String
    redactionText = TB_redactionText.text

    currentRange.text = redactionText
    formRedaction.addRedactedCountByColor LTrim(Str(highlightColor))
End Function

Private Function reset_search_parameters(oRng As Variant)
  With oRng.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .text = ""
    .Replacement.text = ""
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute
  End With
End Function

Private Function save_file(fileSuffix As String)
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    ' XXXXXXXXXXXXXX GET Filename
    ' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    originalDocumentName = formDoc.Name
    
    intPos = InStrRev(originalDocumentName, ".")
    
    strPath = formDoc.Path
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    ' build file name, either from user color selection or if the user has provided a file name use that one.
    newDocumentName = Left(originalDocumentName, intPos - 1) & "-" & Left(fileSuffix, 20) & ".docx"
        
    ' saving new document
    formDoc.SaveAs2 fileName:=strPath & newDocumentName, FileFormat:=wdFormatDocumentDefault
    
    ' open original
    Documents.Open(strPath & originalDocumentName).Activate
    Set formDoc = ActiveDocument
End Function

Private Sub send_finish_log(fileSuffix As String)
    ' trim the last comma and add a point
    sColorConcat = formRedaction.getRedactedColorsWithCount & ". Total redactions: " & formRedaction.getTotalRedactedCount
    
    log_text "***************************************************************"
    log_text "Redacted Version (" & fileSuffix & ") - redacted Colors: " & sColorConcat
    log_text "***************************************************************"

End Sub

Private Sub UserForm_Initialize()
    log_text "Ready to go."
End Sub

Property Set setformRedaction(ByRef redaction As clsRedaction)
    Set formRedaction = redaction
End Property

Property Get getformRedaction() As clsRedaction
    Set getformRedaction = formRedaction
End Property

Property Set setFormDoc(ByRef doc As Document)
    Set formDoc = doc
End Property

Property Get getFormDoc() As Document
    Set getDoc = formDoc
End Property

Public Sub build_user_color_selection_array()
    Dim selectedColorText As String
    Dim toRedactColorsAsIndex() As String
    Dim toRedactColorsAsName() As String
    
    ReDim toRedactColorsAsIndex(Me.LB_colorsToRedact.ListCount - 1)
    ReDim toRedactColorsAsName(Me.LB_colorsToRedact.ListCount - 1)
    
    Dim counter As Integer
    Dim color As Integer
    
    counter = 0
    For color = 0 To Me.LB_colorsToRedact.ListCount - 1
        If Me.LB_colorsToRedact.Selected(color) = True Then
            selectedColorText = Me.LB_colorsToRedact.List(color)
            
            Select Case selectedColorText
                Case "black"
                    toRedactColorsAsIndex(counter) = wdBlack
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "blue"
                    toRedactColorsAsIndex(counter) = wdBlue
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "turquoise"
                    toRedactColorsAsIndex(counter) = wdTurquoise
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "bGreen"
                    toRedactColorsAsIndex(counter) = wdBrightGreen
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "pink"
                    toRedactColorsAsIndex(counter) = wdPink
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "red"
                    toRedactColorsAsIndex(counter) = wdRed
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "yellow"
                    toRedactColorsAsIndex(counter) = wdYellow
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "white"
                    toRedactColorsAsIndex(counter) = wdWhite
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "dBlue"
                    toRedactColorsAsIndex(counter) = wdDarkBlue
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "teal"
                    toRedactColorsAsIndex(counter) = wdTeal
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "green"
                    toRedactColorsAsIndex(counter) = wdGreen
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "violet"
                    toRedactColorsAsIndex(counter) = wdViolet
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "dRed"
                    toRedactColorsAsIndex(counter) = wdDarkRed
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "dYellow"
                    toRedactColorsAsIndex(counter) = wdDarkYellow
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "gray50"
                    toRedactColorsAsIndex(counter) = wdGray50
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
                Case "gray25"
                    toRedactColorsAsIndex(counter) = wdGray25
                    toRedactColorsAsName(counter) = selectedColorText
                    counter = counter + 1
            End Select
        End If
        
    Next color
    
    ReDim Preserve toRedactColorsAsIndex(counter - 1)
    ReDim Preserve toRedactColorsAsName(counter - 1)
    
    Dim fRedaction As clsRedaction
    Set fRedaction = Me.getformRedaction
    
    fRedaction.setRedactColors toRedactColorsAsIndex, toRedactColorsAsName, counter - 1
End Sub

Private Function log_text(text As String)
    logBox.text = logBox.text & text & vbCrLf
End Function
