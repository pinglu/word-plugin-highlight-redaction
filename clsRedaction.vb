' Redaction Class
' this builds:
' - build array of all highlight colors
' - build array of highlight colors that are present in this document
' - build array of story ranges that are present
'
'
Private doc As Document

' see build_highlight_color_names
Private highlightColors() As String ' array of all highlight colors
Private redactedCountByColor() As Long ' array of color id => number of redactions

' see build_story_ranges
Private searchStoryRangeAsIntArray() As Integer
Private redactStoryRangeAsIntArray() As Integer

' see get_used_highlight_colors
Private usedHighlightColorsAsIndex() As String
Private usedHighlightColorsAsName() As String

' array of all the user wants to replace (selected)
' see setRedactColors
Private toRedactColorsAsIndex() As String
Private toRedactColorsAsName() As String

' see build_redacted_colors_with_count_string
Private redactedColorsWithCount As String
Private redactedColorsCount As Long

Private clsForm As userForm1

Public Sub Startup()

    Set doc = ActiveDocument

    Dim searchNumberOfFirstPages As Integer
    searchNumberOfFirstPages = 5
        
    ' initialize all highlight colors
    ' @highlightColors() As String
    ' @redactedCountByColor() as Long
    build_highlight_color_names
    
    ' initialize story ranges (where to search and redact)
    ' @searchStoryRangeAsIntArray() As Integer
    ' @redactStoryRangeAsIntArray() As Integer
    build_story_ranges
    
    ' get colors from document
    ' @usedHighlightColorsAsIndex
    ' @usedHighlightColorsAsName
    get_used_highlight_colors (searchNumberOfFirstPages)
    
End Sub

' ##################################################################################
' ################ Building Arrays of Highlight Color & Redacted Color Count
' ##################################################################################

Public Property Get getHighlightColors() As String()
    getHighlightColors = highlightColors
End Property

Private Function build_highlight_color_names()

    ReDim highlightColors(16)
    ReDim redactedCountByColor(16)
    
    ' indexes are WdColorIndex
    highlightColors(1) = "black"
    highlightColors(2) = "blue"
    highlightColors(3) = "turquoise"
    highlightColors(4) = "bGreen"
    highlightColors(5) = "pink"
    highlightColors(6) = "red"
    highlightColors(7) = "yellow"
    highlightColors(8) = "white"
    highlightColors(9) = "dBlue"
    highlightColors(10) = "teal"
    highlightColors(11) = "green"
    highlightColors(12) = "violet"
    highlightColors(13) = "dRed"
    highlightColors(14) = "dYellow"
    highlightColors(15) = "gray50"
    highlightColors(16) = "gray25"
    
    For i = 1 To 16
        redactedCountByColor(i) = 0
    Next i
    
End Function

' ##################################################################################
' ################ Building Arrays of Story Range (where to search and redact)
' ##################################################################################

Public Property Get getStoryRangeSearchAsIntArray() As Integer()
    getStoryRangeSearchAsIntArray = searchStoryRangeAsIntArray
End Property

Public Property Get getRedactStoryRangeAsIntArray() As Integer()
    getRedactStoryRangeAsIntArray = redactStoryRangeAsIntArray
End Property

Private Function build_story_ranges()
    ' use replace = false for where to search for highlights and replace = true for where to replace
    '     the idea is: in header or footer there might be an explanation of all relevant colors, search there, but don't replace them
    ' @return array of SearchStory Indexes (1:Mainbody, and else if found in document)
    ' loop through story types in word
    ' wdMainTextStory / 1 for Mainbody
    ' wdFootnotesStory / 2 for Footnotes
    ' wdTextFrameStory / 5 for Textboxes
    ' https://learn.microsoft.com/en-us/office/vba/api/word.wdstorytype
    Dim searchStoryCount As Integer
    Dim redactStoryCount As Integer
    Dim story As Range
    
    searchStoryCount = 0
    redactStoryCount = 0
    For Each story In doc.StoryRanges
        Select Case story.storyType
            Case wdMainTextStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            ' only use on replace
            Case wdFootnotesStory
                'searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdTextFrameStory
                'searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            ' Only search when getting highlighting colors
            ' wdEvenPagesHeaderStory: 6
            ' wdPrimaryHeaderStory:   7
            ' wdEvenPagesFooterStory: 8
            ' wdPrimaryFooterStory:   9
            ' wdFirstPageHeaderStory: 10
            ' wdFirstPageFooterStory: 11
            Case wdEvenPagesHeaderStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdPrimaryHeaderStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdEvenPagesFooterStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdPrimaryFooterStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdFirstPageHeaderStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
            Case wdFirstPageFooterStory
                searchStoryCount = update_story_range_search(story.storyType, searchStoryCount)
                'redactStoryCount = update_story_range_redact(story.storyType, redactStoryCount)
        End Select
    Next story
End Function

Private Function update_story_range_search(storyType As Integer, numberOfStoryTypes As Integer) As Integer
    ReDim Preserve searchStoryRangeAsIntArray(numberOfStoryTypes)
    searchStoryRangeAsIntArray(numberOfStoryTypes) = storyType
    numberOfStoryTypes = numberOfStoryTypes + 1
    
    update_story_range_search = numberOfStoryTypes
End Function

Private Function update_story_range_redact(storyType As Integer, numberOfStoryTypes As Integer) As Integer
    ReDim Preserve redactStoryRangeAsIntArray(numberOfStoryTypes)
    redactStoryRangeAsIntArray(numberOfStoryTypes) = storyType
    numberOfStoryTypes = numberOfStoryTypes + 1
    
    update_story_range_redact = numberOfStoryTypes
End Function

' ##################################################################################
' ################ Get all used highlight colors from document
' ##################################################################################

' array of all used in document (present in document)
' > as index
Public Property Get getUsedHighlightColorsAsIndex() As String()
    getUsedHighlightColorsAsIndex = usedHighlightColorsAsIndex
End Property

' > as name
Public Property Get getUsedHighlightColorsAsName() As String()
    getUsedHighlightColorsAsName = usedHighlightColorsAsName
End Property

' get colorId by index
Public Property Get getUsedHighlightColorIdByIndex(index As Integer) As String
    getUsedHighlightColorIdByIndex = usedHighlightColorsAsIndex(index)
End Property

' get name by index
Public Property Get getUsedHighlightColorNameByIndex(index As Integer) As String
    getUsedHighlightColorNameByIndex = usedHighlightColorsAsName(index)
End Property

' build array of all used highlight colors
Private Function get_used_highlight_colors(searchNumberOfFirstPages As Integer)
    Dim currentPosition As Range
    Dim colorCounter As Integer
    colorCounter = 0
    
    Dim currentHighlightColor As String
    Dim multipleHighlightPrompt As String
    ' 16 colors
    ReDim usedHighlightColorsAsIndex(15)
    ReDim usedHighlightColorsAsName(15)
    
    multipleHighlightPrompt = ""
    
    Dim i As Integer
    Dim text As String
    For i = 0 To UBound(searchStoryRangeAsIntArray)
        Set currentPosition = doc.StoryRanges(searchStoryRangeAsIntArray(i))
        
        ' find all colors under searchNumberOfFirstPages
        With currentPosition.Find
            .Highlight = True
            Do While .Execute(FindText:="", Forward:=True, Format:=True) = True And currentPosition.Information(wdActiveEndPageNumber) < searchNumberOfFirstPages
                If currentPosition.HighlightColorIndex = -1 Then
                    text = "Using Custom Highlight Color! Make sure to only use one, as we cannot distingush between different custom highlight colors. Best to use the 16 default ones. Google for Word wdcolorindex"
                    log_text (text)
                End If
                If currentPosition.HighlightColorIndex > 16 Then
                ' only tell user, if its not in main text
                    If currentPosition.storyType <> wdMainTextStory Then
                        multipleHighlightPrompt = multipleHighlightPrompt & "- Page " & currentPosition.Information(wdActiveEndPageNumber) & ": " & Left(currentPosition.text, 100) & vbCrLf
                    End If
                Else
                    currentHighlightColor = LTrim(Str(currentPosition.HighlightColorIndex))
                    If is_in_array(currentHighlightColor, usedHighlightColorsAsIndex) = False Then
                        usedHighlightColorsAsIndex(colorCounter) = currentHighlightColor
                        usedHighlightColorsAsName(colorCounter) = highlightColors(currentHighlightColor)
                        colorCounter = colorCounter + 1
                    End If
                End If
            Loop
        End With
    Next
    
    If colorCounter = 0 Then
        log_text ("No highlights found in first " & searchNumberOfFirstPages & " pages")
    Else
        ReDim Preserve usedHighlightColorsAsIndex(colorCounter - 1)
        ReDim Preserve usedHighlightColorsAsName(colorCounter - 1)
    End If

    If multipleHighlightPrompt <> "" Then
        log_text ("The following Text is highlighted multiple times / nested. This will be ignored and have to be cleaned manually. " & vbCrLf & "Examples: " & vbCrLf & multipleHighlightPrompt)
    End If
End Function

' ##################################################################################
' ################ Building Array of user selected array to redact
' ##################################################################################

Public Property Get getToRedactColorsAsIndex() As String()
    getToRedactColorsAsIndex = toRedactColorsAsIndex
End Property

Public Property Get getToRedactColorsAsName() As String()
    getToRedactColorsAsName = toRedactColorsAsName
End Property

' build array of colors that the user wants to redact
Public Sub setRedactColors(redactColorsIndex() As String, redactColorsName() As String, counter As Integer)    
    ReDim toRedactColorsAsIndex(counter)
    ReDim toRedactColorsAsName(counter)
    
    For i = 0 To counter
        toRedactColorsAsIndex(i) = redactColorsIndex(i)
        toRedactColorsAsName(i) = redactColorsName(i)
    Next i
End Sub

' ##################################################################################
' ################ Redacted Color Count, getter and add 1
' ##################################################################################
Public Property Get getRedactedCountByColor() As Long()
    getRedactedCountByColor = redactedCountByColor
End Property

Public Sub addRedactedCountByColor(color As String)
    redactedCountByColor(color) = redactedCountByColor(color) + 1
End Sub

' ##################################################################################
' ################ build array of redacted colors with count
' ################ string = "yellow (10), red (20)"
' ##################################################################################

Public Property Get getRedactedColorsWithCount() As String
	'first build the array
    build_redacted_colors_with_count_string
    
    getRedactedColorsWithCount = redactedColorsWithCount
End Property

Public Property Get getTotalRedactedCount() As Long
    getTotalRedactedCount = redactedColorsCount
End Property

Public Function resetRedactedColorCounts()
    redactedColorsWithCount = ""
    redactedColorsCount = 0
End Function

' this builds the string colors (count) and the total count
Private Function build_redacted_colors_with_count_string() As String
    Dim i As Integer
    Dim pRedactedColorsCount As Long
    
    For i = 0 To UBound(toRedactColorsAsIndex)
        redactedColorsWithCount = redactedColorsWithCount & toRedactColorsAsName(i) & " (" & redactedCountByColor(toRedactColorsAsIndex(i)) & "), "
        pRedactedColorsCount = pRedactedColorsCount + redactedCountByColor(toRedactColorsAsIndex(i))
    Next i
    
    redactedColorsWithCount = Left(redactedColorsWithCount, Len(redactedColorsWithCount) - 2)
    redactedColorsCount = pRedactedColorsCount
End Function

Public Property Set setForm(myForm As userForm1)
    Set clsForm = myForm
End Property

Private Function log_text(text As String)
    clsForm.logBox.text = logBox.text & text & vbCrLf
End Function
