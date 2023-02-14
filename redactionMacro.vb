Sub RedactionTool()
    '
    ' RedactionTool Makro
    '
    ' @author: Ping Lu <mail at ping dot lu>
    ' @since: 31.01.2023
    ' @version: 1.1
    
    Dim thisDoc As Document
    Set thisDoc = ActiveDocument
    
    ' init userForm
    Dim myForm As userForm1
    Set myForm = New userForm1
    
    ' init and start class colorSelection
    Dim redaction As clsRedaction
    Set redaction = New clsRedaction
    Set redaction.setForm = myForm
    redaction.Startup
    
    ' add found colors from colorSelection to userForm
    addColorsToUserForm myForm, redaction
    addInfoTextToUserForm myForm
    
    Set myForm.setformRedaction = redaction
    Set myForm.setFormDoc = thisDoc
        
    myForm.Show vbModeless
    
End Sub

Sub addColorsToUserForm(myForm As userForm1, redaction As clsRedaction)

    Dim color As Integer
    Dim selectedColor As Integer
    
    ' Adding used highlight colors to form
    For color = 0 To UBound(redaction.getUsedHighlightColorsAsName)
        If redaction.getUsedHighlightColorNameByIndex(color) <> "" Then
            myForm.LB_colorsToRedact.AddItem (redaction.getUsedHighlightColorNameByIndex(color))
        End If
    Next
End Sub

Sub addInfoTextToUserForm(myForm As userForm1)
    Dim text As String
    Dim license As String
    
    text = "Scope: " & vbCrLf & vbCrLf & "This macro will redact highlighted text from a word document and save it. " & vbCrLf & vbCrLf & "On startup it will search the first 5 number of pages (text body, header and footers) for colors, then it will search for story ranges and search for fields (cross references and footnotes). " _
    & vbCrLf & "Once the form is loaded (the window you are reading now), the user will choose the file suffix of the new document, the colors to redact, the redaction text and which page to start from. " & vbCrLf _
    & "The macro will then go through all highlights using While Document.StoryRanges(i).Find.Highlight = True to find all highlighted texts in all storyRanges." & vbCrLf & _
    "It will check if the highlighted color is one of the colors the user has selected, if so, it will check if there are references (footnotes, fields) in the block, if found, it will seperate the range into multiple and replace all highlighted text with the redacted text. It will first go through the main text, then through footnotes and textboxes." & vbCrLf & vbCrLf & _
    "This macro will " & vbCrLf & _
    "- redact text in the main body, tables, footnotes and textfields" & vbCrLf & _
    "- redact text in the main body that have highlights within highlights" & vbCrLf & _
    "- will leave footnotes and cross references (fields) within a highlight (will redact in front of and behind the reference) in the main text" & vbCrLf & vbCrLf & _
    "This macro will NOT: " & vbCrLf & _
    "- redact text in multiple highlights in text fields as vba does not seem to be able to use range within those storyTypes. Will notify user." & vbCrLf & _
    "- will not redact the target of cross references, if those are highlighted. Will notify user." & vbCrLf & _
    "- cannot redact text with custom colors (outside of WdColorIndex)" & vbCrLf & _
    "- it also sometimes redacts one highlight as two consecutive highlights (putting in [redacted][redacted] instead of once [redacted])" & vbCrLf & _
    "- it will not redact text in header of footer (but will search for colors in there). This can be changed in code (see clsRedaction.build_story_ranges)" & vbCrLf & vbCrLf & _
    "This macro is created by Ping Lu (mail@ping.lu) in 2023Q1. For License see below"
    
    myForm.TB_Info.text = text
    
    license = "MIT License" & vbCrLf & vbCrLf & _
    "Copyright (c) 2023 Ping Lu" & vbCrLf & vbCrLf & _
    "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:" & vbCrLf & vbCrLf & _
    "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software." & vbCrLf & _
    "THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE."

    myForm.TB_license.text = license
End Sub
