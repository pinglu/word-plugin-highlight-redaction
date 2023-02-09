Sub RedactionTool()
    '
    ' RedactionTool Makro
    '
    ' @author: Ping Lu <mail@ping.lu>
    ' @since: 31.01.2023
    ' @version: 1.0
    
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
    For color = 0 To UBound(redaction.getUsedHighlightColorsAsName) - 1
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
    "- redact text in the main body that have highlights within highligts" & vbCrLf & _
    "- will leave footnotes and cross references (fields) within a highlight (will redact in front of and behind the reference) in the main text" & vbCrLf & vbCrLf & _
    "This macro will NOT: " & vbCrLf & _
    "- redact text in multiple highlights in textfields or footnotes as vba does not seem to be able to use range within those storyTypes" & vbCrLf & _
    "- cannot redact text with custom colors (outside of WdColorIndex)" & vbCrLf & _
    "- it also sometimes redacts one highlight as two consecutive highlights (putting in [redacted][redacted] instead of once [redacted])" & vbCrLf & _
    "- it will not redact text in header of footer (but will search for colors in there). This can be changed in code (see clsRedaction.build_story_ranges)" & vbCrLf & vbCrLf & _
    "This macro is created by Ping Lu (mail@ping.lu) in 2023Q1. For License see below"
    
    myForm.TB_Info.text = text
    
    license = "Copyright (C) 2023 Ping Lu" & vbCrLf & vbCrLf & _
	"This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version."& vbCrLf & vbCrLf & _
	"This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details."& vbCrLf & vbCrLf & _
	"You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>."

    myForm.TB_license.text = license
End Sub

