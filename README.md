# Word VBA Macro: Highlight Redaction
## This macro will redact highlighted text from a word document and save it to a new file.

### What does it do:
On startup it will search the first 5 number of pages for highlighted colors that are used in the document.
Once the userForm is loaded, the user can choose the file suffix of the redacted document (for saving), the colors to redact, the redaction text and which page to start from.
The macro will then go through all highlights using While Document.StoryRanges(i).Find.Highlight = True to find all highlighted texts in storyRanges (Main, Footnotes, TextFields).
It will check if the highlighted color is one of the colors the user has selected, if so, it will check if there are references (footnotes, fields) in the block, if found, it will seperate the range into multiple and replace all highlighted text with the redacted text. It will first go through the main text, then through footnotes and textboxes.
It will count per color how many redactions have taken place and log it. The user can afterwards create another version right away.
If there are multiple highlights found in a text field, or if there are custom highlight colors present, it will be logged.

ScreenUpdate is disable during redaction to make it a little faster / prevent the screen from flickering.


### This macro will
- redact text in the main body, tables, footnotes and textfields
- redact text in the main body that have highlights within highligts
- will leave footnotes and cross references (fields) within a highlight (will redact in front of and behind the reference) in the main text

### This macro will NOT:
- redact text in multiple highlights in textfields or footnotes as vba does not seem to be able to use range within those storyTypes
- cannot redact text with custom colors (outside of WdColorIndex)
- it also sometimes redacts one highlight as two consecutive highlights (putting in [redacted][redacted] instead of once [redacted])
- it will not redact text in header of footer (but will search for colors in there). This can be changed in code (see clsRedaction.build_story_ranges)
This macro is created by Ping Lu (mail@ping.lu) in 2023Q1. For License see below

### Please Note:
- If an highlight color is not within the first 5 pages and not in the footer or header, it will not be displayed to the user. Change this in: clsRedaction.Startup::searchNumberOfFirstPages
- If the WdColorIndex change, clsRedaction.build_highlight_color_names as well as UserForm1.build_user_color_selection_array has to be updated!
- StoryRanges can be changed here: clsRedaction.build_story_ranges

## Ideas for the future:
- load and save configuration and logs
- let user define the pages it should search for highlight colors (which colors are present)
- let user define which storyTypes it should search for highlight colors
- let user define which storyTypes it should use to redact


