# word-plugin-highlight-redaction
Redacts / Replaces highlighted text in a word file

Adds a new Tab and Ribbon to the Word menu list and when selected adds a Task pane.

The tool works in two steps:
a) search for colors: The user can select where the tool should search for highlight colors that should be replaced (default: Header, Footer, first 5 pages)
b) User has to then select (default, options)
> select which colors to replace (none, multiselect)
> what the version name is ("redacted", string)
> with what string the text should be replaced with (" [ Confidential ] ")
> which page to start and what to replace (2, int)
