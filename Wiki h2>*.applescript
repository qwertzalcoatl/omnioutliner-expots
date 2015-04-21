tell application "TextEdit"
	activate
	set newDocument to make new document
end tell

tell application "OmniOutliner"
	activate
	set outlineDocument to front document
	set thisRow to first row of outlineDocument
	my traverseRow(thisRow, 1, 0, newDocument)
	set text of newDocument to replace_chars(text of newDocument, return & return & return & return, return & return) of me
end tell

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars

on traverseRow(myRow, xCounter, yCounter, myTable)
	set yCounter to yCounter + 1
	tell application "OmniOutliner"
		set myText to text of topic of myRow
		set myComment to text of note cell of myRow
		my editRow(xCounter, yCounter, myText, myComment, myTable)
		if (exists child of myRow) then
			set newRow to first child of myRow
			set xCounter to xCounter + 1
			set yCounter to my traverseRow(newRow, xCounter, yCounter, myTable)
			set xCounter to xCounter - 1
		end if
		if (exists following sibling of myRow) then
			set newRow to following sibling 1 of myRow
			set yCounter to my traverseRow(newRow, xCounter, yCounter, myTable)
		end if
	end tell
	return yCounter
end traverseRow

on editRow(x, y, myText, myComment, myDocument)
	tell application "TextEdit"
		set text of myDocument to Â
			text of myDocument Â
			& my getBeginningOfLine(x) Â
			& myText Â
			& my getComment(myComment) Â
			& return Â
			& my getEndOfLine(x)
	end tell
end editRow

on getComment(comment)
	if (comment is not "") then
		return return & return Â
			& comment Â
			& return & return
	end if
end getComment

on getBeginningOfLine(x)
	set beginningOfLine to ""
	if (x = 1) then
		set beginningOfLine to return & return & "___" & return & return & "h2."
	else
		set x to x - 1
		repeat x times
			set beginningOfLine to "*" & beginningOfLine
		end repeat
	end if
	set beginningOfLine to beginningOfLine & " "
	
	return beginningOfLine
	
end getBeginningOfLine

on getEndOfLine(x)
	if (x < 2) then
		return return
	end if
end getEndOfLine