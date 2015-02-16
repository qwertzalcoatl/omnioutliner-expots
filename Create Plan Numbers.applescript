tell application "Numbers"
	activate
	set newDocument to make new document
	tell newDocument
		set mySheet to sheet 1
		tell mySheet
			set myTable to table 1
			set header column count of myTable to 0
		end tell
	end tell
end tell

tell application "OmniOutliner"
	activate
	set outlineDocument to front document
	set thisRow to first row of outlineDocument
	my traverseRow(thisRow, 1, 0, myTable)
end tell


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

on editRow(x, y, myText, myComment, myTable)
	tell application "Numbers"
		add row below last row of myTable
		set myCell to cell x of row y of myTable
		set value of cell 8 of row y of myTable to myComment
		set value of myCell to myText
		set thisRangeName to Â
			((name of cell x of row y of myTable) & ":" & Â
				(name of cell 7 of row y of myTable))
		tell myTable
			merge range thisRangeName
		end tell
	end tell
end editRow
