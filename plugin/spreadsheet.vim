"=============================================================================
" File: spreadsheet.vim
" Author: Miguel Jaque Barbero <mjaque@ilkebenson.com>
" Last Change: 09.02.2003
" Version: 0.1b
" ChangeLog:
" 	0.1b :	Improved functionality for cursor positioning.
" 	
" 	0.1a :  Added mapping and command for quick cell definition
" 		cell will allow quick insertion of cell labels
" 		command :Cell with or without cell label do the same
"
" 		Added translation for different number sintax
" 		
" Acknowledgement:
" Installation:
" 	Simply copy this file in your .vim/plugin directory.
" Use:
"	This script lets you use vim as a spreadsheet, adding aritmetic
"	calculations to your text files.
"	It provides three simple functions:
"	
"		Get("CellName"): 	  Returns the value of a cell.
"		Set("CellName","Value"):  Sets the value of a cell.
"		Calculate("Operation"):   Returns the result of the operation
"
"	Note: Calculate uses 'bc' command.
"
"	To create a cell simply put the string "(#Name)" BEFORE the WORD you
"	want to become a cell. No closing mark is required, because the first
"	WORD after the mark will be used. 
"
"	You can also use the :Cell <label> command. If no label is given you
"	will be prompted for input.
"		
"		:Cell CellLable
"
"	Or the mapping "cell" at command mode. 
"	
"	For Example: In the following text "Price1", "Price2", "Sum", "VAT"
"	and "Tot" are cell labels:
"
"	BILL
"		Item 1			(#Price1) 1.700,00 euros
"		Item 2			(#Price2)    30,75 euros
"		SUM			(#Sum)        0 euros
"		V.A.T.			(#VAT)        0 euros
"		TOTAL			(#Tot)	      0 euros
"
"	You can then use the functions to create formulas for your file.
"
"	a) Formulas can be directly writen at the command line:
"
"		:call Set("Sum", Calculate(Get("Price1")."+".Get("Price2")))
"
"	b) You can define commands in a script to easily recall commonly used
"	formulas:
"
"		:command VAT :call Set("VAT", Calculate(Get("Sum")."*0.16"))
"
"		And then execute it with:
"
"		:VAT
"
"	c) You can map those commands to access them quicker
"	
"		:map <F2> :VAT<CR>
"
"		And then simply press <F2> to execute it.
"
"	The Calculate function uses the bc operative system command. This is
"	the only way I have found to perform "advanced" operations, as decimal
"	additions. This may not work on non Linux operating systems.
"
"	You can change the default precision of the Calculate function (which
"	is 2 decimal digits) by setting the g:spreadsheet_precision variable
"	in your .vimrc
"
"	The script translates from "your" number sintax to bc required sintax,
"	(which is no thousand separator and "." as decimal separator (this
"	could be different in your sistem)) to a human sintax (by default, "."
"	as the thousand separator and "," as the decimal separator). This
"	values may be changed defining g:spreadsheet_thousand and
"	g:spreadsheet_decimal variables.
"
"	Also you can disable number sintax translation asigning
"	g:spreadsheet_translate the value 0
"
" TODO:
" 	Actually, no error control is implemented. You can try to add two non
" 	numeric words and get a horrible "parse error" result. This will be
" 	solved in next release (whenever it may come)
"
" 	You will also see that no alineation is done. I have no idea about how
" 	to solve this problem. Just edit the line and add or remove any space
" 	as required.
"
"	This script has only been tested in my computer, with my
"	configuration. The use of points, colons,... to define decimal points
"	may be different on other systems. If it doesn't work properly on your
"	system, send me an email and I'll try to do my best.
"
" Known Bugs:
"	ToHumanFormat functions gives unwanted result with negative values
"	less than 1000. Instead of -183,90 if gives -.183,90
" 
" Questions, suggestions, improvements, bugs, hints, acknowledge and support
" should be sent to mjaque@ilkebenson.com
"
" By the way: This script has ABSOLUTELY NO GUARANTEE. It is given as it is,
" with its bugs and mistakes, and you accept to use it at your own risk.
" It is also distributed under the General Public License, as defined by the
" Free Software Foundation (www.fsf.org). So you have the right to use, distribute,
" change and publish it as you wish, as long as it remains free software and the
" GPL license is kept. So don't remove this notice.


" Constant definition
if !exists("g:spreadsheet_precision")
	let g:spreadsheet_precision = 2
endif

if !exists("g:spreadsheet_translate")
	let g:spreadsheet_translate = 1
endif

if !exists("g:spreadsheet_decimal")
	let g:spreadsheet_decimal = ","
endif

if !exists("g:spreadsheet_thousand")
	let g:spreadsheet_thousand = "."
endif


" Highlighting definition
setlocal iskeyword+=(
setlocal iskeyword+=)
setlocal iskeyword+=#
syntax region SpreadSheetMark start=/(#/ end=/)/
highlight SpreadSheetMark ctermfg=DarkBlue guifg=DarkBlue

function Get(cell)
	call GetCursorPosition()	" Remember cursor position
	call search("(#".a:cell.")","w") " Seek cell label
	call search(')')		" To avoid problems with cell labels with numbers
	call search('\d')		" Go to numeric value
	let value = expand("<cWORD>")	" Get the value
	call SetCursorPosition()	" Restore cursor position
	if g:spreadsheet_translate == 1
		let value = TobcFormat(value)
	endif
	return value
endfunction

function Set(cell, value)
	let value = a:value
	if g:spreadsheet_translate == 1
		let value = ToHumanFormat(value)
	endif
	call GetCursorPosition()	" Remember cursor position
	let cell = "(#".a:cell.")"	" Build cell label
	let line = search(cell,"w")	" Seek cell label
	call search(')')		" To avoid problems with cell labels with numbers
	call search('\d')		" Go to numeric value
	let actualValue = expand("<cWORD>")	" Get actual value
	call setline(line, substitute(getline(line), cell.'\s*'.actualValue, cell." ".value,""))	" Change actual value for argument
	call SetCursorPosition()	" Restore cursor position
endfunction

function Calculate(operation)
	call GetCursorPosition()	" Remember cursor position
	let string = 'echo "'.a:operation.'" | bc -l'	" Build operation command
	" Write result in new line
	silent execute "normal :read! ".string."\<cr>"
	let result = getline(".")			" Get result
	" Delete new line
	silent execute "normal dd "
	" Remove unwanted decimals
	let position = match(result, '\.')		" Seek decimal point
	if position != -1				" If exists
		let result = strpart(result, 0, position + g:spreadsheet_precision + 1) " Cut result
	endif
	call SetCursorPosition()	" Restore cursor position
	return result
endfunction

" Functions and commands for quick cell definition.
function AddCell(label)
	let label = a:label
	if label == ""
		let label = input("Cell Label: ")
	endif
	silent execute "normal i(#".label.")"
endfunction

command -nargs=? Cell :call AddCell("<args>")
"command -nargs=? Cell :silent execute "normal i(#<args>)"
map cell :call AddCell("")<CR>

" Functions to adapt local number syntax to bc
" bc, the program used for Calculate() expects numbers to be in the following
" format: 
" 	ddddddddd.ddddd   for example     188282.8833
" 	(This may be different for your system)
" But personally I use dd.ddd,dd (21.344,98) specially in currency values.
" This functions transform to and from bc formats. Use the
" g:spreadsheet_decimal and g:spreadsheet_thousand variables to personalise.

function TobcFormat(string)
	let result = substitute(a:string, "\\".g:spreadsheet_thousand, "", "g")	" Remove all thousand simbols
	let result = substitute(result, g:spreadsheet_decimal,".","")		" Remove the decimal sign
	return result
endfunction

function ToHumanFormat(string)
	let result = substitute(a:string, '\.',g:spreadsheet_decimal,"")
	let position = match(result, g:spreadsheet_decimal)
	if position == -1
		let position = strlen(result)
	endif
	while position > 3
		let position = position - 3
		let result = strpart(result, 0, position).g:spreadsheet_thousand.strpart(result,position)
	endwhile
	return result
endfunction

" The cursor positioning functions GetCursorPosition and SetCursorPosition
" work as a LIFO queue.
" Each cursor position has three values, window, x and y.

let g:spreadsheet_crNumber = 0

function GetCursorPosition()
	let g:spreadsheet_crNumber = g:spreadsheet_crNumber + 1
	let g:spreadsheet_crBuffer_{g:spreadsheet_crNumber} = bufnr("%")
	let g:spreadsheet_crX_{g:spreadsheet_crNumber} = col(".")
	let g:spreadsheet_crY_{g:spreadsheet_crNumber} = line(".")
endfunction

function SetCursorPosition()
	let window = bufwinnr(g:spreadsheet_crBuffer_{g:spreadsheet_crNumber})
	exe "normal \<c-w>".window."w"	
	call cursor(g:spreadsheet_crY_{g:spreadsheet_crNumber}, g:spreadsheet_crX_{g:spreadsheet_crNumber})
	let g:spreadsheet_crNumber = g:spreadsheet_crNumber - 1
endfunction
