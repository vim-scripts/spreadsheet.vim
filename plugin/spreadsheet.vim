"=============================================================================
" File: spreadsheet.vim
" Author: Miguel Jaque Barbero <mjaque@ilkebenson.com>
" Last Change: 09.02.2003
" Version: 0.1
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
"	For Example: In the following text "Price1", "Price2", "Sum", "VAT"
"	and "Tot" are cell labels:
"
"	BILL
"		Item 1			(#Price1) 700.00 euros
"		Item 2			(#Price2)  30.75 euros
"		SUM			(#Sum)      0 euros
"		V.A.T.			(#VAT)      0 euros
"		TOTAL			(#Tot)	    0 euros
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
" Questions, suggestions, improvements, bugs, hints, acknowledge and support
" should be sent to mjaque@ilkebenson.com
"
" By the way: This script has ABSOLUTELY NO GUARANTEE. It is given as it is,
" with its bugs and mistakes, and you accept to use it at your own risk.
" It is also distributed under the General Public License, as defined by the
" Free Software Foundation (www.fsf.org). So you have the right to use, distribute,
" change and publish it as you wish, as long as it remains free software and the
" GPL license is kept. So don't remove this notice.
" 	

if !exists("g:spreadsheet_precision")
  let g:spreadsheet_precision = 2
endif

" Sintax definition
setlocal iskeyword+=(
setlocal iskeyword+=)
setlocal iskeyword+=#
syntax region SpreadSheetMark start=/(#/ end=/)/
highlight SpreadSheetMark ctermfg=DarkBlue guifg=DarkBlue

function Get(cell)
	" Remember cursor position
	let xpos = col(".")
	let ypos = line(".")
	call search("(#".a:cell.")","w") " Seek cell label
	call search(')')		" To avoid problems with cell labels with numbers
	call search('\d')		" Go to numeric value
	let value = expand("<cWORD>")	" Get the value
	call cursor(ypos, xpos)		" Restore cursor position
	return value
endfunction

function Set(cell, value)
	" Remember cursor position
	let xpos = col(".")
	let ypos = line(".")
	let cell = "(#".a:cell.")"	" Build cell label
	let line = search(cell,"w")	" Seek cell label
	call search(')')		" To avoid problems with cell labels with numbers
	call search('\d')		" Go to numeric value
	let actualValue = expand("<cWORD>")	" Get actual value
	call setline(line, substitute(getline(line), cell.'\s*'.actualValue, cell." ".a:value,""))	" Change actual value for argument
	call cursor(ypos, xpos)		" Restore cursor
endfunction

function Calculate(operation)
	" Remember cursor position
	let xpos = col(".")
	let ypos = line(".")
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
	call cursor(ypos, xpos)		" Restore cursor
	return result
endfunction
