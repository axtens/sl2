'stacklang
option explicit

class stack
	dim tos
	dim astack()
	dim stacksize
	
	private sub class_initialize
		stacksize = 10000
		redim astack( stacksize )
		tos = 0
	end sub

	public sub push( x )
		astack(tos) = x
		tos = tos + 1
	end sub
	
	public property get stackempty
		stackempty = ( tos = 0 )
	end property
	
	public property get stackfull
		stackfull = ( tos > stacksize )
	end property
	
	public property get stackroom
		stackroom = stacksize - tos
	end property
	
	public property get stackcount
		stackcount = tos + 1
	end property
	
	public function pop()
		if tos > 0 then
			pop = astack( tos - 1 )
			tos = tos - 1
		else
			wscript.echo "Error: 'pop' but not enough data on stack"
			'~ wscript.quit
		end if
	end function

	public sub resizestack( n )
		redim preserve astack( n )
		stacksize = n
		if tos > stacksize then
			tos = stacksize
		end if
	end sub
	
	public sub rotate
		dim last, i
		dim base
		base = tos - 1
		last = astack( base )
		for i = 1 to base - 1
			astack( i ) = astack( i - 1 )
		next
		astack( 0 ) = last
	end sub
	
	public sub swap
		dim tmp
		if stackcount > 1 then
			tmp = astack( 0 )
			astack( 0 ) = astack( 1 )
			astack( 1 ) = tmp 
		else
			wscript.echo "Error: 'swap' but not enough data on stack"
		end if
	end sub
	
	public sub dup
		dim tmp 
		if stackcount > 0 then
			tmp = pop
			push tmp
			push tmp
		else
			wscript.echo "Error: 'dup' but not enough data on stack"
		end if
	end sub
	
	public sub show
		dim i
		wscript.stdout.write "--["
		for i = 0 to tos - 1
			wscript.stdout.write astack( i )
			if i < tos - 1 then
				wscript.stdout.write ", "
			end if
		next
		wscript.echo "]--"
	end sub 
	
	public property get stack
		stack = rtrim( join( aStack, " " ) )
	end property
	
end class

class machine
	private ip
	private script
	private finished
	private macros
	private subros
	
	sub class_initialize
		set macros = createobject("scripting.dictionary")
		set subros = createobject("scripting.dictionary")
	end sub
	
	public sub macro( key, data )
		if macros.exists( key ) then
			macros(key) = data
		else
			macros.add key, data
		end if
	end sub
	
	public function code( key )
		if macros.exists( key ) then
			code = macros(key)
		else
			code = vbnullstring
		end if
	end function
	
	public sub setscript( s )
		dim ss
		ss = replace( s, vbnewline, " " )
		ss = replace( ss, "  "," " )
		script = split(ss, " " )
		ip = 0
		finished = false
	end sub
	
	public property get nextop
		'~ wscript.stdout.write "(IP=" & ip & ") " & script(ip) & " "
		nextop = script( ip )'mid( script, ip, 1 )
		ip = ip + 1
		if ip > ubound( script ) then
			finished = true
		end if
	end property
	
	public sub prevop
		ip = ip - 1
		if ip = -1 then ip = 0
	end sub
	
	public property get isfinished
		isfinished = finished
	end property
	
	public sub firstop
		ip = 0
	end sub
	
	public function evaluate( CS )
		dim c, tmp, macro, m2, tmp2
		dim macsub
		dim hit, ques
		do while not isfinished
			c = nextop
			wscript.stdout.write c
			select case c
			case "("
				do while true
					c = nextop
					if c = ")" then exit do
				loop
			case "+"
				apply CS, "+", 2
			case "-"
				apply CS, "-", 2
			case "gt"
				apply CS, ">", 2
			case "lt"
				apply CS, "<", 2
			case "eq"
				apply CS, "=", 2
			case "ne"
				apply CS, "<>", 2
			case "and"
				apply CS, "and", 2
			case "dup" 'dup
				'~ if CS.stackcount > 0 then
					'~ tmp = CS.pop
					'~ CS.push tmp
					'~ CS.push tmp
				'~ else
					'~ wscript.echo "Error: ':' but not enough data on stack"
				'~ end if
				CS.dup
			case "rot" 'rotate top n elements of stack
				if CS.stackcount > 1 then
					CS.rotate
				else
					wscript.echo "Error: 'rot' but not enough data on stack"
				end if
			case "drop" ' drop
				if CS.stackcount > 0 then
					CS.pop
				else
					wscript.echo "Error: '_' but not enough data on stack"
				end if
			case "execif" 'test top of stack. Next op should be lowercase a..z being macro
				ques = cs.pop
				if ques then
					macsub = nextop
					wscript.stdout.write " " & macsub & " " 
					if macsub = lcase( macsub ) then
						evalmacro CS, macsub
					else
						RS.push ip
						ip = subros( macsub )
					end if
				else
					nextop
				end if
			case "exec" 'Next op should be lowercase a..z being mac
				macsub = nextop
				wscript.stdout.write " " & macsub & " " 
				if macsub = lcase( macsub ) then
					evalmacro CS, macsub
				else
					RS.push ip
					ip = subros( macsub )
				end if
			case "clear" 'Next op should be lowercase a..z being macro
				macro = nextop
				if macro = lcase( macro ) then 'if instr( "abcdefghijklmnopqrstuvwxyz", macro ) = 0 then
					'~ prevop
					'~ wscript.echo "Error: '?' not followed by macro name"
				'~ else
					wscript.echo "[clear " & macro & "]"
					m.macro macro, "nop"
				end if
			case "not" ' not
				if CS.stackcount > 0 then
					CS.push ( not CS.pop )
				else
					wscript.echo "Error: '~' but not enough data on stack"
				end if
			case "callstart" 'call
				'~ wscript.echo "%%" & ip & "%%"
				RS.push ip
				ip = 0
			case "return" 'return
				if RS.stackcount > 0 then
					ip = RS.pop
				else
					ip = ubound( script )
				end if
				'~ wscript.echo "return's ip", ip
			case "swap" 'swap
				'~ if CS.stackcount > 1 then
					'~ tmp = CS.pop
					'~ tmp2 = CS.pop
					'~ CS.push tmp
					'~ CS.push tmp2
				'~ else
					'~ wscript.echo "Error: '$' but not enough data on stack"
				'~ end if
				CS.swap
			case "PS>RS" 'pop from PS and push to RS
				IF PS.stackcount > 0 then
					tmp = PS.pop
					RS.push tmp 
					'~ wscript.echo tmp,"pushed to return stack"
				else
					wscript.echo "Error: '}' but not enough data on stack"
				end if
			case "RS>PS" 'pop from RS and push to PS
				IF RS.stackcount > 0 then
					tmp = RS.pop
					PS.push tmp
					'~ wscript.echo tmp,"pushed to program stack"
				else
					wscript.echo "Error: '}' but not enough data on return-stack"
				end if
			case "1" 'push 1 on stack
				PS.Push 1
			case "0" 'push 0 on stack
				PS.push 0
			case "print"
				wscript.echo CS.pop
			case "quit"
				wscript.quit 'exit do
			case "jumpstart" 'ip = 1
				m.firstop
			case "nop"
			case else
				PS.push c
			end select
			CS.show
		loop	
		'~ evaluate = CS.pop
	end function
	
	private sub apply( S, op, count )
		if S.stackcount > count then
			'~ wscript.echo "[" & op & "]"
			 S.push Eval( "S.pop " & op & " S.pop" )
		else
			wscript.echo "Error in Apply: '" & op & "' but not enough data on stack"
		end if
	end sub
	
	private sub evalmacro( context, macsub )
		dim m2
		'~ wscript.echo "[evalmacro " & macsub & "]"
		set m2 = new machine
		m2.setscript code(macsub)
		m2.evaluate context
		set m2 = nothing
	end sub

	sub showscript
		wscript.echo "script: " & join( script, " " )
	end sub

	sub subroutine( sName, sCode )
		dim newip
		newip = ubound( script ) + 1
		dim tmp
		if subros.exists( sName ) then 
			wscript.echo "Error: '" & sName & "' already defined as a subroutine"
		else
			subros.add sName, newIp
			script = split( join( script, vbnullchar) & vbnullchar & join( split( sCode, " " ), vbnullchar ), vbnullchar )
			'~ script = script & sCode
		end if
	end sub
	
end class 


dim PS
dim RS

set PS = new stack
set RS = new stack

PS.push 0 'm
PS.push 0 'n
PS.show

rs.push -1

dim m
set m = new machine
m.macro "save", "dup PS>RS swap dup PS>RS swap"
m.macro "rest", "RS>PS RS>PS"
m.macro "recurse", "0 PS>RS return"
'~ m.macro "save", "dup rot swap dup rot swap"
m.macro "test1", "drop 0 eq"
m.macro "test2", "swap drop 0 eq"
m.setscript "exec save exec test1 execif block1 exec rest exec save exec test2 execif BLOCK2 exec rest exec save exec BLOCK3" ' !z!d?A{{!z!e?B{{!z!C*"
m.macro "block1", "exec rest drop 1 + return"
m.subroutine "BLOCK2", "drop 1 + 1 recurse return"
m.subroutine "BLOCK3", "1 - exec recurse swap drop exec recurse return"
'~ m.setscript "drop dup dup + +"
m.evaluate PS
'~ wscript.echo ps.stack
'~ wscript.echo rs.stack
