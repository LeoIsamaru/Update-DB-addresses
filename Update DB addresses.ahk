InputBox, count, Loop Tester, Enter number of loops to do.
Gui, Add, Progress
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 500
Send {f5} ; open the 'goto' box
sleep, 500
send A2		; Escolhe celula ID inicial
Sleep, 500
send {enter}
MainLoop:	;loop geral que pode ser repetido no fim do prog
loop, %count% {
Sleep, 300
send ^c	;copia a ID
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
MouseMove, 140, 190	;posição da box procurar carimbo por ID no ecra do portatil, maximizado
Sleep, 300
MouseClick, left
Sleep, 500
send,^a
Sleep, 300
Send, {Delete}
Sleep, 300
send,^v
Sleep, 300
MouseMove, 250, 190	;posição botao procurar
Sleep, 300
MouseClick, left
Sleep, 300
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 300
send {Right 2}
Sleep, 300
send,^c	;copia o tipo de rua
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
MouseMove, 690, 380	;posicao do carimbo procurado
Sleep, 400
MouseClick, left,,, 2	;abre as infos do carimbo. TABS aqui dentro para garantir sem erros
Sleep, 1000
send,{Tab} ;seleciona a box tipo de rua
Sleep, 500
Send, {Delete}
Sleep, 500
Send, ^v	;cola tipo de rua
Sleep, 500
send,{Down}	;é preciso um DOWN para garantir que encontra o tipo certo na lista, antes de passar para o proximo 
Sleep, 500	;
send,{Tab}	; mudar para nome rua
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 300
send {Right}
send,^c	;copia o nome de rua
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
Send, ^v	;cola tipo de rua
Sleep, 300
send,{Tab}	;muda para tipo de nº
Sleep, 300
send,{Down}	;escolhe SN, para nº certo será preciso modificar algo
Sleep, 300
send,{Tab}
Sleep, 300
send,{Tab 3}
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 300
send {Right}
send,^c	;copia hps privadas/casa
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
Send, ^v	;cola qt hps casas
Sleep, 300
send,{Tab}	;muda hp PRO
Sleep, 300
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 300
send {Right}
send,^c	;copia hp PRO
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
Send, ^v	;cola qt hps PRO
Sleep, 300
send,{Tab}	;muda para cx tipo alimentacao
Sleep, 300
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 300
send {Right}
send,^c	;copia tipo alimentacao
Sleep, 300
WinActivate ahk_class WindowsForms10.Window.8.app.0.141b42a_r6_ad1	;muda para o prog BD
Sleep, 300
send,^a
Sleep, 300	
Send, {Delete}
Sleep, 300
Send, ^v	;cola tipo alimentacao
Sleep, 300
send,{Down}
Sleep, 300
send,{Tab 2}	
Sleep, 500
send, {Enter}
Sleep, 10000 ;10sec
WinActivate ahk_exe EXCEL.EXE	; vai para o ult excel activo
Sleep, 50
send,+{Space}	;shift+pace para selecionar linha
Sleep, 300
send, !4	;alt+4 para selecionar balde e pintar
Sleep, 300
send,{Down 6}
Sleep, 300
send {Right 4}
Sleep, 300
send {enter}
Sleep, 300
send {Home}
Sleep, 300
send,{Down}
Sleep, 300
ClipSaved := ""
tooltip, Loop #: %A_Index% de %count%	;no fim de cada loop mostra qts feitos até ao momento
}	;fim do loop geral
MsgBox, Loop Tester Done
InputBox, count, Loop Tester, Enter number of loops to do.	;cx inicial se for preciso repetir
GoTo, MainLoop
return	;
F3:: Pause
F4:: GoTo, MainLoop
*esc::
   ExitApp