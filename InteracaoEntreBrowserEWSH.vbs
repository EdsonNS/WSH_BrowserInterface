
'1-Interagir Script com Janela do Browser ok

'25/10/2019

Option Explicit  


'---------------------------------
'      Main
'---------------------------------
Call Inicializa() 'Initializes
Dim x : x=0
While ( (LeInterfaceComunication()=VALUE_PAUSED  )  or (LeInterfaceComunication()=VALUE_STARTED   ) ) 'reads in IEx
	While (LeInterfaceComunication()=VALUE_PAUSED  ) 'reads in IEx
		Wscript.Sleep 500
	WEnd
	If (ExisteVidaLaFora) Then 'if there is life out there
		'Processe algo aqui que de tempos em tempos possa verificar se é para continuar ou abortar
		'Process something here that from time to time can check if it is to continue or abort
		x=x+1
		Call escreveIEx(x) 'writes in IEx
		Wscript.Sleep 800
	End If	
	If (ExisteVidaLaFora) Then 'if there is life out there
		'Antes de continuar o processamento verifica novamente se é para continuar ou abortar
		'Before continuing processing, check again whether it is to continue or abort
		Call escreveIEy(x*2)
		Wscript.Sleep 800
	End If	
WEnd

Call Finaliza() 'Ends
MsgBox("FIM") 'The end

'--------------------------------------------
'      Interação entre Script e Internet Explorer
'      Interaction between Script and Internet Explorer
'--------------------------------------------

const ID_COMMUNICATION_INTERFACE = "IdCommunicationInterface"
const ID_DIV_VALUE_X = "divValueX"
const ID_DIV_VALUE_Y = "divValueY"
const VALUE_PAUSED = "pausado"
const VALUE_STOP = "parrar"
const VALUE_STARTED = "started"

Sub Inicializa() 'Initializes
	Call abreIE() 'Opens IE
End Sub

Sub Finaliza() 'Ends
	Call fechaIE() 'Close IE
End Sub

'---------------------------------
'      Internet Explorer
'---------------------------------

Dim IE

Sub abreIE()
	On Error resume Next
	Set IE = CreateObject("InternetExplorer.Application")
	IE.Navigate("about:blank")
	IE.ToolBar = 0
	IE.StatusBar = 0
	IE.Width = 400
	IE.Height = 200
	IE.Left = 0
	IE.Top = 0
	Do While (IE.Busy)
		Wscript.Sleep 200
	Loop    
	IE.Document.Title = "Process Information"
	
	Dim paginaHTML
	paginaHTML="<HTML>" & vbCrLf
	paginaHTML=paginaHTML & "<HEAD>" & vbCrLf	
	
	paginaHTML=paginaHTML & "</HEAD>" & vbCrLf
	paginaHTML=paginaHTML & "<BODY>" & vbCrLf
	paginaHTML=paginaHTML & "<br>Interação X=<span id=" & """" & ID_DIV_VALUE_X & """" & "></span>" & vbCrLf
	paginaHTML=paginaHTML & "<br>Interação Y=<span id=" & """" & ID_DIV_VALUE_Y & """" & "></span>" & vbCrLf

	paginaHTML=paginaHTML & "<input type=" & """" & "hidden" & """" & " id=" & """" & ID_COMMUNICATION_INTERFACE  & """" & " value=" & """" & VALUE_PAUSED  & """" & ">"
	
	paginaHTML=paginaHTML & "<br><input type=" & """" & "button" & """" & "onclick=" 
	paginaHTML=paginaHTML & """" & " document.getElementById('" & ID_COMMUNICATION_INTERFACE & "').value='" & VALUE_STARTED & "'" & """" 
    paginaHTML=paginaHTML & " value=" & """" & "Continuar Interação" & """" & ">" & vbCrLf

	paginaHTML=paginaHTML & "<br><input type=" & """" & "button" & """" & "onclick=" 
	paginaHTML=paginaHTML & """" & " document.getElementById('" & ID_COMMUNICATION_INTERFACE & "').value='" & VALUE_PAUSED  & "'" & """" 
    paginaHTML=paginaHTML & " value=" & """" & "Pausar  a Interação" & """" & ">" & vbCrLf

	paginaHTML=paginaHTML & "<br><input type=" & """" & "button" & """" & "onclick=" 
	paginaHTML=paginaHTML & """" & " document.getElementById('" & ID_COMMUNICATION_INTERFACE & "').value='" & VALUE_STOP & "'" & """" 
    paginaHTML=paginaHTML & " value=" & """" & "Finalizar  o processamento e fechar" & """" & ">" & vbCrLf
	
    paginaHTML=paginaHTML & "</BODY>" & vbCrLf
	paginaHTML=paginaHTML & "</HTML>" & vbCrLf
	IE.Document.Body.InnerHTML = paginaHTML
	IE.Visible = 1  
	Do While (IE.Busy)
		Wscript.Sleep 200
	Loop 	

End Sub

sub escreveIEx(valor) 'writes in IEx
	On Error resume Next
		IE.Document.all(ID_DIV_VALUE_X).innerText = valor
End Sub

sub escreveIEy(valor) 'writes in IEx
	On Error resume Next
	IE.Document.all(ID_DIV_VALUE_Y).innerText = valor
End Sub

Function LeInterfaceComunication() 'reads in IEx
	On Error resume Next
	LeInterfaceComunication = IE.Document.all(ID_COMMUNICATION_INTERFACE).value  
End Function

Sub fechaIE() 'Close IE
	On Error resume Next
	IE.Quit
End Sub

Function ExisteVidaLaFora()
	On Error resume Next
	While (LeInterfaceComunication()=VALUE_PAUSED  ) 'reads in IEx
		Wscript.Sleep 500
	WEnd
	ExisteVidaLaFora = (LeInterfaceComunication()=VALUE_STARTED) 'reads in IEx
End Function