#cs
========================================================

	ELIMINAÇÃO DE LINHAS DUPLICADAS NO TXT EFD CONTRIBUIÇÕES

	Author: Luiz Fernando Cavalcanti

	Created: 27/04/2022

	Edited: 27/04/2022

	Description:
	Extrai um CSV para gerar o bloco F600

========================================================
#ce

#Region ### WRAPPER DIRECTIVES ###

#AutoIt3Wrapper_Icon=IMG\ICONE.ico
#AutoIt3Wrapper_Res_Fileversion=0.1.0
#AutoIt3Wrapper_Res_Productversion=0.1.0
#AutoIt3Wrapper_Res_Field=ProductName|EFD-GERADOR_F600
#AutoIt3Wrapper_Res_LegalCopyright=GPL3 - Author: Luiz Fernando Cavalcanti
#AutoIt3Wrapper_Res_Language=1046
#AutoIt3Wrapper_Res_Description=Carrega linhas de um CSV e substitui o F600 de um arquivo do EFD Contribuições

#AutoIt3Wrapper_Outfile=.\EFD-GERADOR_F600.exe

#AutoIt3Wrapper_au3check_parameters= -w 5


#EndRegion ### WRAPPER DIRECTIVES ###

#region ### INCLUDES ###

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <Array.au3>
#include <File.au3>

#endregion ### INCLUDES ###

#region ### VARIABLES ###


Global $g_aLinhasF[0][6]

Global $g_aLinhasTxtIn[0]
Global $g_aLinhasTxtCsv[0]

Global $g_nPosF001 = 0
Global $g_nPosF990 = 0
Global $g_nPosIniF600 = 0
Global $g_nPosFimF600 = 0

Global $g_bFoundBlocoF = False

Global $g_sArquivoIn = ""
Global $g_sArquivoCsv = ""
Global $g_sArquivoOut = ""
Global $g_oHandleArqIn
Global $g_oHandleArqCsv
Global $g_oHandleProgress

#endregion ### VARIABLES ###

#Region ### START Koda GUI section ###

Opt("GUIOnEventMode", 1)
$fForm = GUICreate("EFD - GERADOR F600", 448, 340, 192, 124)
GUISetOnEvent($GUI_EVENT_CLOSE, "fFormClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "fFormMinimize")

$lExplicacao = GUICtrlCreateLabel("CARREGA UM CSV PARA GERAR LINHAS DO BLOCO F600", 8, 8, 425, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetColor(-1, 0x0078D7)



$lSelecioneArqOri = GUICtrlCreateLabel("Selecione o arquivo original do EFD:", 10, 48, 230, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$iEndArquivoOrigem = GUICtrlCreateInput("", 8, 72, 321, 21)

$bSelectArqOrigem = GUICtrlCreateButton("Selecione", 336, 72, 75, 25)
GUICtrlSetOnEvent(-1, "bSelectArqOrigemClick")



$lSelecioneArqCsv = GUICtrlCreateLabel("Selecione o arquivo CSV do F600:", 10, 128, 200, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$iEndArquivoCsv = GUICtrlCreateInput("", 8, 148, 321, 21)

$bSelectArqCsv = GUICtrlCreateButton("Selecione", 336, 148, 75, 25)
GUICtrlSetOnEvent(-1, "bSelectArqCsvClick")



$lSelecioneArqOut = GUICtrlCreateLabel("Selecione local para salvar o arquivo tratado:", 6, 200, 264, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$iEndArquivoSaida = GUICtrlCreateInput("", 9, 224, 321, 21)

$bSelectArqOut = GUICtrlCreateButton("Selecione", 336, 224, 75, 25)
GUICtrlSetOnEvent(-1, "bSelectArqOutClick")



$bProcessa = GUICtrlCreateButton("PROCESSAR >>", 320, 300, 91, 33)
GUICtrlSetBkColor(-1, 0xC0DCC0)
GUICtrlSetOnEvent(-1, "bProcessaClick")

GUISetState(@SW_SHOW)

#EndRegion ### END Koda GUI section ###

While 1
	Sleep(100)
WEnd

; AO CLICAR PARA FECHAR
Func fFormClose()
	Exit 1
EndFunc

; AO CLICAR PARA MINIMIZAR
Func fFormMinimize()
	WinSetState ($fForm, "", @SW_MINIMIZE)
EndFunc

; AO CLICAR NO BOTÃO SELECIONAR ARQUIVO EFD
Func bSelectArqOrigemClick()

	Local $sFilePath

	$sFilePath = FileOpenDialog("SELECIONE AQUIVO DE ORIGEM",@WorkingDir,"All (*.*)",BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST),"",$fForm)

	If @error Then
		MsgBox(0, @ScriptName, @error & ' - Erro ao localizar arquivo')
		Return
	EndIf

	GUICtrlSetData($iEndArquivoOrigem,$sFilePath)

EndFunc

; AO CLICAR NO BOTÃO DE SELECIONAR ARQUIVO CSV
Func bSelectArqCsvClick()

	Local $sFilePath

	$sFilePath = FileOpenDialog("SELECIONE AQUIVO CSV",@WorkingDir,"CSV file(*.csv)",BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST),"",$fForm)

	If @error Then
		MsgBox(0, @ScriptName, @error & ' - Erro ao localizar arquivo')
		Return
	EndIf

	GUICtrlSetData($iEndArquivoCsv,$sFilePath)
EndFunc

; AO CLICAR NO BOTÃO DE SELECIONAR ARQUIVO DE SAÍDA
Func bSelectArqOutClick()
	Local $sFilePath

	$sFilePath = FileSaveDialog("SELECIONE AQUIVO DE SAÍDA",@WorkingDir,"Text file(*.txt)",BitOR($FD_PATHMUSTEXIST, $FD_PROMPTOVERWRITE),"EFD_F600_SUBSTITUIDO.txt",$fForm)

	If @error Then
		MsgBox(0, @ScriptName, @error & ' - Erro ao selecionar local para salvar o arquivo')
		Return
	EndIf

	GUICtrlSetData($iEndArquivoSaida,$sFilePath)

EndFunc

; AO CLICAR NO BOTÃO PROCESSAR
Func bProcessaClick()

	ProgressOn("PROCESSANDO", "Aguarde..." , "Preparando...", -1, -1, BitOR($DLG_NOTONTOP, $DLG_MOVEABLE))

	$g_sArquivoIn = GUICtrlRead($iEndArquivoOrigem)
	$g_sArquivoCsv = GUICtrlRead($iEndArquivoCsv)
	$g_sArquivoOut = GUICtrlRead($iEndArquivoSaida)

	; Abre o Arquivo de texto de origem
	$g_oHandleArqIn = FileOpen($g_sArquivoIn, $FO_READ)

	; Checa se o arquivo foi lido, se der erro, sai
	If $g_oHandleArqIn = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de origem')
		Return
	EndIf

	; Abre o Arquivo CSV
	$g_oHandleArqCsv = FileOpen($g_sArquivoCsv, $FO_READ)

	; Checa se o arquivo foi lido, se der erro, sai
	If $g_oHandleArqCsv = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo CSV')
		Return
	EndIf

	;Carrega as linhas em Array e depois fecha o Arquivo
	CarregaLinhasEFD()

	;Carrega as linhas em Array e depois fecha o Arquivo
	CarregaLinhasCSV()

	;Processa as linhas do CSV num Array
	ProcessaLinhasCSV()

	;Processa o arquivo do EFD para saber onde começa e termina o bloco F
	ProcessaLinhasEFD()

	GravaArquivo()

	ProgressSet(100, "CONCLUIDO!")
	Sleep(1000)
	ProgressOff()


	MsgBox(BitOR($MB_OK,$MB_ICONINFORMATION),"PROCESSAMENTO CONCLUIDO!","O ARQUIVO:" & @CRLF & $g_sArquivoIn & @CRLF & @CRLF & "FOI PROCESSADO E O RESULTADO GRAVADO NO ARQUIVO:" & @CRLF & $g_sArquivoOut,0,$fForm)

EndFunc

Func CarregaLinhasEFD()

	Local $nLinhasTot = 0
	Local $nLinha

	ProgressSet(1, "Lendo Arquivo de Origem...")

	;determina quantidade de linhas do arquivo e redimensiona array
	$nLinhasTot = _FileCountLines($g_oHandleArqIn)
	ReDim $g_aLinhasTxtIn[$nLinhasTot]

	;Fecha o Arquivo que foi aberto para teste
	FileClose($g_oHandleArqIn)

	; Abre o Arquivo de texto de origem
	$g_oHandleArqIn = FileOpen($g_sArquivoIn, $FO_READ)

	; Checa se o arquivo foi lido, se der erro, sai
	If $g_oHandleArqIn = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo EFD de origem')
		Exit 1
	EndIf

	;Loop para ler as linhas no array
	For $nLinha = 0 to $nLinhasTot-1

		$g_aLinhasTxtIn[$nLinha] =  FileReadLine($g_oHandleArqIn)

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Arquivo EFD')
			Exit 1
		EndIf

	Next

EndFunc

Func CarregaLinhasCSV()

	Local $nLinhasTot = 0
	Local $nLinha

	ProgressSet(10, "Lendo Arquivo CSV...")

	;determina quantidade de linhas do arquivo e redimensiona array
	$nLinhasTot = _FileCountLines($g_oHandleArqCsv)
	ReDim $g_aLinhasTxtCsv[$nLinhasTot]

;~ 	ConsoleWrite($nLinhasTot)

	;Fecha o Arquivo que foi aberto para teste
	FileClose($g_oHandleArqCsv)

	; Abre o Arquivo de texto de origem
	$g_oHandleArqCsv = FileOpen($g_sArquivoCsv, $FO_READ)

	; Checa se o arquivo foi lido, se der erro, sai
	If $g_oHandleArqCsv = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo CSV de origem')
		Exit 1
	EndIf

	;Loop para ler as linhas no array
	For $nLinha = 0 to $nLinhasTot-1

		$g_aLinhasTxtCsv[$nLinha] =  FileReadLine($g_oHandleArqCsv)

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Arquivo CSV')
			Exit 1
		EndIf

	Next

EndFunc

Func ProcessaLinhasCSV()

	ProgressSet(30, "Processando linhas do CSV...")

	Local $nLinhaGeral = 0
	Local $nLinhaF = 0

	Local $nLoopCBloco = 0

	; Array para receber resultado do StringSplit da linha atual
	Local $aSplitLinha[0]

	For $nLinhaGeral = 0 to UBound($g_aLinhasTxtCsv)-1

		;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
		ReDim $aSplitLinha[6]

		;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DOS A100 E ATUALIZA VARIAVEL DE CONTROLE
		$nLinhaF = UBound($g_aLinhasF)
		ReDim $g_aLinhasF[$nLinhaF+1][6]

		;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
		$aSplitLinha = StringSplit($g_aLinhasTxtCsv[$nLinhaGeral],";")

		;4 - INSERIR DADOS NO ARRAY DE LINHAS DO CSV
		For $nLoopCBloco = 1 to 6
			$g_aLinhasF[$nLinhaF][$nLoopCBloco-1] = $aSplitLinha[$nLoopCBloco]
		Next

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do array do CSV')
			Exit 1
		EndIf

	Next

EndFunc

Func ProcessaLinhasEFD()

	ProgressSet(40, "Tratando linhas do arquivo do EFD...")

	Local $nLinhaGeral = 0

	Local $nIdBloco = ""

	For $nLinhaGeral = 0 to UBound($g_aLinhasTxtIn)-1

		$nIdBloco = StringMid($g_aLinhasTxtIn[$nLinhaGeral],2,4)

;~ 		ConsoleWrite($nIdBloco)

		Switch $nIdBloco

			;INICIO DO BLOCO F
			Case $nIdBloco = "F001"

				$g_nPosF001 = $nLinhaGeral

			;BLOCO F600 A SER SUBSTITUIDO
			Case $nIdBloco = "F600"

				;Se já começou o F600 ir atualizando a linha final
				If $g_bFoundBlocoF Then
					$g_nPosFimF600 = $nLinhaGeral
				;Senão marcar a linha inicial e a flag
				Else
					$g_nPosIniF600 = $nLinhaGeral
					$g_bFoundBlocoF = True
				EndIf

			; ULTIMA LINHA DO BLOCO F990
			Case $nIdBloco = "F990"
				$g_nPosF990 = $nLinhaGeral

		EndSwitch

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Bloco A')
			Exit 1
		EndIf

	Next

	If not $g_bFoundBlocoF Then
			MsgBox(0, @ScriptName, @error & ' - Erro o arquivo EFD não continha o bloco F')
			Exit 1
	EndIf

EndFunc

Func GravaArquivo()

	Local Const $SF_ANSI = 1
    Local Const $SF_UTF8 = 4

	Local $nLoopLinha = 0
	Local $nContGeral = 0
	Local $nTotBlocoF = UBound($g_aLinhasF)-1
	Local $nTotArqEfd = UBound($g_aLinhasTxtIn)-2
	Local $sLinha = ""

	Local $sDataRet = ""
	Local $sCNPJ = ""
;~ 	Local $sDataMov = ""

	Local $oHandleArqOut

	ProgressSet(45, "Preparando para gravar aquivo de saída...")

	$oHandleArqOut = FileOpen($g_sArquivoOut,2+512)

	If $oHandleArqOut = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de destino')
		Exit 1
	EndIf

	ProgressSet(50, "Gravando linhas originais até Bloco F600...")
	;GRAVA TODAS AS LINHAS ATE BLOCO F600 COMEÇAR
	For $nLoopLGeral = 0 to $g_nPosIniF600-1
		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($g_aLinhasTxtIn[$nLoopLGeral], $SF_UTF8), $SF_ANSI))
		$nContGeral += 1
	Next


	ProgressSet(70, "Gravando linhas F600...")

	;GRAVA LINHAS F600
	For $nLoopLinha = 0 to $nTotBlocoF

		;DATA RETENCAO
		;31/01/2020
		;1,2 4,2 7,4
		$sDataRet = StringFormat("%02i%02i%04i", StringMid($g_aLinhasF[$nLoopLinha][0],1,2),StringMid($g_aLinhasF[$nLoopLinha][0],4,2),StringMid($g_aLinhasF[$nLoopLinha][0],7,4))

		;CNPJ
		;64.858.525/0078-24
		;1,2 4,3 8,3 12,4 17,2
		$sCNPJ = StringFormat("%02i%03i%03i%04i%02i", StringMid($g_aLinhasF[$nLoopLinha][3],1,2),StringMid($g_aLinhasF[$nLoopLinha][3],4,3),StringMid($g_aLinhasF[$nLoopLinha][3],8,3),StringMid($g_aLinhasF[$nLoopLinha][3],12,4),StringMid($g_aLinhasF[$nLoopLinha][3],17,2))

		$sLinha = "|F600|03|" & $sDataRet & "|" & $g_aLinhasF[$nLoopLinha][1] & "|" & $g_aLinhasF[$nLoopLinha][2] & "||0|" & $sCNPJ & "|" & $g_aLinhasF[$nLoopLinha][4] & "|" & $g_aLinhasF[$nLoopLinha][5] & "|0|"

		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinha, $SF_UTF8), $SF_ANSI))

		$nContGeral += 1

	Next

	ProgressSet(90, "Gravando linha F990")
	;GRAVA LINHA F990 COM CONTADOR ATUALIZADO
	;|F990|90|
	$sLinha = "|F990|" & ($nContGeral-$g_nPosF001+1) & "|"
	FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinha, $SF_UTF8), $SF_ANSI))
	$nContGeral += 1


	ProgressSet(95, "Gravando linhas restantes do arquivo EFD")
	;GRAVA LINHAS RESTANTES
	For $nLoopLGeral = $g_nPosFimF600+2 to $nTotArqEfd
		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($g_aLinhasTxtIn[$nLoopLGeral], $SF_UTF8), $SF_ANSI))
		$nContGeral += 1

	Next


	ProgressSet(99, "Gravando linha 9999")
	;GRAVA LINHA 9999 COM CONTADOR ATUALIZADO
	;|9999|7600|
	$sLinha = "|9999|" & ($nContGeral+1) & "|"
	FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinha, $SF_UTF8), $SF_ANSI))

	FileClose($oHandleArqOut)

EndFunc