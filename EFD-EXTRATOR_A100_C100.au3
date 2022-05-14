#cs
========================================================

	ELIMINAÇÃO DE LINHAS DUPLICADAS NO TXT EFD CONTRIBUIÇÕES

	Author: Luiz Fernando Cavalcanti

	Created: 26/04/2022

	Edited: 10/05/2022

	Description:
	Extrai as linhas A100 e C100 para exportar num arquivo CSV para comparação com planilhas no Excel.

========================================================
#ce

#Region ### WRAPPER DIRECTIVES ###

#AutoIt3Wrapper_Icon=IMG\ICONE.ico
#AutoIt3Wrapper_Res_Fileversion=0.1.1
#AutoIt3Wrapper_Res_Productversion=0.1.1
#AutoIt3Wrapper_Res_Field=ProductName|EFD-EXTRATOR_A100_C100
#AutoIt3Wrapper_Res_LegalCopyright=GPL3 - Author: Luiz Fernando Cavalcanti
#AutoIt3Wrapper_Res_Language=1046
#AutoIt3Wrapper_Res_Description=Busca linhas duplicadas do Bloco A e C do arquivo do EFD e gera novo arquivo tratado

#AutoIt3Wrapper_Outfile=.\EFD-EXTRATOR_A100_C100.exe

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


Global $g_aLinhasA[0][22]
Global $g_aLinhasC[0][30]

Global $g_aLinhasTxtIn[0]

Global $g_sArquivoIn = ""
Global $g_sArquivoOut = ""
Global $g_oHandleArqIn
Global $g_oHandleProgress

#endregion ### VARIABLES ###

#Region ### START Koda GUI section ###

Opt("GUIOnEventMode", 1)
$fForm = GUICreate("EFD - EXTRATOR A100 C100", 448, 284, 192, 124)
GUISetOnEvent($GUI_EVENT_CLOSE, "fFormClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "fFormMinimize")

$lExplicacao = GUICtrlCreateLabel("BUSCA LINHAS A100 e C100 E EXPORTA NUM ARQUIVO CSV", 8, 8, 425, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")
GUICtrlSetColor(-1, 0x0078D7)

$lSelecioneArqOri = GUICtrlCreateLabel("Selecione o arquivo original:", 10, 48, 167, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$iEndArquivoOrigem = GUICtrlCreateInput("", 8, 72, 321, 21)

$bSelectArqOrigem = GUICtrlCreateButton("Selecione", 336, 72, 75, 25)
GUICtrlSetOnEvent(-1, "bSelectArqOrigemClick")

$lSelecioneArqOut = GUICtrlCreateLabel("Selecione local para salvar o arquivo tratado:", 6, 128, 264, 20)
GUICtrlSetFont(-1, 10, 400, 0, "Tahoma")

$iEndArquivoSaida = GUICtrlCreateInput("", 9, 148, 321, 21)

$bSelectArqOut = GUICtrlCreateButton("Selecione", 336, 148, 75, 25)
GUICtrlSetOnEvent(-1, "bSelectArqOutClick")

$bProcessa = GUICtrlCreateButton("PROCESSAR >>", 320, 224, 91, 33)
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

; AO CLICAR NO BOTÃO SELECIONAR ARQUIVO DE ORIGEM
Func bSelectArqOrigemClick()

	Local $sFilePath

	$sFilePath = FileOpenDialog("SELECIONE AQUIVO DE ORIGEM",@WorkingDir,"All (*.*)",BitOR($FD_FILEMUSTEXIST, $FD_PATHMUSTEXIST),"",$fForm)

	If @error Then
		MsgBox(0, @ScriptName, @error & ' - Erro ao localizar arquivo')
		Return
	EndIf

	GUICtrlSetData($iEndArquivoOrigem,$sFilePath)

EndFunc

; AO CLICAR NO BOTÃO DE SELECIONAR ARQUIVO DE SAÍDA
Func bSelectArqOutClick()
	Local $sFilePath

	$sFilePath = FileSaveDialog("SELECIONE AQUIVO DE SAÍDA",@WorkingDir,"CSV file (*.csv)",BitOR($FD_PATHMUSTEXIST, $FD_PROMPTOVERWRITE),"EFD_A100_C100.csv",$fForm)

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
	$g_sArquivoOut = GUICtrlRead($iEndArquivoSaida)

	; Abre o Arquivo de texto de origem
	$g_oHandleArqIn = FileOpen($g_sArquivoIn, $FO_READ)

	; Checa se o arquivo foi lido, se der erro, sai
	If $g_oHandleArqIn = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de origem')
		Return
	Else

		;Carrega as linhas em Array e depois fecha o Arquivo
		CarregaLinhas()


		ProcessaLinhas()

		GravaArquivo()

		ProgressSet(100, "CONCLUIDO!")
		Sleep(1000)
		ProgressOff()


		MsgBox(BitOR($MB_OK,$MB_ICONINFORMATION),"PROCESSAMENTO CONCLUIDO!","O ARQUIVO:" & @CRLF & $g_sArquivoIn & @CRLF & @CRLF & "FOI PROCESSADO E O RESULTADO GRAVADO NO ARQUIVO:" & @CRLF & $g_sArquivoOut,0,$fForm)

	EndIf

EndFunc

Func CarregaLinhas()

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
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de origem')
		Exit 1
	EndIf

	;Loop para ler as linhas no array
	For $nLinha = 0 to $nLinhasTot-1

		$g_aLinhasTxtIn[$nLinha] =  FileReadLine($g_oHandleArqIn)

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Arquivo')
			Exit 1
		EndIf

	Next

EndFunc

Func ProcessaLinhas()

	ProgressSet(40, "Processando linhas A100 e C100...")

	Local $nLinhaGeral = 0
	Local $nLinhaA = 0
	Local $nLinhaC = 0

	Local $nLoopCBloco = 0

	Local $nIdBloco = ""

	; Array para receber resultado do StringSplit da linha atual
	Local $aSplitLinha[0]

	For $nLinhaGeral = 0 to UBound($g_aLinhasTxtIn)-1

		$nIdBloco = StringMid($g_aLinhasTxtIn[$nLinhaGeral],2,4)

		Switch $nIdBloco

			; Se for a linha de cabeçalho do bloco A1
			Case $nIdBloco = "A100"

				;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
				ReDim $aSplitLinha[23]

				;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DOS A100 E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaA = UBound($g_aLinhasA)
				ReDim $g_aLinhasA[$nLinhaA+1][22]

				;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
				$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

				;4 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				For $nLoopCBloco = 1 to 22
					$g_aLinhasA[$nLinhaA][$nLoopCBloco-1] = $aSplitLinha[$nLoopCBloco]
				Next

			; Se for a linha de cabeçalho do bloco A1
			Case $nIdBloco = "C100"

				;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
				ReDim $aSplitLinha[32]

				;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DOS A100 E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaC = UBound($g_aLinhasC)
				ReDim $g_aLinhasC[$nLinhaC+1][30]

				;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
				$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

;~ 				_ArrayDisplay($aSplitLinha)

				;4 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				For $nLoopCBloco = 1 to 30
					$g_aLinhasC[$nLinhaC][$nLoopCBloco-1] = $aSplitLinha[$nLoopCBloco]
				Next

		EndSwitch

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Bloco A')
			Exit 1
		EndIf

	Next

EndFunc

Func GravaArquivo()

	Local Const $SF_ANSI = 1
    Local Const $SF_UTF8 = 4

	Local $nLoopLinha = 0
	Local $nTotBlocoA = UBound($g_aLinhasA)-1
	Local $nTotBlocoC = UBound($g_aLinhasC)-1
	Local $sLinha = ""

	Local $sDataDoc = ""
	Local $sDataMov = ""

	Local $sCabecalhoCSV = "BLOCO;OPERACAO;NUM_DOC;DATA_DOC;DATA_MOV_SERV;VALOR_TOTAL;VALOR_PIS;VALOR_COFINS"

	Local $oHandleArqOut

	ProgressSet(60, "Preparando para gravar aquivo de saída...")

	$oHandleArqOut = FileOpen($g_sArquivoOut,2+512)

	If $oHandleArqOut = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de destino')
		Exit 1
	EndIf

	ProgressSet(65, "Gravando cabeçalho...")

	;GRAVA LINHA DO CABEÇALHO
	FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sCabecalhoCSV, $SF_UTF8), $SF_ANSI))


	ProgressSet(80, "Gravando linhas A100")

	;GRAVA LINHAS A100
	For $nLoopLinha = 0 to $nTotBlocoA



		$sDataDoc = StringFormat("%02i/%02i/%04i", StringMid($g_aLinhasA[$nLoopLinha][10],1,2),StringMid($g_aLinhasA[$nLoopLinha][10],3,2),StringMid($g_aLinhasA[$nLoopLinha][10],5,4))

		$sDataMov = StringFormat("%02i/%02i/%04i", StringMid($g_aLinhasA[$nLoopLinha][11],1,2),StringMid($g_aLinhasA[$nLoopLinha][11],3,2),StringMid($g_aLinhasA[$nLoopLinha][11],5,4))

		$sLinha = "A100;" & $g_aLinhasA[$nLoopLinha][2] & ";" & $g_aLinhasA[$nLoopLinha][8] & ";" & $sDataDoc & ";" & $sDataMov & ";" & $g_aLinhasA[$nLoopLinha][12] & ";" & $g_aLinhasA[$nLoopLinha][16] & ";" & $g_aLinhasA[$nLoopLinha][18]

		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinha, $SF_UTF8), $SF_ANSI))

	Next

	ProgressSet(90, "Gravando linhas C100")

	;GRAVA LINHAS C100
	For $nLoopLinha = 0 to $nTotBlocoC

		$sDataDoc = StringFormat("%02i/%02i/%04i", StringMid($g_aLinhasC[$nLoopLinha][10],1,2),StringMid($g_aLinhasC[$nLoopLinha][10],3,2),StringMid($g_aLinhasC[$nLoopLinha][10],5,4))

		$sDataMov = StringFormat("%02i/%02i/%04i", StringMid($g_aLinhasC[$nLoopLinha][11],1,2),StringMid($g_aLinhasC[$nLoopLinha][11],3,2),StringMid($g_aLinhasC[$nLoopLinha][11],5,4))

		$sLinha = "C100;" & $g_aLinhasA[$nLoopLinha][2] & ";" & $g_aLinhasC[$nLoopLinha][8] & ";" & $sDataDoc & ";" & $sDataMov & ";" & $g_aLinhasC[$nLoopLinha][12] & ";" & $g_aLinhasC[$nLoopLinha][26] & ";" & $g_aLinhasC[$nLoopLinha][27]

		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinha, $SF_UTF8), $SF_ANSI))

	Next


	FileClose($oHandleArqOut)

EndFunc