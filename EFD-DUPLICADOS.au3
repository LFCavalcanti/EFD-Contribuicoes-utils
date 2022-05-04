#cs
========================================================

	ELIMINAÇÃO DE LINHAS DUPLICADAS NO TXT EFD CONTRIBUIÇÕES

	Author: Luiz Fernando Cavalcanti

	Created: 14/04/2022

	Edited: 14/04/2022

	Description:
	Verica linhas duplicadas no arquivo TXT do EFD Contribuições e elimina a duplicidade
	também atualizando os valores da linha de controle do bloco.

========================================================
#ce

#Region ### WRAPPER DIRECTIVES ###

#AutoIt3Wrapper_Icon=IMG\ICONE.ico
#AutoIt3Wrapper_Res_Fileversion=0.1.3
#AutoIt3Wrapper_Res_Productversion=0.1.3
#AutoIt3Wrapper_Res_Field=ProductName|EFD-DUPLICADOS
#AutoIt3Wrapper_Res_LegalCopyright=GPL3 - Author: Luiz Fernando Cavalcanti
#AutoIt3Wrapper_Res_Language=1046
#AutoIt3Wrapper_Res_Description=Busca linhas duplicadas do Bloco A e C do arquivo do EFD e gera novo arquivo tratado

#AutoIt3Wrapper_Outfile=.\EFD-Duplicados.exe

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

;~Posição no Array Original + 21 possíveis posições do arquivo + Posição da ultima cabeça de bloco(A100)
Global $g_aLinhasBlocoA[0][23]

Global $g_aLinhasTxtIn[0]

Global $g_nPosIniBlocoA = 0
Global $g_nPosFimBlocoA = 0

Global $g_bFoundBlocoA = False

Global $g_sArquivoIn = ""
Global $g_sArquivoOut = ""
Global $g_oHandleArqIn
Global $g_oHandleProgress

#endregion ### VARIABLES ###

#Region ### START Koda GUI section ###

Opt("GUIOnEventMode", 1)
$fForm = GUICreate("EFD - LIMPA DUPLICADOS", 448, 284, 192, 124)
GUISetOnEvent($GUI_EVENT_CLOSE, "fFormClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "fFormMinimize")

$lExplicacao = GUICtrlCreateLabel("BUSCA LINHAS DUPLICADAS DO BLOCO A E C -> GERA ARQUIVO TRATADO", 8, 8, 425, 20)
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

	$sFilePath = FileSaveDialog("SELECIONE AQUIVO DE SAÍDA",@WorkingDir,"Text files (*.txt)|All (*.*)",BitOR($FD_PATHMUSTEXIST, $FD_PROMPTOVERWRITE),"EFD_TRATADO.txt",$fForm)

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


		ProcessaLinhasBlocoA()

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

		;$g_aLinhasTxtIn[$nLinha][6] = FileReadLine($g_oHandleArqIn,$nLinha)
		$g_aLinhasTxtIn[$nLinha] =  FileReadLine($g_oHandleArqIn)

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Arquivo')
			Exit 1
		EndIf

	Next

EndFunc

Func ProcessaLinhasBlocoA()

	ProgressSet(40, "Tratando linhas do Bloco A...")

	Local $nLinhaGeral = 0
	Local $nLinhaA = 0
	Local $nLinhaHeadA = 0
	Local $nBloco = 0
	Local $nHeadBloco = 0
	Local $nPosHeadAtual = 0
	Local $nContLinhaGeral = 0

	Local $nLoopCBlocoA = 0

	Local $nIdBloco = ""

	; Pos Array Bloco A, Valor 1, Valor 2, Valor 3
	Local $aLinhasMestre[0][6]

	; Array para receber resultado do StringSplit da linha atual
	Local $aSplitLinha[0]

	For $nLinhaGeral = 0 to UBound($g_aLinhasTxtIn)-1

		$nIdBloco = StringMid($g_aLinhasTxtIn[$nLinhaGeral],2,4)

		Switch $nIdBloco

			Case $nIdBloco = "A001"

				;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
				ReDim $aSplitLinha[3]

				;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DO BLOCO A E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaA = UBound($g_aLinhasBlocoA)
				ReDim $g_aLinhasBlocoA[$nLinhaA+1][23]


				;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
				$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

				;4 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				$g_aLinhasBlocoA[$nLinhaA][0] = $nLinhaGeral
				$g_aLinhasBlocoA[$nLinhaA][1] = $aSplitLinha[2]
				$g_aLinhasBlocoA[$nLinhaA][2] = $aSplitLinha[3]

				If Not($g_bFoundBlocoA) Then
					$g_bFoundBlocoA = True
					$g_nPosIniBlocoA = $nLinhaGeral
				EndIf

			Case $nIdBloco = "A010"

				;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
				ReDim $aSplitLinha[3]

				;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DO BLOCO A E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaA = UBound($g_aLinhasBlocoA)
				ReDim $g_aLinhasBlocoA[$nLinhaA+1][23]

				;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
				$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

				;4 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				$g_aLinhasBlocoA[$nLinhaA][0] = $nLinhaGeral
				$g_aLinhasBlocoA[$nLinhaA][1] = $aSplitLinha[2]
				$g_aLinhasBlocoA[$nLinhaA][2] = $aSplitLinha[3]

			; Se for a linha de cabeçalho do bloco A1
			Case $nIdBloco = "A100"

				;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
				ReDim $aSplitLinha[3]

				;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DO BLOCO A E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaA = UBound($g_aLinhasBlocoA)
				ReDim $g_aLinhasBlocoA[$nLinhaA+1][23]

				;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
				$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

				;4 - ATUALIZA CONTROLES POR SER HEAD DE BLOCO DE NOTA
				$nBloco = 0 ;Contador de linha entre os A100 e todos os A170
				$nHeadBloco = $nLinhaGeral ;Qual a posição do A100 correspondente em relação ao Array Geral

				;5 - GRAVA ARRAY PARA CONTROLAR HEADS A100
				$nLinhaHeadA = UBound($aLinhasMestre)
				ReDim $aLinhasMestre[$nLinhaHeadA+1][6]
				$aLinhasMestre[$nLinhaHeadA][0] = $nLinhaGeral
				$aLinhasMestre[$nLinhaHeadA][1] = 0
				$aLinhasMestre[$nLinhaHeadA][2] = 0
				$aLinhasMestre[$nLinhaHeadA][3] = 0
				$aLinhasMestre[$nLinhaHeadA][4] = 0
				$aLinhasMestre[$nLinhaHeadA][5] = 0

				;6 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				$g_aLinhasBlocoA[$nLinhaA][0] = $nLinhaGeral
				For $nLoopCBlocoA = 2 to 22
					$g_aLinhasBlocoA[$nLinhaA][$nLoopCBlocoA-1] = $aSplitLinha[$nLoopCBlocoA]
				Next
				$g_aLinhasBlocoA[$nLinhaA][22] = $nHeadBloco

			;Se for uma linha de item
			Case $nIdBloco = "A170"

				;Caso a linha de item não seja igual a anterior
				If StringMid($g_aLinhasTxtIn[$nLinhaGeral],8) <> StringMid($g_aLinhasTxtIn[$nLinhaGeral-1],8) Then

					;1 - REDIMENCIONAR ARRAY QUE RECEBE STRING PARA TAMANHO DO BLOCO
					ReDim $aSplitLinha[3]

					;2 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DO BLOCO A E ATUALIZA VARIAVEL DE CONTROLE
					$nLinhaA = UBound($g_aLinhasBlocoA)
					ReDim $g_aLinhasBlocoA[$nLinhaA+1][23]

					;3 - EFETUAR O SPLIT PARA ALIMENTAR O ARRAY
					$aSplitLinha = StringSplit($g_aLinhasTxtIn[$nLinhaGeral],"|")

					;4 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
					$g_aLinhasBlocoA[$nLinhaA][0] = $nLinhaGeral
					For $nLoopCBlocoA = 2 to 19
						$g_aLinhasBlocoA[$nLinhaA][$nLoopCBlocoA-1] = $aSplitLinha[$nLoopCBlocoA]
					Next
					$g_aLinhasBlocoA[$nLinhaA][19] = $nHeadBloco

;~ 					$nPosAtual = UBound($g_aLinhasBlocoA)-1 ;No caso dos A170 sempre vai ser a ultima linha gravada
					$nBloco += 1 ;Avança o contador do bloco A100 atual
					$g_aLinhasBlocoA[$nLinhaA][2] = $nBloco ;Atualiza o contador dentro do bloco da nota já que pode ter linhas duplicadas eliminadas

					;Localiza a linha do Array dos A100 correspondente a A170 atual e converte String para Numero e soma nos Valores 1,2,3
					$nPosHeadAtual = _ArrayBinarySearch($aLinhasMestre,$g_aLinhasBlocoA[$nLinhaA][19],0,UBound($aLinhasMestre)-1,0)

					$aLinhasMestre[$nPosHeadAtual][1] += Number(StringReplace($g_aLinhasBlocoA[$nLinhaA][5], ",", "."))
					$aLinhasMestre[$nPosHeadAtual][2] += Number(StringReplace($g_aLinhasBlocoA[$nLinhaA][10], ",", "."))
					$aLinhasMestre[$nPosHeadAtual][3] += Number(StringReplace($g_aLinhasBlocoA[$nLinhaA][12], ",", "."))
					$aLinhasMestre[$nPosHeadAtual][4] += Number(StringReplace($g_aLinhasBlocoA[$nLinhaA][14], ",", "."))
					$aLinhasMestre[$nPosHeadAtual][5] += Number(StringReplace($g_aLinhasBlocoA[$nLinhaA][16], ",", "."))

				EndIf

			Case $nIdBloco = "A990"

				;1 - MARCA FLAG PARA AUMENTAR O TAMANHO DO ARRAY DO BLOCO A E ATUALIZA VARIAVEL DE CONTROLE
				$nLinhaA = UBound($g_aLinhasBlocoA)
				ReDim $g_aLinhasBlocoA[$nLinhaA+1][23]

				;1 - INSERIR DADOS NO ARRAY DE LINHAS DO BLOCO A
				$g_aLinhasBlocoA[$nLinhaA][0] = $nLinhaGeral
				$g_aLinhasBlocoA[$nLinhaA][1] = "A990"
				$g_aLinhasBlocoA[$nLinhaA][2] = $nLinhaA+1

			Case Else

				If $g_bFoundBlocoA Then
					$g_nPosFimBlocoA = $nLinhaGeral
					$g_bFoundBlocoA = False
				EndIf

		EndSwitch

		If @error Then
			MsgBox(0, @ScriptName, @error & ' - Erro lendo linhas do Bloco A')
			Exit 1
		EndIf

	Next

	ProgressSet(60, "Atualizando valores dos blocos A100...")

	;Zera posição de Head A100
	$nPosHeadAtual = 0
	$nPosAtual = 0

	;Posição maxima do Array de Head A100
	$nContLinhaGeral = UBound($g_aLinhasBlocoA)-1

	;Atualizar valores das linhas A100
	For $nPosAtual = 0 to UBound($aLinhasMestre)-1

		$nPosHeadAtual = _ArrayBinarySearch($g_aLinhasBlocoA,$aLinhasMestre[$nPosAtual][0],0,$nContLinhaGeral,0)

		$g_aLinhasBlocoA[$nPosHeadAtual][12] = StringReplace(StringFormat("%.2f",$aLinhasMestre[$nPosAtual][1]), ".", ",")
		$g_aLinhasBlocoA[$nPosHeadAtual][15] = StringReplace(StringFormat("%.2f",$aLinhasMestre[$nPosAtual][2]), ".", ",")
		$g_aLinhasBlocoA[$nPosHeadAtual][16] = StringReplace(StringFormat("%.2f",$aLinhasMestre[$nPosAtual][3]), ".", ",")
		$g_aLinhasBlocoA[$nPosHeadAtual][17] = StringReplace(StringFormat("%.2f",$aLinhasMestre[$nPosAtual][4]), ".", ",")
		$g_aLinhasBlocoA[$nPosHeadAtual][18] = StringReplace(StringFormat("%.2f",$aLinhasMestre[$nPosAtual][5]), ".", ",")

	Next

EndFunc

Func GravaArquivo()

	Local Const $SF_ANSI = 1
    Local Const $SF_UTF8 = 4

	Local $nLoopLGeral = 0
	Local $nLoopLBlocoA = 0
	Local $nLoopCBlocoA = 0
	Local $nTotBlocoA = UBound($g_aLinhasBlocoA)-1
	Local $sLinhaBlocoA = ""
	Local $sIdBloco = ""

	Local $oHandleArqOut

	ProgressSet(41, "Preparando para gravar aquivo de saída...")

	$oHandleArqOut = FileOpen($g_sArquivoOut,2+512)

	If $oHandleArqOut = -1 Then
		; Mostra aviso e sai com status 1
		MsgBox(0, @ScriptName, 'Erro ao abrir o arquivo TXT de destino')
		Exit 1
	EndIf

	ProgressSet(45, "Gravando linhas originais até Bloco A...")
	;GRAVA TODAS AS LINHAS ATE BLOCO A1 COMEÇAR
	For $nLoopLGeral = 0 to $g_nPosIniBlocoA-1
		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($g_aLinhasTxtIn[$nLoopLGeral], $SF_UTF8), $SF_ANSI))
	Next

	ProgressSet(80, "Gravando linhas TRATADAS do Bloco A...")
	;GRAVA LINHAS ATUALIZADAS DO BLOCO A1
	For $nLoopLBlocoA = 0 to $nTotBlocoA

		$sIdBloco = $g_aLinhasBlocoA[$nLoopLBlocoA][1]

		If $sIdBloco = "A001" Or $sIdBloco = "A010" Or $sIdBloco = "A990" Then
			$sLinhaBlocoA = "|" & $g_aLinhasBlocoA[$nLoopLBlocoA][1] & "|" & $g_aLinhasBlocoA[$nLoopLBlocoA][2] & "|"

		ElseIf $sIdBloco = "A100" Then
			$sLinhaBlocoA = "|"
			For $nLoopCBlocoA = 1 to 21
				$sLinhaBlocoA &= $g_aLinhasBlocoA[$nLoopLBlocoA][$nLoopCBlocoA] & "|"
			Next

		ElseIf $sIdBloco = "A170" Then
			$sLinhaBlocoA = "|"
			For $nLoopCBlocoA = 1 to 18
				$sLinhaBlocoA &= $g_aLinhasBlocoA[$nLoopLBlocoA][$nLoopCBlocoA] & "|"
			Next

		EndIf

		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($sLinhaBlocoA, $SF_UTF8), $SF_ANSI))

	Next

	ProgressSet(90, "Gravando linhas finais do Arquivo...")

	;GRAVA RESTANTE DO ARQUIVO
	For $nLoopLGeral = $g_nPosFimBlocoA to UBound($g_aLinhasTxtIn)-1
		FileWriteLine($oHandleArqOut,BinaryToString(StringToBinary($g_aLinhasTxtIn[$nLoopLGeral], $SF_UTF8), $SF_ANSI))
	Next

	FileClose($oHandleArqOut)

EndFunc