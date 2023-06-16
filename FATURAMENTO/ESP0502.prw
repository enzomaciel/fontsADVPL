
//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
	
//Constantes
#Define STR_PULA		Chr(13)+Chr(10)
	
//{Protheus.doc} ESP0502
//Relatório - Pedido x Item         
//@author Enzo Maciel
//@since 23/05/2023
//@version 1.0
//@example
//	u_ESP0502()


User Function ESP0502()
	Local aArea   := GetArea()
	Local oReport
	Local lEmail  := .F.
	Local cPara   := ""
	Private cPerg := ""
	
	//Definições da pergunta
	cPerg := "ESP0502"

	ValidPerg()
	If !Pergunte(cPerg, .t.)
		Return .f.
	EndIf
	
	//Se a pergunta não existir, zera a variável
	DbSelectArea("SX1")
	SX1->(DbSetOrder(1)) //X1_GRUPO + X1_ORDEM
	If ! SX1->(DbSeek(cPerg))
		cPerg := Nil
	EndIf
	
	//Cria as definições do relatório
	oReport := fReportDef()
	
	//Será enviado por e-Mail?
	If lEmail
		oReport:nRemoteType := NO_REMOTE
		oReport:cEmail := cPara
		oReport:nDevice := 3 //1-Arquivo,2-Impressora,3-email,4-Planilha e 5-Html
		oReport:SetPreview(.F.)
		oReport:Print(.F., "", .T.)
	//Senão, mostra a tela
	Else
		oReport:PrintDialog()
	EndIf
	
	RestArea(aArea)
Return
	
/*-------------------------------------------------------------------------------*
 | Func:  fReportDef                                                             |
 | Desc:  Função que monta a definição do relatório                              |
 *-------------------------------------------------------------------------------*/
	
Static Function fReportDef()
	Local oReport
	Local oSectDad1 := Nil
	Local oSectDad2 := Nil
	Local oBreak := Nil
	
	//Criação do componente de impressão
	oReport := TReport():New(	"ESP0502",;		//Nome do Relatório
								"Pedido x item",;		//Título
								cPerg,;		//Pergunte ... Se eu defino a pergunta aqui, será impresso uma página com os parâmetros, conforme privilégio 101
								{|oReport| fRepPrint(oReport)},;		//Bloco de código que será executado na confirmação da impressão
								)		//Descrição
	oReport:SetTotalInLine(.F.)
	oReport:lParamPage := .T.
	oReport:oPage:SetPaperSize(9) //Folha A4
	oReport:SetLandscape()
	
	//Criando a seção de dados
	oSectDad1 := TRSection():New(	oReport,;		//Objeto TReport que a seção pertence
									"Pedidos",;		//Descrição da seção
									{"QRY_AUX"})		//Tabelas utilizadas, a primeira será considerada como principal da seção
	oSectDad1:SetTotalInLine(.F.)  //Define se os totalizadores serão impressos em linha ou coluna. .F.=Coluna; .T.=Linha
	//Colunas do relatório
	TRCell():New(oSectDad1, "C5_FILIAL", "QRY_AUX", "Codigo", /*Picture*/, 2, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad1, "C5_NUM", "QRY_AUX", "Pedido", /*Picture*/, 6, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad1, "C5_CLIENTE", "QRY_AUX", "Cliente", /*Picture*/, 6, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad1, "A1_NOME", "QRY_AUX", "Nome", /*Picture*/, 50, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad1, "C5_EMISSAO", "QRY_AUX", "DT. Emissao", /*Picture*/, 10, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad1, "C5_FECENT", "QRY_AUX", "DT.Entrega", /*Picture*/, 10, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad1, "C5_NATUREZ", "QRY_AUX", "Natureza", /*Picture*/, 6, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)

	
	//Criando a seção de dados								
	oSectDad2 := TRSection():New(	oReport,;		//Objeto TReport que a seção pertence
									 "Produtos",;		//Descrição da seção
									{"QRY_AUX"})		//Tabelas utilizadas, a primeira será considerada como principal da seção
	oSectDad2:SetTotalInLine(.F.)  //Define se os totalizadores serão impressos em linha ou coluna. .F.=Coluna; .T.=Linha
	oSectDad2:SetLeftMargin(3)
	oSectDad2:SetTitle('')

	//Colunas do relatório
	TRCell():New(oSectDad2, "C6_ITEM", "QRY_AUX", "Item", /*Picture*/, 6, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad2, "C6_PRODUTO", "QRY_AUX", "Produto", /*Picture*/, 14, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad2, "C6_DESCRI", "QRY_AUX", "Descricao", /*Picture*/, 30, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
    TRCell():New(oSectDad2, "C6_QTDVEN", "QRY_AUX", "Quantidade", /*Picture*/, 8, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad2, "C6_PRCVEN", "QRY_AUX", "Preço UN.", /*Picture*/, 8, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	TRCell():New(oSectDad2, "C6_VALOR", "QRY_AUX", "Preço Total", /*Picture*/, 10, /*lPixel*/,/*{|| code-block de impressao }*/,/*cAlign*/,/*lLineBreak*/,/*cHeaderAlign */,/*lCellBreak*/,/*nColSpace*/,/*lAutoSize*/,/*nClrBack*/,/*nClrFore*/,/*lBold*/)
	
	    //Adicione mais campos, caso necessário
	
	//Definindo a quebra
	oBreak := TRBreak():New(oSectDad1,{|| QRY_AUX->(C6_valor) },{|| "Total do Pedido:" })
	oSectDad1:SetHeaderBreak(.F.)
	
	//Definindo a quebra
	oBreak1 := TRBreak():New(oSectDad2,{|| QRY_AUX->(C6_VALOR) },{|| "Total do Geral:" })
	oSectDad2:SetHeaderBreak(.T.)
	
	oSectDad2:SetHeaderSection(.T.)
	oSectDad2:SetLineStyle()
	
	//Totalizadores
	oFunTot1 := TRFunction():New(oSectDad2:Cell("C6_VALOR"),"Total Geral","SUM",oBreak,,"@E 99999999.99")
	oFunTot1:SetEndReport(.F.)
	
    //Aqui, farei uma quebra  por seção
	oSectDad1:SetPageBreak(.T.)
	oSectDad1:SetTotalText(" ")			
	
Return oReport
	
/*-------------------------------------------------------------------------------*
 | Func:  fRepPrint                                                              |
 | Desc:  Função que imprime o relatório                                         |
 *-------------------------------------------------------------------------------*/
	
Static Function fRepPrint(oReport)
	Local aArea    := GetArea()
	Local cQryAux  := ""
	Local oSectDad1 := Nil
	Local oSectDad2 := Nil
	Local nAtual   := 0
	Local nTotal   := 0
	Local cVend    := ""
	
	//Pegando as seções do relatório
	oSectDad1 := oReport:Section(1)
	oSectDad2 := oReport:Section(2)
	
	//Montando consulta de dados
	cQryAux := ""
	cQryAux += "SELECT C5_FILIAL, C5_NUM, C5_CLIENTE, A1_COD,A1_NOME,C6_NUM, C6_PRODUTO, C6_ITEM,C6_DESCRI,C6_QTDVEN,C6_PRCVEN,C6_VALOR, C5_EMISSAO, C5_FECENT, C5_NATUREZ FROM SC5990 SC5"	    + STR_PULA
	cQryAux += "		 INNER JOIN SA1990 SA1 ON ( SC5.C5_CLIENTE = SA1.A1_COD AND SA1.D_E_L_E_T_ = ' ' )"		+ STR_PULA
    cQryAux += "		 INNER JOIN SC6990 SC6 ON (SC5.C5_NUM = SC6.C6_NUM AND SC6.D_E_L_E_T_ = ' ')"		+ STR_PULA
	cQryAux += "WHERE SC5.D_E_L_E_T_ = ' ' "		+ STR_PULA
	If !Empty(mv_par01) .And. !Empty(mv_par02)
		If(mv_par01 > mv_par02)
			MsgInfo("Parametro PEDIDO incorreto","ERRO PARAMETRO")
		EndIf
        cQryAux += "AND SC5.C5_NUM BETWEEN '" + mv_par01 + "' AND " + mv_par02 + " "		+ STR_PULA
    EndIf
	If !Empty(mv_par03) .And. !Empty(mv_par04)
		If(mv_par03 > mv_par04)
			MsgInfo("Parametro DATA incorreto","ERRO PARAMETRO")
		EndIf
		cQryAux += "AND SC5.C5_EMISSAO BETWEEN '" + DtoS(mv_par03) + "' AND " + DtoS(mv_par04) + " "		+ STR_PULA
	EndIf 
	If !Empty(mv_par05) .And. !Empty(mv_par06)
		If(mv_par05 > mv_par06)
			MsgInfo("Parametro FILIAL incorreto","ERRO PARAMETRO")
		EndIf
		cQryAux += "AND SC5.C5_FILIAL BETWEEN '" + mv_par05 + "' AND " + mv_par06 + " "		+ STR_PULA
	EndIf 


	cQryAux += "AND SA1.A1_MSBLQL <> '1'"		+ STR_PULA
	cQryAux += "ORDER BY C5_FILIAL, C5_NUM"		+ STR_PULA
	cQryAux := ChangeQuery(cQryAux)
	
	//Executando consulta e setando o total da régua
	TCQuery cQryAux New Alias "QRY_AUX"
	Count to nTotal
	oReport:SetMeter(nTotal)
	TCSetField("QRY_AUX", "A1_ULTCOM", "D")
	
	//Enquanto houver dados
	//oSectDad1:Init()
	QRY_AUX->(DbGoTop())
	While ! QRY_AUX->(Eof())
	
		If oReport:Cancel()
			Exit
		EndIf
	
		//inicializo a primeira seção
		oSectDad1:Init()
			
		//Incrementando a régua
		nAtual++
		oReport:SetMsgPrint("Imprimindo registro "+cValToChar(nAtual)+" de "+cValToChar(nTotal)+"...")
		oReport:IncMeter()
		
		cVend := QRY_AUX->(C5_NUM)
		
		//Imprimindo a linha atual
		oSectDad1:PrintLine()
		
		oSectDad2:Init()
		
		While QRY_AUX->(C6_NUM) == cVend
		
			nAtual++
			oReport:SetMsgPrint("Imprimindo registro "+cValToChar(nAtual)+" de "+cValToChar(nTotal)+"...")
			oReport:IncMeter()
			oSectDad2:Printline()
		
			QRY_AUX->(DbSkip())
		EndDo	
	EndDo
	oSectDad1:Finish()
	oSectDad2:Finish()
	QRY_AUX->(DbCloseArea())
	
	RestArea(aArea)
Return

Static Function ValidPerg()
	Local aArea:= GetArea()                                                                                                                          
	Local aRegs
	Local i, j
	
	DbSelectArea("SX1")
	DbSetOrder(1)
	cPerg:= Padr(cPerg, 10)
	
	aRegs:= {}
	aAdd(aRegs, {cPerg,"01","Pedido De           ?","","","mv_ch1","C",06,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
	aAdd(aRegs, {cPerg,"02","Pedido Até          ?","","","mv_ch2","C",06,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
    aAdd(aRegs, {cPerg,"03","Pedido De           ?","","","mv_ch3","D",10,0,0,"G","","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	aAdd(aRegs, {cPerg,"04","Pedido Até          ?","","","mv_ch4","D",10,0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	aAdd(aRegs, {cPerg,"05","Pedido De           ?","","","mv_ch5","C",06,0,0,"G","","mv_par05","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
	aAdd(aRegs, {cPerg,"06","Pedido Até          ?","","","mv_ch6","C",06,0,0,"G","","mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})

	For i:= 1 To Len(aRegs)
		If !DbSeek(cPerg + aRegs[i,2])
			RecLock("SX1", .t.)
			For j:= 1 To fCount()
				FieldPut(j, aRegs[i,j])
			Next j
			MsUnlock()
			DbCommit()
		Endif
	Next i
		
	RestArea(aArea)		
Return
