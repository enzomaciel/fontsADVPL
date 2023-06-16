//Bibliotecas
#Include "Protheus.ch"
#Include "TopConn.ch"
	
//Constantes
#Define STR_PULA		Chr(13)+Chr(10)
	
//{Protheus.doc} ESP0503
//Relatório - Pedido x Item         
//@author Enzo Maciel
//@since 23/05/2023
//@version 1.1
//@example
//	u_ESP0503()


User Function ESP0503()
	Local aArea   := GetArea()
	Local oReport
	Local lEmail  := .F.
	Local cPara   := ""
	Private cPerg := ""
	
	//Definições da pergunta
	cPerg := "ESP0503"

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
		oReport:nDevice := 1 //1-Arquivo,2-Impressora,3-email,4-Planilha e 5-Html
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
	Local oSectDad3 := Nil
	//Criação do componente de impressão
	oReport := TReport():New(	"ESP0503",;		//Nome do Relatório
								"Pedido x item | " + "Periodo: " + DTOC(mv_par03) + " / " + DTOC(mv_par04),;		//Título
								cPerg,;		//Pergunte ... Se eu defino a pergunta aqui, será impresso uma página com os parâmetros, conforme privilégio 101
								{|oReport| fRepPrint(oReport)},;
							)
	oReport:nFontBody := 10
	oReport:HideParamPage()
	oReport:SetLandscape(.T.)
	//Criando a seção 1
	oSectDad1 := TRSection():New(oReport,;		//Objeto TReport que a seção pertence
									"Filial")//Tabelas utilizadas, a primeira será considerada como principal da seção
	oSectDad1:SetLineStyle(.T.)
	oSectDad1:SetLinesBefore(1) 
	TRCell():New(oSectDad1 , "FILIAL_VAR" , "" , "FILIAL" , /*<cPicture>*/ , 2 , .T. , /*<bBlock>*/ ,/*<cAlign>*/ , /*<lLineBreak>*/ , /*<cHeaderAlign>*/ , /*<lCellBreak>*/, /*<nColSpace>*/ , .T. , /*<nClrBack>*/ , /*<nClrFore>*/ , /*<lBold>*/ )
	oSectDad2:= TRSection():New(oReport,;
								"CLIENTE",;
								{"QRY_AUX"})
	oSectDad2:SetLinesBefore(1) 
	oSectDad2:SetLineStyle(.T.)
	oSectDad2:SetLeftMargin(2)
	TRCell():New(oSectDad2,"CLIENTE_VAR","","CLIENTE",,6)
	oSectDad3:= TRSection():New(oReport,;
								"PRODUTO",;
								{"QRY_AUX"})
	oSectDad3:SetLeftMargin(4)
	oSectDad3:SetLineStyle(.T.)		
	TRCell():New(oSectDad3,"PRODUTO_VAR","","PRODUTO",,16)										
	TRCell():New(oSectDad3,"QTD_VAR","","QTD",,15)
	TRCell():New(oSectDad3,"VL_VAR","","VALOR",,15)
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
	Local oSectDad3 := Nil
	Local nAtual   := 0
	Local nTotal   := 0
	Local aListFilial := {}
	Local aListCliente := {}
	Local aListProdutos := {}
	Local n := 0
	Local j:= 0
	Local z:= 0
	Local x
	Local valorTotal := 0
	

	//Pegando as seções do relatório
	oSectDad1 := oReport:Section(1)
	oSectDad2 := oReport:Section(2)
	oSectDad3 := oReport:Section(3)
	
	//Montando consulta de dados
	cQryAux := ""
	cQryAux += "SELECT C6_FILIAL,C6_CLI,A1_NOME,A1_COD ,C6_PRODUTO,C6_DESCRI,C6_QTDVEN,C6_VALOR,C6_ENTREG FROM SC6990 SC6, SA1990 SA1 WHERE A1_COD = C6_CLI" +STR_PULA
	If !Empty(mv_par01) .And. !Empty(mv_par02)
		If(mv_par01 > mv_par02)
			MsgInfo("Parametro PEDIDO incorreto","ERRO PARAMETRO")
		EndIf
		cQryAux += "AND SC6.C6_PRODUTO BETWEEN '" + mv_par01 + "' AND " + mv_par02 + " "		+ STR_PULA
	EndIf
	If !Empty(mv_par03) .And. !Empty(mv_par04)
		If(mv_par03 > mv_par04)
			MsgInfo("Parametro DATA incorreto","ERRO PARAMETRO")
		EndIf
		cQryAux += "AND SC6.C6_ENTREG BETWEEN '" + DtoS(mv_par03) + "' AND " + DtoS(mv_par04) + " "		+ STR_PULA
	EndIf 
	If !Empty(mv_par05) .And. !Empty(mv_par06)
		If(mv_par05 > mv_par06)
			MsgInfo("Parametro FILIAL incorreto","ERRO PARAMETRO")
		EndIf
			cQryAux += "AND SC6.C6_FILIAL BETWEEN '" + mv_par05 + "' AND " + mv_par06 + " "		+ STR_PULA
	EndIf 
	cQryAux += "GROUP BY C6_FILIAL,C6_CLI,A1_NOME,A1_COD ,C6_PRODUTO,C6_DESCRI,C6_QTDVEN,C6_VALOR,C6_ENTREG"
	cQryAux := ChangeQuery(cQryAux)

	//Executando consulta e setando o total da régua
	TCQuery cQryAux New Alias "QRY_AUX"
	Count to nTotal
	oReport:SetMeter(nTotal)
	QRY_AUX->(DbGoTop())
	While ! QRY_AUX->(Eof())
		//Processo de separação de filial
        FILIAL_VAR := QRY_AUX->(C6_FILIAL) 
		CLIENTE_VAR := QRY_AUX->(C6_CLI)
		PRODUTO_VAR := QRY_AUX->(C6_PRODUTO)
		EMISSAO_VAR := QRY_AUX->(C6_ENTREG)
		NOME_VAR := QRY_AUX->(A1_NOME)
		QUANTIDADE_VAR := QRY_AUX->(C6_QTDVEN)
		VALOR_VAR := QRY_AUX->(C6_VALOR)
        if(len(aListFilial)== 0)
             AAdd(aListFilial,{FILIAL_VAR,NOME_VAR})
        elseif(ASCAN(aListFilial,{|X| x[1] == FILIAL_VAR}) == 0)	
          	AAdd(aListFilial,{FILIAL_VAR,NOME_VAR})
        else 
        endif  
		//Processo de separação do Cliente
		if(len(aListCliente)== 0)
             AAdd(aListCliente,{CLIENTE_VAR,FILIAL_VAR})
        elseif(ASCAN(aListCliente,{|X| x[1] == CLIENTE_VAR}) == 0)	
          	AAdd(aListCliente,{CLIENTE_VAR,FILIAL_VAR})
        else 
        endif
		//Processo de separação do Produto  
		if(len(aListProdutos) ==0)
			AAdd(aListProdutos,{PRODUTO_VAR,CLIENTE_VAR,FILIAL_VAR,QUANTIDADE_VAR,VALOR_VAR})
		elseif (ASCAN(aListProdutos,{|X| X[1] == PRODUTO_VAR}) == 0 .OR. ASCAN(aListProdutos,{|X| X[2] == CLIENTE_VAR}) == 0 )
			AAdd(aListProdutos,{PRODUTO_VAR,CLIENTE_VAR,FILIAL_VAR,QUANTIDADE_VAR,VALOR_VAR})
		else
			aCont := len(aListProdutos) -1
			aTotal:= 0
			aQuant := 0
			For x:= 1 to Len(aListProdutos)
				bproduto := aListProdutos[x,1]
				bcliente:= aListProdutos[x,2]
				bfilial := aListProdutos[x,3]
				if(CLIENTE_VAR == bcliente .AND. FILIAL_VAR == bfilial .AND. PRODUTO_VAR == bproduto)
					aTotal := VALOR_VAR + aListProdutos[x,5]
					aQuant := QUANTIDADE_VAR + aListProdutos[x,4]
					ADEL(aListProdutos,x)
					ASiZE(aListProdutos,aCont)
					AAdd(aListProdutos,{PRODUTO_VAR,CLIENTE_VAR,FILIAL_VAR,aQuant,aTotal})
				elseif(CLIENTE_VAR == bcliente)
					AAdd(aListProdutos,{PRODUTO_VAR,CLIENTE_VAR,FILIAL_VAR,QUANTIDADE_VAR,VALOR_VAR})
				endif
			Next
		endif
        QRY_AUX->(DbSkip())
    EndDo
	//Enquanto houver dados
	//oSectDad1:Init()

	While n < len(aListFilial)
		n++
	//inicializo a primeira seção
	 	If oReport:Cancel()
             Exit 
        EndIf
		//Incrementando a régua
		nAtual++
		oReport:SetMsgPrint("Imprimindo Produtos "+cValToChar(len(aListFilial))+" de "+cValToChar(len(aListFilial))+"...")
		oReport:IncMeter(n)
		oSectDad1:Init(.F.)	
        oSectDad1:Cell("FILIAL_VAR"):SetValue(cValToChar(aListFilial[n,1]))
		oSectDad1:PrintHeader()
		oSectDad1:PrintLine()
		oReport:ThinLine()
		while j <len(aListCliente)
			j++
			If(aListFilial[n,1] == aListCliente[j,2])
				valorTotal := 0
				oSectDad2:Init(.F.)
				oSectDad2:Cell("CLIENTE_VAR"):SetValue(cValToChar(aListCliente[j,1]))
				oSectDad2:PrintHeader()
				oSectDad2:PrintLine()
				oSectDad3:PrintHeader()
				while z <len(aListProdutos)
					z++
					If(aListFilial[n,1]==aListProdutos[z,3] .AND. aListCliente[j,1] == aListProdutos[z,2])
						oSectDad3:Init(.F.)
						oSectDad3:Cell("PRODUTO_VAR"):SetValue(cValToChar(aListProdutos[z,1]))
						oSectDad3:Cell("QTD_VAR"):SetValue(cValToChar(aListProdutos[z,4]))
						oSectDad3:Cell("VL_VAR"):SetValue("R$ "+ Alltrim(Transform(NOROUND(aListProdutos[z,5],2), "@E 999,999,999.99")))
						valorTotal := valorTotal + aListProdutos[z,5]
						oSectDad3:PrintLine()
					EndIf
				EndDo
				oReport:PrintText("                                                       Valor Total: R$" + Alltrim(Transform(NOROUND(valorTotal,2), "@E 999,999,999.99")))
				oReport:SkipLine(1)
				oReport:ThinLine()
				z:= 0
			EndIf
		EndDo
		
		j:=0
    EndDo
	
	oReport:FatLine()
	
	oSectDad1:Finish()
	oSectDad2:Finish()
	oSectDad3:Finish()
	oReport:SkipLine(2) 
	QRY_AUX->(DbCloseArea())
	RestArea(aArea)
Return
/*-------------------------------------------------------------------------------*
 | Func:  ValidPerg                                                              |
 | Desc:  Função criando as perguntas                                            |
 *-------------------------------------------------------------------------------*/
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
