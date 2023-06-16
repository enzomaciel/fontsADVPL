#Include "Protheus.ch"
#Include "TopConn.ch"
#include "TOTVS.ch"

#Define STR_PULA Chr(13)+Chr(10)

User Function ESPD27()
    Local aArea := GetArea()
    Local oReport
    Local lEmail := .F.
    Local cPara := ""
    Private cPerg := ""

    //Definicao da pergunta
    cPerg := "ESPD27"

    ValidPerg()
    If !Pergunte(cPerg,.t.)
        return .f.
    EndIf 

    //Se a pergunta nao existir, zera a variavel
    DBSelectArea("SX1")
    SX1->(DbSetOrder(1))
    If !SX1->(DbSeek(cPerg))
        cPerg := Nil
    EndIf 

    //Cria as definicoes do relatorio
    oReport := fReportDef()

    If lEmail
        oReport:nRemoteType := NO_REMOTE
        oReport:cEmail := cPara
        oReport:nDevice := 3
        oReport:SetPreview(.F.)
        oReport:Print(.F.,"",.T.)
    Else 
        oReport:PrintDialog()
    EndIf 
    RestArea(aArea)
Return 

Static Function fReportDef()
    Local oReport
    Local oSectDad1 := Nil 

    //Criando o componente de impressao 
    oReport := TReport():New("ESPD27",;
                            "Ranking de Produtos",;
                            cPerg,;
                            {|oReport| fRepPrint(oReport)},;
                            )
    oReport:SetTotalInLine(.F.)
    oReport:lParamPage := .T.
    oReport:oPage:SetPaperSize(9)
    oReport:SetLandscape()

    //Criando a sessao de dados
    oSectDad1 := TRSection():New(   oReport,;
                                    "Pedidos",;
                                    {"QRY_AUX"})
    oSectDad1:SetTotalInLine(.F.)
    TRCell():New(oSectDad1,"COD_VAR","","CODIGO",,14,,,,,,,,,,,)
    TRCell():New(oSectDad1,"DESC_VAR","","DESCRICAO",,30,,,,,,,,,,,)
    TRCell():New(oSectDad1,"QUAN_VAR","","QUANTIDADE",,14,,,"L",,,,,,,,)
    TRCell():New(oSectDad1,"VL_VAR","","VL.TOTAL",,14,,,"R",,,,,,,,)
    oSectDad1:SetPageBreak(.T.)
    oSectDad1:SetTotalText(" ")
return oReport

Static Function fRepPrint(oReport)
    Local aArea := GetArea()
    Local cQryAux := ""
    Local oSectDad1 := Nil 
    Local nTotal := 0
    Local nAtual := 0
    Local alistprod := {}
    Local n := 0
    Local valorTotal := 0

    //Pegando as sessao do relatorio
    oSectDad1 := oReport:Section(1)

    //Montando consulta de dados
    cQryAux := ""
    If !Empty(mv_par03) .And. !Empty(mv_par04)
        If(mv_par03 > mv_par04)
            MsgInfo("Parametro DATA incorreto","ERRO PARAMETRO")
        EndIf
        cQryAux += "SELECT B1_COD,B1_DESC,C6_PRODUTO,C6_ENTREG,C6_FILIAL,SUM(C6_QTDVEN) AS QUANTIDADE,SUM(C6_VALOR) AS VALOR FROM SC6990 SC6, SB1990 SB1 WHERE B1_COD = C6_PRODUTO" + STR_PULA
        cQryAux += "AND SC6.C6_ENTREG BETWEEN '" + DtoS(mv_par03) + "' AND " + DtoS(mv_par04) + " "		+ STR_PULA
    Else 
	    cQryAux += "SELECT B1_COD,B1_DESC,C6_PRODUTO,C6_FILIAL,SUM(C6_QTDVEN) AS QUANTIDADE,SUM(C6_VALOR) AS VALOR FROM SC6990 SC6, SB1990 SB1 WHERE B1_COD = C6_PRODUTO" + STR_PULA
    EndIf
    If !Empty(mv_par01) .And. !Empty(mv_par02)
        If(mv_par01 > mv_par02)
            MsgInfo("Parametro PEDIDO incorreto","ERRO PARAMETRO")
        EndIf
        cQryAux += "AND SB1.B1_COD BETWEEN '" + mv_par01 + "' AND " + mv_par02 + " "		+ STR_PULA
    EndIf
    If !Empty(mv_par05) .And. !Empty(mv_par06)
        If(mv_par05 > mv_par06)
            MsgInfo("Parametro FILIAL incorreto","ERRO PARAMETRO")
        EndIf
        cQryAux += "AND SC5.C6_FILIAL BETWEEN '" + mv_par05 + "' AND " + mv_par06 + " "		+ STR_PULA
    EndIf 
    If !Empty(mv_par03) .And. !Empty(mv_par04)
        cQryAux += "GROUP BY C6_PRODUTO,B1_COD,B1_DESC,C6_FILIAL,C6_ENTREG" + STR_PULA
    Else 
	   cQryAux += "GROUP BY C6_PRODUTO,B1_COD,B1_DESC,C6_FILIAL" + STR_PULA
    EndIf
	cQryAux := ChangeQuery(cQryAux)

    //Executando consulta e setando  total da regua
    TCQuery cQryAux New Alias "QRY_AUX"
    Count to nTotal
    oReport:SetMeter(nTotal)

    //Enquanto houver dados
    QRY_AUX->(DbGoTop())
    // Separando os valores
    While ! QRY_AUX->(Eof())
        
        COD_VAR := QRY_AUX->(B1_COD)  
        DESC_VAR := QRY_AUX->(B1_DESC)
        QUANT_VAR := QRY_AUX->(QUANTIDADE)
        VL_VAR := QRY_AUX->(VALOR)
        if(len(alistprod)==0)
            AAdd(alistprod,{COD_VAR,DESC_VAR,QUANT_VAR,VL_VAR})
        elseif(ASCAN(alistprod[1],COD_VAR) == 0)
            AAdd(alistprod,{COD_VAR,DESC_VAR,QUANT_VAR,VL_VAR})
        else 
            aCont := len(alistprod) -1
            aVar := ASCAN(alistprod[1],COD_VAR)
            aTotal := VL_VAR + alistprod[aVar,4]
            aQuant := QUANT_VAR + alistprod[aVar,3]
            ADEL(alistprod,ASCAN(alistprod[1],COD_VAR))
            ASiZE(alistprod,aCont)
            AAdd(alistprod,{COD_VAR,DESC_VAR,aQuant,aTotal})
        endif    
        QRY_AUX->(DbSkip())
    EndDo
    // Imprimindo os valores
    While n < len(alistprod)
        n++
        If oReport:Cancel()
             Exit 
        EndIf 
        oSectDad1:Init()
        nAtual++
        oReport:SetMsgPrint("Imprimindo relatorio "+cValToChar(nAtual)+" de "+cValToChar(nTotal)+"...")
        oReport:IncMeter()
        oSectDad1:Cell("COD_VAR"):SetValue(alistprod[n,1])
        oSectDad1:Cell("DESC_VAR"):SetValue(alistprod[n,2])
        oSectDad1:Cell("QUAN_VAR"):SetValue(alistprod[n,3])
        oSectDad1:Cell("VL_VAR"):SetValue("R$ "+ Alltrim(Transform(NOROUND(alistprod[n,4],2), "@E 999,999,999.99")))
        valorTotal := valorTotal + alistprod[n,4]
        oSectDad1:PrintLine()
    EndDo
    oReport:PrintText("                                                       Valor Total: R$" + Alltrim(Transform(NOROUND(valorTotal,2), "@E 999,999,999.99")))
	oReport:SkipLine(1)
    oSectDad1:Finish()
    QRY_AUX->(DbCloseArea())
    RestArea(aArea)
Return

Static Function ValidPerg()
    Local aArea:= GetArea()
    Local aRegs
    Local i,j 

    DbSelectArea("SX1")
    DbSetOrder(1)
    cPerg:=Padr(cPerg,10)

    aRegs:= {}
	aAdd(aRegs, {cPerg,"01","Pedido De           ?","","","mv_ch1","C",06,0,0,"G","","mv_par01","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
	aAdd(aRegs, {cPerg,"02","Pedido Até          ?","","","mv_ch2","C",06,0,0,"G","","mv_par02","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
    aAdd(aRegs, {cPerg,"03","Pedido De           ?","","","mv_ch3","D",10,0,0,"G","","mv_par03","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	aAdd(aRegs, {cPerg,"04","Pedido Até          ?","","","mv_ch4","D",10,0,0,"G","","mv_par04","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
	aAdd(aRegs, {cPerg,"05","Pedido De           ?","","","mv_ch5","C",06,0,0,"G","","mv_par05","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})
	aAdd(aRegs, {cPerg,"06","Pedido Até          ?","","","mv_ch6","C",06,0,0,"G","","mv_par06","","","","","","","","","","","","","","","","","","","","","","","","","SC5","","","","",""})

    For i:= 1 To Len(aRegs)
        If !DbSeek(cPerg + aRegs[i,2])
            RecLock("SX1",.t.)
            For j:= 1 To fCount()
                FieldPut(j, aRegs[i,j])
            Next j 
            MsUnlock()
            DbCommit()
        EndIf
    Next i 
    RestArea(aArea)
Return




