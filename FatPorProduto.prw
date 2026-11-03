#Include "Protheus.ch"

Static cGrpPerg := Upper("pFatProd")


/*=============================
  {ADVPL/TLPP}
  @descrição: Montagem do objeto de impressão
  @autor Marcelo Iuspa em https://github.com/marceloiuspa/iprotools)
  @data  29/10/2026
  =============================*/

User Function RFatProd //Trocar nome para não ter repetidos com mesmo nome
oReport:= ReportDef()
oReport:PrintDialog()
Return


/*=============================
  {ADVPL/TLPP}
  @descrição: Definições do relatório e apresentação da interface
  @autor Gerado por iProTools (download patch em https://github.com/marceloiuspa/iprotools)
  @data  29/10/2026
  =============================*/

Static Function ReportDef()
Local oReport
Local oSection1
Local cRel     := "FatProd"
Local cSecao   := "FatProd"
Local cTitulo  := "246 Relatorio gerado pelo iProTools"
Local cDescr   := Nil // Depois vou preencher com oReport:SetTitle

oReport:= TReport():New(cRel,cTitulo, {|| Perg(.T.)}, {|oReport| ReportPrint(oReport)}, cDescr)

oReport:SetLandscape()
oReport:SetTotalInLine(.F.)

oSection1 := TRSection():New(oReport, cSecao)
oSection1:SetHeaderPage(.F.)

Return(oReport)

/*=============================
  {ADVPL/TLPP}
  @descrição: Impressão do relatório
  @autor Gerado por iProTools (download patch em https://github.com/marceloiuspa/iprotools)
  @data  29/10/2026
  =============================*/

Static Function ReportPrint(oReport)
Local oSection1  := Nil
Local aSavAre    := GetArea()
Local lExcel     := oReport:nDevice == 4  // IMP_EXCEL

Private cAliasTop  := "QRY" + StrTran(Time(), ":", "")


//oReport:nFontBody := 10

Perg(.F.)

If ! OpenQuery(cAliasTop)
	MsgInfo("Sem dados para imprimir")
	Return
Endif

oReport:SetTitle("Relatório de Vendas por Período entre " + Dtoc(MV_PAR01, "dd/mm/yyyy") + " e " + Dtoc(MV_PAR02, "dd/mm/yyyy"))

oSection1 := oReport:Section(1)

/*
METHOD New(oParent,cName,cAlias,cTitle,cPicture,nSize,lPixel,bBlock,cAlign,lLineBreak,cHeaderAlign,lCellBreak,nColSpace,lAutoSize,nClrBack,nClrFore,lBold)
*/

TRCell():New(oSection1, "D2_COD"        , cAliasTop  , "Produto"         , ""                ,   18,/*lPixel*/, {|| D2_COD}       , /*cAlign*/)
TRCell():New(oSection1, "B1_DESC"       , cAliasTop  , "Descricao"       , ""                ,   52,/*lPixel*/, {|| B1_DESC}      , /*cAlign*/)
TRCell():New(oSection1, "B1_TIPO"       , cAliasTop  , "Tipo"            , ""                ,   18,/*lPixel*/, {|| B1_TIPO}      , /*cAlign*/)
TRCell():New(oSection1, "B1_GRUPO"      , cAliasTop  , "Grupo"           , ""                ,   18,/*lPixel*/, {|| B1_GRUPO}     , /*cAlign*/)
TRCell():New(oSection1, "BM_DESC"       , cAliasTop  , "Desc Grupo"      , ""                ,   32,/*lPixel*/, {|| BM_DESC}      , /*cAlign*/)
TRCell():New(oSection1, "D2_UM"         , cAliasTop  , "Unidade"         , ""                ,   18,/*lPixel*/, {|| D2_UM}        , /*cAlign*/)
TRCell():New(oSection1, "SAIDAS"        , cAliasTop  , "Qtde. Notas"     , "@E 99,999,999"   ,   10,/*lPixel*/, {|| SAIDAS}       , "RIGHT")
TRCell():New(oSection1, "D2_QUANT"      , cAliasTop  , "Quantidade"      , "@E 999,999.99"   ,   18,/*lPixel*/, {|| D2_QUANT}     , "RIGHT")
TRCell():New(oSection1, "D2_TOTAL"      , cAliasTop  , "Vlr.Total"       , "@E 999,999.99"   ,   18,/*lPixel*/, {|| D2_TOTAL}     , "RIGHT")
TRCell():New(oSection1, "A1_EST"        , cAliasTop  , "Estado"          , ""                ,   18,/*lPixel*/, {|| A1_EST}       , /*cAlign*/)

If MV_PAR13
	oBreak := TRBreak():New(oSection1,{|| A1_EST},NIL,.F.)
Else
	oBreak := TRBreak():New(oSection1,{|| EOF()},NIL,.F.)
Endif

TRFunction():New(oSection1:Cell("SAIDAS"   ),,"SUM",oBreak,,,,.F.,.F.,)
TRFunction():New(oSection1:Cell("D2_QUANT" ),,"SUM",oBreak,,,,.F.,.F.,)
TRFunction():New(oSection1:Cell("D2_TOTAL" ),,"SUM",oBreak,,,,.F.,.F.,)

oSection1:Init()
oReport:SetMeter((cAliasTop)->(LastRec()))

Do While ! oReport:Cancel() .And. ! Eof()
	If oReport:Cancel()
		Exit
	EndIf

	oReport:IncMeter()

	oSection1:PrintLine()

	If lExcel
		oSection1:lHeaderSection := .F.
	Endif

	dbSkip()
Enddo

oSection1:Finish()

(cAliasTop)->(dbCloseArea())
RestArea(aSavAre)

Return
/*  Programa ADVPL montado pelo aplicativo iProTools patch disponível em
    https://github.com/marceloiuspa/iprotools
    Data 29/10/2026 */

Static Function OpenQuery(cAliasTop)

cQuery := " SELECT D2_COD, B1_DESC, B1_TIPO, B1_GRUPO, BM_DESC, D2_UM, COUNT(*) SAIDAS, SUM(D2_QUANT) D2_QUANT, SUM(D2_TOTAL) D2_TOTAL, A1_EST "                                                                                            + ;
          " FROM " + RetSqlName("SD2") + " "                                                                                                                                                                                                + ;
          " INNER JOIN " + RetSqlName("SF2") + " SF2 ON F2_DOC   = D2_DOC     AND D2_SERIE = F2_SERIE AND F2_CLIENTE = D2_CLIENTE AND F2_LOJA = D2_LOJA AND F2_FILIAL = '" + xFilial("SF2") + "' AND SF2.D_E_L_E_T_ = ' ' "                 + ;
          " INNER JOIN " + RetSqlName("SA1") + " SA1 ON A1_COD   = F2_CLIENTE AND A1_LOJA = F2_LOJA AND A1_FILIAL = '" + xFilial("SA1") + "' AND SA1.D_E_L_E_T_ = ' ' "                                                                     + ;
          " INNER JOIN " + RetSqlName("SB1") + " SB1 ON B1_COD   = D2_COD     AND B1_FILIAL = '" + xFilial("SB1") + "' AND SB1.D_E_L_E_T_ = ' ' "                                                                                           + ;
          " INNER JOIN " + RetSqlName("SBM") + " SBM ON B1_GRUPO = BM_GRUPO   AND BM_FILIAL = '" + xFilial("SBM") + "' AND SBM.D_E_L_E_T_ = ' ' "                                                                                           + ;
          " WHERE D2_TIPO = 'N' AND D2_SERVIC = ' ' "                                                                                                                                                                                       + ;
          " AND D2_TES IN ( SELECT F4_CODIGO FROM " + RetSqlName("SF4") + " SF4 WHERE F4_PODER3  = 'N' AND F4_TIPO = 'S' AND F4_DUPLIC  = 'S' AND F4_ESTOQUE = 'S' AND SF4.D_E_L_E_T_ = ' '  AND F4_FILIAL = '" + xFilial("SF4") + "') "    + ;
          " GROUP BY D2_COD, B1_DESC, B1_TIPO, B1_GRUPO, BM_DESC, D2_UM, A1_EST "

	If MV_PAR13
		cQuery += " ORDER BY A1_EST, SUM(D2_TOTAL) DESC "
	Else
		cQuery += " ORDER BY SUM(D2_TOTAL) DESC "
	Endif

dbUseArea(.T., "TOPCONN", TcGenQry(,,cQuery), cAliasTop, .F., .T.)

SX3->(dbSetOrder(2)) // X3_CAMPO
aEval(dbStruct(), {|z| If(SX3->(dbSeek(z[1])) .And. SX3->X3_TIPO $ "N;D;M", TcSetField(cAliasTop, z[1], SX3->X3_TIPO, SX3->X3_TAMANHO, SX3->X3_DECIMAL), Nil)})
SX3->(dbSetOrder(1)) // X3_ARQUIVO+X3_ORDEM

Return(! Eof())

/*=============================
  {ADVPL/TLPP}
  @descrição: Parâmetros do relatório
  @autor Marcelo Iuspa
  @data 29/10/2026
  =============================*/
Static Function Perg(lAsk)
Local aRetorn := {}
Local aParam  := {}
Local nLoop   := Nil

Default lAsk  := .T.

aAdd(aParam, {1, "Data de Emissao De"           ,  Stod("20260101")       ,,,          , ".T.",  48, .F.})  // MV_PAR01
aAdd(aParam, {1, "Data de Emissao Até"          ,  Stod("20261231")       ,,,          , ".T.",  48, .F.})  // MV_PAR02

aAdd(aParam, {1, "Estado do cliente De"         , "     "                 ,,, "12"     , ".T.",  30, .F.})  // MV_PAR03
aAdd(aParam, {1, "Estado do cliente Até"        , "zzzzz"                 ,,, "12"     , ".T.",  30, .F.})  // MV_PAR04

aAdd(aParam, {1, "Tipo de Produto (MP,PA,.) De" , "  "                    ,,, "02"     , ".T.",  12, .F.})  // MV_PAR05
aAdd(aParam, {1, "Tipo de Produto (MP,PA,.) Até", "zz"                    ,,, "02"     , ".T.",  12, .F.})  // MV_PAR06

aAdd(aParam, {1, "Grupo de Estoque De"          , "    "                  ,,, "SBM"    , ".T.",  24, .F.})  // MV_PAR07
aAdd(aParam, {1, "Grupo de Estoque Até"         , "zzzz"                  ,,, "SBM"    , ".T.",  24, .F.})  // MV_PAR08

aAdd(aParam, {1, "Codigo do Produto De"         , "               "       ,,, "SB1"    , ".T.",  90, .F.})  // MV_PAR09
aAdd(aParam, {1, "Codigo do Produto Até"        , "zzzzzzzzzzzzzzz"       ,,, "SB1"    , ".T.",  90, .F.})  // MV_PAR10

aAdd(aParam, {1, "Unidade de Medida De"         , "  "                    ,,, "SAH"    , ".T.",  12, .F.})  // MV_PAR11
aAdd(aParam, {1, "Unidade de Medida Até"        , "zz"                    ,,, "SAH"    , ".T.",  12, .F.})  // MV_PAR12

aAdd( aParam, {4,"Quebrar por Estado"           , .F., "Se marcado, totaliza por estado (UF)"  , 110, .T., .F.}) // MV_PAR13

If lAsk
	If ! ParamBox(aParam,'Parâmetros',@aRetorn, , , , , , , cGrpPerg, .T., .T.)
		Return(.F.)
	Endif
Else
	For nLoop := 1 to Len(aParam)
		&("MV_PAR"+AllTrim(STRZERO(nLoop,2,0))) := ParamLoad(__cUserID+"_" + cGrpPerg,aParam,nLoop,aParam[nLoop,3],Iif(Len(aParam[nLoop])>9,aParam[nLoop,10],.F.))
	Next
Endif
Return
