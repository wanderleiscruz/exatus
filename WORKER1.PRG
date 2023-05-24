//*********************************************************************************
// Programa.: CARTEIRA DE AÇÕES                                                   *
// Autor....: Wanderlei Souza Cruz                                                *
// Data.....: 21/01/2003                                                          *
// Aviso....: Copyright (c) 2003, Exatus Com.Proc.Dados Ltda, All Rights Reserved *
//*********************************************************************************
#include "FiveWin.ch"
#include "TSButton.ch"
#include "Report.ch"
#include "InKey.ch"
#include "winapi.ch"
#include "BMPGET.CH"
#include "FILEXLS.CH"
#include "sqllib.ch"
#include "FILEIO.CH"
#define  HKEY_CURRENT_USER 2147483649

Procedure MMMM(QZ1)
  QZ=QZ1
  ECORRETORA=.F.
  Do While .T.
    SysRefresh()
    If QZ<>3
      CLOSE DATABASE
      SELE 1
      If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"EMP")
        MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
        CLOSE DATABASE
        QUIT
      Else
        ARQ=LOCALCART+"EMPRESA1.IND"
        SET INDEX TO &ARQ
        SET ORDER TO 1
        BrowseC(1,"Manutenção dos Investidores", "Investidores", { || buscaEMP() },{ ||NCADEMP("I") },{ ||NCADEMP("A") })
      Endif
      SET ORDER TO 2
    Else
      SELE 29
      SET ORDER TO 1
      BrowseC(1,"Manutenção dos Investidores", "Investidores", { || buscaEMP() },{ ||NCADEMP("I") },{ ||NCADEMP("A") })
      Return
    Endif
    COD1=A01CODE
    M->CALCATE=A01CALCATE
    If EMPTY(M->CALCATE)
      M->CALCATE=M->DAT_HOJE
    Endif
    If A01CODNAC="S"
      ECORRETORA=.T.
    Endif
    _DATA=STRZERO(VAL(TRIM(EMP->a01data)),4)
    If VAL(COD1)=0
      CLOSE DATABASE
      MSGINFO("Codigo do Cliente igual a zero !",WSIST)
      LOOP
    Endif
    EMP1=ALLTRIM(A01NOME)
    _CGCCPF=STRZERO(VAL(a01CGCCPF),14)
    _LOCAL=LOCALCART+A01CODE+"/"
    _data=strzero(val(a01data),4)
    If VERDIRETORIO(_LOCAL)
      M->LOCAL=_LOCAL
      M->CONTAB=M->LOCAL
      EXIT
    Else
      MSGINFO("Não foi verificado o Diretório !",WSIST)
      LOOP
    Endif
  ENDDO
  If QZ<>3
    NOMEARQS(QZ)
    CLOSE DATABASE
  Endif
Return

Procedure NOMEARQS(QZ)
  _LOCAL=Strtran(_LOCAL,"/","\")
  M->CONTAB=_LOCAL
  M->LOCAL=_LOCAL
  _CGCCPF=STRZERO(VAL(a01CGCCPF),14)
  COD1=A01CODE
  M->CALCATE=A01CALCATE
  If EMPTY(M->CALCATE)
    M->CALCATE=M->DAT_HOJE
  Endif
  DECLARE VLABEL[ADIR(M->LOCAL+"TMP*.*")]
  ADIR(M->LOCAL+"TMP*.*",vlabel)
  VLABEL:=ASORT(VLABEL)
  I=0
  FOR I=1 TO LEN(VLABEL)
    SysRefresh()
    NARQ=M->LOCAL+VLABEL[I]
    Erase &NARQ
  NEXT I
  WEMP2=EMP2+TRIM(EMP1)+" ref:"+_DATA+" Local: "+_LOCAL
  If QZ=2
    SET MESSAGE OF oWnd TO WEMP2 CENTERED
  Endif
  _DATA=STRZERO(VAL(_DATA),4)
  arq01 = LOCALCART+"EMPRESAS.DBF"     // Cadastro de Empresas
  arq02 = LOCALCART+"OPERADOR.DBF"     // Cadastro de Operadores
  arq03 = _LOCAL+"PAPE2006.DBF"        // CadastPapeis da Carteira
  arq04 = _LOCAL+"MERC2006.DBF"        // Cadastro de Mercados
  ARQ05 = _LOCAL+"RDAR2006.DBF"        // Resumo de Auração de Ganhos Renda Variavel
  ARQ06 = LOCALCART+"BDIS.DBF"         // Arquivo de BDI completo
  ARQ07 = _LOCAL                       // Arquivos de Notas em XML
  ARQ08 = LOCALCART+"PARCMV.DBF"     // Parametrizaçào do arquivo CMVD
  arq10 = _LOCAL+"MOVT2006.DBF"        // Movimento Diario (Extra)
  arq11 = _LOCAL+"MOVN2006.DBF"        // Movimento das Notas
  arq12 = _LOCAL+"NOTA2006.DBF"        // Notas
  arq13 = _LOCAL+"INTE2006.DBF"        // Integracao Contabil
  arq14 = LOCALCART+"PAPEIS.DBF"       // Cadastro de Papeis p/Mov
  arq15 = _LOCAL+"BOLSA.DBF"           // Arquivo de Notas de Bolsa
  arq16 = _LOCAL+"BMF.DBF"             // Arquivo de Notas de BMF
  ind0101 = LOCALCART+"EMPRESA1.IND"
  ind0201 = LOCALCART+"OPERADOR.IND"
  ind0301 = _LOCAL+"PAPE2006.IND"
  ind0401 = _LOCAL+"MERC2006.IND"
  ind0501 = _LOCAL+"RDAR2006.IND"
  ind0801 = LOCALCART+"PARCMV.IND"
  ind1001 = _LOCAL+"MVTO2006.IND"
  ind1101 = _LOCAL+"MVTC2006.IND"
  ind1201 = _LOCAL+"NOTA2006.IND"
  ind1301 = _LOCAL+"INTE2006.IND"
  ind1401 = LOCALCART+"PAPEIS.IND"
  ind1501 = _LOCAL+"BOLSA.IND"
  ind1601 = _LOCAL+"BMF.IND"
  chv0101 = "a01code"
  CHV0102 = "VAL(A01CGCCPF)"
  chv0103 = "a01nome"
  CHV0104 = "STRZERO(A01CORP,3)+STRZERO(A01COD1,8)"
  CHV0105 = "STRZERO(A01CORP,3)+STRZERO(A01COD1I,8)"
  CHV0106 = "STRZERO(A01COR2,3)+STRZERO(A01COD2,8)"
  CHV0107 = "STRZERO(A01COR2,3)+STRZERO(A01COD2I,8)"
  CHV0108 = "STRZERO(A01COR3,3)+STRZERO(A01COD3,8)"
  CHV0109 = "STRZERO(A01COR3,3)+STRZERO(A01COD3I,8)"
  CHV0110 = "STRZERO(A01COR4,3)+STRZERO(A01COD4,8)"
  CHV0111 = "STRZERO(A01COR4,3)+STRZERO(A01COD4I,8)"
  CHV0112 = "STRZERO(A01COR5,3)+STRZERO(A01COD5,8)"
  CHV0113 = "STRZERO(A01COR5,3)+STRZERO(A01COD5I,8)"
  chv0201 = "a02codo"
  chv0301 = "a03codp"
  chv0302 = "a03desc"
  chv0303 = "a03codt+a03codp"
  chv0401 = "a04codt"
  chv0501 = "RIGHT(a05mesano,4)+LEFT(A05MESANO,2)"
  CHV0801 = "a08cod"
  chv1001 = "a10codp+DTOS(a10data)+a10oper"
  chv1101 = "a11nota+DTOS(a11data)+a11codt"
  chv1201 = "a12nota"
  CHV1202 = "DTOS(A12DATA)+A12NOTA"
  chv1301 = "DTOS(a13data)+a13dcto"
  chv1401 = "a03codp"
  chv1402 = "a03desc"
  chv1403 = "a03codt+a03codp"
  chv1501 = "STRZERO(NRNOTA,9)+STRZERO(FOLHA,3)+DTOC(PREGAO)"
  chv1601 = "STRZERO(NRNOTA,9)+STRZERO(FOLHA,3)+DTOC(PREGAO)"
Return

Procedure CRIARQC()
  If ECORRETORA
    Return
  Endif
  CLOSE DATABASE
  If !FILE(ARQ05)
    mat_stru={}
    AADD(mat_stru,{"a05mesano","C",6,0})
    //   Operações Comuns
    AADD(mat_stru,{"a05c01","N",19,2}) // Mercado a Vista - Ações
    AADD(mat_stru,{"a05c02","N",19,2}) // Mercado a Vista - Ouro
    AADD(mat_stru,{"a05c03","N",19,2}) // Mercado a Vista - Ouro ativo financeiro fora de Bolsa
    AADD(mat_stru,{"a05c04","N",19,2}) // Mercado Opções - Ações
    AADD(mat_stru,{"a05c05","N",19,2}) // Mercado Opções - Ouro
    AADD(mat_stru,{"a05c06","N",19,2}) // Mercado Opções - Fora de Bolsa
    AADD(mat_stru,{"a05c07","N",19,2}) // Mercado Opções - Outros
    AADD(mat_stru,{"a05c08","N",19,2}) // Mercado Futuro - Dolar dos EUA
    AADD(mat_stru,{"a05c09","N",19,2}) // Mercado Futuro - Indices
    AADD(mat_stru,{"a05c10","N",19,2}) // Mercado Futuro - Juros
    AADD(mat_stru,{"a05c11","N",19,2}) // Mercado Futuro - Outros
    AADD(mat_stru,{"a05c12","N",19,2}) // Mercado a Termo - Ações/Ouro
    AADD(mat_stru,{"a05c13","N",19,2}) // Mercado a Termo - Outros
    // Operações Day-Trade
    AADD(mat_stru,{"a05cD01","N",19,2}) // Mercado a Vista - Ações
    AADD(mat_stru,{"a05cD02","N",19,2}) // Mercado a Vista - Ouro
    AADD(mat_stru,{"a05cD03","N",19,2}) // Mercado a Vista - Ouro ativo financeiro fora de Bolsa
    AADD(mat_stru,{"a05cD04","N",19,2}) // Mercado Opções - Ações
    AADD(mat_stru,{"a05cD05","N",19,2}) // Mercado Opções - Ouro
    AADD(mat_stru,{"a05cD06","N",19,2}) // Mercado Opções - Fora de Bolsa
    AADD(mat_stru,{"a05cD07","N",19,2}) // Mercado Opções - Outros
    AADD(mat_stru,{"a05cD08","N",19,2}) // Mercado Futuro - Dolar dos EUA
    AADD(mat_stru,{"a05cD09","N",19,2}) // Mercado Futuro - Indices
    AADD(mat_stru,{"a05cD10","N",19,2}) // Mercado Futuro - Juros
    AADD(mat_stru,{"a05cD11","N",19,2}) // Mercado Futuro - Outros
    AADD(mat_stru,{"a05cD12","N",19,2}) // Mercado a Termo - Ações/Ouro
    AADD(mat_stru,{"a05cD13","N",19,2}) // Mercado a Termo - Outros
    // CALCULOS
    AADD(MAT_STRU,{"A05PANTV","N",19,2})   // PREJ.A COMPENSAR VISTA (ANTERIOR)
    AADD(MAT_STRU,{"A05PANTO","N",19,2})   // PREJ.A COMPENSAR DT (ANTERIOR)
    AADD(mat_stru,{"a05PREJA","N",19,2})   // Prejuizo a compensar proximo mes
    AADD(mat_stru,{"a05PREJAD","N",19,2})  // Prejuizo a compensar proximo mes - Day trade
    AADD(mat_stru,{"a05aliq","n",5,2})     // Aliquota Operações Comun
    AADD(mat_stru,{"a05aliqd","n",5,2})    // Aliquota Day Trade
    AADD(mat_stru,{"a05IMP","N",19,2})     // Imposto devido op.Comuns
    AADD(mat_stru,{"a05IMPD","N",19,2})    // Imposto devido op.Day Trade
    // CONSOLIDAÇÃO
    AADD(mat_stru,{"A05IMPDEVI","N",19,2})  // Imposto devido total
    AADD(mat_stru,{"A05IRDAYME","N",19,2})   // IR fonte de day-trade do mes
    AADD(mat_stru,{"A05IRDAYAN","N",19,2})   // IR a Compensar Anteriores
    AADD(mat_stru,{"A05IRFONTE","N",19,2})  // IR a Compensar
    AADD(MAT_STRU,{"A05IRDANT","N",19,2})     // IR A VISTA 11.033/2004
    AADD(MAT_STRU,{"A05IMPAPAG","N",19,2})  // IMPOSTO A PAGAR
    AADD(MAT_STRU,{"A05IRPAGO","N",19,2})     // IR PAGO
    DBCREATE(arq05,mat_stru)
    USE &arq05
    INDEX ON &chv0501 TAG CHAVE1 TO &ind0501
  Endif
  ARQ_BOL=_LOCAL+"BOLSA.DBF"
  ARQ_BOL1=_LOCAL+"BOLSA.IND"
  If !FILE(ARQ_BOL1) .And. FILE(ARQ_BOL)
    USE &ARQ_BOL
    INDEX ON &chv1501 TAG CHAVE1 TO &ARQ_BOL1
    USE
  Endif
  If USEREDE(ARQ05,.F.,10,"IRRF")
    If !FILE(IND0501)
      INDEX ON &chv0501 TAG CHAVE1 TO &ind0501
    Else
      SET INDEX TO &ind0501
    Endif
    FOR ZS=1 TO 12
      SysRefresh()
      SEEK STRZERO(VAL(_DATA),4)+STRZERO(ZS,2)
      If EOF()
        ADIREG(0)
        REPLACE A05MESANO WITH STRZERO(ZS,2)+STRZERO(VAL(_DATA),4)
      Endif
    NEXT ZS
    SEEK STRZERO(VAL(_DATA)-1,4)+"12"
    If EOF()
      ADIREG(0)
      REPLACE A05MESANO WITH "12"+STRZERO(VAL(_DATA)-1,4)
    Endif
    CLOSE DATABASE
  Endif
  If !FILE(ARQ04)
    // Cadastro de Tabelas *
    mat_stru = {}
    AADD(mat_stru,{"a04codt","C",02,0})
    AADD(mat_stru,{"a04desc","C",25,0})
    AADD(mat_stru,{"a04cpad","C",06,0})
    AADD(mat_stru,{"a04cpac","C",06,0})
    AADD(mat_stru,{"a04cpah","C",40,0})
    AADD(mat_stru,{"a04vdad","C",06,0})
    AADD(mat_stru,{"a04vdac","C",06,0})
    AADD(mat_stru,{"a04vdah","C",40,0})
    AADD(mat_stru,{"a04tmad","C",06,0})
    AADD(mat_stru,{"a04tmac","C",06,0})
    AADD(mat_stru,{"a04tmah","C",40,0})
    AADD(mat_stru,{"a04tmed","C",06,0})
    AADD(mat_stru,{"a04tmec","C",06,0})
    AADD(mat_stru,{"a04tmeh","C",40,0})
    AADD(mat_stru,{"a04bmad","C",06,0})
    AADD(mat_stru,{"a04bmac","C",06,0})
    AADD(mat_stru,{"a04bmah","C",40,0})
    AADD(mat_stru,{"a04bmed","C",06,0})
    AADD(mat_stru,{"a04bmec","C",06,0})
    AADD(mat_stru,{"a04bmeh","C",40,0})
    AADD(mat_stru,{"a04lucd","C",06,0})
    AADD(mat_stru,{"a04lucc","C",06,0})
    AADD(mat_stru,{"a04luch","C",40,0})
    AADD(mat_stru,{"a04pred","C",06,0})
    AADD(mat_stru,{"a04prec","C",06,0})
    AADD(mat_stru,{"a04preh","C",40,0})
    DBCREATE(arq04,mat_stru)
    USE &arq04
    INDEX ON &chv0401 TAG CHAVE1 TO &ind0401
    DECLARE HIST[29]
    HIST[1]="COMPRA OPCOES"
    HIST[2]="COMPRA VISTA"
    HIST[3]="COMPRA TERMO"
    HIST[4]="VENDA OPCOES"
    HIST[5]="VENDA VISTA"
    HIST[6]="VENDA TERMO"
    HIST[7]="ENTRADA SUBSCRICAO"
    HIST[8]="ENTRADA OUTROS"
    HIST[9]="SAIDA OUTROS"
    HIST[10]="BONIFICAÇÃO"
    HIST[11]="ENTRADA REF.TRANSFERENCIA"
    HIST[12]="SAIDA TRANSFERENCIA"
    HIST[13]="IMPLIT"
    HIST[14]="LUCRO OPCOES NORMAIS"
    HIST[15]="LUCRO VISTA NORMAIS"
    HIST[16]="LUCRO TERMO NORMAIS"
    HIST[17]="LUCRO OPCOES DAY-TRADE"
    HIST[18]="LUCRO VISTA DAY-TRADE"
    HIST[19]="LUCRO TERMO DAY-TRADE"
    HIST[20]="PREJUIZO OPCOES NORMAIS"
    HIST[21]="PREJUIZO VISTA NORMAIS"
    HIST[22]="PREJUIZO TERMO NORMAIS"
    HIST[23]="PREJUIZO OPCOES DAY-TRADE"
    HIST[24]="PREJUIZO VISTA DAY-TRADE"
    HIST[25]="PREJUIZO TERMO DAY-TRADE"
    HIST[26]="EXERC.OPÇÕES COMPRADAS"
    HIST[27]="EXERC.OPÇÕES VENDIDAS"
    HIST[28]="OPÇÕES NÃO EXERCIDAS CPA"
    HIST[29]="OPÇÕES NÃO EXERCIDAS VDA"
    FOR ZS=1 TO 29
      SysRefresh()
      SEEK STRZERO(ZS,2)
      If EOF()
        ADIREG(0)
        REPLACE A04CODT WITH STRZERO(ZS,2)
        REPLACE A04DESC WITH ANSITOOEM(HIST[ZS])
      Endif
    NEXT ZS
    CLOSE DATABASE
  Endif
  If !FILE(ARQ08)
    // Cadastro de Tabelas *
    mat_stru = {}
    AADD(mat_stru,{"a08cod","C",03,0})
    AADD(mat_stru,{"a08desc","C",50,0})
    AADD(MAT_STRU,{"a08qfzr","N",2,0})
    DBCREATE(arq08,mat_stru)
    CLOSE DATABASE
    USE &arq08
    INDEX ON &chv0801 TAG CHAVE1 TO &ind0801
    SET ORDER TO 1
    DECLARE HIST[115]
    HIST[1]="DEBITO POR RETIRADA"
    HIST[2]="CREDITO POR RETIRADA"
    HIST[3]="CREDITO POR CANCELAMENTO DE RETIRADA"
    HIST[4]="DEBITO POR CANCELAMENTO DE RETIRADA"
    HIST[5]="DEBITO POR TRANSF.DE ACOES"
    HIST[6]="CREDITO POR TRANSF.DE ACOES"
    HIST[7]="DEBITO POR CANCELAMENTO DE TRANSF.DE ACOES"
    HIST[8]="CREDITO POR CANCELAMENTO DE TRANSF.DE ACOES"
    HIST[9]="DEBITO POR TRANSF.INTERPRACA"
    HIST[10]="CREDITO POR TRANSF.INTERPRACA"
    HIST[11]="DEBITO POR CRIACAO DE CART.SELECIONADA DE ACOES"
    HIST[12]="DEBITO POR CRIACAO DE CART.SELECIONADA DE ACOES"
    HIST[13]="DEBITO POR CANCELAMENTO DE TRANSF.INTERPRACA"
    HIST[14]="CREDITO POR CANCELAMENTO DE TRANSF.INTERPRACA"
    HIST[15]="DEBITO POR TRANSF.DE DIREITO"
    HIST[16]="CREDITO POR TRANSF.DE DIREITO"
    HIST[17]="DEBITO POR CANCELAMENTO DE TRANSF.DE DIREITO"
    HIST[18]="CREDITO POR CANCELAMENTO DE TRANSF.DE DIREITO"
    HIST[19]="DEBITO POR BLOQUEIO JUDICIAL"
    HIST[20]="CREDITO POR BLOQUEIO JUDICIAL"
    HIST[21]="DEBITO POR TRANSF.COM BLOQUEIO JUDICIAL"
    HIST[22]="CREDITO POR TRANSF.COM BLOQUEIO JUDICIAL"
    HIST[23]="DEBITO DE DIREITO POR BLOQUEIO JUDICIAL"
    HIST[24]="CREDITO DE DIREITO POR BLOQUEIO JUDICIAL"
    HIST[25]="CREDITO POR DEPOSITO"
    HIST[26]="DEBITO POR DEPOSITO"
    HIST[27]="CREDITO POR ESTORNO DE DEPOSITO"
    HIST[28]="DEBITO POR ESTORNO DE DEPOSITO"
    HIST[29]="DIVIDENDO CALCULADO"
    HIST[30]="DIVIDENDO PAGO"
    HIST[31]="CREDITO POR BONIFICACAO PREVISTA"
    HIST[32]="CREDITO POR BONIFICACAO CREDITADA"
    HIST[33]="CREDITO POR DIREITO DE SUBSCRICAO"
    HIST[34]="CREDITO POR RECIBO DE SUBSCRICAO (SEM USO)"
    HIST[35]="CREDITO POR ATUALIZACAO DE SALDOS"
    HIST[36]="DEBITO POR DIVIDENDO EXCLUIDO"
    HIST[37]="DEBITO POR CONCILIACAO DE INTERFACE"
    HIST[38]="CREDITO POR CONCILIACAO DE INTERFACE"
    HIST[39]="RETIRADA POR CONCILIACAO DE INTERFACE"
    HIST[40]="DEBITO POR CANCELAMENTO DE DIVIDENDOS"
    HIST[41]="DEBITO POR CANCELAMENTO DE BONIFICACAO"
    HIST[42]="DEBITO POR CANCELAMENTO DE SUBSCRICAO"
    HIST[43]="CREDITO POR GRUPAMENTO"
    HIST[44]="CREDITO POR CISAO"
    HIST[45]="CREDITO POR DESDOBRO"
    HIST[46]="CREDITO POR INCORPORACAO"
    HIST[47]="CREDITO POR FUSAO"
    HIST[48]="DEBITO POR MANIF.DE EXERCICIO DE SUBSC"
    HIST[49]="CREDITO POR CANCELAMENTO DA MANIF.DO EX.DE SUBSC"
    HIST[50]="DEBITO POR PEDIDO DE EMISSAO DE CESSAO DE DIREITO"
    HIST[51]="CREDITO POR CANC.DE PEDIDO DE CESSAO DE DIREITO"
    HIST[52]="DEBITO POR CANC.DE RECIBO DE SUBSCRICAO"
    HIST[53]="DEBITO POR BAIXA DE FRACAO"
    HIST[54]="CREDITO DE FRACAO"
    HIST[55]="COMPLEMENTO DE FRACAO"
    HIST[56]="DEBITO DE FRACAO (R$)"
    HIST[57]="DEBITO POR LIQ.FISICA DE ACOES" // VENDA
    HIST[58]="CREDITO POR LIQ.FISICA DE ACOES" // COMPRA
    HIST[59]="DEBITO FINANC.POR EX.DE SUBSCRICAO"
    HIST[60]="DEBITO POR RETORNO DE PTA"
    HIST[61]="CREDITO POR ATUALIZACAO DE PTA"
    HIST[62]="CREDITO DE RECIBO DE SUBSCRICAO"
    HIST[63]="DIVIDENDO TRANSF.DA 267"
    HIST[64]="DIVIDENDO TRANSF.P/COMPRADOR"
    HIST[65]="SUBS.TRANSF. DA 267"
    HIST[66]="SUBS.TRANSF.PARA COMPRADOR"
    HIST[67]="BONIFICACAO TRANSF.DA 267"
    HIST[68]="BONIFICACAO TRANSF.P/COMPRADOR"
    HIST[69]="DEBITO POR EXCLUSAO DO SALDO COM DIREITO"
    HIST[70]="CREDITO POR INCLUSAO DO SALDO EX.DIREITO"
    HIST[71]="EXCLUSAO DE CESSAO JA EMITIDA"
    HIST[72]="SUBSCRICAO EXERCIDA (SEM USO)"
    HIST[73]="CREDITO DE RECIBO PREVISTO"
    HIST[74]="DIREITOS NAO EXERCIDOS"
    HIST[75]="DEBITO POR BONIF.PREVISTA"
    HIST[76]="DEBITO POR BONIF.CREDITADA"
    HIST[77]="CREDITO POR EXCLUSAO DE SALDO COM"
    HIST[78]="DEBITO POR INCLUSAO DE SALDO EX"
    HIST[79]="DEBITO POR ATUALIZACAO DE SALDO"
    HIST[80]="DEBITO POR AGRUPAMENTO"
    HIST[81]="DEBITO POR CISAO"
    HIST[82]="DEBITO POR INCORPORACAO"
    HIST[83]="DEBITO POR FUSAO "
    HIST[84]="DEBITO POR DESDOBRO"
    HIST[85]="CREDITO POR DESDOBRO PREVISTO"
    HIST[86]="DEBITO POR DESDOBRO PREVISTO "
    HIST[87]="DEBITO POR DESDOBRO TRANSFERIDO"
    HIST[88]="CREDITO POR DESDOBRO TRANSFERIDO"
    HIST[89]="DESDOBRO CANCELADO"
    HIST[90]="RESTITUICAO DE CAPITAL PREVISTA"
    HIST[91]="RESTITUICAO DE CAPITAL CREDITADO"
    HIST[92]="RESTITUICAO DE CAPITAL EXCLUIDA"
    HIST[93]="RESTITUICAO DE CAPITAL CANCELADA"
    HIST[94]="RESTITUICAO DE CAPITAL TRANSF.(DEBITO)"
    HIST[95]="RESTITUICAO DE CAPITAL TRANSF.(CREDITO)"
    HIST[96]="BONIFICACAO EM DINHEIRO PREVITA"
    HIST[97]="BONIFICACAO EM DINHEIRO CREDITADO"
    HIST[98]="BONIFICACAO EM DINHEIRO EXCLUIDO"
    HIST[99]="BONIFICACAO EM DINHEIRO CANCELADO"
    HIST[100]="BONIFICACAO EM DINHEIRO TRANSF.(DEBITO)"
    HIST[101]="BONIFICACAO EM DINHEIRO TRANSF.(CREDITO)"
    HIST[102]="JUROS DE CAPITAL PROPRIO PREVISTO"
    HIST[103]="JUROS S/CAPITAL PROPRIO CREDITADO"
    HIST[104]="JUROS S/CAPITAL PROPRIO EXCLUIDO"
    HIST[105]="JUROS S/CAPITAL PROPRIO CANCELADO"
    HIST[106]="JUROS S/CAPITAL PROPRIO TRANSF(DEBITO)"
    HIST[107]="JUROS S/CAPITAL PROPRIO TRANSF(CREDITO)"
    HIST[108]="RENDIMENTOS PREVISTOS"
    HIST[109]="RENDIMENTOS CREDITADOS"
    HIST[110]="RENDIMENTOS EXCLUIDOS"
    HIST[111]="RENDIMENTOS CANCELADOS"
    HIST[112]="RENDIMENTOS TRANSF.(DEBITO)"
    HIST[113]="RENDIMENTOS TRANSF.(CREDITO)"
    HIST[114]="DEBITO POR RESGATE DA CARTEIRA SELECIONADA DE ACOES"
    HIST[115]="CREDITO POR RESGATE DA CARTEIRA SELECIONADA DE ACOES"
    FOR ZS=1 TO 115
      SysRefresh()
      SEEK STRZERO(ZS,3)
      If EOF()
        ADIREG(0)
        REPLACE A08COD WITH STRZERO(ZS,3)
        REPLACE A08DESC WITH ANSITOOEM(HIST[ZS])
        REPLACE A08QFZR WITH 1
      Endif
    NEXT ZS
    RELEASE HIST
    CLOSE DATABASE
  Endif
  If !FILE(ARQ03)
    // Cadastro de Papeis
    mat_stru = {}
    AADD(mat_stru,{"a03codp","C",12,0})  // Codigo do papel
    AADD(mat_stru,{"a03desc","C",40,0})  // Descricao
    AADD(mat_stru,{"a03tipo","C",01,0})  // Tipo de papel (1-A Vista 2-Opcao 3-A Termo)
    AADD(mat_stru,{"a03codt","C",02,0})  // Codigo da tabela
    AADD(mat_stru,{"a03lmil","C",01,0})  // Lote de mil (S/N)
    AADD(mat_stru,{"a03qant","N",19,0})  // Quantidade anterior
    AADD(mat_stru,{"a03vant","N",19,2})  // Financeiro anterior
    AADD(mat_stru,{"a03qcpa","N",19,0})  // Quantidade compra
    AADD(mat_stru,{"a03vcpa","N",19,2})  // Financeiro compra
    AADD(mat_stru,{"a03qvda","N",19,0})  // Quantidade venda
    AADD(mat_stru,{"a03vvda","N",19,2})  // Financeiro venda
    AADD(mat_stru,{"a03qtma","N",19,0})  // Quantidade transferencia (+)
    AADD(mat_stru,{"a03vtma","N",19,2})  // Financeiro transferencia (+)
    AADD(mat_stru,{"a03qtme","N",19,0})  // Quantidade transferencia (-)
    AADD(mat_stru,{"a03vtme","N",19,2})  // Financeiro transferencia (-)
    AADD(mat_stru,{"a03qbma","N",19,0})  // Quantidade baixa (+)
    AADD(mat_stru,{"a03vbma","N",19,2})  // Financeiro baixa (+)
    AADD(mat_stru,{"a03qbme","N",19,0})  // Quantidade baixa (-)
    AADD(mat_stru,{"a03vbme","N",19,2})  // Financeiro baixa (-)
    AADD(mat_stru,{"a03vluc","N",19,2})  // Financeiro lucro
    AADD(mat_stru,{"a03vpre","N",19,2})  // Financeiro prejuizo
    AADD(mat_stru,{"a03qatu","N",19,0})  // Quantidade atual
    AADD(mat_stru,{"a03pmed","N",19,8})  // Preco medio
    AADD(mat_stru,{"a03vatu","N",19,2})  // Financeiro atual
    AADD(mat_stru,{"a03pmer","N",19,8})  // Preco mercado
    AADD(mat_stru,{"a03vmer","N",19,2})  // Financeiro mercado
    AADD(mat_stru,{"a03vari","N",19,2})  // Variacao
    AADD(mat_stru,{"a03DTBD","D",8,0})   // Data do Preço Médio
    AADD(MAT_STRU,{"a03pop","C",1,0})    // S = POP
    DBCREATE(arq03,mat_stru)
    USE &arq03
    INDEX ON &chv0301 TAG CHAVE1 TO &ind0301
    INDEX ON &chv0302 TAG CHAVE2 TO &ind0301
    INDEX ON &chv0303 TAG CHAVE3 TO &ind0301
    CLOSE DATABASE
  Endif
  CLOSE DATABASE
  If !FILE(ARQ10)
    // Movimento Diario/Extra
    mat_stru = {}
    AADD(MAT_STRU,{"A10REG","N",12,0})   // Registro para a Internet
    AADD(mat_stru,{"a10codp","C",12,0})  // Codigo do papel
    AADD(mat_stru,{"a10codt","C",02,0})  // Codigo da tabela
    AADD(mat_stru,{"a10nota","C",06,0})  // Numero da nota
    AADD(mat_stru,{"a10data","D",08,0})  // Data do movimento
    AADD(mat_stru,{"a10oper","C",01,0})  // Operacao (C)ompra (V)enda (D)aixa (T)ran
    AADD(mat_stru,{"a10qtde","N",19,0})  // Quantidade
    AADD(mat_stru,{"a10valr","N",19,2})  // Financeiro
    AADD(mat_stru,{"a10puni","N",19,8})  // Preco unitario
    AADD(mat_stru,{"a10dtval","D",8,0})  // Data da Liquidação
    AADD(mat_stru,{"a10flag","C",01,0})  // Flag de controle
    AADD(mat_stru,{"A10resu","N",19,2})  // Resultado
    AADD(MAT_STRU,{"A10CUSTO","N",19,2}) // Custo para Calculo do Resultado
    aadd(mat_stru,{"A10HIST","C",38,0})  // Histórico do movimento
    aadd(mat_stru,{"A10CODH","N",3,0})   // Código do Histórico do movimento
    AADD(MAT_STRU,{"A10TERMO","C",1,0})  // Se for Termo terá aqui a letra T
    DBCREATE(arq10,mat_stru)
    USE &arq10
    INDEX ON &chv1001 TAG CHAVE1 TO &ind1001
    CLOSE DATABASE
  Endif
  If !FILE(ARQ11)
    // Movimento Notas *
    mat_stru = {}
    AADD(mat_stru,{"a11reg","N",12,0})   // Nr. do Registro p/ Internet
    AADD(mat_stru,{"a11nota","C",06,0})  // Numero da nota
    AADD(mat_stru,{"a11codp","C",12,0})  // Codigo do papel
    AADD(mat_stru,{"a11codt","C",02,0})  // Codigo do Mercado
    AADD(mat_stru,{"a11data","D",08,0})  // Data do movimento
    AADD(mat_stru,{"a11oper","C",01,0})  // Operacao (C)ompra (V)enda (B)aixa (T)ran
    AADD(mat_stru,{"a11prazo","N",3,0})  // Prazo
    AADD(mat_stru,{"a11qtde","N",19,0})  // Quantidade
    AADD(mat_stru,{"a11valr","N",19,2})  // Financeiro
    AADD(mat_stru,{"a11puni","N",15,8})  // Preco unitario
    AADD(mat_stru,{"a11vcor","N",19,2})  // Valor da corretagem
    AADD(mat_stru,{"a11vliq","N",19,2})  // Valor liquido
    AADD(mat_stru,{"a11dtval","D",8,0})  // Data da Liquidação
    AADD(mat_stru,{"a11dt","C",01,0})    // se for day Trade
    aadd(mat_stru,{"A11HIST","C",38,0})  // Histórico do movimento
    aadd(mat_stru,{"A11CODH","N",3,0})   // Código do Histórico do movimento
    AADD(mat_stru,{"a11flag","C",01,0})  // Flag de controle
    DBCREATE(arq11,mat_stru)
    USE &arq11
    INDEX ON &chv1101 TAG CHAVE1 TO &ind1101
    CLOSE DATABASE
  Endif
  If !FILE(ARQ12)
    // Dados da Nota *
    mat_stru = {}
    AADD(mat_stru,{"a12nota",  "C",06,0})  // NR da Nota
    AADD(mat_stru,{"a12data",  "D",08,0})  // Data da Nota
    AADD(mat_stru,{"a12txliq", "N",19,2})  // Tx.Liquidação CBLC
    AADD(mat_stru,{"a12txreg", "N",19,2})  // Tx.Registro CBLC
    AADD(mat_stru,{"a12corret","N",19,2})  // Corretagem
    AADD(mat_stru,{"a12txteop","N",19,2})  // Taxa de Termo/Opções - BOVESPA/SOMA
    AADD(mat_stru,{"a12txana", "N",19,2})  // Taxa ANA - BOVESPA/SOMA
    AADD(mat_stru,{"a12emol",  "N",19,2})  // Taxa Emolumentos - BOVESPA/SOMA
    AADD(mat_stru,{"a12IRRF",  "N",19,2})  // IRRF s/operações
    AADD(MAT_STRU,{"a12IRDay", "N",19,2})  // IR s/Day Trade
    AADD(MAT_STRU,{"A12OUTRAS","N",19,2})  // Outras Despesas
    AADD(MAT_STRU,{"a12LIQ","N",19,2})     // Liquido da Nota
    AADD(MAT_STRU,{"A12FLAG","C",1,0})     // Flag
    DBCREATE(arq12,mat_stru)
    USE &arq12
    INDEX ON &chv1201 TAG CHAVE1 TO &ind1201
    INDEX ON &CHV1202 TAG CHAVE2 TO &IND1201
    CLOSE DATABASE
  Endif
  If !FILE(ARQ13)
    // Integracao Contabil *
    mat_stru = {}
    AADD(mat_stru,{"a13data","D",08,0})
    AADD(mat_stru,{"a13dcto","C",06,0})
    AADD(mat_stru,{"a13cdeb","C",30,0})
    AADD(mat_stru,{"a13ccre","C",30,0})
    AADD(mat_stru,{"a13valr","N",19,2})
    AADD(mat_stru,{"a13hist","C",60,0})
    AADD(mat_stru,{"a13flag","C",01,0})
    DBCREATE(arq13,mat_stru)
    USE &arq13
    INDEX ON &chv1301 TAG CHAVE1 TO &ind1301
    CLOSE DATABASE
  Endif
Return(.T.)

FUNCTION INDICESC()
  If ECORRETORA
    Return
  Endif
  CLOSE DATABASE
  If .NOT. FILE(IND0401) .OR. PCOUNT()<>0
    Erase &IND0401
    If .NOT. USEREDE(ARQ04,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ04,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,4 )}              , ;
      IND0401,"Indexação")
  Endif
  CLOSE DATABASE
  If USEREDE(ARQ04,.T.,10,"TEMP")
    If FILE(IND0401)
      SET INDEX TO &IND0401
    Else
      INDEX ON &chv0401 TAG CHAVE1 TO &ind0401
    Endif
    If EOF()
      DECLARE HIST[25]
      HIST[1]="COMPRA OPCOES"
      HIST[2]="COMPRA VISTA"
      HIST[3]="COMPRA TERMO"
      HIST[4]="VENDA OPCOES"
      HIST[5]="VENDA VISTA"
      HIST[6]="VENDA TERMO"
      HIST[7]="ENTRADA SUBSCRICAO"
      HIST[8]="ENTRADA OUTROS"
      HIST[9]="SAIDA OUTROS"
      HIST[10]="BONIFICAÇÃO"
      HIST[11]="ENTRADA REF.TRANSFERENCIA"
      HIST[12]="SAIDA TRANSFERENCIA"
      HIST[13]="IMPLIT"
      HIST[14]="LUCRO OPCOES NORMAIS"
      HIST[15]="LUCRO VISTA NORMAIS"
      HIST[16]="LUCRO TERMO NORMAIS"
      HIST[17]="LUCRO OPCOES DAY-TRADE"
      HIST[18]="LUCRO VISTA DAY-TRADE"
      HIST[19]="LUCRO TERMO DAY-TRADE"
      HIST[20]="PREJUIZO OPCOES NORMAIS"
      HIST[21]="PREJUIZO VISTA NORMAIS"
      HIST[22]="PREJUIZO TERMO NORMAIS"
      HIST[23]="PREJUIZO OPCOES DAY-TRADE"
      HIST[24]="PREJUIZO VISTA DAY-TRADE"
      HIST[25]="PREJUIZO TERMO DAY-TRADE"
      FOR ZS=1 TO 25
        SysRefresh()
        SEEK STRZERO(ZS,2)
        If EOF()
          ADIREG(0)
          REPLACE A04CODT WITH STRZERO(ZS,2)
          REPLACE A04DESC WITH HIST[ZS]
        Endif
      NEXT ZS
    Endif
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND0301) .OR. PCOUNT()<>0
    Erase &IND0301
    If .NOT. USEREDE(ARQ03,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ03,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,7 )}              , ;
      IND0301,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND1101) .OR. PCOUNT()<>0
    Erase &IND1101
    If .NOT. USEREDE(ARQ11,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ11,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,10 )}              , ;
      IND1101,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND1201) .OR. PCOUNT()<>0
    Erase &IND1201
    If .NOT.USEREDE(ARQ12,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ12,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,11 )}              , ;
      IND1201,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND1301) .OR. PCOUNT()<>0
    Erase &IND1301
    If .NOT. USEREDE(ARQ13,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ13,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,12 )}              , ;
      IND1301,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND1001) .OR. PCOUNT()<>0
    Erase &IND1001
    If .NOT. USEREDE(ARQ10,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ10,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,13 )}              , ;
      IND1001,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 1
  If .NOT. FILE(IND0501) .OR. PCOUNT()<>0
    Erase &IND0501
    If .NOT. USEREDE(ARQ05,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ05,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,5 )}              , ;
      IND0501,"Indexação")
  Endif
  CLOSE DATABASE
  SELE 6
  If .NOT. FILE(IND0801) .OR. PCOUNT()<>0
    Erase &IND0801
    If .NOT. USEREDE(ARQ08,.T.,10,"TEMP")
      MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
      MsgStop("Não foi possível acesso ao arquivo "+ARQ08,WSIST)
      QUIT
    Endif
    PACK
    MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
      INDICC( oMeter, oText, oDlg, @lEnd,6 )}              , ;
      IND0801,"Indexação")
  Endif
  CLOSE DATABASE
Return

FUNCTION INDICC( oMeter, oText, oDlg, lEnd,CHAVE )
  oMeter:nTotal := lastrec()
  If CHAVE=1
    INDEX ON A01CODE          TAG CODIGO TO &IND0101
    INDEX ON VAL(A01CGCCPF)   TAG CGCCPF TO &IND0101
    INDEX ON LEFT(A01NOME,15) TAG NOME   TO &IND0101
    INDEX ON &CHV0104 TAG CHAVE4  TO &IND0101
    INDEX ON &CHV0105 TAG CHAVE5  TO &IND0101
    INDEX ON &CHV0106 TAG CHAVE6  TO &IND0101
    INDEX ON &CHV0107 TAG CHAVE7  TO &IND0101
    INDEX ON &CHV0108 TAG CHAVE8  TO &IND0101
    INDEX ON &CHV0109 TAG CHAVE9  TO &IND0101
    INDEX ON &CHV0110 TAG CHAVE10  TO &IND0101
    INDEX ON &CHV0111 TAG CHAVE11 TO &IND0101
    INDEX ON &CHV0112 TAG CHAVE12 TO &IND0101
    INDEX ON &CHV0113 TAG CHAVE13 TO &IND0101 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=2
    INDEX ON &CHV1401 TAG CHAVE1 TO &IND1401
    INDEX ON &CHV1402 TAG CHAVE2 TO &IND1401
    INDEX ON &CHV1403 TAG CHAVE3 TO &IND1401 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=3
    ARQ="USUARIOS.IND"
    INDEX ON CODIGO TAG CODIGO TO &ARQ EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=4
    INDEX ON &chv0401 TAG CHAVE1 TO &IND0401 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=5
    INDEX ON &chv0501 TAG CHAVE1 TO &IND0501 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=6
    INDEX ON &chv0801 TAG CHAVE1 TO &IND0801 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=7
    INDEX ON &chv0301 TAG CHAVE1 TO &IND0301
    INDEX ON &chv0302 TAG CHAVE2 TO &IND0301
    INDEX ON &chv0303 TAG CHAVE3 TO &IND0301 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=10
    INDEX ON &chv1101 TAG CHAVE1 TO &IND1101 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=11
    INDEX ON &chv1201 TAG CHAVE2 TO &IND1201
    INDEX ON &CHV1202 TAG CHAVE3 TO &IND1201 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=12
    INDEX ON &chv1301 TAG CHAVE3 TO &IND1301 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  ElseIf CHAVE=13
    INDEX ON &chv1001 TAG CHAVE4 TO &IND1001 EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
  Endif
Return( NIL )

FUNCTION VERCGCCPF()
  If !CGC(Z01CGCCPF) .And. !CIC(Z01CGCCPF)
    MSGINFO("CNPJ/CPF Inválido !",WSIST)
    Return .F.
  Else
    Return .T.
  Endif

FUNCTION CGC
  PARA C_GC
  M->D_1=0
  M->D_4=0
  M->X_X=1
  FOR M->CON_TA=1 TO LEN(M->C_GC)-2
    SysRefresh()
    If AT(SUBS(M->C_GC,M->CON_TA,1),"/-.")=0
      M->D_1=M->D_1+VAL(SUBS(M->C_GC,M->CON_TA,1))*(IIF(M->X_X<5,6,14)-M->X_X)
      M->D_4=M->D_4+VAL(SUBS(M->C_GC,M->CON_TA,1))*(IIF(M->X_X<6,7,15)-M->X_X)
      M->X_X=M->X_X+1
    Endif
  NEXT
  M->RES_TO=M->D_1-(INT(M->D_1/11)*11)
  M->DIG_ITO=IIF(M->RES_TO<2,0,11-M->RES_TO)
  M->D_4=M->D_4+2*M->DIG_ITO
  M->RES_TO=M->D_4-(INT(M->D_4/11)*11)
  M->DIG_ITO=VAL(STR(M->DIG_ITO,1)+STR(IIF(M->RES_TO<2,0,11-M->RES_TO),1))
  If DIG_ITO<>VAL(SUBS(M->C_GC,LEN(M->C_GC)-1,2))
    Return .F.
  Else
    Return .T.
  Endif

FUNCTION CIC
  PARA C_IC
  M->D_1=0
  M->D_2=0
  M->X_X=1
  FOR M->CON_TA=1 TO LEN(M->C_IC)-2
    SysRefresh()
    If AT(SUBS(M->C_IC,M->CON_TA,1),"/-.")=0
      M->D_1=M->D_1+(11-M->X_X)*VAL(SUBS(M->C_IC,M->CON_TA,1))
      M->D_2=M->D_2+(12-M->X_X)*VAL(SUBS(M->C_IC,M->CON_TA,1))
      M->X_X=M->X_X+1
    Endif
  NEXT
  M->RES_TO=M->D_1-(INT(M->D_1/11)*11)
  M->DIG_ITO=IIF(M->RES_TO<2,0,11-M->RES_TO)
  M->D_2=M->D_2+2*M->DIG_ITO
  M->RES_TO=M->D_2-(INT(M->D_2/11)*11)
  M->DIG_ITO=VAL(STR(M->DIG_ITO,1)+STR(IIF(M->RES_TO<2,0,11-M->RES_TO),1))
  If M->DIG_ITO<>VAL(SUBS(M->C_IC,LEN(M->C_IC)-1,2))
    Return .F.
  Else
    Return .T.
  Endif

FUNCTION EXISTE(Arg1)
  LOCAL Local1
  If (dbSeek(Arg1[1]))
    FOR Local1:= 2 to Len(Arg1)
      SysRefresh()
      Arg1[Local1]:= fieldget(Local1)
    NEXT
  Endif
Return Found()

function BrowseC(aq, cTitle, cListName, bSearch, bNew, bModify, bDelete, bList, bedit )

  local oDlg, oFont
  local btnNew, btnModify, btnDelete, btnSearch, btnList, btnEnd,btnitens
  local n,BORDEM
  local CONTINUA:=.T.
  Public oLbx
  Public aq1

  aq1=aq

  MFONTES()

  DEFAULT cTitle  := "Browse", cListName := "Registros",;
    bNew    := { || oLbx:RecAdd(),RecModify(AQ,oLbx)},;
    bDelete := { || RecDelete( oLbx ),oLbx:Refresh(),oLbx:Setfocus() },;
    bModify := { || RecModify(AQ, oLbx),oLbx:Setfocus() },;
    bList   := { || CReport(aq, oLbx,clistname,oLbx:Setfocus() )},BEDIT:=.F.

  If AQ#8
    DEFINE DIALOG oDlg FROM 1, 5 TO 27, 80 TITLE cTitle FONT oFontTit BRUSH oBrushTela STYLE WS_CAPTION
  ElseIf AQ=8
    DEFINE DIALOG oDlg FROM 1, 5 TO 27, 80 TITLE "Itens Substrição/Bonificação/Outros" FONT oFontTit  BRUSH oBrushTela STYLE WS_CAPTION
  Endif
  GO BOTT
  FOR I=1 TO 15
    SysRefresh()
    If !BOF()
      SKIP -1
    Endif
  NEXT I
  If AQ=1
    If QZ<>3
      @ 1, 1 LISTBOX oLbx FIELDS EMP->A01CODE,STRZERO(EMP->A01CORRET,1),OEMTOANSI(EMP->A01NOME),EMP->A01DATA,EMP->A01CGCCPF HEADERS "Código","Status","Nome","Ref.","CNPJ/CPF " SIZE 275, 137  OF oDlg
    Else
      @ 1, 1 LISTBOX oLbx FIELDS CARTEMP->A01CODE,STRZERO(CARTEMP->A01CORRET,1),OEMTOANSI(CARTEMP->A01NOME),CARTEMP->A01DATA,CARTEMP->A01CGCCPF HEADERS "Código","Status","Nome","Ref.","CNPJ/CPF" SIZE 275, 137 OF oDlg
    Endif
  ElseIf AQ=2
    @ 1,1 LISTBOX oLbx FIELDS OPER->CODIGO,OPER->NOME  HEADERS "Código","Nome" SIZE 275, 137  OF oDlg
  ElseIf AQ=4
    @ 1,1 LISTBOX oLbx FIELDs PAPEL->A03CODP,PAPEL->A03DESC,PAPEL->a03lmil,IIF(PAPEL->A03TIPO="1","Vista",IIF(PAPEL->A03TIPO="2","Opcao","Termo")),DTOC(PAPEL->a03DTBD)  HEADERS "Código","Descrição","Em Mil","Mercado","BDI de" SIZE 275,137 OF oDlg
  ElseIf AQ=5
    @ 1,1 LISTBOX oLbx FIELDs SALD->A03CODP,SALD->A03DESC,SALD->a03lmil,IIF(SALD->A03TIPO="1","Vista",IIF(SALD->A03TIPO="2","Opcao","Termo"))  HEADERS "Código","Descrição","Em Mil","Mercado" SIZE 275,137 OF oDlg
  ElseIf AQ=6
    FZ=.T.
    @ 1,1 LISTBOX oLbx FIELDs NOTA->A12NOTA,TRANSFORM(NOTA->A12DATA,"@D"),TRANSFORM(NOTA->A12LIQ,"@E 999,999,999.99") HEADERS "Nota","Data","Liquido" SIZE 275,137 OF oDlg
  ElseIf AQ=7
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(LOG->DATA,"@D"),LOG->HORA,LOG->USUARIO,LOG->ROTN HEADERS "Data","Hora","Usuário","Rotina Utilizada" SIZE 275,137 OF oDlg
  ElseIf AQ=8
    @ 1,1 LISTBOX oLbx FIELDS IIF(ITENS->A11DT="I","S","N"),TRANSFORM(ITENS->A11DATA,"@D"),IIF((ITENS->A11OPER="C" .OR. ITENS->A11OPER="1"),"Entrada","Saida"),ITENS->A11CODP,PESQ(ITENS->A11CODP,"2",1,"A03DESC"),TRANSFORM(ITENS->A11QTDE,"@E 999,999,999,999"),TRANSFORM(ITENS->A11VALR,"@E 999,999,999,999.99") HEADERS "Auto","Data","C/V","Papel","Descricao","Quantidade","Financeiro" SIZE 275,137 OF oDlg
  ElseIf AQ=9
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(NOTAS->PREGAO,"@D"),STRZERO(NOTAS->NRNOTA,9),TRANSFORM(NOTAS->VALOR,"@E 999,999,999.99"),NOTAS->DC HEADERS "Pregao","Nr.Nota","Valor","D/C" SIZE 275,137 OF odlg
  ElseIf AQ=10
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(NOTAS->PREGAO,"@D"),STRZERO(NOTAS->NRNOTA,9),TRANSFORM(NOTAS->VALOR,"@E 999,999,999.99"),NOTAS->DC HEADERS "Pregao","Nr.Nota","Valor","D/C" SIZE 275,137 OF odlg
  ElseIf AQ=11
    @ 1,1 LISTBOX oLbx FIELDS EXTRATO->MESANO,EXTRATO->TIPOEX,EXTRATO->CORRETORA HEADERS "Ref.","Tipo Contas","Corretora" SIZE 275,137 OF odlg
  ElseIf AQ=12
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(IRRF->A05mesano,"@R 99-9999") HEADERS "Mes/Ano" SIZE 275,137 OF oDlg
  ElseIf AQ=13
    @ 1,1 LISTBOX oLbx FIELDS OEMTOANSI(CONTA->A04DESC),CONTA->a04cpah,CONTA->a04vdah HEADERS "Histórico","Debita","Credita" SIZE 275,137 OF odlg
  ElseIf AQ=14
    @ 1,1 LISTBOX oLbx FIELDS CMVD->A08COD,CMVD->A08DESC HEADERS "Código","Descrição" SIZE 275,137 OF odlg
  Else
    @ 1, 1 LISTBOX oLbx FIELDS SIZE 275, 137  OF oDlg
  Endif

  oLbx:bLDblClick = { | nRow, nCol | Eval( bModify ) }

  If AQ<>9 .And. aq<>10 .And. aq<>11 .And. AQ<>13

    @ 155, 10 SBUTTON pixels btnNew    RESOURCE "INCLUI" PROMPT "&Inclui" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Inclui Novos Registros (dados)"
    @ 155, 50 SBUTTON PIXELS btnModify RESOURCE "ALTERA" PROMPT "&Altera" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Altera os Registros (dados)"
    @ 155, 90 SBUTTON PIXELS btnDelete RESOURCE "EXCLUI" PROMPT "&Exclui" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Exclui os Registros (dados)"
    If AQ#5 .And. aq#12
      @ 155,130 SBUTTON PIXELS btnList  RESOURCE "IMPRIME" PROMPT "I&mprime" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Imprime relatório"
    Endif
    If AQ<>11
      @ 155,170 SBUTTON PIXELS btnSearch RESOURCE "PROC" PROMPT "&Busca" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF odlg TOOLTIP "Faz pesquisa dos Dados"
    Endif
    If AQ=6
      @ 155,210 SBUTTON PIXELS btnItens RESOURCE "INCLUI" PROMPT "&Operações" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Operações da Nota" ACTION (CADITENS(),oLbx:refresh(),oLbx:SetFocus())
      @ 155,250 SBUTTON PIXELS btnEnd   RESOURCE "SAIR2" PROMPT "&Sair"  FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Retorna ao Menu Principal"
    Endif
    If AQ=1
      @ 155,210 SBUTTON PIXELS btnEnd RESOURCE "SAIR2" PROMPT "&Ok" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Dados do Investidor Acima escolhido"
    ElseIf AQ#6 .And. AQ#13
      @ 155,210 SBUTTON PIXELS btnEnd RESOURCE "SAIR2" PROMPT "&Sair" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Retorna ao Menu Principal"
    ElseIf AQ=13
      @ 155,90 SBUTTON PIXELS btnEnd RESOURCE "SAIR2" PROMPT "&Sair" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Retorna ao Menu Principal"
    Endif
    If AQ=8
      btnSearch:Disable()
    Endif
    If AQ#13
      btnNew:bAction    = { || Eval( bNew ),oLbx:Refresh(), oLbx:SetFocus() }
    Endif
    btnDelete:bAction = If( bDelete != nil,{ || Eval( bDelete ),oLbx:Refresh(), oLbx:SetFocus() },)
    btnModify:bAction = If( bModify != nil,{ || Eval( bModify ),oLbx:Refresh(), oLbx:SetFocus() },)
    If AQ<>11 .And. AQ<>13
      btnSearch:bAction = If( bSearch != nil,{ || Eval( bSearch ),oLbx:Refresh(), oLbx:SetFocus() },)
    Endif
    If AQ#5 .And. aq#12
      btnList:bAction   = { || Eval( bList ), oLbx:Refresh() }
    Endif
  ElseIf aq=9 .or. aq=10 .OR. AQ=11
    If AQ=9
      @ 155,10 SBUTTON PIXELS          RESOURCE "IMPRIME" PROMPT "&Visualiza" ACTION (canexo:=_LOCAL+"NOTASBOV\"+TRIM(NOTAS->NOMEPDF)+".PDF",RUNFILE(CANEXO)) FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Imprime relatório"
    ElseIf AQ=10
      @ 155,10 SBUTTON PIXELS          RESOURCE "IMPRIME" PROMPT "&Visualiza" ACTION (canexo:=_LOCAL+"NOTASBMF\"+TRIM(NOTAS->NOMEPDF)+".PDF",RUNFILE(CANEXO)) FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Imprime relatório"
    ElseIf AQ=11
      @ 155,10 SBUTTON PIXELS          RESOURCE "IMPRIME" PROMPT "&Visualiza" ACTION (canexo:=_LOCAL+"EXTRATOS\E"+STRZERO(EXTRATO->NOMEPDF,7)+".PDF",RUNFILE(CANEXO)) FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Imprime relatório"
    Endif
    @ 155,90 SBUTTON PIXELS btnEnd   RESOURCE "SAIR2"   PROMPT "&Sair"  FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Retorna ao Menu Principal"
  Endif
  If AQ=13
    @ 155, 10 SBUTTON PIXELS btnModify RESOURCE "ALTERA" PROMPT "&Altera" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Altera os Registros (dados)"
    @ 155, 50 SBUTTON PIXELS btnList   RESOURCE "IMPRIME" PROMPT "I&mprime" FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Imprime relatório"
    @ 155,90 SBUTTON PIXELS btnEnd     RESOURCE "SAIR2"   PROMPT "&Sair"  FONT oFontbut TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 OF oDlg TOOLTIP "Retorna ao Menu Principal"
    btnModify:bAction = If( bModify != nil,{ || Eval( bModify ),oLbx:Refresh(), oLbx:SetFocus() },)
    btnList:bAction   = { || Eval( bList ), oLbx:Refresh() }
  Endif

  btnEnd:bAction    = { || oDlg:End()}
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
Return nil
//----------------------------------------------------------------------------//

function CReport( AQ,oLbx,ctitle)
  local n
  LOCAL oFont1, oFont2
  PUBLIC oRpt

  DEFINE FONT oFont1 NAME "ARIAL" SIZE 0,-10
  DEFINE FONT oFont2 NAME "ARIAL" SIZE 0,-7
  If aq=12
    Return
  Endif
  If AQ=6
    IMPNOTA1()
    Return
  Endif
  reg=recno()
  If EOF()
    MSGINFO("Não Existe Dados para Impressão",WSIST)
    Return
  Endif
  GO TOP
  If PREV

  REPORT oRpt TITLE " "," "," "," ",EMP1,CTitle FONT oFont1,oFont2;
    PREVIEW;
    FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
    CAPTION ctitle;

  Else

  REPORT oRpt TITLE " "," "," "," ",EMP1,CTitle FONT oFont1,oFont2;
    FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
    CAPTION ctitle;

  Endif

  If AQ=1
    COLUMN TITLE "Cod." DATA EMP->A01CODE FONT 2 GRID
    COLUMN TITLE "Nome" DATA EMP->A01NOME FONT 2 GRID
  ElseIf AQ=2
    COLUMN TITLE "Código" DATA OPER->CODIGO FONT 2 GRID
    COLUMN TITLE "Nome" DATA OPER->NOME FONT 2 GRID
  ElseIf AQ=4
    DECLARE TPS[3]
    TPS[1]="Vista"
    TPS[2]="Opção"
    TPS[3]="Termo"
    COLUMN TITLE "Código" DATA PAPEL->A03CODP FONT 2 GRID
    COLUMN TITLE "Descrição" DATA LEFT(PAPEL->A03DESC,12) FONT 2 GRID
    COLUMN TITLE "Tipo" DATA SUBSTR(PAPEL->A03DESC,13,12) FONT 2 GRID
    COLUMN TITLE "Ult.BDI" DATA DTOC(PAPEL->A03DTBD) FONT 2 GRID
    COLUMN TITLE "MERCADO" DATA IIF(VAL(PAPEL->A03TIPO)#0,TPS[VAL(PAPEL->A03TIPO)],"") FONT 2 GRID
  ElseIf AQ=8
    COLUMN TITLE "Data" DATA ITENS->A11DATA GRID
    COLUMN TITLE "Movto" DATA IIF((ITENS->A11OPER="C" .OR. ITENS->A11OPER="1"),"Entrada","Saida") GRID
    column TITLE "Papel" DATA ITENS->A11CODP+"-"+PESQ(ITENS->A11CODP,"2",1,"A03DESC") GRID
    COLUMN TITLE "Quantidade" DATA ITENS->A11QTDE PICTURE "@E 999,999,999,999" GRID
    COLUMN TITLE "Financeiro" DATA ITENS->A11VALR PICTURE "@E 999,999,999,999.99" RIGHT TOTAL GRID
  ElseIf AQ=9
    COLUMN TITLE "Nota nr" DATA NOTA->A12NOTA FONT 2 GRID
    COLUMN TITLE "Data " DATA DTOC(NOTA->A12DATA) FONT 2 GRID
    COLUMN TITLE "Liquido" DATA NOTA->A12LIQ PICTURE "@E 999,999,999.99" TOTAL FONT 2 GRID
  ElseIf AQ=13
    COLUMN TITLE "Ordem" DATA CONTA->a04codt FONT 2 GRID
    COLUMN TITLE "Descrição" DATA OEMTOANSI(CONTA->A04DESC) FONT 2 GRID
    COLUMN TITLE "Debita" DATA LEFT(CONTA->A04CPAH,20) FONT 2 GRID
    COLUMN TITLE "Credita" DATA LEFT(CONTA->A04VDAH,20) FONT 2 GRID
  ElseIf AQ=14
    COLUMN TITLE "Código" DATA CMVD->A08COD FONT 2 GRID
    COLUMN TITLE "Descrição" DATA CMVD->A08DESC FONT 2 GRID
    DECLARE HIST[7]
    HIST[1]="(1) - Descartar"
    HIST[2]="(2) - Outras"
    HIST[3]="(3) - Subscrição"
    HIST[4]="(4) - Bonificação"
    HIST[5]="(5) - Transferencias"
    HIST[6]="(6) - Implit"
    HIST[7]="(7) - Dividendos"
    COLUMN TITLE "Movimento" DATA HIST[CMVD->a08qfzr] FONT 2 GRID
  Endif
  oRpt:CELLVIEW()
  ENDREPORT

  ACTIVATE REPORT oRpT ON STARTPAGE IMPRIMECAB()

  GO reg
  oFont1:End()
  oFont2:End()
Return nil
//----------------------------------------------------------------------------//
FUNCTION IMPRIMECAB()
  oRpt:saybitmap(0.2,0.2,"logo.fig",1.625,0.653)
  Return
  oRpt:SayBitmap(0.5,0.5,"LOGO.FIG",1.94,0.4)
Return

static function FInfo1( cAlias, n )
Return { || ( cAlias )->( FieldName( n ) ) }

static function FInfo2( cAlias, n )
Return { || ( cAlias )->( FieldGet( n ) ) }
//----------------------------------------------------------------------------//
static function RecModify(AQ, oLbx )
  local n := 1
  local nCols := ( oLbx:cAlias )->( FCount() )
  local u := ( oLbx:cAlias )->( FieldGet( 1 ) )
  while n <= nCols .And. oLbx:lEditCol( n, @u )
    SysRefresh()
    oLbx:DrawSelect()
    If ( oLbx:cAlias )->( RLock() )
      ( oLbx:cAlias )->( FieldPut( n, u ) )
      UNLOCK
    Else
      MsgAlert( "Registro Bloqueado !","Verifique" )
      n = nCols + 1
    Endif
    n++
    If n <= nCols
      u = ( oLbx:cAlias )->( FieldGet( n ) )
    Endif
  end
  SysRefresh()
Return nil
//----------------------------------------------------------------------------//
static function EditCell( oLbx, nRow, nCol )
  local nColumn := oLbx:nAtCol( nCol )
  local u
  u = ( oLbx:cAlias )->( FieldGet( nColumn ) )
  If oLbx:lEditCol( nColumn, @u )
    If ( oLbx:cAlias )->( RLock() )
      ( oLbx:cAlias )->( FieldPut( nColumn, u ) )
      UNLOCK
      oLbx:DrawSelect()
    Else
      MsgAlert( "Registro Bloqueado por outro Usuário! ","Verifique")
    Endif
  Endif
Return nil
//----------------------------------------------------------------------------//
static function KeyDown( oLbx, nKey, nFlags )
  do case
  endcase
Return nil
//----------------------------------------------------------------------------//
static function KeyChar( oLbx, nKey, nFlags )
  do case
    case nKey == K_ENTER
      RecModify( oLbx )
  endcase
Return nil
//----------------------------------------------------------------------------//

static function Rec1Delete( oLbx )
  If MsgYesNo( "Exclui este Registro ?", WSIST)
    If REGLOCK(10)
      DELETE
    Else
      MsgInfo("Registro Bloqueado por outro Usuário",WSIST)
    Endif
    oLbx:Refresh()
  Endif
Return nil

FUNCTION VERMC(X,QUAL)
  MFONTES()
  If X=17
    RETORNO=.F.
  Endif
  If X=21
    Z11HIST=PESQ(STRZERO(Z11CODH,3),"5",1,"A08DESC")
    ohist:refresh()
    If EMPTY(Z11HIST)
      Return .F.
    Endif
    Return .T.
  Endif
  If X=34
    RETORNO=.F.
  Endif
  If X=18
    SELE 1
    If Z03lmil="S"
      Z03VMER = Z03qatu * (Z03pmer/1000)
    Else
      Z03vmer = Z03qatu * Z03pmer
    Endif
    Z03vari=Z03vmer - Z03vatu
    Return
  Endif
  If X=19
    CONTINUA=.F.
    Return
  Endif
  If X=20
    Z11DTVAL=Z12DATA+Z11PRAZO
    oz11dtval:refresh()
    Return .T.
  Endif
  If X=1
    Z04CODT=STRZERO(VAL(Z04CODT),2)
  ElseIf X=2
    If Z03LMIL
      Z03LMIL="S"
    Else
      Z03LMIL="N"
    Endif
    Z03TIPO=STRZERO(ZCLAS,1)
  ElseIf X=3
    If ADIREG(0)
      Return .T.
    Else
      MSGINFO("Não foi possível Adicionar o Novo Registro ",WSIST)
      Return .T.
    Endif
  ElseIf X=4
    SET ORDER TO 1
    Z01CODE=STRZERO(VAL(Z01CODE),6)
    SEEK Z01CODE
    If !EOF()
      MSGINFO("Investidor Já Existente !",WSIST)
      Return .F.
    Endif
  ElseIf X=7
    SELE 2
    SET ORDER TO 1
    SEEK Z03CODP
    If EOF() .And. QUAL=1
      WVERPAP()
      Z03CODP=A03CODP
      SELE 1
      Return .F.
    Endif
    If QUAL=1
      NPAP=A03DESC
      @   3.57,18 SAY A03DESC FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
    Endif
    If QUAL=2
      @   1.74,9 SAY A03DESC OF oDlg FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11
    Endif
    Z03CODT=A03CODT
    Z03DESC=A03DESC
    Z03TIPO=A03TIPO
    Z03LMIL=A03LMIL
    Z03TIPO=A03TIPO
    SELE 3
    SET ORDER TO 1
    SEEK Z03CODT
    DECLARE TPS[5]
    TPS[1]="Vista"
    TPS[2]="Opcão"
    TPS[3]="Termo"
    TPS[4]="POP"
    TPS[5]="N.Consta"
    If VAL(Z03TIPO)<1 .OR. VAL(Z03TIPO)>3
      Z03TIPO="4"
    Endif
    If !EOF()
      If QUAL=1
        @   3.57,18 SAY NPAP FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
        @   5.31,18 SAY IIF(VAL(Z03TIPO)>0 .And. VAL(Z03TIPO)<4,TPS[VAL(Z03TIPO)],"") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
      ElseIf QUAL=2
        @   2.61,9 SAY IIF(VAL(Z03TIPO)>0 .And. VAL(Z03TIPO)<4,TPS[VAL(Z03TIPO)],"") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
      Endif
    Else
      If QUAL=1
        @   5.31,18 SAY IIF(VAL(Z03TIPO)>0 .And. VAL(Z03TIPO)<4,TPS[VAL(Z03TIPO)],"") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
      ElseIf QUAL=2
        @   2.61,9 SAY IIF(VAL(Z03TIPO)>0 .And. VAL(Z03TIPO)<4,TPS[VAL(Z03TIPO)],"") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11 OF oDlg
      Endif
    Endif
    SELE 1
    Return .T.
  ElseIf X=8
    SELE 2
    SET ORDER TO 1
    SEEK Z03CODP
    DECLARE TPS[5]
    TPS[1]="Vista"
    TPS[2]="Opcão"
    TPS[3]="Termo"
    TPS[4]="POP"
    TPS[5]="N.Consta"
    Z03TIPO=A03TIPO
    If VAL(Z03TIPO)<1 .OR. VAL(Z03TIPO)>3
      Z03TIPO="5"
    Endif
    If !EOF()
      @   1.74,9 SAY A03DESC FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,11
      @   2.61,9 SAY IIF(VAL(Z03TIPO)>0 .And. VAL(Z03TIPO)<4,TPS[VAL(Z03TIPO)],"") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 80,11
    Endif
    SELE 1
    Return .T.
  ElseIf X=9
    Z12nota=STRZERO(VAL(Z12NOTA),6)
  ElseIf X=10
    AA=YEAR(QUAL)
    If AA#VAL(_DATA)
      MSGINFO("Ano Invalido ! ",WSIST)
      Return .F.
    Endif
  ElseIf X=11
    SELE 2
    SET ORDER TO 1
    SEEK Z11CODP
    If EOF()
      WVERPAP()
      Z11CODP=A03CODP
      SELE 3
      Return .F.
    Endif
    Z11DESCR=A03DESC
    Z11CODT=A03CODT
    If A03TIPO="2"
      Z11OPCAO="S"
    Else
      Z11OPCAO="N"
    Endif
    @   5.31,14 SAY Z11DESCR FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) OF oDlgit SIZE 120,11
    ZMERC=A03TIPO
    SELE 3
  ElseIf X=12
    SELE 2
    SET ORDER TO 1
    SEEK Z11CODP
    If EOF()
      WVERPAP()
      Z11CODP=A03CODP
      SELE 3
      Return .F.
    Endif
    Z11DESCR=A03DESC
    Z11CODT=A03CODT
    @   5.31,14 SAY Z11DESCR FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) OF oDlgit SIZE 120,11
    ZMERC=A03TIPO
    If A03TIPO="2"
      Z11DTVAL=A03DTOP
      If Z11DTVAL<Z12DATA
        MSGINFO("Opção Vencida !",WSIST)
        SELE 3
        Return .F.
      Endif
    Endif
    SELE 3
  ElseIf X=13
    CONTINUA=.F.
  ElseIf X=14
    Z11NOTA=Z12NOTA
    Z11DATA=Z12DATA
    If PAPEL->A03TIPO="2"
      Z11DTVAL=PAPEL->A03DTOP
    Endif
    If ZCLASIT=1
      Z11OPER="C"
    Else
      Z11OPER="V"
    Endif
    Z11ORDEM=RECNO()
    If PESQ(Z11CODP,"2",1,"A03TIPO")="1"
      Z11DTVAL=Z12DATA
    Endif
  ElseIf X=15
    If Z11VENC>=Z12DATA
      Return .T.
    Else
      MSGINFO("A data do Termo deve ser maior que a data da operacao",WSIST)
      Return .F.
    Endif
  ElseIf X=16
    CLOSE DATABASE
    Return
  Endif
Return .T.

FUNCTION CREAS(ARQ1,ARQ2)
  CREATE &ARQ1 FROM &ARQ2
Return

FUNCTION BUSCAEMP()
  TEXTO="Investidor:"
  VARIAVEL=SPACE(30)
  MASCARA="@!"
  REG=RECNO()
  Do While .T.
    SysRefresh()
    If !PROCURA()
      EXIT
    Endif
    CHAVE=VARIAVEL
    If EMPTY(CHAVE)
      EXIT
    Endif
    If VAL(CHAVE)>0
      SET ORDER TO 1
      SEEK STRZERO(VAL(chave),6)
    Else
      SET ORDER TO 3
      SEEK TRIM(CHAVE)
    Endif
    If EOF() .OR. DELE()
      MSGALERT("Não Encontrado !",WSIST)
      GO REG
      WCODI=SPACE(30)
    Else
      EXIT
    Endif
  ENDDO
Return .T.

Procedure NCADEMP(F)
  LOCAL oDlg,oBrush,oCbx
  PUBLIC ZCLAS
  If F="I"
    GO BOTT
    SKIP
  Endif
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  If F="I"
    GO BOTT
    Z01CODE=STRZERO(VAL(A01CODE)+1,6)
    Z01TIPO=1
  Endif
  MFONTES()
  DEFINE DIALOG oDlg FROM 1,1 TO 32,67 TITLE "Dados do Investidor" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.07,0.5 say " " BOX RAISED COLOR CLR_BLACK,nRGB(248,248,248) SIZE 256,210
  @   0.87,01 SAY "Cod Exatus:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   0.87,25 SAY "Status:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,01 SAY "Nome:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,01 SAY "Tipo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,01 SAY "CNPJ/CPF:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,23 SAY "Ano Ref:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   4.35,01 SAY "Senha:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   4.35,23 SAY "Calc.até:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,01 SAY "Frase:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,01 SAY "Validade:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,23 SAY "Dia Vencto:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.96,01 SAY "Vencto:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.96,23 SAY "Pagto:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.83,01 SAY "Valor:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   8.70,01 SAY "Email:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   9.57,01 SAY "É Corretora? :" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  10.44,01 SAY "Corretora Adm:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  10.44,14 SAY "Cod.na Cor:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  10.44,29 SAY "Cta.Invest:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  11.31,01 SAY "Corretora 02 :" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  11.31,14 SAY "Cod.na Cor:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  11.31,29 SAY "Cta.Invest:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  12.18,01 SAY "Corretora 03 :" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  12.18,14 SAY "Cod. na Cor:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  12.18,29 SAY "Cta.Invest:" FONT oFontsay  COLOR CLR_BLACK,nRGB(248,248,248)
  @  13.05,01 SAY "Corretora 04 :" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  13.05,14 SAY "Cod. na Cor:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  13.05,29 SAY "Cta.Invest:" FONT oFontsay  COLOR CLR_BLACK,nRGB(248,248,248)
  If F="I"
    @ 01,07 GET Z01CODE PICTURE "999999" FONT oFontget SIZE 27,11 VALID VERMC(4)
  Else
    @ 0.87,9.5 SAY Z01CODE FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  Endif
  @ 0.4,12 CHECKBOX Z01INATI PROMPT "Inativo" COLOR CLR_BLACK,nRGB(248,248,248) SIZE 60,6 FONT oFontget OF oDlg
  @ 01,22 get Z01CORRET PICTURE "9" FONT oFontget SIZE 10,11
  @ 02,07 GET Z01NOME   PICTURE "@!" FONT oFontGET SIZE 200,11
  @ 02.8,07 COMBOBOX oCbx1 VAR Z01TIPO ITEMS {"Fisica","Juridica"} OF oDlg UPDATE font OFONTGET SIZE 43,50
  @ 0.4,18 CHECKBOX Z01CONFER PROMPT "Autorizado pela Corretora" COLOR CLR_BLACK,nRGB(248,248,248) SIZE 120,6 FONT oFontget OF odlg
  @ 04,07 GET Z01CGCCPF PICTURE "@!" FONT oFontget SIZE 60,11 VALID VERCGCCPF()
  @ 04,23 GET o01data VAR Z01DATA   PICTURE "99999" FONT oFontget SIZE 30,11 NO MODIFY
  @ 05,07 GET Z01SENHA   FONT oFontget SIZE 80,11
  @ 05,23 GET Z01CALCATE FONT oFontget SIZE 45,11 VALID VERM(6)
  @ 06,07 get z01frase   FONT oFontget SIZE 200,11
  @ 07,07 GET Z01VALID   FONT oFontget SIZE 45,11
  @ 07,23 GET Z01DIAVENC FONT oFontget SIZE 20,11 VALID Z01DIAVENC>0 .And. Z01DIAVENC<32
  @ 08,07 GET Z01EMITIDO FONT oFontget SIZE 45,11
  @ 08,23 GET Z01PAGO    FONT oFontget SIZE 45,11
  @ 09,07 GET Z01VALMES  FONT oFontget SIZE 60,11 RIGHT PICTURE "@E 999,999,999.99"
  @ 10,07 GET Z01EMAIL   FONT oFontget SIZE 200,11
  @ 11,07 GET Z01CODNAC  FONT oFontget SIZE 8,11 PICTURE "!"
  @ 12,07 get Z01CORP PICTURE "999" FONT oFontget SIZE 20,11 RIGHT
  @ 12,16 GET Z01COD1 PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 12,26 GET Z01COD1I PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 13,07 get Z01COR2 PICTURE "999" FONT oFontget SIZE 20,11 RIGHT
  @ 13,16 GET Z01COD2 PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 13,26 GET Z01COD2I PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 14,07 get Z01COR3 PICTURE "999" FONT oFontget SIZE 20,11 RIGHT
  @ 14,16 GET Z01COD3 PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 14,26 GET Z01COD3I PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 15,07 get Z01COR4 PICTURE "999" FONT oFontget SIZE 20,11 RIGHT
  @ 15,16 GET Z01COD4 PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 15,26 GET Z01COD4I PICTURE "9999999" FONT oFontget SIZE 40,11 RIGHT
  @ 218,80 SBUTTON PIXEL RESOURCE "OFF","ON","DISABLE" NOBOX ADJUST PROMPT "&Confirma" TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (IIF(F="I",VERMC(3),""),PUTREC(),oDlg:END())
  @ 218,150 SBUTTON PIXEL RESOURCE "OFF","ON","DISABLE" NOBOX ADJUST PROMPT "C&ancela"  TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (oDlg:End())
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
Return

Procedure CADPAP()
  CLOSE DATABASE
  SELE 1
  If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
    MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1401
  Endif
  SELE 2
  If .NOT. USEREDE(ARQ04,.F.,10,"MERC")
    MsgStop("O arquivo MERCADOS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0401
  Endif
  SELE 1
  BrowseC(4,"Papeis","Tabela de Papeis", { || buscaPAP() },{ ||NCADPAP("I") },{ ||NCADPAP("A") })
  USE
  CLOSE DATABASE
Return

FUNCTION BUSCAPAP()
  TEXTO="Papel:"
  VARIAVEL=SPACE(20)
  MASCARA="@!"
  REG=RECNO()
  Do While .T.
    SysRefresh()
    If !PROCURA()
      EXIT
    Endif
    CHAVE=VARIAVEL
    If EMPTY(CHAVE)
      EXIT
    Endif
    If AT("1",CHAVE)#0 .OR. AT("2",CHAVE)#0 .OR. AT("3",CHAVE)#0 .OR. AT("4",CHAVE)#0 .OR. AT("5",CHAVE)#0 .OR. AT("6",CHAVE)#0 .OR. AT("7",CHAVE)#0 .OR. AT("8",CHAVE)#0 .OR. AT("9",CHAVE)#0 .OR. AT("0",CHAVE)#0
      SET ORDER TO 1
      SEEK TRIM(CHAVE)
    Else
      SET ORDER TO 2
      SEEK TRIM(CHAVE)
    Endif
    If EOF() .OR. DELE()
      MSGALERT("Não Encontrado !",WSIST)
      GO REG
    Else
      EXIT
    Endif
  ENDDO
Return

Procedure NCADPAP(F)
  LOCAL oDlg,oBrush,oCbx
  PUBLIC ZCLAS
  If F="I"
    DECLARE _var1 [FCOUNT()],_var2 [FCOUNT()],_var3 [FCOUNT()]
    AFIELDS(_var1,_var2,_var3)
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      If _var2 [t] = "C"
        &_var4 = SPACE(_var3 [t])
      ElseIf _var2 [t] = "N"
        &_var4 = 0
      ElseIf _var2 [t] = "D"
        &_var4 = CTOD(SPACE(08))
      ElseIf _var2 [t] = "M"
        &_var4 = SPACE(10)
      ElseIf _var2 [t] = "L"
        &_var4 = SPACE(01)
      Endif
    NEXT t
    Z03LMIL="S"
  Else
    If !REGLOCK()
      MSGINFO("Registro sendo alterado por outro Usuário",WSIST)
      Return
    Else
      DECLARE _var1 [FCOUNT()]
      AFIELDS(_var1)
      FOR t = 1 TO FCOUNT()
        SysRefresh()
        _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
        _var5 = _var1 [t]
        &_var4 = &_var5
      NEXT t
    Endif
  Endif
  If Z03LMIL="S"
    Z03LMIL=.T.
  Else
    Z03LMIL=.F.
  Endif
  MFONTES()
  ZCLAS=VAL(Z03TIPO)
  DEFINE DIALOG oDlg FROM 0,0 TO 21,37 TITLE "Papeis" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 139,128
  @   0.87,01 SAY "Código" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,01 SAY "Descrição" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3,01 SAY "Tipo" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,01 SAY "Venc.Opções" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,01 SAY "Pap.Vista" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.0,01 SAY "POP" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.8,01 SAY "Preço Merc." FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  If F="I"
    @ 01,6 GET Z03CODP PICTURE "@!" VALID VSTCON(Z03CODP) FONT oFontget SIZE 50,11
  Else
    @ 01,6 GET Z03CODP PICTURE "@!" WHEN A="S" FONT oFontget SIZE 50,11
  Endif
  @ 02,6 GET Z03DESC PICTURE "@!" FONT oFontget SIZE 90,11
  @ 3,6 COMBOBOX oCbx VAR ZCLAS items { "(1) - Vista","(2) - Opção","(3) - Termo"} of oDlg UPDATE font oFontGET size 80,90
  @ 4.7,1 CHECKBOX Z03LMIL PROMPT "Cotado em lote de Mil"  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 130,11 OF oDlg
  @ 6,6 GET Z03DTOP PICTURE "@D" FONT oFontget SIZE 45,11
  @ 7,6 GET Z03CODPO PICTURE "@!" FONT oFontget SIZE 50,11
  @ 8,6 GET Z03POP PICTURE "!" FONT oFontget SIZE 10,11 VALID Z03POP$"SN"
  @ 9,6 GET Z03PM PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11

  @ 140,032 SBUTTON PROMPT "&Confirma" PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" FONT oFontbut ADJUST TEXT ON_CENTER OF oDlg size 40,12 ACTION (IIF(F="I",VERMC(3),""),VERMC(2),PUTREC(),oDlg:end())
  @ 140,077 SBUTTON PROMPT "C&ancela" PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" FONT oFontbut ADJUST TEXT ON_CENTER OF oDlg SIZE 40,12 ACTION (OdLG:eND())
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
Return

Procedure CADDIV()
  CLOSE DATABASE
  SELE 2
  If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
    MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1401
  Endif
  SELE 3
  If .NOT. USEREDE(ARQ11,.F.,10,"ITENS")
    MsgStop("O arquivo ITENS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1101
  Endif
  SELE 4
  If .NOT. USEREDE(ARQ04,.F.,10,"TABELA")
    MsgStop("O arquivo TABELA não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0401
  Endif
  SELE 5
  If !USEREDE(ARQ08,.F.,10,"CMVD")
    MsgStop("O arquivo TABELA CMVD não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0801
  Endif
  SELE 3
  SET FILTER TO VAL(A11NOTA)=0
  BrowseC(8,"Movimento Extra","Movimento Extra", { || buscaNOT() },{ ||NCADITENS1("I") },{ ||NCADITENS1("A") })
  USE
  CLOSE DATABASE
  CALCSALD()
  IMPRES(0)
  CURSORARROW()
Return

Procedure NCADITENS1(Fit)
  public oDlgit,oBrushit,oCbxix,oCbxix1
  PUBLIC ZCLASit,ZCLASIT1
  If Fit="I"
    GO BOTT
    SKIP
  Endif
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  ZHIST=VAL(Z11FLAG)
  If ZHIST=0
    ZHIST=1
  Endif
  MFONTES()
  DEFINE DIALOG oDlgit FROM 0,0 TO 21,41.5 TITLE "Item Dividendo/Bonificação/Outros" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 158,130
  @   0.87,01 SAY "E/S" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,01 SAY "Código" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,01 SAY "Descrição" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,01 SAY "Quantidade" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   4.35,01 SAY "Financeiro" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,01 SAY "Data" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,01 SAY "Tp.Mov" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.96,01 SAY "Historico" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.83,01 SAY "Desc.Hist" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  Z12DATA=Z11DATA
  If Z11OPER="C" .OR. Z11OPER="1"
    ZCLASIT=1
  Else
    ZCLASIT=2
  Endif
  @ 0.8,7 COMBOBOX oCbxix VAR ZCLASit items { "(1) - Entrada","(2) - Saída"} of oDlgit UPDATE font oFontGET size 80,90
  @ 2,7 GET Z11CODP PICTURE "@!" FONT oFontget SIZE 40,11 VALID VERMC(11)
  @ 2.61,7 SAY PESQ(Z11CODP,"2",1,"A03DESC") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,11
  @ 4,7 GET Z11QTDE PICTURE "@E 999,999,999,999" RIGHT FONT oFontget SIZE 60,11
  @ 5,7 GET Z11VALR PICTURE "@E 999,999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 6,7 GET Z11DATA PICTURE "@D" FONT oFontget SIZE 40,11 VALID VERMC(10,Z11DATA)
  @ 6.7,7 COMBOBOX oCbx VAR zhist items {"Outros","Subscrição","Bonificação","Transferencia","Implit"} OF oDlgit FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 70,60
  @ 8,7 GET Z11CODH PICTURE "999" FONT oFontget SIZE 20,11 RIGHT VALID VERMC(21)
  @ 9,7 GET ohist VAR Z11HIST PICTURE "@!" FONT oFontget SIZE 100,11
  @ 141.5,043 SBUTTON PIXELS FONT oFontbut prompt "&Confirma" TEXT ON_CENTER NOBOX RESOURCE "OFF","ON","DISABLE" OF oDlgit size 40,12 ACTION (IIF(FIT="I",VERMC(3),""),IIF(ZCLASIT=1,Z11OPER:="C",Z11OPER:="V"),Z11REG:=RECNO(),Z11FLAG:=STRZERO(ZHIST,1),PUTREC(),conti(3),oDlgit:end())
  @ 141.5,088 SBUTTON PIXELS FONT oFontbut PROMPT "C&ancela" TEXT ON_CENTER NOBOX RESOURCE "OFF","ON","DISABLE" OF OdLGit SIZE 40,12 ACTION (oDlgit:End())
  odlgIT:lhelpicon:=.f.
  ACTIVATE DIALOG oDlgit CENTERED
Return

Procedure CADSALD()
  // CALCULO DO MOVIMENTO *
  CLOSE DATABASE
  SELE 2
  If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
    MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1401
  Endif
  SELE 3
  If .NOT. USEREDE(ARQ04,.F.,10,"MERC")
    MsgStop("O arquivo MERCADOS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0401
  Endif
  SELE 1
  If .NOT. USEREDE(ARQ03,.F.,10,"SALD")
    MsgStop("O arquivo SALDOS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0301
  Endif
  SELE 1
  BrowseC(5,"Saldo","Saldo", { || buscaSAL() },{ ||NCADSAL("I") },{ ||NCADSAL("A") })
  USE
  CLOSE DATABASE
  CALCSALD()
  IMPRES(0)
  CLOSE DATABASE
  CURSORARROW()
Return

FUNCTION BUSCASAL()
  TEXTO="Papel:"
  VARIAVEL=SPACE(20)
  MASCARA="@!"
  REG=RECNO()
  Do While .T.
    SysRefresh()
    If !PROCURA()
      EXIT
    Endif
    CHAVE=VARIAVEL
    If EMPTY(CHAVE)
      EXIT
    Endif
    If AT("1",CHAVE)#0 .OR. AT("2",CHAVE)#0 .OR. AT("3",CHAVE)#0 .OR. AT("4",CHAVE)#0 .OR. AT("5",CHAVE)#0 .OR. AT("6",CHAVE)#0 .OR. AT("7",CHAVE)#0 .OR. AT("8",CHAVE)#0 .OR. AT("9",CHAVE)#0 .OR. AT("0",CHAVE)#0
      SET ORDER TO 1
      SEEK TRIM(CHAVE)
    Else
      SET ORDER TO 2
      SEEK TRIM(CHAVE)
    Endif
    If EOF() .OR. DELE()
      MSGALERT("Não Encontrado !",WSIST)
      GO REG
    Else
      EXIT
    Endif
  ENDDO
Return

Procedure NCADSAL(F)
  PUBLIC oDlg,oBrush,oCbx
  PUBLIC ZCLAS
  SELE 1
  If F="I"
    GO BOTT
    SKIP
  Endif
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  MFONTES()
  DEFINE DIALOG oDlg FROM 0,0 TO 31,60 TITLE "Saldos/Movto" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 230,210
  @   0.87,02 SAY "Código"    FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,02 SAY "Descrição" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,02 SAY "Mercado"   FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,02 SAY "MOVIMENTO                QUANTIDADE                           FINANCEIRO" FONT oFontsayB COLOR CLR_BLACK,nRGB(248,248,248) SIZE 215,11
  @   4.35,02 SAY "Saldo Anterior" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,02 SAY "Compras" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,02 SAY "Vendas" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.96,02 SAY "(Transferencia +)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.83,02 SAY "(Transferencia -)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   8.70,02 SAY "Baixas +" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   9.57,02 SAY "Baixas -" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  10.44,02 say "Lucros" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  11.31,02 SAY "Prejuízos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  12.18,02 SAY "Saldo Atual" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  13.05,02 SAY "Preco Medio" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  If F="I"
    @ 01,9 GET Z03CODP PICTURE "@!" VALID VSTCON(Z03CODP) .And. VERMC(7,1) FONT oFontget SIZE 50,11
  Else
    @ 01,9 GET Z03CODP PICTURE "@!" WHEN A="S" FONT oFontget SIZE 50,11
    VERMC(7,2)
    VERMC(8)
  Endif
  @ 05,9  GET Z03QANT PICTURE "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 05,19 GET Z03VANT PICTURE "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 06,9  GET Z03QCPA PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 06,19 GET Z03VCPA PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 07,9  GET Z03qVDA PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 07,19 GET Z03VVDA PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 08,9  GET Z03QTMA PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 08,19 GET Z03VTMA PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 09,9  GET Z03QTME PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 09,19 GET Z03VTME PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 10,9  GET Z03QBMA PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 10,19 GET Z03VBMA PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 11,9  GET Z03QBME PICT "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 11,19 GET Z03VBME PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 12,19 GET Z03VLUC PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 13,19 GET Z03VPRE PICT "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 14,9  GET Z03QATU PICTURE "@E 99,999,999,999" RIGHT FONT oFontget SIZE 70,11
  @ 14,19 GET Z03VATU PICTURE "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  @ 15,9  GET Z03PMER PICTURE "@E 9,999,999,999.99" RIGHT FONT oFontget SIZE 70,11
  //
  @ 220,078 SBUTTON PIXELS FONT oFontbut PROMPT "&Confirma" NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST TEXT ON_CENTER OF oDlg size 40,12 ACTION (IIF(F="I",VERMC(3),""),VERMC(18),PUTREC(),oDlg:end())
  @ 220,128 SBUTTON PIXELS FONT oFontbut PROMPT "C&ancela" NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST TEXT ON_CENTER  OF oDlg SIZE 40,12 ACTION (oDlg:End())
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
Return

Procedure WVERPAP()
  local oLbxver
  public oDlg3
  SELE 2
  SET ORDER TO 1
  GO TOP
  DEFINE DIALOG oDlg3 FROM 1, 5 TO 26, 79 TITLE "Papeis" BRUSH oBrushTela STYLE WS_CAPTION
  @ 1, 1 LISTBOX oLbxver FIELDS PAPEL->A03CODP,PAPEL->A03DESC,IIF(PAPEL->A03TIPO="1","Vista",IIF(PAPEL->A03TIPO="2","Opcao","Termo")) HEADERS "Código","Descrição","Mercado" SIZE 265, 137 OF oDlg3
  oLbxver:bLDblClick = { | nRow, nCol | oDlg3:end() }

  @ 165,085 SBUTTON PIXELS PROMPT "&Procura"  FONT oFontbut NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST TEXT ON_CENTER OF oDlg3 SIZE 40,12 ACTION (BUSCApap(),oLbxver:Refresh(),oLbxver:Setfocus())
  @ 165,145 SBUTTON PIXELS PROMPT "&Confirma" FONT oFontbut NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST TEXT ON_CENTER OF oDlg3 SIZE 40,12 ACTION oDlg3:end()
  odlg3:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg3
Return

Procedure CADNOTAS()
  CLOSE DATABASE
  SELE 1
  If .NOT. USEREDE(ARQ12,.F.,10,"NOTA")
    MsgStop("O arquivo NOTAS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1201
  Endif
  SELE 2
  If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
    MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1401
  Endif
  SELE 3
  If .NOT. USEREDE(ARQ11,.F.,10,"ITENS")
    MsgStop("O arquivo ITENS não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1101
  Endif
  SELE 4
  If .NOT. USEREDE(ARQ04,.F.,10,"TABELA")
    MsgStop("O arquivo TABELA não está  disponível",WSIST)
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0401
  Endif
  SELE 1
  SET ORDER TO 2
  BrowseC(6,"Notas","Notas", { || buscaNOT() },{ ||NCADNOT("I") },{ ||NCADNOT("A") },{ ||DELNOTA()})
  USE
  CLOSE DATABASE
  CALCSALD()
  IMPRES(0)
  CLOSE DATABASE
  CURSORARROW()
Return

Procedure delnota()
  If !MSGNOYES("Excluir esta nota ?",WSIST)
    Return
  Endif
  Z12NOTA=NOTA->A12NOTA
  SELE 3
  SET FILTER TO ITENS->A11NOTA=Z12NOTA
  GO TOP
  Do While .NOT. EOF()
    SysRefresh()
    If REGLOCK(10)
      DELE
    Endif
    SKIP
  ENDDO
  SET FILTER TO
  SELE 1
  If REGLOCK(10)
    DELE
  Endif
Return

FUNCTION BUSCANOT()
  TEXTO="Nota Nr:"
  VARIAVEL=0
  MASCARA="999999"
  REG=RECNO()
  Do While .T.
    SysRefresh()
    If !PROCURA()
      EXIT
    Endif
    CHAVE=VARIAVEL
    If EMPTY(CHAVE)
      EXIT
    Endif
    CHAVE=STRZERO(CHAVE,6)
    SET ORDER TO 1
    SEEK CHAVE
    If EOF() .OR. DELE()
      MSGALERT("Não Encontrado !",WSIST)
      GO REG
    Else
      EXIT
    Endif
  ENDDO
Return

Procedure NCADNOT(F)
  LOCAL oDlg,oBrush
  SET ORDER TO 1
  If F="I"
    GO BOTT
    SKIP
  Endif
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  MFONTES()
  DEFINE DIALOG oDlg FROM 0,0 TO 26,45 TITLE "Resumo da Nota" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 170,165
  @   0.87,02 SAY "Nota Nr." FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,02 SAY "Data Pregão" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,02 SAY "Tx.Liquid.CBLC" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,02 SAY "Tx.Reg.CBLC" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   4.35,02 SAY "Tx.Termo/Opções-Bovespa" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,02 SAY "Tx.Ana-Bovespa" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.09,02 SAY "Emolumentos-Bovespa" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   6.96,02 SAY "Corretagem" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   7.83,02 SAY "IRRF Descontado" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   8.70,02 SAY "IRRF Projetado" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   9.57,02 SAY "IRRF S/Day Trade" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @  10.44,02 SAY "Liquido" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  If F="I"
    @ 01,13 GET Z12NOTA  PICTURE "999999" VALID VAL(Z12NOTA)>0 .And. VSTCON(STRZERO(VAL(Z12NOTA),6)) FONT oFontget SIZE 25,11
  Else
    @ 01,13 GET Z12NOTA  PICTURE "999999" WHEN A="S" FONT oFontget SIZE 25,11
  Endif
  @ 02,13 GET Z12DATA    PICTURE "@D" FONT oFontget SIZE 35,11 VALID VERMC(10,Z12DATA) .And. Z12DATA<VALIDOATE
  @ 03,13 GET Z12TXLIQ   PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 04,13 GET Z12TXREG   PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 05,13 GET Z12TXTEOP  PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 06,13 GET Z12TXANA   PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 07,13 GET Z12EMOL    PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 08,13 GET Z12CORRET  PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 09,13 GET Z12IRRF    PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 10,13 GET Z12OUTRAS  PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 11,13 GET Z12IRDAY   PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 12,13 GET Z12LIQ     PICTURE "@E 999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 178,048 SBUTTON PIXELS FONT oFontbut NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST TEXT ON_CENTER PROMPT "&Confirma" OF oDlg size 40,12 ACTION (IIF(F="I",VERMC(3),""),VERMC(9),PUTREC(),oDlg:end())
  @ 178,093 SBUTTON PIXELS FONT oFontbut NOBOX RESOURCE "OFF","ON","DISABLE" adjust TEXT ON_CENTER PROMPT "C&ancela" OF ODlg SIZE 40,12 ACTION (oDlg:End())
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
  SET ORDER TO 2
Return

Procedure CADITENS()
  public oDlgi, oFonti, oBrushi
  Public oLbxi
  PUBLIC Z12FLAG1
  Z12NOTA=NOTA->A12NOTA
  Z12DATA=NOTA->A12DATA
  Z12IRRF=NOTA->A12IRRF
  TOTDEZP=NOTA->A12TXLIQ+NOTA->A12TXREG+NOTA->A12CORRET+NOTA->A12TXANA+NOTA->A12EMOL
  TOTDEZOPCAO=NOTA->A12TXTEOP
  WREG=RECNO()
  SELE 3
  SET FILTER TO A11NOTA=Z12NOTA
  MFONTES()
  DEFINE DIALOG oDlgi FROM 1, 5 TO 27, 90 TITLE "Operações da Nota "+Z12NOTA+" Pregão: "+DTOC(Z12DATA) FONT oFontTit BRUSH oBrushTela STYLE WS_CAPTION
  GO BOTT
  FOR I=1 TO 15
    SysRefresh()
    If !BOF()
      SKIP -1
    Endif
  NEXT I
  DECLARE TPS[3]
  TPS[1]="Vista"
  TPS[2]="Opção"
  TPS[3]="Termo"
  @ 1,1 LISTBOX oLbxi FIELDs IIF((ITENS->A11OPER="C" .OR. ITENS->A11OPER="1"),"Compra","Venda"),DTOC(ITENS->A11DTVAL),ITENS->A11CODP,IIF(PESQ(ITENS->A11CODP,"2",1,"A03TIPO")$"123",TPS[VAL(PESQ(ITENS->A11CODP,"2",1,"A03TIPO"))],""),LEFT(PESQ(ITENS->A11CODP,"2",1,"A03DESC"),20),TRANSFORM(ITENS->A11QTDE,"@E 999,999,999,999"),TRANSFORM(ITENS->A11VALR,"@E 999,999,999,999.99") HEADERS "C/V","Liquidação","Papel","Mercado","Descricao","Quantidade","Financeiro" SIZE 320,137 OF oDlgi
  @ 156,090 SBUTTON PIXELS RESOURCE "INCLUI" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 PROMPT "&Inclui" OF oDlgi ACTION (NCADITEM("I"),CONTI(1),CONTI(3),oLbxi:Refresh(),oLbxi:SetFocus())
  @ 156,130 SBUTTON PIXELS RESOURCE "ALTERA" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 PROMPT "&Altera" OF oDlgi ACTION (NCADITEM("A"),oLbxi:Refresh(),oLbxi:SetFocus())
  @ 156,170 SBUTTON PIXELS RESOURCE "EXCLUI" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 PROMPT "&Exclui" OF oDlgi ACTION (Rec1Delete(oLbxi),oLbxi:Refresh(),oLbxi:SetFocus())
  @ 156,210 SBUTTON PIXELS RESOURCE "SAIR2"  TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) SIZE 35,35 PROMPT "&Sair"   OF oDlgi ACTION (oDlgi:end())
  odlgI:lhelpicon:=.f.
  ACTIVATE DIALOG oDlgi CENTERED
  SELE 3
  GO TOP
  TOTOP=0
  TOTVIS=0
  Do While !EOF()
    SysRefresh()
    If A11NOTA=Z12NOTA
      If PESQ(A11CODP,"2",1,"A03POP")="S"
        SKIP
        LOOP
      Endif
      If VAL(PESQ(A11CODP,"2",1,"A03TIPO"))=2 .or.VAL(PESQ(A11CODP,"2",1,"A03TIPO"))=3
        TOTOP=TOTOP+A11VALR
      Endif
      TOTVIS=TOTVIS+A11VALR
    Endif
    SKIP
  ENDDO
  GO TOP
  Do While !EOF()
    SysRefresh()
    If A11NOTA=Z12NOTA .And. REGLOCK(10)
      If PESQ(A11CODP,"2",1,"A03POP")="S"
        SKIP
        LOOP
      Endif
      TP=PESQ(A11CODP,"2",1,"A03TIPO")
      If val(TP)=2 .or. val(TP)=3
        REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)+(A11VALR/TOTOP*TOTDEZOPCAO)
      Else
        REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)
      Endif
      If A11OPER="C" .OR. A11OPER="1"
        REPLACE A11VLIQ WITH A11VALR+A11VCOR
      Else
        REPLACE A11VLIQ WITH A11VALR-A11VCOR
      Endif
      UNLOCK
    Endif
    SKIP
  ENDDO
  GO TOP
  LIQ=0
  Do While !EOF()
    SysRefresh()
    If PESQ(A11CODP,"2",1,"A03POP")="S"
      SKIP
      LOOP
    Endif
    If A11NOTA=Z12NOTA .And. PESQ(A11CODP,"2",1,"A03TIPO")#"3"
      If A11OPER="V"
        LIQ=LIQ+A11VLIQ
      Else
        LIQ=LIQ-A11VLIQ
      Endif
    Endif
    If A11NOTA=Z12NOTA .And. PESQ(A11CODP,"2",1,"A03TIPO")="3"
      LIQ=LIQ-a11VCOR
    Endif
    SKIP
  ENDDO
  SELE 3
  SET FILTER TO
  SELE 1
  If REGLOCK(10)
    REPLACE A12LIQ WITH LIQ-A12IRRF
  Endif
  UNLOCK
Return

Procedure NCADITEM(Fit)
  public oDlgit,oBrushit,oCbxix,oCbxix1
  PUBLIC ZCLASit,ZCLASIT1
  If FIT="I"
    GO BOTT
    SKIP
  Endif
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  If Fit="I"
    ZCLASit=1
    ZCLASIT1=1
  Endif
  If Z11OPER="C"
    ZCLASIT=1
  Else
    ZCLASIT=2
  Endif
  MFONTES()
  DEFINE DIALOG oDlgit FROM 0,0 TO 15.5,37 TITLE "Nota Nr:"+Z12NOTA+" Pregão: "+DTOC(Z12DATA) BRUSH oBrushTela STYLE WS_CAPTION
  ZMERC=" "
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 141,87
  @   0.87,01 SAY "C/V" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,01 SAY "Código/Prazo" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,01 SAY "Descrição" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   3.48,01 SAY "Quantidade" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   4.35,01 SAY "Financeiro" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   5.22,01 SAY "Liquidação" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  @ 0.8,7 COMBOBOX oCbxix VAR ZCLASit items { "(1) - Compra","(2) - Venda"} of oDlgit UPDATE font oFontGET size 84,90
  @ 2,7 GET Z11CODP PICTURE "@!" FONT oFontget SIZE 40,11 VALID VERMC(12)
  @ 2.61,9.2 SAY PESQ(Z11CODP,"2",1,"A03DESC") FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 80,11
  @ 2,15 GET Z11PRAZO PICTURE "999" FONT oFontget SIZE 20,11 WHEN ZMERC="3" VALID VERMC(20)
  @ 4,7 GET Z11QTDE PICTURE "@E 999,999,999,999" RIGHT FONT oFontget SIZE 60,11
  @ 5,7 GET Z11VALR PICTURE "@E 999,999,999,999.99" RIGHT FONT oFontget SIZE 60,11
  @ 6,7 GET oz11dtval var Z11DTVAL PICTURE "@D" FONT oFontget SIZE 45,11 WHEN ZMERC="3"
  If FIT="A"
    @ 100,030 SBUTTON PIXELS TEXT ON_CENTER PROMPT "&Confirma" ADJUST NOBOX RESOURCE "OFF","ON","DISABLE" OF oDlgit size 40,12 ACTION (IIF(FIT="I",VERMC(3),""),VERMC(14),z11reg:=recno(),PUTREC(),conti(3),oDlgi:Refresh(),oDlgit:end())
    @ 100,080 SBUTTON PIXELS TEXT ON_CENTER PROMPT "C&ancela" ADJUST NOBOX RESOURCE "OFF","ON","DISABLE" OF ODlgit SIZE 40,12 ACTION (oDlgit:End())
  Else
    @ 100,008 SBUTTON PIXELS TEXT ON_CENTER PROMPT "&Confirma" ADJUST NOBOX RESOURCE "OFF","ON","DISABLE" OF oDlgit size 40,12 ACTION (IIF(FIT="I",VERMC(3),""),VERMC(14),z11reg:=recno(),PUTREC(),conti(3),oDlgi:Refresh(),oDlgit:end())
    @ 100,053 SBUTTON PIXELS TEXT ON_CENTER PROMPT "C&ancela" ADJUST NOBOX RESOURCE "OFF","ON","DISABLE"  OF ODlgit SIZE 40,12 ACTION (oDlgit:End())
    @ 100,098 SBUTTON PIXELS TEXT ON_CENTER PROMPT "&Sair" ADJUST NOBOX RESOURCE "OFF","ON","DISABLE" OF oDlgit size 40,12 ACTION (OdLGIT:END(),VERMC(13))
  Endif
  odlgIT:lhelpicon:=.f.
  ACTIVATE DIALOG oDlgit CENTERED
  oLbxi:Refresh()
Return

FUNCTION CONTI(X)
  CONTINUA=.T.
  If X=1
    Do While CONTINUA
      SysRefresh()
      NCADITEM("I")
    ENDDO
  Endif
  If X=3
    SELE 3
    GO BOTT
    FOR I=1 TO 15
      SysRefresh()
      If !BOF()
        SKIP -1
      Endif
    NEXT I
  Endif
Return

Procedure IMPNOTA1()
  MWREG=RECNO()
  If MSGYESNO("Relatório de Notas ?",WSIST)
    go top
    FZ=.F.
    CReport(9,oLbx,"Resumo das Notas")
    sele nota
    go mwreg
    FZ=.T.
    Return
  Endif
  IMPNOTA(1)
  SELE NOTA
  GO MWREG
Return

function verl()
  If !FZ
    Return 0
  Endif
  Z12NOTA=NOTA->A12NOTA
  SELE 3
  SET FILTER TO ITENS->A11NOTA=Z12NOTA
  GO TOP
  ZLIQN=0
  Do While .NOT. EOF()
    SysRefresh()
    If ITENS->A11OPER="C" .OR. ITENS->A11OPER="1"
      ZLIQN=ZLIQN-ITENS->A11VLIQ
    Else
      ZLIQN=ZLIQN+ITENS->A11VLIQ
    Endif
    SKIP
  ENDDO
  SET FILTER TO
  SELE 1
Return ZLIQN

function IMPNOTA(QN)
  local n
  LOCAL oFont1, oFont2
  PUBLIC oRpt
  FZ=.F.
  DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-8
  DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-8
  DEFINE FONT oFont3 NAME "VERDANA" SIZE 0,-8
  DEFINE PEN oPen1 WIDTH 0.1
  DEFINE PEN oPen2 WIDTH 0.1
  Z12NOTA=NOTA->A12NOTA
  Z12DATA=NOTA->A12DATA
  Z12IRRF=NOTA->A12IRRF
  reg=recno()
  SELE 3
  SET FILTER TO A11NOTA=Z12NOTA
  GO TOP
  COUNT ALL TO TOTOPER FOR A11NOTA=Z12NOTA
  GO TOP
  If EOF()
    MSGINFO("Não Existe Registros",WSIST)
    Return
  Endif
  GO TOP
  CTITLE="Nota de Corretagem Nr: "+Z12NOTA+" De "+DTOC(Z12DATA)+" Imposto retido: "+ALLTRIM(TRANSFORM(Z12IRRF,"@E 999,999,999.99"))
  DECLARE TPS[3]
  TPS[1]="Vista"
  TPS[2]="Opção"
  TPS[3]="Termo"

  If PREV

  REPORT oRpt TITLE " "," "," "," ",EMP1,CTITLE FONT oFont1,oFont2,oFont3;
    PEN oPen1,oPen2;
    PREVIEW;
    FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3),"DT=Day Trade";
    CAPTION ctitle;

  Else

  REPORT oRpt TITLE " "," "," "," ",EMP1,CTitle FONT oFont1,oFont2,oFont3;
    PEN oPen1,oPen2;
    FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3),"DT=Day Trade";
    CAPTION ctitle;

  Endif

  COLUMN TITLE "C/V" DATA ITENS->A11OPER FONT 2 SIZE 4 GRID
  COLUMN TITLE "Papel" DATA ITENS->A11CODP FONT 2 SIZE 9 GRID
  COLUMN TITLE "Descrição" DATA PESQ(ITENS->A11CODP,"2",1,"A03DESC") FONT 2 SIZE 14 GRID
  If QN=1
    COLUMN TITLE "Mercado" DATA IIF(PESQ(ITENS->A11CODP,"2",1,"A03TIPO")$"123",TPS[VAL(PESQ(ITENS->A11CODP,"2",1,"A03TIPO"))],"") FONT 2 SIZE 10 GRID
  Else
    COLUMN TITLE "Mercado" DATA IIF(PESQ(ITENS->A11CODP,"2",1,"A03TIPO")$"123",TPS[VAL(PESQ(ITENS->A11CODP,"2",1,"A03TIPO"))],"") FONT 2 SIZE 10 GRID
  Endif

  COLUMN TITLE "Quantidade" DATA ITENS->A11QTDE PICTURE "@E 999,999,999,999" RIGHT FONT 2 SIZE 12 GRID
  COLUMN TITLE "Financeiro" DATA ITENS->A11VALR PICTURE "@E 999,999,999.99" RIGHT FONT 2 TOTAL SIZE 12 GRID
  COLUMN TITLE "Despesas" DATA ITENS->A11VCOR PICTURE "@E 999,999,999.99" FONT 2 RIGHT TOTAL SIZE 12 GRID
  COLUMN TITLE "Liquido s/I.R" DATA IIF(PESQ(ITENS->A11CODP,"2",1,"A03TIPO")#"3",(IIF(ITENS->A11OPER="C" .OR. ITENS->A11OPER="1",ITENS->A11VLIQ*-1,ITENS->A11VLIQ)), (IIF(ITENS->A11OPER="C" .OR. ITENS->A11OPER="1",ITENS->A11VCOR*-1,ITENS->A11VCOR)) ) PICTURE "@E 999,999,999.99" RIGHT FONT 2 TOTAL SIZE 12 GRID
  oRpt:cGrandTotal := "Total Geral " + strzero(totoper,3)+" Operação(oes)"
  oRpt:CELLVIEW()
  oRpt:lpagetotal:=.F.

  ENDREPORT

  ACTIVATE REPORT oRpT ON STARTPAGE IMPRIMECAB()

  SELE 3
  SET FILTER TO
  SELE 1
  GO reg
  oFont1:End()
  oFont2:End()
  oFont3:End()
  FZ=.T.
Return nil

FUNCTION PERIODO()
  local oD1lg
  PUBLIC RETORNO
  MFONTES()
  RETORNO=.T.
  DEFINE DIALOG oD1lg FROM 1, 1 TO 13.5 , 55 TITLE "Período" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,9.0 say "" box raised COLOR CLR_BLACK,nRGB(248,248,248) size 135,50 OF oD1lg
  @   0.87,09.5 say "Informe o Período Desejado" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,11
  @   1.74,9.5 say "De:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,9.5 SAY "Ate:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @ 2,11 GET DE FONT oFontget SIZE 40,11
  @ 3,11 GET ATE FONT oFontget SIZE 40,11
  @ 0.7,0.5 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER size 46,65 OF oD1lg
  @ 70,060 SBUTTON PIXELS ADJUST PROMPT "&Confirma" TEXT ON_CENTER NOBOX  RESOURCE "OFF","ON","DISABLE" FONT oFontbut OF oD1lg SIZE 40,12 ACTION oD1lg:end()
  @ 70,110 SBUTTON PIXELS ADJUST PROMPT "C&ancela" TEXT ON_CENTER NOBOX RESOURCE "OFF","ON","DISABLE" FONT oFontbut OF oD1lg SIZE 40,12  ACTION (VERMC(17),oD1lg:End())
  od1lg:lhelpicon:=.f.
  ACTIVATE DIALOG oD1lg CENTERED
Return RETORNO

Procedure AUTOPU(QUEMPED,QUALARQ)
  CLOSE DATABASE
  NAOFEZBDI=.T.
  MZARQ=LOCALCART+"BDI.DAT"
  Erase &MZARQ
  If !FILE(LOCALCART+"BDI.DAT")
    mat_stru = {}
    AADD(MAT_STRU,{"NADA","C",4,0})
    AADD(MAT_STRU,{"NMERC","C",16,0})
    AADD(MAT_STRU,{"NADA5","C",14,0})
    AADD(MAT_STRU,{"DESC","C",20,0})
    AADD(MAT_STRU,{"NADA2","C",3,0})
    AADD(mat_stru,{"CODP","C",12,0})
    AADD(MAT_STRU,{"NADA3","C",3,0})
    AADD(MAT_STRU,{"MERCADO","C",11,0})
    AADD(mat_stru,{"NADA1","C",40,0})
    AADD(mat_stru,{"PRECO","C",11,0})
    AADD(MAT_STRU,{"NADA4","C",87,0})
    AADD(MAT_STRU,{"VENCOP","C",8,0})
    DBCREATE(LOCALCART+"BDI.DAT",mat_stru)
  Endif
  SELE 1
  If !USEREDE(LOCALCART+"BDI.DAT",.T.,10,"BDI")
    If QUEMPED=0
      MsgStop("O arquivo BDI não está  disponível",WSIST)
    Endif
    CLOSE DATABASE
    Return
  Endif
  If QUEMPED=0
    cOutFile:=cGetFile("Arq.Internet (BDI*.*) | BDI*.* |","Arquivo a ser utilizado")
    ARQ=ALLTRIM(COUTFILE)
    If AT(".",ARQ)=0
      ARQ=ARQ+"."
    Endif
    If !FILE(ARQ)
      MSGINFO("Arquivo "+arq+" Inexistente !",WSIST)
      CLOSE DATABASE
      Return
    Endif
  Else
    ARQ=QUALARQ
  Endif
  APPEND FROM &ARQ SDF
  SELE 2
  If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
    If QUEMPED=0
      MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
    Endif
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind1401
  Endif
  SET ORDER TO 1
  SELE 3
  If !USEREDE(ARQ06,.F.,10,"BDIS")
    If QUEMPED=0
      MsgStop("O arquivo BDDIS não está  disponível",WSIST)
    Endif
    CLOSE DATABASE
    Return
  Else
    SET INDEX TO &ind0601
  Endif
  SELE 1
  GO TOP
  CURSORWAIT()
  AA=SUBSTR(NADA5,3,4)
  MM=SUBSTR(NADA5,7,2)
  DD=SUBSTR(NADA5,09,2)
  NRABDI=CTOD(DD+"/"+MM+"/"+AA)
  Do While .NOT. EOF()
    SysRefresh()
    If MERCADO="FRACIONARIO" .OR. (LEFT(BDI->CODP,1)$"0123456789 ") .OR. (LEFT(BDI->DESC,1)$"0123456789")
      SKIP
      LOOP
    Endif
    If BDI->MERCADO="VISTA  "
      TPM="1"
    ElseIf BDI->MERCADO="OPC"
      TPM="2"
    ElseIf BDI->MERCADO="TERMO"
      TPM="3"
    Else
      SELE 1
      SKIP
      LOOP
    Endif
    SELE 2
    SET ORDER TO 1
    SEEK BDI->CODP
    If EOF()
      ADIREG(0)
      UNLOCK
    Endif
    SELE 3
    SEEK BDI->CODP+DTOS(NRABDI)
    If VAL(BDI->PRECO)#0 .OR. EOF()
      If EOF()
        ADIREG(0)
      Else
        REGLOCK(10)
      Endif
      REPLACE DATA WITH NRABDI
      REPLACE CODPAPEL WITH BDI->CODP
      REPLACE PRECOM WITH VAL(BDI->PRECO)/100
    Endif
    SELE 2
    If A03DTBD>=NRABDI
      SELE 1
      SKIP
      LOOP
    Endif
    NAOFEZBDI=.F.
    If REGLOCK(10)
      REPLACE A03CODP WITH BDI->CODP
      REPLACE A03DESC WITH BDI->DESC
      REPLACE A03TIPO WITH TPM
      If VAL(A03TIPO)=3
        REPLACE A03CODPO WITH LEFT(ALLTRIM(A03CODP),LEN(ALLTRIM(A03CODP))-1)
      Endif
      If VAL(A03TIPO)=2
        REPLACE A03DTOP WITH CTOD(RIGHT(BDI->VENCOP,2)+"/"+SUBSTR(BDI->VENCOP,5,2)+"/"+LEFT(BDI->VENCOP,4))
        DO CASE
          CASE AT("ON ",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"3"
          CASE AT("PN ",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"4"
          CASE AT("PNA",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"5"
          CASE AT("PNB",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"6"
          CASE AT("PNC",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"7"
          CASE AT("PND",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"8"
          CASE AT("PNE",A03DESC)#0
            REPLACE A03CODPO WITH LEFT(A03DESC,4)+"9"
        ENDCASE
        If EMPTY(A03CODPO)
          REPLACE A03CODPO WITH LEFT(A03DESC,4)
        Endif
      Endif
      If BDI->MERCADO="VISTA      "
        REPLACE A03CODT WITH "01"
      ElseIf BDI->MERCADO="TERMO      "
        REPLACE A03CODT WITH "03"
      ElseIf BDI->MERCADO="OPC"
        REPLACE A03CODT WITH "02"
      Endif
      If AT("*",A03DESC)>0
        REPLACE A03LMIL WITH "S"
      Else
        REPLACE A03LMIL WITH "N"
      Endif
      If VAL(BDI->PRECO)#0
        REPLACE A03PM WITH VAL(BDI->PRECO)/100
        REPLACE A03DTBD WITH NRABDI
      Endif
      UNLOCK
    Endif
    SELE 1
    SKIP
  ENDDO
  CURSORARROW()
  CLOSE DATABASE
  MAR=LOCALCART+"BDI.DAT"
  Erase &MAR
Return

Procedure IMPCART(QUEMPEDE,NOMEPDF)
  local n
  LOCAL oFont1, oFont2
  PUBLIC oRpt
  DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-8
  DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-8
  DEFINE FONT oFont3 NAME "VERDANA" SIZE 0,-8
  DEFINE PEN oPen1 WIDTH 0.1
  DEFINE PEN oPen2 WIDTH 0.1
  CLOSE DATABASE
  SELE 3
  If USEREDE(ARQ03,.F.,10,"SALDO")
    SET INDEX TO &ind0301
    SET ORDER TO 3
  Else
    If QUEMPEDE=0
      MSGSTOP("Não foi possível abrir arquivo de Carteira",WSIST)
    Endif
    CLOSE DATABASE
    Return .F.
  Endif
  SET FILTER TO A03QATU#0
  GO TOP
  If EOF()
    If QUEMPEDE=0
      MSGINFO("Não Existe Dados para Impressão",WSIST)
      CLOSE DATABASE
      Return
    Endif
  Endif

  CTITLE="Demonstrativo Carteira - Ano: "+_DATA
  DECLARE TPS[5]
  TPS[1]="Vista"
  TPS[2]="Opcão"
  TPS[3]="Termo"
  TPS[4]="POP"
  TPS[5]="N.Consta"

  If QUEMPEDE=1
    oReg := TReg32():Create(,"SOFTWARE\Dane Prairie Systems\Win2PDF")
    oReg:Set("PDFFilename",NOMEPDF)
    oReg:Close()
    PRINT oPrn NAME CTITLE TO TEMPDF1
    ENDPRINT
  Endif

  GO TOP
  TOTCAR=0
  Do While !EOF()
    SysRefresh()
    If (VAL(A03TIPO)>4 .OR. VAL(A03TIPO)<1) .And. REGLOCK(10)
      REPLACE A03TIPO WITH "1"
      REPLACE A03CODT WITH "01"
    Endif
    TOTCAR=TOTCAR+A03VMER
    SKIP
  ENDDO
  GO TOP

  If PREV .And. QUEMPEDE=0

    REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTITLE FONT oFont1,oFont2,oFont3;
      PEN oPen1,oPen2;
      PREVIEW;
      FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3),"* % da Carteira (Preço de Mercado)";
      CAPTION ctitle;

    Else

    REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTitle FONT oFont1,oFont2,oFont3;
      PEN oPen1,oPen2;
      FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3),"* % da Carteira (Preço de Mercado)";
      CAPTION ctitle;

    Endif

    COLUMN TITLE "Código" DATA SALDO->A03codp FONT 2 SIZE 8 GRID
    COLUMN TITLE "Descrição/Tipo" DATA TRIM(LEFT(SALDO->A03desc,12))+" "+SUBSTR(SALDO->A03DESC,13,12) FONT 2 SIZE 17 GRID
    COLUMN TITLE "Q.Atual" DATA SALDO->A03QATU PICTURE "@E 999,999,999,999" RIGHT FONT 2 TOTAL SIZE 12 GRID
    COLUMN TITLE "Pço Aquisição" DATA IIF(SALDO->A03LMIL="S",(SALDO->A03VATU/SALDO->A03QATU*1000),(SALDO->A03VATU/A03QATU)) PICTURE "@E 999,999,999.99" RIGHT SIZE 10 FONT 2 GRID
    COLUMN TITLE "Financeiro" DATA SALDO->A03VATU PICTURE "@E 999,999,999,999.99" RIGHT SIZE 13 FONT 2 TOTAL GRID
    //  COLUMN TITLE "Pço Merc." DATA SALDO->A03pmer PICTURE "@E 999,999,999.99" FONT 2 RIGHT SIZE 10 GRID
    //  COLUMN TITLE "Financeiro" DATA SALDO->A03VMER PICTURE "@E 999,999,999,999.99" RIGHT TOTAL FONT 2 SIZE 13 GRID
    //  COLUMN TITLE "Variação R$" DATA SALDO->A03vari PICTURE "@E 999,999,999,999.99" RIGHT FONT 2 TOTAL SIZE 13 GRID
    //  COLUMN TITLE "* %" DATA SALDO->A03VMER/TOTCAR*100 PICTURE "@E 999.99" RIGHT FONT 2 TOTAL SIZE 7 GRID

    //  COLUMN TITLE "BDI de" DATA SALDO->A03DTBD PICTURE "@D" FONT 2 SIZE 9 RIGHT GRID

    GROUP ON SALDO->A03TIPO HEADER "Mercado: "+TPS[VAL(oRpt:agroups[1]:cvalue)] FOOTER "Total do Mercado" FONT 1

    oRpt:cPageTotal:="Sub-Total"
    oRpt:cGrandTotal:="Total Geral "
    oRpt:CELLVIEW()
    oRpt:lpagetotal:=.F.
    ENDREPORT

    ACTIVATE REPORT oRpT ON STARTPAGE IMPRIMECAB() FOR SALDO->A03QATU#0 .And. SALDO->A03POP<>"S"

    oFont1:End()
    oFont2:End()
    CLOSE DATABASE
  Return NIL

  Procedure ATBDI(QUEMPED)
    CURSORWAIT()
    SELE 1
    If USEREDE(ARQ03,.F.,10,"SALDO")
      SET INDEX TO &ind0301
    Else
      If QUEMPED=0
        MSGSTOP("Não foi possível abrir arquivo de Carteira",WSIST)
      Endif
      CLOSE DATABASE
      Return .F.
    Endif
    SELE 2
    If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
      If QUEMPED=0
        MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
      Endif
      CLOSE DATABASE
      Return
    Else
      SET INDEX TO &ind1401
      SET ORDER TO 1
    Endif
    SELE 3
    If !USEREDE(ARQ06,.F.,10,"BDIS")
      If QUEMPED=0
        MsgStop("O arquivo BDIS não está  disponível",WSIST)
      Endif
      CLOSE DATABASE
      Return
    Else
      SET INDEX TO &ind0601
      SET ORDER TO 1
    Endif
    SELE 1
    GO TOP
    Do While !EOF()
      SysRefresh()
      sele 1
      CHAVE=A03CODP
      SELE 3
      SET SOFTSEEK ON
      SEEK CHAVE+DTOS(M->CALCATE)
      SET SOFTSEEK OFF
      If CODPAPEL<>CHAVE .And. !BOF()
        SKIP -1
      Endif
      If CODPAPEL<>CHAVE
        SELE 1
        SKIP
        LOOP
      Endif
      SELE 2
      SEEK CHAVE
      If !EOF()
        m03desc=a03desc
        m03lmil=a03lmil
        m03pmer=BDIS->PRECOM
        M03CODT=A03CODT
        M03TIPO=A03TIPO
        M03DTBD=BDIS->DATA
        M03POP=A03POP
        SELE 1
        If REGLOCK(10)
          If a03lmil = "S" .And. M03PMER#0
            REPLACE A03vmer WITH a03qatu * (m03pmer/1000)
            REPLACE A03pmer WITH m03pmer
          ElseIf A03LMIL#"S" .And. M03PMER#0
            REPLACE A03vmer WITH A03qatu * m03pmer
            REPLACE A03pmer WITH M03PMER
          Endif
          REPLACE A03DESC WITH M03DESC
          REPLACE A03LMIL WITH M03LMIL
          REPLACE A03vari WITH A03vmer - A03vatu
          REPLACE A03CODT WITH M03CODT
          REPLACE A03TIPO WITH M03TIPO
          REPLACE A03DTBD WITH M03DTBD
          REPLACE A03POP  WITH M03POP
          If A03POP="S"
            REPLACE A03TIPO WITH "4"
            REPLACE A03CODT WITH "04"
          Endif
        Endif
      Endif
      SELE 1
      SKIP
    ENDDO
    CLOSE DATABASE
    CURSORARROW()
  Return

  Procedure CALCSALD()
    MSGRUN("Aguarde Apurando Resultados !",WSIST,{ ||CALCSAL2(0)})
    CLOSE DATABASE
    ATBDI(0)
    CLOSE DATABASE
  Return

  Procedure IMPMOV(TIPO,QUEMPED,NOMEPDF)
    LOCAL oFont1,oFont2,oPen1,oPen2
    PUBLIC oRpt
    DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-8
    DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-8
    DEFINE FONT oFont3 NAME "VERDANA" SIZE 0,-8
    DEFINE PEN oPen1 WIDTH 0.1
    DEFINE PEN oPen2 WIDTH 0.1
    DE=CTOD("01/01/"+STRZERO(VAL(_DATA),4))
    ATE=CTOD("31/12/"+STRZERO(VAL(_DATA),4))
    If !AUTOMATICO
      PERIODO()
      If !RETORNO
        CLOSE DATABASE
        Return
      Endif
    Endif
    SELE 2
    If USEREDE(ARQ03,.F.,10,"SALDO")
      SET INDEX TO &ind0301
    Else
      If QUEMPED=0
        MSGSTOP("Não foi possível abrir arquivo de Carteira","Carteira")
      Endif
      CLOSE DATABASE
      Return
    Endif
    // Movimento Diario *
    SELE 4
    If .NOT. USEREDE(ARQ10,.F.,10,"DIARIO")
      If QUEMPED=0
        MsgStop("O arquivo MOVTO DIARIO não está  disponível","Movto Diario")
      Endif
      CLOSE DATABASE
      Return
    Else
      SET INDEX TO &ind1001
    Endif
    If EOF()
      If QUEMPED=0
        MSGINFO("Não Existe Dados para Impressão","Wincart")
        CLOSE DATABASE
        Return
      Endif
    Endif
    Do While .T.
      SysRefresh()
      M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
      ARQ_PRN=_LOCAL+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
      ARQ_PRN1:=ARQ_PRN+".DAT"
      ARQ_PRN2:=ARQ_PRN+".KEY"
      ARQ_PRN3:=ARQ_PRN+".TMP"
      If !FILE(ARQ_PRN1)
        EXIT
      Endif
    ENDDO
    SELE 20
    MAT_STRU:={}
    AADD(MAT_STRU,{"DATA","D",8,0})
    AADD(MAT_STRU,{"PAPEL","C",30,0})
    AADD(MAT_STRU,{"CODT","C",2,0})
    AADD(MAT_STRU,{"HISTGE","C",40,0})
    AADD(MAT_STRU,{"QUANTID","N",19,0})
    AADD(MAT_STRU,{"FINANCE","N",19,2})
    AADD(MAT_STRU,{"CUSTO","N",19,2})
    AADD(MAT_STRU,{"RESULTA","N",19,2})
    AADD(MAT_STRU,{"QUANTIS","N",19,0})
    AADD(MAT_STRU,{"FINANCS","N",19,2})
    AADD(MAT_STRU,{"VOPCOES","N",19,2})
    AADD(MAT_STRU,{"VVISTA","N",19,2})
    AADD(MAT_STRU,{"VTERMO","N",19,2})
    AADD(MAT_STRU,{"VDAYTR","N",19,2})
    AADD(MAT_STRU,{"VDAYTRO","N",19,2})
    DBCREATE(ARQ_PRN1,MAT_STRU)
    USE
    SELE 20
    If !USEREDE(ARQ_PRN1,.F.,10,"TMP")
      If QUEMPED=0
        MSGSTOP("Não foi possiver utilizar os arquivos Temporarios","Tente Mais Tarde")
      Endif
      CLOSE DATABASE
      Erase &ARQ_PRN1
      Return
    Endif
    SELE 4
    GO TOP
    Do While .NOT. EOF()
      SysRefresh()
      SELE 2
      SEEK DIARIO->A10CODP
      SELE 20
      ADIREG(0)
      REPLACE DATA WITH DE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Saldo Anterior"
      //  REPLACE QUANTIS WITH SALDO->A03QANT
      //  REPLACE FINANCS WITH SALDO->A03VANT
      REPLACE QUANTID WITH SALDO->A03QANT
      REPLACE FINANCE WITH SALDO->A03VANT
      SELE 4
      VISTA=0
      OPCOES=0
      TERMO=0
      DTVISTA=0
      DTOPCOES=0
      LUCROS=0
      PREJUIZOS=0
      ANTES=A10CODP
      Do While A10CODP=ANTES .And. !EOF()
        SysRefresh()
        SELE 20
        ADIREG(0)
        REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
        REPLACE DATA WITH DIARIO->A10DATA
        REPLACE HISTGE WITH DTOC(DIARIO->A10DATA)
        If SALDO->A03TIPO="3"
          REPLACE HISTGE WITH DTOC(DIARIO->A10DTVAL)
        Endif
        If DIARIO->A10oper = "1"
          REPLACE HISTGE WITH TRIM(HISTGE)+" "+"(CPA) NC:"+DIARIO->A10NOTA
        ElseIf DIARIO->A10oper = "2"
          REPLACE HISTGE WITH TRIM(HISTGE)+" "+"(VDA) NC:"+DIARIO->A10NOTA
        Endif
        If DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
          REPLACE HISTGE WITH TRIM(HISTGE)+" DAY TRADE"
        Endif
        If DIARIO->A10TERMO="S"
          REPLACE HISTGE WITH TRIM(HISTGE)+" TERMO"
        Endif
        If VAL(DIARIO->A10NOTA)=0
          Z13HIST=""
          ZLAG=DIARIO->A10FLAG
          If ZLAG="1"
            Z13HIST=" OUTROS "
          ElseIf ZLAG="2"
            Z13HIST=" SUBSCRICAO"
          ElseIf ZLAG="3"
            Z13HIST=" BONIFICACAO"
          ElseIf ZLAG="4"
            Z13HIST=" TRANSFERENCIA"
          ElseIf ZLAG="5"
            Z13HIST=" IMPLIT"
          ElseIf VAL(ZLAG)=0
            If DIARIO->A10OPER="1"
              Z13HIST=" (ENTRADA)"
            Else
              Z13HIST=" (SAIDA)"
            Endif
          Endif
          REPLACE HISTGE WITH DTOC(DIARIO->A10DATA)+Z13HIST
          REPLACE DATA WITH DIARIO->A10DATA
          If DIARIO->A10CODH#0
            REPLACE HISTGE WITH DTOC(DIARIO->A10DATA)+" "+STRZERO(DIARIO->A10CODH,3)+" "+LEFT(DIARIO->A10HIST,23)
          Endif
        Endif
        If DIARIO->A10FLAG="E"
          REPLACE HISTGE WITH DTOC(DIARIO->A10DATA)+" EXERCICIO OPCOES"
        Endif
        REPLACE DATA WITH DIARIO->A10DATA
        REPLACE QUANTID WITH DIARIO->A10QTDE
        REPLACE FINANCE WITH DIARIO->A10VALR
        REPLACE CUSTO WITH DIARIO->A10CUSTO
        REPLACE RESULTA WITH DIARIO->A10RESU
        If SALDO->A03TIPO="1" .And. DIARIO->A10FLAG<>"D"
          REPLACE VVISTA WITH DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
          REPLACE VDAYTR WITH DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
          REPLACE VTERMO WITH DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG<>"D"
          REPLACE VOPCOES WITH DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
          REPLACE VDAYTRO WITH DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
          REPLACE VTERMO WITH DIARIO->A10RESU
        Endif
        If DIARIO->A10RESU>0
          LUCROS=LUCROS+DIARIO->A10RESU
        Else
          PREJUIZOS=PREJUIZOS+(DIARIO->A10RESU*-1)
        Endif
        If SALDO->A03TIPO="1" .And. DIARIO->A10FLAG<>"D"
          VISTA=VISTA+DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
          DTVISTA=DTVISTA+DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
          TERMO=TERMO+DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG<>"D"
          OPCOES=OPCOES+DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
          DTOPCOES=DTOPCOES+DIARIO->A10RESU
        ElseIf SALDO->A03TIPO="3" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
          TERMO=TERMO+DIARIO->A10RESU
        Endif
        SELE 4
        SKIP
      ENDDO
      SELE 20
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "* R E S U M O * "+_DATA
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Saldo Anterior:"
      REPLACE QUANTID WITH SALDO->A03QANT
      REPLACE FINANCE WITH SALDO->A03VANT
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Compras(Entradas):"
      REPLACE QUANTID WITH SALDO->A03QCPA
      REPLACE FINANCE WITH SALDO->A03VCPA
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Vendas(Saidas):"
      REPLACE QUANTID WITH SALDO->A03QVDA
      REPLACE FINANCE WITH SALDO->A03VVDA
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Resultados:"
      REPLACE FINANCE WITH LUCROS-PREJUIZOS
      REPLACE VOPCOES WITH OPCOES
      REPLACE VVISTA WITH VISTA
      REPLACE VDAYTR WITH DTVISTA
      REPLACE VDAYTRO WITH DTOPCOES
      REPLACE VTERMO WITH TERMO
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      REPLACE HISTGE WITH "Saldo Atual:"
      REPLACE QUANTID WITH SALDO->A03QATU
      REPLACE FINANCE WITH SALDO->A03VATU
      ADIREG(0)
      REPLACE DATA WITH ATE
      REPLACE PAPEL WITH SALDO->A03CODP+"  "+SALDO->A03desc
      SELE 4
    ENDDO
    SELE 20
    GO TOP
    CTITLE="Movimento da Carteira "+_DATA
    If QUEMPED=1
      oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
      oReg:Set("PDFFilename",NOMEPDF)
      oReg:Close()
      PRINT oPrn NAME CTITLE TO TEMPDF1
      ENDPRINT
    Endif

    If PREV .And. QUEMPED=0

      REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTITLE FONT oFont1,oFont2,oFont3;
        PEN oPen1,oPen2;
        PREVIEW;
        FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
        CAPTION ctitle;

      Else

      REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTitle FONT oFont1,oFont2,oFont3;
        PEN oPen1,oPen2;
        FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
        CAPTION ctitle;

      Endif

      COLUMN TITLE "Histórico" DATA HISTGE FONT 2 SIZE 30 grid
      COLUMN TITLE "Quantidade" DATA QUANTID PICTURE "@ZE 999,999,999,999" RIGHT FONT 2 size 12 grid
      COLUMN TITLE "Financeiro" DATA FINANCE PICTURE "@ZE 999,999,999,999.99" RIGHT FONT 2 SIZE 12 grid
      COLUMN TITLE "Res.Vista"    DATA VVISTA PICTURE "@ZE 999,999,999,999.99"  RIGHT FONT 2 SIZE 12 grid
      COLUMN TITLE "Res.Opções"   DATA VOPCOES PICTURE "@ZE 999,999,999,999.99" RIGHT FONT 2 SIZE 12 grid
      COLUMN TITLE "Res.Termo"   DATA VTERMO PICTURE "@ZE 999,999,999,999.99" RIGHT FONT 2 SIZE 12 grid
      COLUMN TITLE "Res.DayTrade" DATA VDAYTR PICTURE "@ZE 999,999,999,999.99"  RIGHT FONT 2 SIZE 12 grid
      COLUMN TITLE "Res.DT-Opções" DATA VDAYTRO PICTURE "@ZE 999,999,999,999.99"  RIGHT FONT 2 SIZE 12 grid

      oRpt:cPageTotal:="Sub-Totais"
      oRpt:cGrandTotal:="Totais(Gerais)"
      oRpt:lpagetotal:=.F.

      GROUP ON PAPEL HEADER "Movimentação: "+oRpt:agroups[1]:cvalue FOOTER " " FONT 1 GRID

      oRpt:CELLVIEW()

      ENDREPORT

      ACTIVATE REPORT oRpT ON STARTPAGE IMPRIMECAB() FOR DATA>=DE .And. DATA<=ATE

      oFont1:End()
      oFont2:End()
      CLOSE DATABASE

      Erase &ARQ_PRN1
      Erase &ARQ_PRN2
      Erase &ARQ_PRN3
    Return

    Procedure DCONT(QUEMPED,NOMEPDF)
      local n
      LOCAL oFont1, oFont2
      PUBLIC oRpt
      DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-10
      DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-10
      SELE 2
      If .NOT. USEREDE(ARQ13,.F.,10,"LANC")
        If QUEMPED=0
          MsgStop("O arquivo LANC.CONTABEIS não está  disponível",WSIST)
        Endif
        CLOSE DATABASE
        Return
      Else
        SET INDEX TO &ind1301
      Endif
      DE=CTOD("01/01/"+STRZERO(VAL(_DATA),4))
      ATE=CTOD("31/12/"+STRZERO(VAL(_DATA),4))
      If !AUTOMATICO
        PERIODO()
        If !RETORNO
          CLOSE DATABASE
          Return
        Endif
      Endif
      SELE 2
      GO TOP
      If EOF()
        If QUEMPED=0
          MSGINFO("Não Existe Dados para Impressão",WSIST)
          CLOSE DATABASE
          Return
        Endif
      Endif
      CTITLE="Demonstrativo Contábil de "+DTOC(DE)+" a "+DTOC(ATE)
      DEFINE PEN oPen1 WIDTH 0.1
      DEFINE PEN oPen2 WIDTH 0.1

      If QUEMPED=1
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        oReg:Set("PDFFilename",NOMEPDF)
        oReg:Close()
        PRINT oPrn NAME CTITLE TO TEMPDF1
        ENDPRINT
      Endif

      If PREV .And. QUEMPED=0

        REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTITLE FONT oFont1,oFont2 PEN oPen1,Open2;
          PREVIEW;
          FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
          CAPTION ctitle;

        Else

        REPORT oRpt TITLE " "," "," "," ",TRIM(EMP1),CTitle FONT oFont1,oFont2;
          FOOTER "CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net","Emitido em "+DTOC(DAT_HOJE)+"                  Página: "+strzero(oRpt:nPage,3);
          CAPTION ctitle;

        Endif

        COLUMN TITLE "Data" DATA A13DATA FONT 2 SIZE 9 GRID
        COLUMN TITLE "Debita" DATA A13CDEB FONT 2 SIZE 15 GRID
        COLUMN TITLE "Credita" DATA A13CCRE FONT 2 SIZE 15 GRID
        COLUMN TITLE "Historico" DATA OEMTOANSI(A13HIST) FONT 2 GRID
        COLUMN TITLE "Valor" DATA A13VALR PICTURE "@E 999,999,999.99" FONT 2 SIZE 14 GRID

        GROUP ON MONTH(A13DATA) HEADER "Mes: "+strzero(VAL(oRpt:agroups[1]:cvalue),2)+"/"+_DATA FOOTER " " FONT 1

        oRpt:cPageTotal:="Sub-Total"
        oRpt:cGrandTotal := "Total Geral "
        oRpt:lpagetotal:=.F.

        oRpt:CELLVIEW()

        ENDREPORT

        ACTIVATE REPORT oRpT ON STARTPAGE IMPRIMECAB() FOR A13VALR#0 .And. A13DATA>=DE .And. A13DATA<=ATE

        oFont1:End()
        oFont2:End()
        CLOSE DATABASE
        Return nil

        CLOSE DATABASE
      Return

      Procedure INTSINA(QUEMPED,NOMEPDF)
        local n
        LOCAL oFont1, oFont2
        PUBLIC oRpt
        DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-10
        DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-10
        SELE 2
        If .NOT. USEREDE(ARQ13,.F.,10,"LANC")
          If QUEMPED=0
            MsgStop("O arquivo LANC.CONTABEIS não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1301
        Endif
        DE=CTOD("01/01/"+STRZERO(VAL(_DATA),4))
        ATE=CTOD("31/12/"+STRZERO(VAL(_DATA),4))
        If !AUTOMATICO
          PERIODO()
          If !RETORNO
            CLOSE DATABASE
            Return
          Endif
        Endif
        coutfile=space(120)
        cOutFile:=cGetdir("Pasta para Gravar o Arquivo",curdir())
        coutfile=alltrim(coutfile)+"\INT"+DTOS(DATE())+".TXT"
        DTLANC=CTOD("")
        DTLANC=OUTROS("Data dos Lancamentos: ",DTLANC,"@D")
        DTLANC1=DTLANC
        DTLANC=DTOC(DTLANC)
        DTLANC=LEFT(DTLANC,2)+SUBSTR(DTLANC,4,2)+RIGHT(DTLANC,4)
        If !MsgYesno( "Gerar arquivo "+coutfile+" ?","Selecione")
          CLOSE DATABASE
          Return
        Endif
        SELE 2
        GO TOP
        If EOF()
          MSGINFO("Não Existe Dados para integrar neste periodo",WSIST)
          CLOSE DATABASE
          Return
        Endif
        FNAME = coutfile
        ARQ = FCREATE (FNAME)
        Do While !EOF()
          If A13DATA<DE .OR. A13DATA>ATE
            SKIP
            LOOP
          Endif
          If EMPTY(A13CDEB) .And. EMPTY(A13CCRE)
            SKIP
            LOOP
          Endif
          LINHA:="00110"+DTLANC+"00"+LEFT(A13CDEB,4)+"999"+LEFT(A13CCRE,4)+"999"+STRZERO(A13VALR*100,17)+LEFT(OEMTOANSI(A13HIST),50)+LEFT(OEMTOANSI(A13HIST),50)+"982"+space(40)+CHR(13)+CHR(10)
          FWRITE (ARQ, LINHA, LEN (LINHA))
          SELE 2
          SKIP
        ENDDO
        FWRITE (ARQ,CHR(26),1)
        FCLOSE(ARQ)
        CLOSE DATABASE
        MSGINFO("Gerado arquivo: "+FNAME," ")
      Return

      Procedure CARTEIRANET()
        If AUTOMATICO
          EMP2="Verificando Carteira.Net"
          SET MESSAGE OF oWnd to emp2 CENTERED
        Endif
        SELE 29
        If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
          Return
        Else
          arqind=localcart+"empresa1.ind"
          set index to &arqind
        Endif
        SET INDEX TO
        GO TOP
        TEMNOVO=.F.
        Do While .NOT. EOF()
          SysRefresh()
          SELE 29
          If A01CALCATE=M->DAT_HOJE .And. REGLOCK(10)
            REPLACE A01CALCATE WITH CTOD("")
          Endif
          WMREG=RECNO()
          If A01VALID<DATE() .And. REGLOCK(10) .And. a01codnac<>"S"
            REPLACE A01INATI WITH .T.
            UNLOCK
          Endif
          If VAL(A01CODE)=0 .And. A01CODNAC="S"
            ARQ=LOCALCART+"EMPRESA1.IND"
            Erase &ARQ
            INDICES()
            SELE 29
            USE
            If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
              MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
              CLOSE DATABASE
              QUIT
            Else
              ARQ=LOCALCART+"EMPRESA1.IND"
              SET INDEX TO &ARQ
            Endif
            SET ORDER TO 1
            GO BOTT
            Do While val(a01code)=0 .And. !eof()
              SysRefresh()
              skip -1
            ENDDO
            WCODIG=VAL(A01CODE)+1
            GO WMREG
            REGLOCK(10)
            REPLACE A01CODE WITH STRZERO(WCODIG,6)
            REPLACE A01INATI WITH .F.
            _LOCAL=LOCALCART+A01CODE+"/"
            ARQTES=LOCALCART+A01CODE+"/"
            VERDIRETORIO(ARQTES)
            ARQTES=ARQTES+"NOTAS/"
            VERDIRETORIO(ARQTES)
          Endif
          If VAL(A01CODE)=0
            ARQ=LOCALCART+"EMPRESA1.IND"
            Erase &ARQ
            INDICES()
            SELE 29
            USE
            If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
              MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
              CLOSE DATABASE
              QUIT
            Else
              ARQ=LOCALCART+"EMPRESA1.IND"
              SET INDEX TO &ARQ
            Endif
            SET ORDER TO 1
            GO BOTT
            Do While VAL(A01CODE)=0 .And. !BOF()
              SysRefresh()
              skip -1
            ENDDO
            WCODIG=VAL(A01CODE)+1
            GO WMREG
            REGLOCK(10)
            REPLACE A01CODE WITH STRZERO(WCODIG,6)
            REPLACE A01DATA WITH STRZERO(YEAR(DATE()),4)+" "
            REPLACE A01INATI WITH .F.
            REPLACE A01VALID WITH CTOD("")
            REPLACE A01DIAVENC WITH DAY(DATE())
            If EMPTY(A01EMITIDO)
              REPLACE A01EMITIDO WITH DATE()+5
            Endif
            REPLACE A01VALID WITH A01EMITIDO
            REPLACE A01VALMES WITH 60.00
            REPLACE A01CORRET WITH 0
            cod1=a01code
            _LOCAL=LOCALCART+A01CODE+"/"
            ARQTES=TRIM(LOCALCART)+A01CODE+"/"
            VERDIRETORIO(ARQTES)
            _LOCAL=arqtes
            LOCALQTIPO="C"
            M->CALCATE=A01CALCATE
            If EMPTY(M->CALCATE)
              M->CALCATE=M->DAT_HOJE
            Endif
            _DATA=STRZERO(VAL(TRIM(A01DATA)),4)
            EMP1=A01NOME
            ARQTES1=ARQTES+"NOTASBOV/"
            VERDIRETORIO(ARQTES1)
            ARQTES1=ARQTES+"NOTASBMF/"
            VERDIRETORIO(ARQTES1)
            ARQTES1=ARQTES+"EXTRATOS/"
            VERDIRETORIO(ARQTES1)
            NOMEARQS(1)
            CLOSE DATABASE
            a01code=cod1
            _LOCAL=LOCALCART+A01CODE+"/"
            CRIARQC()
            INDICESC()
            CLOSE DATABASE
            ARQ=LOCALCART+STRZERO(VAL(a01code),6)+"/ATUALIZ.INC"
            ARQ=ARQTES+"ATUALIZ.INC"
            MZLINHA="Atualizado em "+DTOC(DATE())+" as "+TIME()+" Até: "+DTOC(M->CALCATE)+CHR(13)+CHR(10)
            MEMOWRIT(ARQ,MZLINHA)
            LOCALQTIPO="W"
            ARQ=LOCALCART+"EMPRESA1.IND"
            Erase &ARQ
            INDICES()
            CLOSE DATABASE
            ABREARQS()
            SELE 29
            GO WMREG
            ENVIAEMAIL("bene@exatus.net","Cliente Novo no Carteira.Net","Cliente: "+A01CODE+"-"+A01NOME,"","","carteiranet@exatus.net","rrsecyie")
            ENVIAEMAIL(lower(alltrim(a01email)),"Liberado - CarteiraNet - Sistema de Cálculo da Carteira de Investimentos","Caro Sr(a)."+A01NOME+CHR(13)+CHR(10)+CHR(13)+CHR(10)+"A utilização do sistema CarteiraNet - Sistema de Cálculo da Carteira de Investimento."+chr(13)+chr(10)+"foi liberado para utilização até o dia "+DTOC(A01VALID)+" por favor emita o boleto de cobrança com validade de 30 dias, acesse em: https://exatus.net/carteiranet/login.asp "+CHR(13)+CHR(10)+;
              CHR(13)+CHR(10)+"Obrigado"+CHR(13)+CHR(10)+"CarteiraNet"+CHR(13)+CHR(10)+"Exatus.Net - www.exatus.net","","","carteiranet@exatus.net","rrsecyie")
            TEMNOVO=.T.
          Endif
          SKIP
        ENDDO
        SELE 29
        GO TOP
        Do While .NOT. EOF()
          SysRefresh()
          COD1=A01CODE
          If VAL(A01CODE)=0
            SKIP
            LOOP
          Endif
          M->CALCATE=A01CALCATE
          If EMPTY(M->CALCATE)
            M->CALCATE=M->DAT_HOJE
          Endif
          If VAL(A01DATA)<>YEAR(M->CALCATE) .And. REGLOCK(10)
            REPLACE A01DATA WITH STRZERO(YEAR(M->CALCATE),4)
          Endif
          MWREC=RECNO()
          If A01CORRET=1 .And. A01CODNAC="S"
            ARQTES=LOCALCART+A01CODE+"/"
            DECLARE NOMAARQUIVO[ADIR(ARQTES+"*.ZIP")]
            ADIR(ARQTES+"*.ZIP",NOMAARQUIVO)
            CURSORWAIT()
            WFAZER=LEN(NOMAARQUIVO)
            FOR ZII=1 TO WFAZER
              SysRefresh()
              NARQ=ARQTES+NOMAARQUIVO[ZII]
              COMANDO=WCORRENTE+'brazip.exe EH "'+NARQ+'" "'+ARQTES+'"'
              CURSORWAIT()
              WAITRUN(COMANDO)
              CURSORWAIT()
              Erase &NARQ
            NEXT ZII
            FEZPROV=.F.
            NRAPROV=""
            FEZBDI=.F.
            NRABDI=CTOD("")
            RECBDI()
            GO MWREC
            FEZNOTASBMF=.F.
            NRABMF=CTOD("")
            GO MWREC
            FEZEXTRATO=.F.
            GO MWREC
            FEZNOTABOV=.F.
            qualpregao=ctod("")
            NRANOTA=QUALPREGAO
            GO MWREC
            REGLOCK(10)
            REPLACE A01CORRET WITH 0
            UNLOCK
            COMMIT
            If FEZBDI .OR. FEZNOTASBMF .OR. FEZEXTRATO.OR. FEZNOTABOV .OR. FEZPROV
              MQUEMRECEBE=LOWER(ALLTRIM(A01EMAIL))
              MTITULO="Rotinas Efetuadas"
              mtexto=IIF(FEZBDI,"Recebido BDI de "+DTOC(NRABDI)+CHR(13)+CHR(10),"")
              MTEXTO=MTEXTO+IIF(FEZNOTASBMF,"Recebido Notas BMF de "+DTOC(NRABMF)+CHR(13)+CHR(10),"")
              MTEXTO=MTEXTO+IIF(FEZEXTRATO,"Recebido Extratos de Contas Correntes"+CHR(13)+CHR(10),"")
              MTEXTO=MTEXTO+IIF(FEZNOTABOV,"Recebido Notas Bovespa de "+DTOC(NRANOTA)+CHR(13)+CHR(10),"")
              MTEXTO=MTEXTO+IIF(FEZPROV,"Recebido Proventos de "+NRAPROV+CHR(13)+CHR(10),"")
              MANEXO1=""
              MANEXO=""
              ENVIAEMAIL(TRIM(LOWER(MQUEMRECEBE))+";wanderlei@exatus.net",MTITULO,MTEXTO,MANEXO,MANEXO1,"bolsasnet@exatus.net","orckmwfd")
            Endif
          Endif
          SKIP
        ENDDO
        SELE 29
        GO TOP
        Do While .NOT. EOF()
          SysRefresh()
          MWREC=RECNO()
          If (A01CORRET=1 .OR. A01CORRET=2 .OR. A01CORRET=3) .And. A01CODNAC#"S" .And. REGLOCK(10)
            QFZ=A01CORRET
            ARQTES=LOCALCART+A01CODE+"/"
            M->CALCATE=A01CALCATE
            If EMPTY(M->CALCATE)
              M->CALCATE=M->DAT_HOJE
            Endif
            _LOCAL=arqtes
            _DATA=STRZERO(VAL(TRIM(A01DATA)),4)
            EMP1=A01NOME
            SELE 29
            If VAL(A01CODE)=0
              wreg=recno()
              ARQ=LOCALCART+"EMPRESA1.IND"
              Erase &ARQ
              INDICES()
              SELE 29
              USE
              If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
                MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
                CLOSE DATABASE
                QUIT
              Else
                ARQ=LOCALCART+"EMPRESA1.IND"
                SET INDEX TO &ARQ
              Endif
              SET ORDER TO 1
              GO BOTT
              Do While val(a01code)=0 .And. !bof()
                SysRefresh()
                skip -1
              ENDDO
              WCODIG=VAL(A01CODE)+1
              go wreg
              reglock(10)
              replace a01code with strzero(wcodig,6)
              _LOCAL=LOCALCART+A01CODE+"/"
              ARQTES=TRIM(LOCALCART)+A01CODE+"/"
              VERDIRETORIO(ARQTES)
              _LOCAL=arqtes
              LOCALQTIPO="C"
              M->CALCATE=A01CALCATE
              If EMPTY(M->CALCATE)
                M->CALCATE=M->DAT_HOJE
              Endif
              _DATA=STRZERO(VAL(TRIM(A01DATA)),4)
              EMP1=A01NOME
              ARQTES1=ARQTES+"NOTASBOV/"
              VERDIRETORIO(ARQTES1)
              ARQTES1=ARQTES+"NOTASBMF/"
              VERDIRETORIO(ARQTES1)
              ARQTES1=ARQTES+"EXTRATOS/"
              VERDIRETORIO(ARQTES1)
              NOMEARQS(1)
              CLOSE DATABASE
              SELE 29
              USE
              If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
                MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
                CLOSE DATABASE
                QUIT
              Else
                ARQ=LOCALCART+"EMPRESA1.IND"
                SET INDEX TO &ARQ
              Endif
              go wreg
              a01code=cod1
              _LOCAL=LOCALCART+A01CODE+"/"
            Endif
            sele 29
            _LOCAL=LOCALCART+A01CODE+"/"
            NOMEARQS(1)
            If QFZ=3
              CLOSE DATABASE
              SELE 1
              If USEREDE(ARQ03,.T.,10,"SALD")
                SET INDEX TO &ind0301
                reindex
                USE
              Endif
              If USEREDE(ARQ10,.T.,10,"DIARIO")
                SET INDEX TO &IND1001
                REINDEX
                USE
              Endif
              If USEREDE(ARQ11,.T.,10,"ITENS")
                SET INDEX TO &IND1101
                REINDEX
                USE
              Endif
              If USEREDE(ARQ12,.T.,10,"NOTAS")
                SET INDEX TO &IND1201
                REINDEX
                USE
              Endif
              QFZ=1
              CLOSE DATABASE
            Endif
            If QFZ=1
              CALCSAL2(1)
              CLOSE DATABASE
              ATBDI(1)
              CLOSE DATABASE
              IMPRES(0)
              CLOSE DATABASE
            ElseIf QFZ=2
              CLOSE DATABASE
            Endif
            CLOSE DATABASE
            ABREARQS()
            sele 29
            USE
            If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
              MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
              CLOSE DATABASE
              QUIT
            Else
              ARQ=LOCALCART+"EMPRESA1.IND"
              SET INDEX TO &ARQ
            Endif
            SELE 29
            GO MWREC
            ARQ=LOCALCART+STRZERO(VAL(a01code),6)+"/ATUALIZ.INC"
            MZLINHA="Atualizado em "+DTOC(DATE())+" as "+TIME()+" Até: "+DTOC(M->CALCATE)+CHR(13)+CHR(10)
            reglock(10)
            REPLACE A01CORRET WITH 0
            MEMOWRIT(ARQ,MZLINHA)
          Endif
          SELE 29
          SKIP
        ENDDO
        If AUTOMATICO
          EMP2=SPACE(40)
          SET MESSAGE OF oWnd to emp2 CENTERED
        Endif
        SELE 29
        SET ORDER TO 1
      Return

      Procedure RECBDI()
        NAOFEZBDI=.T.
        ARQTES=LOCALCART+A01CODE+"/"
        DECLARE NOMAARQUIVO[ADIR(ARQTES+"*.DAT")]
        ADIR(ARQTES+"*.DAT",NOMAARQUIVO)
        CURSORWAIT()
        WFAZER=LEN(NOMAARQUIVO)
        FOR ZII=1 TO WFAZER
          SysRefresh()
          FEZBDI=.T.
          NRABDI=CTOD("")
          AUTOPU(1,ARQTES+NOMAARQUIVO[ZII])
          ABREARQS()
          SELE 29
          If NAOFEZBDI=.F.
            GO TOP
            Do While .NOT. EOF()
              SysRefresh()
              If A01CORRET=0 .And. A01CODNAC<>"S" .And. !A01INATI .And. REGLOCK(10)
                REPLACE A01CORRET WITH 2
              Endif
              SKIP
            ENDDO
          Endif
          M->LOC=LOCALCART+"BDIS/"
          VERDIRETORIO(M->LOC)
          M->LOC=M->LOC+DTOS(NRABDI)+".BDI"
          MARQDE=ALLTRIM(ALLTRIM(ARQTES)+NOMAARQUIVO[ZII])
          LZCOPYFILE(MARQDE,M->LOC)
          M->LOC=ALLTRIM(Strtran(MARQDE,"/","\"))
          LOC=M->LOC
          Erase &LOC
        NEXT ZII
      Return

      Procedure CALCSAL2(QUEMPED)
        //  Cadastro de Notas *
        SELE 1
        If .NOT. USEREDE(ARQ12,.F.,10,"NOTA")
          MsgStop("O arquivo NOTAS não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1201
        Endif
        // Cadastro de Papeis (Saldo)*
        SELE 2
        If .NOT. USEREDE(ARQ03,.T.,10,"SALD")
          If QUEMPED=0
            MsgStop("O arquivo SALDOS não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind0301
        Endif
        // Movimento Diário *
        SELE 4
        If .NOT. USEREDE(ARQ10,.T.,10,"DIARIO")
          If QUEMPED=0
            MsgStop("O arquivo MOVTO DIARIO não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1001
        Endif
        // Movimento de Notas *
        SELE 5
        If .NOT. USEREDE(ARQ11,.T.,10,"ITENS")
          If QUEMPED=0
            MsgStop("O arquivo ITENS não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1101
        Endif
        // Integracao Contabil *
        SELE 6
        If .NOT. USEREDE(ARQ13,.T.,10,"LANC")
          If QUEMPED=0
            MsgStop("O arquivo LANC.CONTABEIS não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1301
        Endif
        SELE 7
        If !USEREDE(ARQ14,.F.,10,"PAPEIS")
          If QUEMPED=0
            MsgStop("O arquivo PAPEIS CUSTO MÉDIO não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1401
        Endif
        SELE 8
        If !USEREDE(ARQ04,.F.,10,"CONTAS")
          If QUEMPED=0
            MsgStop("O arquivo CONTAS não está  disponível",WSIST)
          Endif
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind0401
        Endif
        CURSORWAIT()
        //* CALCULA DESPESAS DAS NOTAS QUE VIEREM DA INTERNET
        SELE 1
        GO TOP
        Do While .NOT. EOF()
          SysRefresh()
          If A12FLAG="G" .And. REGLOCK(10)
            REPLACE A12FLAG WITH "F"
            Z12NOTA=A12NOTA
            TOTDEZP=A12TXLIQ+A12TXREG+A12CORRET+A12TXANA+A12EMOL
            TOTDEZOPCAO=A12TXTEOP
            Z12IRRF=A12IRRF
            SELE 5
            GO TOP
            TOTOP=0
            TOTVIS=0
            Do While !EOF()
              SysRefresh()
              If A11NOTA=Z12NOTA
                If VAL(PESQ(A11CODP,"7",1,"A03TIPO"))=2 .or. VAL(PESQ(A11CODP,"7",1,"A03TIPO"))=3
                  TOTOP=TOTOP+A11VALR
                Endif
                TOTVIS=TOTVIS+A11VALR
              Endif
              SKIP
            ENDDO
            GO TOP
            Do While !EOF()
              SysRefresh()
              If A11NOTA=Z12NOTA .And. REGLOCK(10)
                TP=PESQ(A11CODP,"7",1,"A03TIPO")
                If val(TP)=2 .or. val(TP)=3
                  REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)+(A11VALR/TOTOP*TOTDEZOPCAO)
                Else
                  REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)
                Endif
                If A11OPER="C" .OR. A11OPER="1"
                  If TP#"3"
                    REPLACE A11VLIQ WITH A11VALR+A11VCOR
                  Else
                    REPLACE A11VLIQ WITH A11VCOR
                  Endif
                Else
                  If TP#"3"
                    REPLACE A11VLIQ WITH A11VALR-A11VCOR
                  Else
                    REPLACE A11VLIQ WITH A11VCOR*-1
                  Endif
                Endif
                UNLOCK
              Endif
              SKIP
            ENDDO
            GO TOP
            LIQ=0
            Do While !EOF()
              SysRefresh()
              If A11NOTA=Z12NOTA .And. REGLOCK(10)
                If A11OPER="V"
                  LIQ=LIQ+A11VLIQ
                Else
                  LIQ=LIQ-A11VLIQ
                Endif
              Endif
              SKIP
            ENDDO
            LIQ=LIQ-Z12IRRF
            SELE 1
            If REGLOCK(10)
              REPLACE A12LIQ WITH LIQ
            Endif
            UNLOCK
          Endif
          SKIP
        ENDDO
        SELE 2
        GO TOP
        //*  ZERA MOVIMENTO **
        Do While ! EOF()
          SysRefresh()
          REPLACE a03qcpa WITH 0,a03vcpa WITH 0.00,a03qvda WITH 0,a03vvda WITH 0.00,;
            a03qtma WITH 0,a03vtma WITH 0.00,a03qtme WITH 0,a03vtme WITH 0.00,;
            a03qbma WITH 0,a03vbma WITH 0.00,a03qbme WITH 0,a03vbme WITH 0.00,;
            a03vluc WITH 0.00,a03vpre WITH 0.00,;
            a03qatu WITH a03qant,a03vatu WITH a03vant
          SKIP
        ENDDO
        //* ELIMINA MOVIMENTO DE NOTAS ANTERIORES **
        SELE 4
        ZAP
        //* BUSCA MOVIMENTO DE NOTAS PARA O CALCULO **
        SELE 5
        GO TOP
        Do While ! EOF()
          SysRefresh()
          If reglock(10)
            replace a11reg with recno()
          Endif
          If A11DATA>M->CALCATE
            SKIP
            LOOP
          Endif
          If VAL(A11NOTA)#0
            SELE 1
            SEEK ITENS->A11NOTA
            If EOF()
              SELE 5
              SKIP
              LOOP
            Endif
            SELE 5
          Endif
          If VAL(A11OPER)=0
            Z10OPER=IIF((A11OPER="C" .OR. A11OPER="1"),"1","2")
          Else
            Z10OPER=A11OPER
          Endif
          SELE 7
          SEEK ITENS->A11CODP
          SELE 5
          CDP=A11CODP
          ZDT=A11DATA
          If M->CALCATE>=A11DTVAL .And. PAPEIS->A03TIPO="3"
            CDP=PAPEIS->A03CODPO
            ZDT=A11DTVAL
          Endif
          CHAVE=CDP+DTOS(ZDT)+Z10OPER
          SELE 4
          SEEK CHAVE
          If EOF()
            ADIREG(0)
          Endif
          If PAPEIS->A03TIPO="3"
            REPLACE A10TERMO WITH "S"
            REPLACE A10DATA WITH ITENS->A11DTVAL
            REPLACE A10DTVAL WITH ITENS->A11DATA
          Endif
          REPLACE A10CODP WITH CDP
          If ITENS->A11REG=0
            REPLACE A10CODT WITH "03"
          ElseIf M->CALCATE<ITENS->A11DTVAL .And. PAPEIS->A03TIPO="3"
            REPLACE A10CODT WITH "03"
          Else
            REPLACE A10CODT WITH "00"
          Endif
          REPLACE A10NOTA WITH ITENS->A11NOTA
          REPLACE A10DATA WITH ZDT
          REPLACE A10OPER WITH Z10OPER
          REPLACE A10QTDE WITH A10QTDE+ITENS->A11QTDE
          REPLACE A10FLAG WITH ITENS->A11FLAG
          REPLACE A10HIST WITH ITENS->A11HIST
          REPLACE A10CODH WITH ITENS->A11CODH
          If VAL(ITENS->A11NOTA)#0
            If PAPEIS->A03TIPO#"3"
              REPLACE A10VALR WITH A10VALR+ITENS->A11VLIQ
            Else
              REPLACE A10VALR WITH A10VALR+ITENS->A11VALR+ITENS->A11VCOR
            Endif
          Else
            REPLACE A10VALR WITH A10VALR+ITENS->A11VALR
          Endif
          If PAPEIS->A03LMIL="S"
            REPLACE A10PUNI WITH A10VALR / (a10QTDE/1000)
          Else
            REPLACE A10PUNI WITH A10VALR / a10qtde
          Endif
          If PAPEIS->A03TIPO="2"
            REPLACE A10DTVAL WITH PAPEIS->A03DTOP
          Endif
          SELE 2
          SEEK DIARIO->A10CODP
          If EOF()
            ADIREG(0)
            REPLACE A03CODP WITH ITENS->A11CODP
            REPLACE A03DESC WITH PAPEIS->A03DESC
            REPLACE A03TIPO WITH PAPEIS->A03TIPO
            REPLACE A03CODT WITH PAPEIS->A03CODT
            REPLACE A03LMIL WITH PAPEIS->A03LMIL
          Endif
          SELE 5
          SKIP
        ENDDO
        //** VERIFICAR QUEM É DAYTRADE
        SELE 4
        GO TOP
        Do While .NOT. EOF()
          SysRefresh()
          If A10OPER="2" .OR. A10CODT="03" .OR. VAL(A10NOTA)=0
            SKIP
            LOOP
          Endif
          Z10REG=A10REG
          Z10CODP=A10CODP
          Z10CODT=A10CODT
          Z10NOTA=A10NOTA
          Z10DATA=A10DATA
          Z10OPER=A10OPER
          Z10QTDE=A10QTDE
          Z10VALR=A10VALR
          Z10PUNI=A10PUNI
          Z10DTVAL=A10DTVAL
          Z10RESU=0
          Z10TERMO=A10TERMO
          Z10HIST=A10HIST
          Z10CODH=A10CODH
          REGCOMPRA=RECNO()
          SKIP
          REGVENDA=RECNO()
          If !EOF() .And. Z10CODP=A10CODP .And. A10CODT#"03" .And. Z10DATA=A10DATA
            //** DAYTRADE COMPRA IGUAL A VENDA
            If Z10QTDE=A10QTDE
              REPLACE A10RESU WITH A10VALR-Z10VALR
              REPLACE A10CUSTO WITH A10VALR
              REPLACE A10CODT WITH "00"
              REPLACE A10FLAG WITH "D"
              REPLACE A10TERMO WITH Z10TERMO
              SKIP -1
              REPLACE A10CODT WITH "00"
              REPLACE A10FLAG WITH "D"
              SKIP 2
              LOOP
              //** DAYTRADE COMPRA MAIOR QUE A VENDA
            ElseIf Z10QTDE>A10QTDE
              QDIF=Z10QTDE-A10QTDE
              CUSTO=ROUND((Z10VALR/Z10QTDE*A10QTDE),2)
              VDIF=Z10VALR-CUSTO
              ADIREG(0)
              REPLACE A10CODP WITH Z10CODP
              REPLACE A10CODT WITH "03"
              REPLACE A10NOTA WITH Z10NOTA
              REPLACE A10DATA WITH Z10DATA
              REPLACE A10OPER WITH "1"
              REPLACE A10QTDE WITH QDIF
              REPLACE A10VALR WITH VDIF
              If PESQ(A10CODP,"7",1,"A03lmil")="S"
                REPLACE A10PUNI WITH A10valr / (A10qtde/1000)
              Else
                REPLACE A10PUNI WITH A10VALR / A10QTDE
              Endif
              REPLACE A10DTVAL WITH Z10DTVAL
              REPLACE A10FLAG WITH " "
              GO REGCOMPRA
              REPLACE A10CODT WITH "00"
              REPLACE A10QTDE WITH A10QTDE-QDIF
              REPLACE A10VALR WITH A10VALR-VDIF
              If PESQ(A10CODP,"7",1,"A03lmil")="S"
                REPLACE A10PUNI WITH A10valr / (A10qtde/1000)
              Else
                REPLACE A10PUNI WITH A10VALR / A10QTDE
              Endif
              Z10VALR=A10VALR
              REPLACE A10FLAG WITH "D"
              GO REGVENDA
              REPLACE A10RESU WITH A10VALR-Z10VALR
              REPLACE A10CUSTO WITH A10VALR
              REPLACE A10CODT WITH "00"
              REPLACE A10FLAG WITH "D"
              REPLACE A10TERMO WITH Z10TERMO
              SKIP
              LOOP
            Else
              //** DAYTRADE COMPRA MENOR QUE A VENDA
              QDIF=A10QTDE-Z10QTDE
              CUSTO=ROUND((A10VALR/A10QTDE*Z10QTDE),2)
              VDIF=A10VALR-CUSTO
              ADIREG(0)
              REPLACE A10CODP WITH Z10CODP
              REPLACE A10CODT WITH "03"
              REPLACE A10NOTA WITH Z10NOTA
              REPLACE A10DATA WITH Z10DATA
              REPLACE A10OPER WITH "2"
              REPLACE A10QTDE WITH QDIF
              REPLACE A10VALR WITH VDIF
              If PESQ(A10CODP,"7",1,"A03lmil")="S"
                REPLACE A10PUNI WITH A10valr / (A10qtde/1000)
              Else
                REPLACE A10PUNI WITH A10VALR / A10QTDE
              Endif
              REPLACE A10DTVAL WITH Z10DTVAL
              REPLACE A10FLAG WITH " "
              GO REGCOMPRA
              REPLACE A10CODT WITH "00"
              REPLACE A10FLAG WITH "D"
              Z10VALR=A10VALR
              GO REGVENDA
              REPLACE A10QTDE WITH A10QTDE-QDIF
              REPLACE A10VALR WITH A10VALR-VDIF
              If PESQ(A10CODP,"7",1,"A03lmil")="S"
                REPLACE A10PUNI WITH A10valr / (A10qtde/1000)
              Else
                REPLACE A10PUNI WITH A10VALR / A10QTDE
              Endif
              REPLACE A10RESU WITH A10VALR-Z10VALR
              REPLACE A10CUSTO WITH A10VALR
              REPLACE A10CODT WITH "00"
              REPLACE A10FLAG WITH "D"
              REPLACE A10TERMO WITH Z10TERMO
              SKIP
              LOOP
            Endif
          Endif
        ENDDO
        //** CALCULAR SALDOS
        calcsal1(1)
        //** BAIXAR AS OPCOES
        SELE 2
        GO TOP
        Z10REG=0
        Z10RESU=0
        Z10CUSTO=0
        Z10DTVAL=CTOD("")
        Z10TERMO=" "
        Z10HIST=""
        Z10CODH=0
        Do While !EOF()
          SysRefresh()
          If PESQ(A03CODP,"7",1,"A03TIPO")#"2"
            SKIP
            LOOP
          Endif
          If RIGHT(ALLTRIM(A03CODP),1)="E"
            CDOP=LEFT(ALLTRIM(A03CODP),LEN(ALLTRIM(A03CODP))-1)
            _qatu_   = A03qatu
            _vatu_   = A03vatu
            REG=RECNO()
            SELE 2
            SEEK CDOP
            PM=A03VATU/A03QATU*_QATU_
            PS=A03VATU/A03QATU
            QU=_QATU_
            If PM<0
              PM=PM*-1
            Endif
            GO REG
            Z10CODP=PESQ(CDOP,"7",1,"A03CODPO")
            Z10CODT=PESQ(CDOP,"7",1,"A03CODT")
            Z10NOTA="999999"
            Z10DATA=PESQ(CDOP,"7",1,"A03DTOP")
            Z10FLAG="E"
            If _QATU_<0
              Z10OPER="2"
              _QATU_=_QATU_*-1
              _VATU_=_VATU_*-1
            Else
              Z10OPER="1"
            Endif
            Z10QTDE=_QATU_
            Z10VALR=_VATU_+PM
            If PESQ(CDOP,"7",1,"A03lmil")="S"
              Z10PUNI = Z10valr / (Z10qtde/1000)
            Else
              Z10PUNI = Z10VALR / Z10QTDE
            Endif
            SELE 4
            If Z10QTDE<0
              Z10QTDE=Z10QTDE*-1
            Endif
            If Z10VALR<0
              Z10VALR=Z10VALR*-1
            Endif
            ADIREG(0)
            PUTREC()
            Z10VALR=_VATU_
            If Z10OPER="1"
              Z10OPER="2"
            Else
              Z10OPER="1"
            Endif
            Z10CODP=CDOP+"E"
            If Z10QTDE<0
              Z10QTDE=Z10QTDE*-1
            Endif
            If Z10VALR<0
              Z10VALR=Z10VALR*-1
            Endif
            ADIREG(0)
            PUTREC()
            Z10CODP=CDOP
            Z10QTDE=QU
            Z10VALR=QU*PS
            If Z10QTDE<0
              Z10QTDE=Z10QTDE*-1
            Endif
            If Z10VALR<0
              Z10VALR=Z10VALR*-1
            Endif
            If PESQ(Z10CODP,"7",1,"A03lmil")="S"
              Z10PUNI = Z10valr / (Z10qtde/1000)
            Else
              Z10PUNI = Z10VALR / Z10QTDE
            Endif
            ADIREG(0)
            PUTREC()
            SELE 2
            GO REG
          Endif
          SKIP
        ENDDO
        SELE 2
        GO TOP
        Do While !EOF()
          SysRefresh()
          REPLACE a03qcpa WITH 0,a03vcpa WITH 0.00,a03qvda WITH 0,a03vvda WITH 0.00,;
            a03qtma WITH 0,a03vtma WITH 0.00,a03qtme WITH 0,a03vtme WITH 0.00,;
            a03qbma WITH 0,a03vbma WITH 0.00,a03qbme WITH 0,a03vbme WITH 0.00,;
            a03vluc WITH 0.00,a03vpre WITH 0.00,;
            a03qatu WITH a03qant,a03vatu WITH a03vant
          SKIP
        ENDDO
        CALCSAL1(1)
        SELE 2
        GO TOP
        //* BAIXA OPÇÕES QUE VIRARAM PÓ
        Do While !EOF()
          SysRefresh()
          If PESQ(A03CODP,"7",1,"A03DTOP")>M->CALCATE .OR. PESQ(A03CODP,"7",1,"A03TIPO")#"2"
            SKIP
            LOOP
          Endif
          _qatu_   = A03qatu
          _vatu_   = A03vatu
          REG=RECNO()
          Z10CODP=A03CODP
          Z10CODT=A03CODT
          Z10NOTA="999999"
          Z10DATA=PESQ(Z10CODP,"7",1,"A03DTOP")
          Z10FLAG="P"
          If _QATU_<0
            Z10OPER="1"
            _QATU_=_QATU_*-1
            _VATU_=_VATU_*-1
          Else
            Z10OPER="2"
          Endif
          Z10QTDE=_QATU_
          Z10VALR=0.00
          Z10PUNI=0
          SELE 4
          ADIREG(0)
          PUTREC()
          SELE 2
          GO REG
          SKIP
        ENDDO
        SELE 2
        GO TOP
        Do While !EOF()
          SysRefresh()
          REPLACE a03qcpa WITH 0,a03vcpa WITH 0.00,a03qvda WITH 0,a03vvda WITH 0.00,;
            a03qtma WITH 0,a03vtma WITH 0.00,a03qtme WITH 0,a03vtme WITH 0.00,;
            a03qbma WITH 0,a03vbma WITH 0.00,a03qbme WITH 0,a03vbme WITH 0.00,;
            a03vluc WITH 0.00,a03vpre WITH 0.00,;
            a03qatu WITH a03qant,a03vatu WITH a03vant
          SKIP
        ENDDO
        CALCSAL1(0)
        //** LANCAMENTOS CONTABEIS
        SELE 6
        ZAP
        SELE 4
        GO TOP
        Do While ! EOF()
          SysRefresh()
          If A10QTDE=0 .And. REGLOCK(10)
            DELE
            SKIP
            LOOP
          Endif
          CPAVDA=""
          Z13HIST=""
          CTA="00"
          WMERCADO=PESQ(A10CODP,"7",1,"A03TIPO")
          WMER=WMERCADO
          If VAL(A10FLAG)=0
            If A10OPER="1"
              Z13HIST="CPA "
              If WMERCADO="1"
                CTA="02"
              ElseIf WMERCADO="2"
                CTA="01"
              ElseIf WMERCADO="3"
                CTA="03"
              Endif
            Else
              Z13HIST="VDA "
              If WMERCADO="1"
                CTA="05"
              ElseIf WMERCADO="2"
                CTA="04"
              ElseIf WMERCADO="3"
                CTA="06"
              Endif
            Endif
          ElseIf A10FLAG="1"
            Z13HIST="OUTROS "
            If A10OPER="1"
              CTA="08"
            Else
              CTA="09"
            Endif
          ElseIf A10FLAG="2"
            Z13HIST="SUBSCRICAO"
            CTA="07"
          ElseIf A10FLAG="3"
            Z13HIST="BONIFICACAO"
            CTA="10"
          ElseIf A10FLAG="4"
            Z13HIST="TRANSFERENCIA"
            If A10OPER="1"
              CTA="11"
            Else
              CTA="12"
            Endif
          ElseIf A10FLAG="5"
            Z13HIST="IMPLIT"
            CTA="13"
          Endif
          If A10CODH#0
            Z13HIST=A10HIST
          Endif
          If WMERCADO="1"
            WMERCADO="Vista "
          ElseIf WMERCADO="2"
            WMERCADO="Opções "
          ElseIf WMERCADO="3"
            WMERCADO="Termo "
          Endif
          If A10TERMO="S"
            WMERCADO="Termo "
          Endif
          Z13hist = WMERCADO+Z13HIST+ALLTRIM(TRANS(A10qtde,"@E 999,999,999,999"))+;
            " "+ALLTRIM(PESQ(A10CODP,"7",1,"A03desc"))+"("+ALLTRIM(A10CODP)+")"
          If A10FLAG="E"
            If A10OPER="1"
              CTA="26"
            Else
              CTA="27"
            Endif
            Z13HIST = "EXERCICIO DE OPÇÕES "+ALLTRIM(TRANS(A10qtde,"@E 999,999,999,999"))+;
              " "+ALLTRIM(PESQ(A10CODP,"7",1,"A03desc"))+"("+ALLTRIM(A10CODP)+")"
          Endif
          If A10CODH#0
            Z13HIST=A10HIST
          Endif
          If A10FLAG="E" .And. WMERCADO="Vista "
          Else
            SELE 6
            ADIREG(0)
            REPLACE A13DATA WITH DIARIO->A10DATA
            REPLACE A13VALR WITH DIARIO->A10VALR
            REPLACE A13HIST WITH ANSITOOEM(Z13HIST)
            REPLACE A13CDEB WITH PESQ(CTA,"8",1,"A04CPAH")
            REPLACE A13CCRE WITH PESQ(CTA,"8",1,"A04VDAH")
          Endif
          CTA="00"
          SELE 4
          If A10RESU#0
            If A10OPER="1"
              CPAVDA="CPA "
            Else
              CPAVDA="VDA "
            Endif
            If A10FLAG="D"
              CPAVDA=CPAVDA+" DT."
            Endif
            WMERCADO=WMER
            If WMERCADO="1"
              WMERCADO="Vista "
            ElseIf WMERCADO="2"
              WMERCADO="Opções "
            ElseIf WMERCADO="3"
              WMERCADO="Termo "
            Endif
            If A10TERMO="S"
              WMERCADO="Termo "
            Endif
            If A10RESU>0
              If WMER="1"
                CTA="15"
                If A10FLAG="D"
                  CTA="18"
                Endif
              ElseIf WMER="2"
                CTA="14"
                If A10FLAG="D"
                  CTA="17"
                Endif
              ElseIf WMER="3"
                CTA="16"
                If A10FLAG="D"
                  CTA="19"
                Endif
              Endif
              Z13hist = WMERCADO+"LUCRO "+CPAVDA+ALLTRIM(TRANS(A10qtde,"@E 999,999,999,999"))+;
                " "+ALLTRIM(PESQ(A10CODP,"7",1,"A03DESC"))+"("+ALLTRIM(A10CODP)+")"
            Else
              If WMER="1"
                CTA="21"
                If A10FLAG="D"
                  CTA="24"
                Endif
              ElseIf WMER="2"
                CTA="20"
                If A10FLAG="D"
                  CTA="23"
                Endif
              ElseIf WMER="3"
                CTA="22"
                If A10FLAG="D"
                  CTA="25"
                Endif
              Endif
              Z13hist = WMERCADO+"PREJ."+CPAVDA+ALLTRIM(TRANS(A10qtde,"@E 999,999,999,999"))+;
                " "+ALLTRIM(PESQ(A10CODP,"7",1,"A03DESC"))+"("+ALLTRIM(A10CODP)+")"
            Endif
            If A10FLAG="P"
              If A10OPER="1"
                CTA="28"
              Else
                CTA="29"
              Endif
              Z13HIST = "OPÇÕES NÃO EXERCIDAS "+ALLTRIM(TRANS(A10qtde,"@E 999,999,999,999"))+;
                " "+ALLTRIM(PESQ(A10CODP,"7",1,"A03desc"))+"("+ALLTRIM(A10CODP)+")"
            Endif
            If A10CODH#0
              If A10RESU=0
                Z13HIST=A10HIST
              Else
                Z13HIST=IIF(A10RESU>0,"LUCRO "+A10HIST,"PREJ."+A10HIST)
              Endif
            Endif
            SELE 6
            ADIREG(0)
            REPLACE A13DATA WITH DIARIO->A10DATA
            REPLACE A13CDEB WITH PESQ(CTA,"8",1,"A04CPAH")
            REPLACE A13CCRE WITH PESQ(CTA,"8",1,"A04VDAH")
            If DIARIO->A10RESU>0
              REPLACE A13VALR WITH DIARIO->A10RESU
            Else
              REPLACE A13VALR WITH DIARIO->A10RESU*-1
            Endif
            REPLACE A13HIST WITH ANSITOOEM(Z13HIST)
          Endif
          SELE 4
          SKIP
        ENDDO
        CURSORARROW()
      Return

      Procedure CALCSAL1(QZL)
        //* RECALCULANDO SALDOS **
        SELE 4
        GO TOP
        Do While .NOT. EOF()
          SysRefresh()
          If PESQ(DIARIO->A10CODP,"7",1,"A03TIPO")#"2" .And. QZL=1
            SKIP
            LOOP
          Endif
          SELE 2
          SEEK DIARIO->A10CODP
          If EOF()
            ADIREG(0)
            REPLACE A03CODP WITH DIARIO->A10CODP
          Endif
          _qatu_   = A03qatu
          _vatu_   = A03vatu
          _cust_   = 0.00
          _resu_   = 0.00
          _compras = 0.00
          _qcompras= 0.00
          _vendas  = 0.00
          _qvendas = 0.00
          _vlucro  = 0.00
          _vpreju  = 0.00
          SELE 4
          If A10FLAG="D"
            If A10OPER="1"
              _QCOMPRAS = A10QTDE
              _COMPRAS  = A10VALR
            Else
              _QVENDAS = A10qtde
              _VENDAS  = A10valr
            Endif
            If A10RESU # 0
              If A10RESU > 0.0
                _vlucro = _vlucro + A10RESU
              Else
                _vpreju = _vpreju + (A10RESU * -1)
              Endif
            Endif
            SELE 2
            REPLACE A03QCPA WITH A03QCPA+_QCOMPRAS
            REPLACE A03VCPA WITH A03VCPA+_COMPRAS
            REPLACE A03QVDA WITH A03QVDA+_QVENDAS
            REPLACE A03VVDA WITH A03VVDA+_VENDAS
            REPLACE A03VLUC WITH A03VLUC+_VLUCRO
            REPLACE A03VPRE WITH A03VPRE+_VPREJU
            SELE 4
            SKIP
            LOOP
          Endif
          If A10oper = "1"  // Compra
            If A10FLAG<>"5"
              If _QATU_<0 .And. (A10QTDE+_QATU_)<=0
                _pmed_ = (_vatu_/_qatu_)
                _vvda_ = A10qtde * ABS(_pmed_)
                _cust_ = A10valr
                _resu_ = _vvda_ - _cust_
                _qatu_ = _qatu_ + A10qtde
                _vatu_ = _vatu_ + _vvda_
              ElseIf _qatu_ >= 0.00
                _cust_ = 0.00
                _resu_ = 0.00
                _qatu_ = _qatu_ + A10qtde
                _vatu_ = _vatu_ + A10valr
              ElseIf (A10qtde + _qatu_) > 0.00
                _pmed_ = (A10valr/A10qtde)
                _cust_ = ABS(_qatu_) * _pmed_
                _resu_ = ABS(_vatu_) - _cust_
                _qatu_ = _qatu_ + A10qtde
                _vatu_ = _qatu_ * _pmed_
              Endif
            Endif
            If A10FLAG="5"
              REPLACE A10RESU WITH 0
              REPLACE A10CUSTO WITH 0
              _QATU_ = _QATU_ + A10QTDE
              _VATU_ = _VATU_
              _RESU_ = 0
            Else
              _QCOMPRAS = _QCOMPRAS+ A10QTDE
              _COMPRAS  = _COMPRAS + A10VALR
              If _resu_ # 0
                REPLACE A10RESU WITH _RESU_
                REPLACE A10CUSTO WITH _CUST_
                If _resu_ > 0.00
                  _vlucro = _vlucro + _resu_
                Else
                  _vpreju = _vpreju + (_resu_ * -1)
                Endif
              Endif
            Endif
          ElseIf A10oper = "2"  // Venda
            If A10FLAG<>"5"
              If _qatu_ <= 0.00
                _cust_ = 0.00
                _resu_ = 0.00
                _qatu_ = _qatu_ - A10qtde
                _vatu_ = _vatu_ - A10valr
              ElseIf A10qtde > _qatu_
                _pmed_ = (A10valr/A10qtde)
                _vvda_ = ABS(_qatu_) * _pmed_
                _cust_ = _vatu_
                _resu_ = _vvda_ - _cust_
                _qatu_ = _qatu_ - A10qtde // ESTAVA POSITIVO.
                _vatu_ = _qatu_ * _pmed_
              Else
                _pmed_ = _vatu_/_qatu_
                _cust_ = A10qtde * _pmed_
                _resu_ = A10valr - _cust_
                _qatu_ = _qatu_ - A10qtde
                _vatu_ = _vatu_ - _cust_
              Endif
            Endif
            If A10FLAG="5"
              REPLACE A10RESU WITH 0
              REPLACE A10CUSTO WITH 0
              _QATU_ = _QATU_ - A10QTDE
              _VATU_ = _VATU_
              _RESU_ = 0
            Else
              _QVENDAS = _QVENDAS + A10qtde
              _VENDAS  = _VENDAS  + A10valr
              If _resu_ # 0.00
                REPLACE A10RESU WITH _RESU_
                REPLACE A10CUSTO WITH _CUST_
                If _resu_ > 0.00
                  _vlucro = _vlucro + _resu_
                Else
                  _vpreju = _vpreju +(_resu_*-1)
                Endif
              Endif
            Endif
          Endif
          SELE 2
          REPLACE A03QCPA WITH A03QCPA+_QCOMPRAS
          REPLACE A03VCPA WITH A03VCPA+_COMPRAS
          REPLACE A03QVDA WITH A03QVDA+_QVENDAS
          REPLACE A03VVDA WITH A03VVDA+_VENDAS
          REPLACE A03VLUC WITH A03VLUC+_VLUCRO
          REPLACE A03VPRE WITH A03VPRE+_VPREJU
          REPLACE A03qatu WITH _qatu_
          REPLACE A03vatu WITH _vatu_
          SELE 4
          SKIP
        ENDDO
        SELE 4
      Return

      Procedure CADIMP()
        CLOSE DATABASE
        SELE 1
        If !USEREDE(ARQ05,.F.,10,"IRRF")
          Return
        Else
          SET INDEX TO &IND0501
        Endif
        BrowseC(12,"Cadastro de IRs","IRRF", { || buscaIR() },{ ||NCADIR("I") },{ ||NCADIR("A") })
        CLOSE DATABASE
      Return

      Procedure BUSCAIR()
        TEXTO="MES/ANO:"
        VARIAVEL=SPACE(6)
        MASCARA="@R 99-9999"
        REG=RECNO()
        Do While .T.
          SysRefresh()
          If !PROCURA()
            EXIT
          Endif
          CHAVE=VARIAVEL
          If EMPTY(CHAVE)
            EXIT
          Endif
          CHAVE=STRZERO(VAL(CHAVE),6)
          SEEK RIGHT(CHAVE,4)+LEFT(CHAVE,2)
          If EOF() .OR. DELE()
            MSGALERT("Não Encontrado !",WSIST)
            GO REG
            VARIAVEL=SPACE(6)
          Else
            EXIT
          Endif
        ENDDO
      Return .T.

      Procedure NCADIR(FS)
        LOCAL oDlg1,oBrush
        PUBLIC F
        F=FS
        DECLARE _VAR1[FCOUNT()]
        AFIELDS(_VAR1)
        If F="I"
          GO BOTT
          SKIP
        Endif
        FOR t = 1 TO FCOUNT()
          SysRefresh()
          _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
          _var5 = _var1 [t]
          &_var4 = &_var5
        NEXT t
        RELEASE _VAR1
        MFONTES()
        DEFINE DIALOG oDlg1 FROM 1,1 TO 32,68 TITLE "Dados para Calculo do IR" BRUSH oBrushTela STYLE WS_CAPTION
        @   0.07,0.5 say " " BOX RAISED COLOR CLR_BLACK,nRGB(248,248,248) SIZE 260,210
        @   0.87,01 SAY "Mes/Ano:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   0.87,20 SAY "Normais" FONT oFontsayB COLOR CLR_BLACK,nRGB(248,248,248)
        @   0.87,33.4 SAY "Day-Trade" FONT oFontsayB COLOR CLR_BLACK,nRGB(248,248,248) SIZE 35,11
        @   1.74,01 SAY "a Vista - Ouro:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,01 SAY "a Vista - Ouro ativo Fora de Bolsa:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 180,11
        @   3.48,01 SAY "Opções Ouro:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   4.35,01 SAY "Opções - Ouro Fora de Bolsa:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   5.22,01 SAY "Opções - Outros:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   6.09,01 SAY "Futuro - Dolar EUA:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   6.96,01 SAY "Futuro - Indices:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   7.83,01 SAY "Futuro - Juros:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   8.70,01 SAY "Futuro - Outros:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   9.57,01 SAY "Termo - Ouro:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

        @  10.8,01 SAY "Base Negativa Anterior:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @  12.18,01 SAY "IR Anteriores:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

        @ 13.05,01 SAY "Imposto Pago:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

        @ 01,05 GET Z05mesano FONT oFontget SIZE 32,11 PICTURE "@R 99-9999"
        @ 02,15 GET Z05C02 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 02,25 GET Z05CD02 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 03,15 GET Z05C03 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 03,25 GET Z05CD03 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 04,15 GET Z05C05 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 04,25 GET Z05CD05 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 05,15 GET Z05C06 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 05,25 GET Z05CD06 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 06,15 GET Z05C07 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 06,25 GET Z05CD07 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 07,15 GET Z05C08 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 07,25 GET Z05CD08 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 08,15 GET Z05C09 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 08,25 GET Z05CD09 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 09,15 GET Z05C10 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 09,25 GET Z05CD10 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 10,15 GET Z05C11 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 10,25 GET Z05CD11 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 11,15 GET Z05C13 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 11,25 GET Z05CD13 PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT

        @ 12.5,15 GET Z05PREJA PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 12.5,25 GET Z05PREJAD PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT
        @ 14,15 GET Z05IRFONTE PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT

        @ 15,15 GET Z05IRPAGO PICTURE "@E 999,999,999.99" FONT oFontget SIZE 60,11 RIGHT

        @ 16.7,18 SBUTTON NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST PROMPT "&Confirma" TEXT ON_CENTER OF oDlg1 SIZE 40,12 ACTION (IIF(F="I",ADIREG(0),REGLOCK(10)),PUTREC(),oDlg1:END()) FONT oFontbut
        @ 16.7,28 SBUTTON NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST PROMPT "C&ancela" TEXT ON_CENTER OF oDlg1 SIZE 40,12 ACTION (oDlg1:End()) FONT oFontbut
        odlg1:lhelpicon:=.f.
        ACTIVATE DIALOG oDlg1 CENTERED
      Return

      Procedure IMPRES(QUEMPED)
        LOCAL oFont1,oFont2,oPen1,oPen2
        PUBLIC oRpt
        DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-7
        DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-7
        DEFINE FONT oFont3 NAME "VERDANA" SIZE 0,-7
        DEFINE PEN oPen1 WIDTH 0.1
        DEFINE PEN oPen2 WIDTH 0.1
        CURSORWAIT()
        CLOSE DATABASE
        If QUEMPED=0
          SELE 1
          If USEREDE(ARQ12,.F.,10,"NOTAS")
            SET INDEX TO &IND1201
          Else
            CLOSE DATABASE
            Return
          Endif
          SELE 2
          If USEREDE(ARQ03,.F.,10,"SALDO")
            SET INDEX TO &ind0301
          Else
            CLOSE DATABASE
            Return
          Endif
          SELE 4
          If .NOT. USEREDE(ARQ10,.F.,10,"DIARIO")
            CLOSE DATABASE
            Return
          Else
            SET INDEX TO &ind1001
          Endif
        Endif
        SELE 3
        If USEREDE(ARQ05,.F.,10,"IRRF")
          SET INDEX TO &ind0501
        Else
          If QUEMPED=1
            MSGSTOP("Não foi possível abrir arquivo de Carteira "+ARQ05,"Carteira")
          Endif
          CLOSE DATABASE
          Return
        Endif
        If QUEMPED=0
          If EOF()
            CLOSE DATABASE
            Return
          Endif
          Do While .T.
            SysRefresh()
            M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
            ARQ_PRN=_LOCAL+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
            ARQ_PRN1:=ARQ_PRN+".DAT"
            ARQ_PRN2:=ARQ_PRN+".KEY"
            ARQ_PRN3:=ARQ_PRN+".TMP"
            If !FILE(ARQ_PRN1)
              EXIT
            Endif
          ENDDO
          SELE 20
          MAT_STRU:={}
          AADD(MAT_STRU,{"MESANO","C",6,0})
          AADD(MAT_STRU,{"RESULTA","N",19,2})
          AADD(MAT_STRU,{"VOPCOES","N",19,2})
          AADD(MAT_STRU,{"VVISTA","N",19,2})
          AADD(MAT_STRU,{"VTERMO","N",19,2})
          AADD(MAT_STRU,{"VDAYTR","N",19,2})
          AADD(MAT_STRU,{"VDAYTRO","N",19,2})
          AADD(MAT_STRU,{"IRDAYT","N",19,2})
          AADD(MAT_STRU,{"IRANTE","N",19,2})
          AADD(MAT_STRU,{"VENDAS","N",19,2})
          DBCREATE(ARQ_PRN1,MAT_STRU)
          USE
          SELE 20
          If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
            CLOSE DATABASE
            Erase &ARQ_PRN1
            Return
          Endif
          INDEX ON MESANO TO &ARQ_PRN2
          SELE 4
          GO TOP
          CURSORWAIT()
          Do While .NOT. EOF()
            SysRefresh()
            SELE 2
            SEEK DIARIO->A10CODP
            WMESANO=STRZERO(MONTH(DIARIO->A10DATA),2)+STRZERO(YEAR(DIARIO->A10DATA),4)
            SELE 20
            SEEK WMESANO
            If EOF()
              ADIREG(0)
            Endif
            REPLACE MESANO WITH WMESANO
            If (DIARIO->A10OPER="V" .OR. DIARIO->A10OPER="2") .And. VAL(DIARIO->A10NOTA)#0
              REPLACE VENDAS WITH VENDAS+DIARIO->A10VALR
            Endif
            REPLACE RESULTA WITH RESULTA+DIARIO->A10RESU
            If SALDO->A03TIPO="1" .And. DIARIO->A10FLAG<>"D"
              REPLACE VVISTA WITH VVISTA+DIARIO->A10RESU
            ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
              REPLACE VDAYTR WITH VDAYTR+DIARIO->A10RESU
            ElseIf SALDO->A03TIPO="1" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
              REPLACE VTERMO WITH VTERMO+DIARIO->A10RESU
            ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG<>"D"
              REPLACE VOPCOES WITH VOPCOES+DIARIO->A10RESU
            ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10FLAG="D" .And. DIARIO->A10TERMO#"S"
              REPLACE VDAYTRO WITH VDAYTRO+DIARIO->A10RESU
            ElseIf SALDO->A03TIPO="2" .And. DIARIO->A10TERMO="S" .And. DIARIO->A10FLAG="D"
              REPLACE VTERMO WITH VTERMO+DIARIO->A10RESU
            Endif
            SELE 4
            SKIP
          ENDDO
          SELE 1
          GO TOP
          Do While .NOT. EOF()
            SysRefresh()
            WMESANO=STRZERO(MONTH(NOTAS->A12DATA),2)+STRZERO(YEAR(NOTAS->A12DATA),4)
            SELE 20
            SEEK WMESANO
            If EOF()
              ADIREG(0)
            Endif
            REPLACE MESANO WITH WMESANO
            REPLACE IRDAYT WITH IRDAYT+NOTAS->a12IRDay
            REPLACE IRANTE WITH IRANTE+NOTAS->A12IRRF+NOTAS->A12OUTRAS
            SELE 1
            SKIP
          ENDDO
          SELE 20
          GO TOP
          If !EOF()
            MM1=CTOD("01/"+LEFT(MESANO,2)+"/"+RIGHT(MESANO,4))-15
          Else
            MM1=CTOD("01/"+STRZERO(MONTH(M->CALCATE))+"/"+STRZERO(YEAR(M->CALCATE),4))
          Endif
          MM=STRZERO(MONTH(MM1),2)+STRZERO(YEAR(MM1),4)
          MMULT=STRZERO(MONTH(M->CALCATE),2)+STRZERO(YEAR(M->CALCATE),4)
          MEZES=0
          SELE 3
          Do While MM<>MMULT .And. MEZES<300
            SysRefresh()
            MEZES+=1
            SEEK RIGHT(MM,4)+LEFT(MM,2)
            If EOF()
              ADIREG(0)
              REPLACE a05mesano WITH MM
            Endif
            MM1=CTOD("01/"+LEFT(MM,2)+"/"+RIGHT(MM,4))+34
            MM=STRZERO(MONTH(MM1),2)+STRZERO(YEAR(MM1),4)
          ENDDO
          SELE 20
          GO TOP
          Do While .NOT. EOF()
            SysRefresh()
            SELE 3
            SEEK RIGHT(TMP->MESANO,4)+LEFT(TMP->MESANO,2)
            If EOF()
              ADIREG(0)
              REPLACE A05MESANO WITH TMP->MESANO
            Endif
            If REGLOCK(10)
              REPLACE a05c01 WITH TMP->VVISTA
              REPLACE a05c04 WITH TMP->VOPCOES
              REPLACE a05c12 WITH TMP->VTERMO
              REPLACE a05cD01 WITH TMP->VDAYTR
              REPLACE a05cD04 WITH TMP->VDAYTRO
              REPLACE A05IRDAYME WITH TMP->IRDAYT
              REPLACE A05IRDANT WITH TMP->IRANTE
            Endif
            SELE 20
            SKIP
          ENDDO
          SELE 3
          GO TOP
          SKIP
          Do While !EOF()
            SysRefresh()
            WMESANO=LEFT(A05MESANO,2)+RIGHT(A05MESANO,4)
            SELE 20
            SEEK WMESANO
            WVENDA=VENDAS
            SELE 3
            RLIQAVISTA=a05c02+a05c03+a05c04+a05c05+a05c06+a05c07+a05c08+a05c09+a05c10+a05c11+a05c12+a05c13
            RLIQDAYTRA=a05cD02+a05cD03+a05cD04+a05cD05+a05cD06+a05cD07+a05cD08+a05cD09+a05cD10+a05cD11+a05cD12+a05cD13
            If WVENDA>=20000
              RLIQAVISTA=RLIQAVISTA+A05C01
              RLIQDAYTRA=RLIQDAYTRA+A05CD01
            Endif
            If wvenda<20000 .And. A05C01<0
              RLIQAVISTA=RLIQAVISTA+A05C01
            Endif
            If WVENDA<20000 .And. A05CD01<0
              RLIQDAYTRA=RLIQDAYTRA+A05CD01
            Endif
            MM1=CTOD("01/"+LEFT(a05mesano,2)+"/"+RIGHT(a05mesano,4))-15
            MM=STRZERO(MONTH(MM1),2)+STRZERO(YEAR(MM1),4)
            MM2=RIGHT(MM,4)+LEFT(MM,2)
            REGAGORA=RECNO()
            PREJANT=PESQ(MM2,"3",1,"a05PREJA")
            PREJANTD=PESQ(MM2,"3",1,"a05PREJAD")
            IRACOMP=PESQ(MM2,"3",1,"A05IRFONTE")
            VENDAS=PESQ(MM2,"20",1,"VENDAS")
            IRMENOR=PESQ(MM2,"3",1,"A05IMPAPAG")
            If IRMENOR>10
              IRMENOR=0
            Endif
            GO REGAGORA
            REGLOCK(10)
            REPLACE A05PANTV WITH PREJANT
            REPLACE A05PANTO WITH PREJANTD
            REPLACE A05IRDAYAN WITH IRACOMP
            BASEV=RLIQAVISTA-PREJANT
            BASED=RLIQDAYTRA-PREJANTD
            If BASEV<=0
              REPLACE A05PREJA WITH BASEV*-1
              BASEV=0
            Else
              REPLACE A05PREJA WITH 0
            Endif
            If BASED<=0
              REPLACE A05PREJAD WITH BASED*-1
              BASED=0
            Else
              REPLACE A05PREJAD WITH 0
            Endif
            If VAL(RIGHT(a05mesano,4))>2005
              REPLACE A05ALIQ WITH 15
              REPLACE A05ALIQD WITH 20
            Endif
            REPLACE A05IMP WITH BASEV*A05ALIQ/100
            REPLACE A05IMPD WITH BASED*A05ALIQD/100
            REPLACE A05IMPDEVI WITH A05IMP+A05IMPD
            IRAPAGAR=A05IMPDEVI-A05IRDAYME-A05IRDAYAN-A05IRDANT
            If IRAPAGAR<=0
              REPLACE A05IRFONTE WITH IRAPAGAR*-1
              IRAPAGAR=0
            Else
              REPLACE A05IRFONTE WITH 0
            Endif
            REPLACE A05IMPAPAG WITH IRAPAGAR+IRMENOR
            SKIP
          ENDDO
          CLOSE DATABASE
          Erase &ARQ_PRN1
          Erase &ARQ_PRN2
          Erase &ARQ_PRN3
          Return
        Endif
        SELE 3
        GO TOP
        CURSORWAIT()
        If QUEMPED=1 .OR. QUEMPED=2
          CTITLE="Resumo de Apuração de ganhos renda Variável - Ano: "+_DATA
          DEFINE PEN oPen1  WIDTH 1
          DEFINE PEN oPen2  WIDTH 2
          If QUEMPED=2
            oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
            oReg:Set("PDFFilename",NOMEPDF)
            oReg:Close()
            PRINT oPrn NAME CTITLE TO TEMPDF1
            ENDPRINT
            PRINT oPrn NAME CTITLE
          Else
            PRINT oPrn NAME CTITLE PREVIEW
          Endif
          DEFINE FONT oFont08 NAME "VERDANA" BOLD  SIZE 0,-5  OF oPrn
          DEFINE FONT oFont13 NAME "VERDANA" BOLD  SIZE 0,-7  OF oPrn
          DEFINE FONT oFont15 NAME "VERDANA" BOLD  SIZE 0,-6  OF oPrn
          DEFINE FONT oFont10 NAME "courier new" SIZE 0,-6 OF oPrn
          Do While !EOF()
            SysRefresh()
            If VAL(RIGHT(a05mesano,4))=YEAR(M->CALCATE)
              PAGE
              RLIQAVISTA=a05c01+a05c02+a05c03+a05c04+a05c05+a05c06+a05c07+a05c08+a05c09+a05c10+a05c11+a05c12+a05c13
              RLIQDAYTRA=a05cD01+a05cD02+a05cD03+a05cD04+a05cD05+a05cD06+a05cD07+a05cD08+a05cD09+a05cD10+a05cD11+a05cD12+a05cD13
              BASEV=RLIQAVISTA-A05PANTV
              BASED=RLIQDAYTRA-A05PANTO
              If BASEV<=0
                BASEV=0
              Endif
              If BASED<=0
                BASEV=0
              Endif
              Caja(2.8,2.7,4.7,18,oPrn,0,,oPen2)
              oPrn:cmsay(3,3,"RESUMO DE APURAÇÃO DE GANHOS - RENDA VARIAVEL",oFont13)
              oPrn:cmsay(3.3,3,"IMPOSTO DE RENDA - PESSOA FISICA",oFont13)
              oPrn:cmsay(3.6,3,"EXERCICIO: "+STRZERO(VAL(RIGHT(A05MESANO,4))+1,4),oFont13)
              oPrn:cmsay(3.9,3,"Ano Calendário: "+RIGHT(A05MESANO,4),oFont13)
              oprn:cmsay(4.2,3,"GANHOS LIQUIDOS OU PERDAS Ref: "+LEFT(A05MESANO,2)+"/"+RIGHT(A05MESANO,4)+"  "+OEMTOANSI(EMP1),oFont13)
              Caja(4.9,2.7,5.4,18,oPrn,0,,oPen2)
              oprn:cmsay(5,3,"Tipo de Mercado/Ativo",oFont13)
              oprn:cmsay(5,10,"OPERAÇÕES COMUNS",oFont13)
              oprn:cmsay(5,15,"OPERAÇÕES DAY TRADE",oFont13)
              Caja(5.4,2.7,15,18,oPrn,0,,oPen2)
              oprn:cmsay(05.5,3,"Mercado à vista - Ações",oFONT08)
              oprn:cmsay(06.0,3,"Mercado à vista - Ouro",oFONT08)
              oprn:cmsay(06.5,3,"Mercado à vista - Ouro-Ativo Financ.Fora de Bolsa",oFONT08)
              oprn:cmsay(07.0,3,"Mercado de Opções - Ações",oFONT08)
              oprn:cmsay(07.5,3,"Mercado de Opções - Ouro",oFONT08)
              oprn:cmsay(08.0,3,"Mercado de Opções - Fora de Bolsa",oFONT08)
              oprn:cmsay(08.5,3,"Mercado de Opções - Outros",oFONT08)
              oprn:cmsay(09.0,3,"Mercado Futuro - Dolar dos EUA",oFONT08)
              oprn:cmsay(09.5,3,"Mercado Futuro - Indices",oFONT08)
              oprn:cmsay(10.0,3,"Mercado Futuro - Juros",oFONT08)
              oprn:cmsay(10.5,3,"Mercado Futuro - Outros",oFONT08)
              oprn:cmsay(11.0,3,"Mercado a Termo -Ações/Ouro",oFONT08)
              oprn:cmsay(11.5,3,"Mercado a Termo - Outros",oFONT08)
              oprn:cmsay(12.0,3,"RESULTADO LÍQUIDO DO MES",ofont08)
              oprn:cmsay(12.5,3,"Resultado negativo até o mes anterior",oFONT08)
              oprn:cmsay(13.0,3,"BASE DE CÁLCULO DO IMPOSTO",oFONT08)
              oPrn:cmsay(13.5,3,"Prejuizo a compensar",oFONT08)
              oPrn:cmsay(14.0,3,"Aliquota do Imposto",oFont08)
              oPrn:cmsay(14.5,3,"IMPOSTO DEVIDO",oFont08)
              oPrn:cmsay(05.5,10,TRANSFORM(a05c01,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(06.0,10,TRANSFORM(a05c02,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(06.5,10,TRANSFORM(a05c03,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(07.0,10,TRANSFORM(a05c04,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(07.5,10,TRANSFORM(a05c05,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(08.0,10,TRANSFORM(a05c06,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(08.5,10,TRANSFORM(a05c07,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(09.0,10,TRANSFORM(a05c08,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(09.5,10,TRANSFORM(a05c09,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(10.0,10,TRANSFORM(a05c10,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(10.5,10,TRANSFORM(a05c11,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(11.0,10,TRANSFORM(a05c12,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(11.5,10,TRANSFORM(a05c13,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(12.0,10,TRANSFORM(RLIQAVISTA,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(12.5,10,TRANSFORM(A05PANTV,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(13.0,10,TRANSFORM(BASEV,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(13.5,10,TRANSFORM(a05PREJA,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(14.0,10,"          15%",oFont10)
              oprn:cmsay(14.5,10,TRANSFORM(a05IMP,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(05.5,15,TRANSFORM(a05cD01,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(06.0,15,TRANSFORM(a05cD02,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(06.5,15,TRANSFORM(a05cD03,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(07.0,15,TRANSFORM(a05cD04,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(07.5,15,TRANSFORM(a05cD05,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(08.0,15,TRANSFORM(a05cD06,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(08.5,15,TRANSFORM(a05cD07,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(09.0,15,TRANSFORM(a05cD08,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(09.5,15,TRANSFORM(a05cD09,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(10.0,15,TRANSFORM(a05cD10,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(10.5,15,TRANSFORM(a05cD11,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(11.0,15,TRANSFORM(a05cD12,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(11.5,15,TRANSFORM(a05cD13,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(12.0,15,TRANSFORM(RLIQDAYTRA,"@ZE 999,999,999.99"),oFont10)
              oprn:cmsay(12.5,15,TRANSFORM(A05PANTO,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(13.0,15,TRANSFORM(BASED,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(13.5,15,TRANSFORM(a05PREJAD,"@ZE 999,999,999.99"),oFont10)
              oPrn:cmsay(14.0,15,"          20%",oFont10)
              oprn:cmsay(14.5,15,TRANSFORM(A05IMPD,"@ZE 999,999,999.99"),ofont10)
              Caja(16.0,5.5,19.5,14.5,oPrn,0,,oPen2)
              oprn:cmsay(16.5,6.5,"Total do imposto devido",oFONT08)
              oprn:cmsay(17.0,6.5,"IR fonte day-trade do mes",oFONT08)
              oprn:cmsay(17.5,6.5,"IR fonte (lei 11.033/2004)",ofont08)
              oprn:cmsay(18.0,6.5,"I.R. a compensar meses anteriores",oFONT08)
              oprn:cmsay(18.5,6.5,"Imposto a Pagar",oFONT08)
              oPrn:cmsay(19.0,6.5,"Imposto Pago",oFONT08)
              oprn:cmsay(16.0,9.0,"CONSOLIDAÇÃO DO MES",ofont13)
              oprn:cmsay(16.5,12.5,TRANSFORM(A05IMPDEVI,"@E 999,999,999.99"),ofont10)
              oprn:cmsay(17.0,12.5,TRANSFORM(A05IRDAYME,"@E 999,999,999.99"),ofont10)
              oprn:cmsay(17.5,12.5,TRANSFORM(A05IRDANT,"@E 999,999,999.99"),ofont10)
              oprn:cmsay(18.0,12.5,TRANSFORM(A05IRDAYAN,"@E 999,999,999.99"),ofont10)
              oprn:cmsay(18.5,12.5,TRANSFORM(A05IMPAPAG,"@E 999,999,999.99"),ofont10)
              oprn:cmsay(19.0,12.5,TRANSFORM(A05IRPAGO,"@E 999,999,999.99"),ofont10)
              oPrn:cmsay(23.1,0.5,"CarteiraNet - Desenvolvido por: Exatus.net - http://www.exatus.net",oFont08)
              oPrn:cmsay(23.5,0.5,"Emitido em "+DTOC(DAT_HOJE)+"   Página: "+strzero(1,3),ofont08)
              ENDPAGE
            Endif
            SKIP
          ENDDO
          ENDPRINT
        Endif
        CLOSE DATABASE
      Return

      Procedure Linha(nArriba,nIzq,nAbajo,nDerecha,oPrn,oPen)
        Local xCor := {} , yCor := {}
        xCor := oPrn:Cmtr2Pix(nArriba,nIzq)
        yCor := oPrn:Cmtr2Pix(nAbajo,nDerecha)
        oPrn:Line(xCor[1],xCor[2],yCor[1],yCor[2],oPen)
      Return NIL

      Func SayBitmap(nArriba,nIzq,nAncho,nAlto,xBitmap,oPrn)
        //-------------------------------------------------------------*
        PUBLIC xCor := {}
        PUBLIC nPixWidth, nPixHeight
        xCor := oPrn:Cmtr2Pix(nArriba,nIzq)
        nAncho := nAncho/2.54
        nAlto  := nAlto /2.54
        nPixWidth  := INT(oPrn:nLogPixelY()*nAncho)
        nPixHeight := INT(oPrn:nLogPixelX()*nAlto)
        oPrn:SayBitmap(xCor[1],xCor[2],xBitmap,nPixWidth ,nPixHeight)
      Return NIL

      Procedure RUNFILE(FNAME)
        Do While .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_RUN="c:\BOLSANET\ARQ_R"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_RUN1:=ARQ_RUN+".BAT"
          If !FILE(ARQ_RUN1)
            EXIT
          Endif
        ENDDO
        CMACROS="@ECHO OFF"+CHR(13)+CHR(10)+"START "+FNAME+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(ARQ_RUN1,CMACROS)
        WAITRUN(ARQ_RUN1,0)
        Erase &ARQ_RUN1
      Return NIL

      Procedure CADCONT()
        CLOSE DATABASE
        SELE 1
        If .NOT. USEREDE(ARQ04,.F.,10,"CONTA")
          MsgStop("O arquivo de Contabilidade não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind0401
        Endif
        BrowseC(13,"Contas","Contas", { || buscaConta() },{ ||NCADConta("I") },{ ||NCADCONTA("A") })
        USE
        CLOSE DATABASE
      Return

      FUNCTION BUSCACONTA()
      Return

      Procedure NCADCONTA(F)
        LOCAL oDlg,oBrush,oCbx
        PUBLIC ZCLAS
        If F="I"
          DECLARE _var1 [FCOUNT()],_var2 [FCOUNT()],_var3 [FCOUNT()]
          AFIELDS(_var1,_var2,_var3)
          FOR t = 1 TO FCOUNT()
            SysRefresh()
            _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
            If _var2 [t] = "C"
              &_var4 = SPACE(_var3 [t])
            ElseIf _var2 [t] = "N"
              &_var4 = 0
            ElseIf _var2 [t] = "D"
              &_var4 = CTOD(SPACE(08))
            ElseIf _var2 [t] = "M"
              &_var4 = SPACE(10)
            ElseIf _var2 [t] = "L"
              &_var4 = SPACE(01)
            Endif
          NEXT t
        Else
          If !REGLOCK()
            MSGINFO("Registro sendo alterado por outro Usuário",WSIST)
            Return
          Else
            DECLARE _var1 [FCOUNT()]
            AFIELDS(_var1)
            FOR t = 1 TO FCOUNT()
              SysRefresh()
              _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
              _var5 = _var1 [t]
              &_var4 = &_var5
            NEXT t
          Endif
        Endif
        MFONTES()
        DEFINE DIALOG oDlg FROM 0,0 TO 10,37 TITLE "Conta" BRUSH oBrushTela STYLE WS_CAPTION
        Z04DESC=OEMTOANSI(Z04DESC)
        @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 140,49
        @   0.87,01 SAY "Descrição" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   1.74,01 SAY "Debita" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,01 SAY "Credita" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        A="N"
        @ 01,6 GET Z04DESC PICTURE "@!" FONT oFontget SIZE 90,11 NO MODIFY
        @ 02,6 GET Z04CPAH FONT oFontget SIZE 90,11
        @ 03,6 GET Z04VDAH FONT oFontget SIZE 90,11

        @ 60,030 SBUTTON PROMPT "&Confirma" PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" FONT oFontbut ADJUST TEXT ON_CENTER OF oDlg size 40,12 ACTION (Z04DESC:=ANSITOOEM(Z04DESC),REGLOCK(10),PUTREC(),oDlg:end())
        @ 60,075 SBUTTON PROMPT "C&ancela" PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" FONT oFontbut ADJUST TEXT ON_CENTER OF oDlg SIZE 40,12 ACTION (OdLG:eND())
        odlg:lhelpicon:=.f.
        ACTIVATE DIALOG oDlg CENTERED
      Return

      Procedure INTNOTA()
        cOutFile:=cGetFile("Arquivo (*.*) | *.* |","Arquivo a ser utilizado")
        ARQ=ALLTRIM(COUTFILE)
        If !FILE(ARQ)
          MSGINFO("Arquivo "+arq+" Inexistente !",WSIST)
          CLOSE DATABASE
          Return
        Endif
        CLOSE DATABASE
        SELE 1
        If .NOT. USEREDE(ARQ12,.F.,10,"NOTA")
          MsgStop("O arquivo NOTAS não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1201
        Endif
        SELE 2
        If .NOT. USEREDE(ARQ14,.F.,10,"PAPEL")
          MsgStop("O arquivo PAPEIS não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1401
        Endif
        SELE 3
        If .NOT. USEREDE(ARQ11,.F.,10,"ITENS")
          MsgStop("O arquivo ITENS não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind1101
        Endif
        SELE 4
        If .NOT. USEREDE(ARQ04,.F.,10,"TABELA")
          MsgStop("O arquivo TABELA não está  disponível",WSIST)
          CLOSE DATABASE
          Return
        Else
          SET INDEX TO &ind0401
        Endif
        oText:=TTxtfile():NEW(ARQ)
        CONTAD=oText:RecCount()
        If VALTYPE(CONTAD)<>"N"
          CONTAD=1
        Endif
        dnota=oText:Readline()
        oText:Skip()
        SELE 1
        ADIREG(0)
        REPLACE A12NOTA WITH STRZERO(VAL(SUBSTR(DNOTA,1,10)),6)
        REPLACE A12DATA WITH CTOD(SUBSTR(DNOTA,11,10))
        REPLACE A12TXLIQ WITH VAL(SUBSTR(DNOTA,21,15))/100
        REPLACE A12TXREG WITH VAL(SUBSTR(DNOTA,36,15))/100
        REPLACE A12TXTEOP WITH VAL(SUBSTR(DNOTA,51,15))/100
        REPLACE A12TXANA WITH VAL(SUBSTR(DNOTA,66,15))/100
        REPLACE A12EMOL WITH VAL(SUBSTR(DNOTA,81,15))/100
        REPLACE A12CORRET WITH VAL(SUBSTR(DNOTA,96,15))/100
        REPLACE A12IRRF WITH VAL(SUBSTR(DNOTA,111,15))/100
        REPLACE A12OUTRAS WITH VAL(SUBSTR(DNOTA,126,15))/100
        REPLACE A12IRDAY WITH VAL(SUBSTR(DNOTA,141,15))/100
        REPLACE A12LIQ WITH VAL(SUBSTR(DNOTA,156,15))/100
        Z12IRRF=NOTA->A12IRRF
        Z12NOTA=A12NOTA
        TOTDEZP=NOTA->A12TXLIQ+NOTA->A12TXREG+NOTA->A12CORRET+NOTA->A12TXANA+NOTA->A12EMOL
        TOTDEZOPCAO=NOTA->A12TXTEOP
        FOR WZI=2 TO CONTAD
          linha=oText:Readline()
          Z11CODP=ALLTRIM(SUBSTR(LINHA,2,10))
          Z11CODT=PESQ(Z11CODP,"2",1,"A03CODT")
          Z11MIL=PESQ(Z11CODP,"2",1,"a03lmil")
          SELE 3
          ADIREG(0)
          REPLACE A11NOTA WITH STRZERO(VAL(SUBSTR(DNOTA,1,10)),6)
          REPLACE A11DATA WITH CTOD(SUBSTR(DNOTA,11,10))
          REPLACE A11REG WITH RECNO()
          REPLACE A11CODP WITH Z11CODP
          REPLACE A11OPER WITH LEFT(LINHA,1)
          REPLACE A11PRAZO WITH VAL(SUBSTR(LINHA,12,3))
          REPLACE A11QTDE WITH VAL(SUBSTR(LINHA,15,15))
          REPLACE A11VALR WITH VAL(SUBSTR(LINHA,30,15))/100
          REPLACE A11VLIQ WITH A11VALR
          REPLACE A11DTVAL WITH CTOD(SUBSTR(LINHA,45,10))
          REPLACE A11CODT WITH Z11CODT
          If Z11MIL="S"
            REPLACE A11PUNI WITH A11VALR/A11QTDE*1000
          Else
            REPLACE A11PUNI WITH A11VALR/A11QTDE
          Endif
          oText:Skip()
        NEXT WZI

        SELE 3
        SET FILTER TO A11NOTA=Z12NOTA
        GO TOP
        TOTOP=0
        TOTVIS=0
        Do While !EOF()
          SysRefresh()
          If A11NOTA=Z12NOTA
            If PESQ(A11CODP,"2",1,"A03POP")="S"
              SKIP
              LOOP
            Endif
            If VAL(PESQ(A11CODP,"2",1,"A03TIPO"))=2 .or.VAL(PESQ(A11CODP,"2",1,"A03TIPO"))=3
              TOTOP=TOTOP+A11VALR
            Endif
            TOTVIS=TOTVIS+A11VALR
          Endif
          SKIP
        ENDDO
        GO TOP
        Do While !EOF()
          SysRefresh()
          If A11NOTA=Z12NOTA .And. REGLOCK(10)
            If PESQ(A11CODP,"2",1,"A03POP")="S"
              SKIP
              LOOP
            Endif
            TP=PESQ(A11CODP,"2",1,"A03TIPO")
            If val(TP)=2 .or. val(TP)=3
              REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)+(A11VALR/TOTOP*TOTDEZOPCAO)
            Else
              REPLACE A11VCOR WITH (A11VALR/TOTVIS*TOTDEZP)
            Endif
            If A11OPER="C" .OR. A11OPER="1"
              REPLACE A11VLIQ WITH A11VALR+A11VCOR
            Else
              REPLACE A11VLIQ WITH A11VALR-A11VCOR
            Endif
            UNLOCK
          Endif
          SKIP
        ENDDO
        GO TOP
        LIQ=0
        Do While !EOF()
          SysRefresh()
          If PESQ(A11CODP,"2",1,"A03POP")="S"
            SKIP
            LOOP
          Endif
          If A11NOTA=Z12NOTA .And. PESQ(A11CODP,"2",1,"A03TIPO")#"3"
            If A11OPER="V"
              LIQ=LIQ+A11VLIQ
            Else
              LIQ=LIQ-A11VLIQ
            Endif
          Endif
          If A11NOTA=Z12NOTA .And. PESQ(A11CODP,"2",1,"A03TIPO")="3"
            LIQ=LIQ-a11VCOR
          Endif
          SKIP
        ENDDO
        SELE 3
        SET FILTER TO
        SELE 1
        If REGLOCK(10)
          REPLACE A12LIQ WITH LIQ-A12IRRF
        Endif
        UNLOCK
        CLOSE DATABASE
        otext:close()
        CALCSALD()
        IMPRES(0)
        CURSORARROW()
        MSGINFO("Nota Importada !",wsist)
      Return

      Procedure ASSINANET(PARAM)
        If !FILE("CONEX.DAD")
          Return
        Endif
        RESTORE FROM CONEX.DAD ADDITIVE
        CONEXON=trim(Decrypt(CONEXON,"leopardo"))
        CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
        CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
        CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
        CONEXINTO=90
        CURSORWAIT()
        SQL CONNECT ON CONEXON;
          PORT CONEXPORT;
          DATABASE CONEXDTBAS;
          USER CONEXUSER;
          PASSWORD CONEXSENHA;
          OPTIONS SQL_NO_WARNING;
          LIB 'MySQL';
          INTO CONEXINTO

        If SQL_ErrorNO() > 0
          Return
        Endif
        SQLSETCONNECTION(CONEXINTO)
        SET PACKETSIZE TO 25
        SELE 35
        If !USEREDE("instituicao",.F.,10,"CORRETOR","M")
          SELE 30
          Return
        Endif
        SEENVIAEMAIL=PARAM
        FEZALGO=.F.
        SELE 35
        GO TOP
        CURSORWAIT()
        WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
        VERDIRETORIO(WEBLOCAL1)
        quemenviou=""
        emailquemenviou=""
        Do While .NOT. EOF()
          SELE 35
          wcpfcnpj=cnpjcpf
          wqtd_ass_if=qtdass
          WNRC="SELECT sql_rowid FROM INSTITUICAO WHERE nome like '%"+trim(nome)+"%' and cnpjcpf="+WCPFCNPJ
          mz=SQLArrayAssoc(WNRC)
          wcodif=MZ[1,1]
          WCODIF=STRZERO(VAL(WCODIF),10)
          SysRefresh()
          MCAMINHO=CAMINHO
          M1CAMINHO=CAMINHO
          If SUBSTR(CAMINHO,2,1)<>":"
            MCAMINHO=ALLTRIM(CURDRIVE()+":"+caminho)
          Endif
          mcaminho=alltrim(mcaminho)+"diversos\"
          cmacro="MD "+ALLTRIM(MCAMINHO)
          MZLINHA=CMACRO+CHR(13)+CHR(10)+"DEL "+WEBLOCAL1+"AGUABC.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(WEBLOCAL1+"AGUABC.TXT",MZLINHA)
          MEMOWRIT(WEBLOCAL1+"FZCON.BAT",MZLINHA)
          WAITRUN(WEBLOCAL1+"FZCON.BAT",0)
          CURSORWAIT()
          Do While FILE(WEBLOCAL1+"AGUABC.TXT")
            SysRefresh()
          ENDDO
          CURSORWAIT()
          ARQF3=ALLTRIM(MCAMINHO)+"FEITO.ASS"
          ARQF30=ALLTRIM(M1CAMINHO)+"FEITO.ASS"
          If FILE(ARQF3)
            DECLARE VLABEL[ADIR(mcaminho+"*.PDF")]
            ADIR(mcaminho+"*.PDF",VLABEL)
            FOR OI=1 TO LEN(VLABEL)
              NZO=mcaminho+ALLTRIM(VLABEL[OI])
              NZO1=LEFT(NZO,LEN(NZO)-3)+"TXT"
              NZO0=alltrim(m1caminho)+ALLTRIM(VLABEL[OI])
              NZO10=LEFT(NZO0,LEN(NZO0)-3)+"TXT"
              NOMAARQUIVO:=DIRECTORY(NZO)
              If NOMAARQUIVO[1][2]<500
                Erase &NZO
                Erase &NZO1
              Endif
              LZCOPYFILE(nzo,nzo0)
              LZCOPYFILE(nzo1,nzo10)
              LZCOPYFILE(arqf3,arqf30)
              Erase &nzo
              Erase &nzo1
            NEXT OI
            Erase &arqf3
          Endif
          MCAMINHO=ALLTRIM(M1CAMINHO)
          cmacro="MD "+ALLTRIM(MCAMINHO)
          MZLINHA=CMACRO+CHR(13)+CHR(10)+"DEL "+WEBLOCAL1+"AGUABC.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(WEBLOCAL1+"AGUABC.TXT",MZLINHA)
          MEMOWRIT(WEBLOCAL1+"FZCON.BAT",MZLINHA)
          WAITRUN(WEBLOCAL1+"FZCON.BAT",0)
          CURSORWAIT()
          Do While FILE(WEBLOCAL1+"AGUABC.TXT")
            SysRefresh()
          ENDDO
          CURSORWAIT()
          DARQ=WEBLOCAL1+"FZCON.BAT"
          Erase &DARQ
          DECLARE VLABEL1[ADIR(MCAMINHO+"*.XML")]
          ADIR(MCAMINHO+"*.XML",vlabel1)
          I=0
          FOR I=1 TO LEN(VLABEL1)
            SysRefresh()
            NARQ=TRIM(MCAMINHO)+VLABEL1[I]
            Erase &NARQ
          NEXT I
          ARQF3=ALLTRIM(CAMINHO)+"FEITO.ASS"
          If !FILE(ARQF3)
            SKIP
            LOOP
          Endif
          TLINHA=MEMOREAD(ARQF3)
          quemenviou=SINAL(SUBSTR(TLINHA,1,AT(CHR(10),TLINHA)-1))
          emailquemenviou=SINAL(SUBSTR(TLINHA,AT(CHR(10),TLINHA)+1,LEN(TLINHA)))
          If LOWER(emailquemenviou)="roberto.amaral@moneycorp.com"
            emailquemenviou="onboardingbrazil@moneycorp.com"
          Endif
          If EMPTY(QUEMENVIOU)
            QUEMENVIOU="Robot"
          Endif
          If at("@",emailquemenviou)=0
            emailquemenviou=""
          Endif
          M->LOCAL21=TRIM("/"+CURDIR())
          SELE 35
          WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
          VERDIRETORIO(WEBLOCAL1)
          WEBLOCAL2=ALLTRIM(CAMINHO)+"ARQF\"
          VERDIRETORIO(WEBLOCAL2)
          WEBLOCAL3=alltrim(CAMINHO)+"DOCTOS\"
          VERDIRETORIO(WEBLOCAL3)
          DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
          ADIR(WEBLOCAL1+"*.*",VLABEL1)
          FOR O=1 TO LEN(VLABEL1)
            SysRefresh()
            ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
            Erase &ARQDE
          NEXT O
          DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
          ADIR(WEBLOCAL1+"*.*",VLABEL1)
          If LEN(VLABEL1)>0
            MSGINFO("Verifique pasta temporaria, não apagou todos arquivos",wsist)
            EXIT
          Endif
          SELE 35
          WEBLOCAL=ALLTRIM(CAMINHO)
          WEBCOR=NOME
          WEBASSUNTO=ASSUNTO
          If EMPTY(WEBASSUNTO)
            WEBASSUNTO="AssinaNet - Documento para ser assinado eletronicamente"
          Endif
          WEBCORPO=OEMTOANSI(CORPOEMAIL)
          WEBPROCE=OEMTOANSI(CORPOPROCESSO)
          WEBCORPODOCTO1=OEMTOANSI(CORPODOCTO1)
          WEBCORPODOCTO2=OEMTOANSI(CORPODOCTO2)
          WEBCORPODOCTO3=OEMTOANSI(CORPODOCTO3)
          WEBCORPODOCTO4=OEMTOANSI(CORPODOCTO4)
          WEBCORPODOCTO5=OEMTOANSI(CORPODOCTO5)
          WEBCORPODOCTO6=OEMTOANSI(CORPODOCTO6)
          WEBCORPODOCTO7=OEMTOANSI(CORPODOCTO7)
          WEBCORPODOCTO8=OEMTOANSI(CORPODOCTO8)
          WEBSENHA=SENHAEMAIL
          WEBLOGIN=USUARIOEMAIL
          WEBSENHA=SENHAEMAIL
          WEBSERV=ALLTRIM(SMTP)
          WEBEMAIL=ALLTRIM(EMAIL)
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.zip")]
          ADIR(WEBLOCAL+"*.zip",vlabel)
          I=0
          FOR I=1 TO LEN(VLABEL)
            SysRefresh()
            CURSORWAIT()
            NARQ=WEBLOCAL+VLABEL[I]
            lCHDIR(weblocal)
            COMANDO=WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
            WAITRUN(COMANDO,0)
            lCHDIR(wcorrente)
            Erase &NARQ
          NEXT I
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
          ADIR(WEBLOCAL+"*.PDF",vlabel)
          FOR I=1 TO LEN(VLABEL)
            NARQ=WEBLOCAL+VLABEL[I]
            If LEFT(VLABEL[I],2)="._"
              Erase &NARQ
            Endif
          NEXT I
          I=0
          WDADOS=WEBLOCAL1+"ENVIADO.TXT"
          DECLARE TX1[1]
          TX1[1]="Assinanet: "+WEBCOR
          LOGFILE(WEBLOCAL2+"ROTINAS.TXT",TX1)
          TEVE=.F.
          RELBC()
          If SEENVIAEMAIL="S" .And. FEZALGO .And. FILE(WDADOS) .And. !empty(emailquemenviou)
            ENVIAEMAIL(ALLTRIM(emailquemenviou),"AssinaNet - Log dos Documentos enviados","Em anexo relatorio",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          Endif
          If SEENVIAEMAIL="S" .And. FEZALGO .And. FILE(WDADOS)
            ENVIAEMAIL("bene@exatus.net","AssinaNet "+WEBCOR,"Em anexo relatorio dos documentos Enviados",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            //    enviaaviso("Assinanet documentos Enviados "+WEBCOR,"19861349898")
          Endif
          If SEENVIAEMAIL="S" .And. FEZALGO .And. FILE(WDADOS)
            //    ENVIAEMAIL("wanderlei.cruz@starmoney.com.br","AssinaNet "+WEBCOR,"Em anexo relatorio dos documentos Enviados",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          Endif
          LCHDIR(M->LOCAL21)
          // If TEVE=.F.
          Erase &ARQF3
          // Endif
          SELE 35
          SKIP
          If FILE(WDADOS)
            Erase &WDADOS
          Endif
        ENDDO
        cursorarrow()
        SELE 35
        USE
        SQLDISCONNECT(CONEXINTO)
        SELE 30
        CURSORARROW()
      Return

      Procedure RELBC()
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
        ADIR(WEBLOCAL+"*.PDF",VLABEL)
        FOR OI=1 TO LEN(VLABEL)
          NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
          NZO1=LEFT(NZO,LEN(NZO)-3)+"TXT"
          NOMAARQUIVO:=DIRECTORY(NZO)
          If NOMAARQUIVO[1][2]<500
            Erase &NZO
            Erase &NZO1
          Endif
        NEXT OI
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
        ADIR(WEBLOCAL+"*.PDF",VLABEL)
        OZ=0
        If LEN(VLABEL)>0
          FEZALGO=.T.
        Endif
        MAISNOME=DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)
        FOR OI=1 TO LEN(VLABEL)
          NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
          NZO1=LEFT(NZO,LEN(NZO)-3)+"TXT"
          If FILE(NZO1)
            TEVE=.T.
            MZLINHA=MEMOREAD(NZO1)
          Endif
        NEXT OI
        If !TEVE
          FOR OI=1 TO LEN(VLABEL)
            NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
            NZO1=WEBLOCAL+STRZERO(OI,5)+".PDF"
            FRENAME(NZO,NZO1)
          NEXT OI
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
          ADIR(WEBLOCAL+"*.PDF",VLABEL)
          FOR OI=1 TO LEN(VLABEL)
            NARQ=WEBLOCAL+ALLTRIM(VLABEL[OI])
            M->LOCAL3=ALLTRIM(WEBLOCAL2)+MAISNOME+ALLTRIM(VLABEL[OI])
            LZCOPYFILE(NARQ,M->LOCAL3)
            M->LOCAL3=ALLTRIM(WEBLOCAL1)+MAISNOME+ALLTRIM(VLABEL[OI])
            LZCOPYFILE(NARQ,M->LOCAL3)
            LCHDIR(WEBLOCAL1)
            CURSORWAIT()
            DARQ=WEBLOCAL1+"FZCON.BAT"
            MZLINHA=WCORRENTE+"\PDFTEXT.EXE -eol dos -layout "+M->LOCAL3+CHR(13)+CHR(10)+"DEL "+WEBLOCAL1+"AGUABC.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
            MEMOWRIT(WEBLOCAL1+"AGUABC.TXT",MZLINHA)
            MEMOWRIT(WEBLOCAL1+"FZCON.BAT",MZLINHA)
            WAITRUN(WEBLOCAL1+"FZCON.BAT",0)
            CURSORWAIT()
            Do While FILE(WEBLOCAL1+"AGUABC.TXT")
              SysRefresh()
            ENDDO
            CURSORWAIT()
            Erase &DARQ
            Erase &NARQ
          NEXT OI
          Do While .T.
            SysRefresh()
            M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
            NARQ3=weblocal1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".DAT"
            If !FILE(NARQ3)
              EXIT
            Endif
          ENDDO
          MAT_STRU={}
          AADD(MAT_STRU,{"LINHA","C",254,0})
          DBCREATE(NARQ3,MAT_STRU,"DBFCDX")
          SELE 33
          If !USEREDE(NARQ3,.T.,10,"TMP")
            SELE 35
            Return
          Endif
        Endif
        TEVEERRO=.F.
        ABREASSINA()
        If TEVEERRO
          SELE 31
          USE
          SELE 32
          USE
          SELE 34
          USE
          SELE 36
          USE
          SELE 38
          USE
          SELE 39
          USE
          SELE 35
          Return
        Endif
        WCODINT=0
        If FILE(WEBLOCAL+"AT_ASSINA.CSV")
          ATCADASS()
        Endif
        If TEVE   //  Inclusão pela Capa
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.TXT")]
          ADIR(WEBLOCAL+"*.TXT",VLABEL)
          FOR OI=1 TO LEN(VLABEL)
          WIF="0"
          NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
          NZO1=LEFT(NZO,LEN(NZO)-3)+"PDF"
          NZO2=WEBLOCAL+STRZERO(OI,5)+".PDF"
          M->LOCAL3=ALLTRIM(WEBLOCAL2)+MAISNOME+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(NZO,M->LOCAL3)
          FRENAME(NZO1,NZO2)
          M->LOCAL3=ALLTRIM(WEBLOCAL2)+MAISNOME+LEFT(vlabel[OI],LEN(VLABEL[OI])-3)+"PDF"
          LZCOPYFILE(NZO2,M->LOCAL3)
          NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
          NZO3=LEFT(NZO,LEN(NZO)-3)+"TXT"
          NZO4=WEBLOCAL+STRZERO(OI,5)+".TXT"
          FRENAME(NZO3,NZO4)
          MZLINHA=Strtran(MEMOREAD(NZO4),CHR(10),"")+CHR(13)+CHR(13)
          Erase &NZO4
          Erase &NZO
          LCHDIR(WEBLOCAL1)
          CURSORWAIT()
          ZPFCNPJ=LECAPA("CNPJ/CPF:")
          ZTCONTR=LECAPA("DATA:")
          ZEAIS=LECAPA("R$:")
          ZBS=LECAPA("OBS:")
          ZNPJINT=LECAPA("INTERMEDIARIO:")
          ZNPJINT=STRZERO(VAL(ZNPJINT),14)
          ZIPCORRETORA=LECAPA("IPCORRETORA:")
          ZPFASSINOU=ALLTRIM(LECAPA("PFASSINOU:"))
          ZIPPF=LECAPA("01IPPF:")
          ZRCONTR=val(LECAPA("01NRDOCTO:"))
          ZCORRETASSINA=LECAPA("01INSTASSINA:") // inclui a assinatura de quem Aceitou o Cadastro
          ZACEITEPOR=LECAPA("01ACEITEPOR:")
          ZACEITEPORCPF=LECAPA("01ACEITEPORCPF:")
          zaceiteem=LECAPA("01ACEITEEM:")
          If EMPTY(ZACEITEEM)
            ZACEITEEM="NOW()"
          Else
            zaceiteem="'"+ZACEITEEM+"'"
          Endif
          W01TPDOCTO=VAL(TRIM(LECAPA("01TIPODOCUMENTO:")))
          wmcpfoperador=LECAPA("01CPFOPERADOR:")
          wmnomeoper=LECAPA("01NOMEOPERADOR:")
          ZIPOPERADOR=LECAPA("01IPOPERADOR:")
          ZCOMPLETO=LECAPA("01COMPLETO:")
          ZCHAVEFICHA=LECAPA("01CHAVEFICHA:")
          ZHIDDEN=LECAPA("01HIDDEN:")
          ZEAIS=Strtran(ZEAIS,".","")
          ZEAIS=Strtran(ZEAIS,",",".")
          ZEAIS=VAL(ZEAIS)
          ZPFCNPJ=STRZERO(VAL(ZPFCNPJ),14)
          WNOMEEMPR=""
          WWCPFCNPJ=""

          //   MSGINFO(ZPFCNPJ)
          //   MSGINFO(ZTCONTR)
          //   MSGINFO(ZEAIS)
          //   MSGINFO(ZBS)
          //   MSGINFO(ZNPJINT)
          //   MSGINFO(ZRCONTR)

          //quit

          CADASTROCAPA() // FAZ O CADASTRO ANTES DE RECONHECER
          If VAL(ZPFCNPJ)=0
            NZO=WEBLOCAL+ALLTRIM(VLABEL[OI])
            Erase &NZO
            Erase &NZO2
          Endif
          If VAL(ZPFCNPJ)#0
            SELE 35
            ZEQUENC=SEQUENCIA+1
            If REGLOCK(10)
              REPLACE SEQUENCIA WITH ZEQUENC
            Endif
            SELE 36
            BEGIN TRANSACTION
            ADIREG(0)
            REPLACE NOMEARQ WITH VLABEL[OI]
            REPLACE ENVIADO WITH DTOC(DATE())+"-"+TIME()
            REPLACE USUARIO WITH QUEMENVIOU
            REPLACE NOMEIf WITH WEBCOR
          END TRANSACTION
          COMMIT
          SELE 34
          wfiltass="codif="+wcodif+" and CNPJCPF="+ZPFCNPJ
          SQL FILTER TO wfiltass
          GO TOP
          WCODCLIIF=0
          If val(CNPJCPF)=val(ZPFCNPJ) .And. WCODIF=CODIF
            WCODCLIIF=CLIIF->SQL_ROWID
            MWNOVO=CLIIF->NOVO
            WIF=STRZERO(WCODCLIIF,10)
          Endif
          WCODINT=0
          If VAL(ZNPJINT)#0
            SELE 34
            wfiltass="CNPJCPF="+ZNPJINT+" and codif="+wcodif
            SQL FILTER TO wfiltass
            GO TOP
            WCODINT=0
            If val(CNPJCPF)=val(ZNPJINT)
              WCODINT=CLIIF->SQL_ROWID
            Endif
          Endif
          SELE 32
          NARQ=NZO2
          ZRQPDF=WEBLOCAL3+STRZERO(ZEQUENC,10)+".PDF"
          ZRQTXT=WEBLOCAL3+STRZERO(ZEQUENC,10)+".TXT"
          ZRQPDF1=WEBLOCAL3+STRZERO(ZEQUENC,10)+".PDF1"
          ZARQH=WEBLOCAL1+STRZERO(ZEQUENC,10)+".PDF"
          ZARQH1=WEBLOCAL1+STRZERO(ZEQUENC,10)+".PDF1"
          cAnexo=LOWER(NARQ)
          wcAnexo=Strtran(cAnexo,".pdf","_C.pdf")
          comando=WCORRENTE+"\gswin32c -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/prepress -dNOPAUSE -dQUIET -dBATCH -sOutputFile="+wcAnexo+" "+cAnexo
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+weblocal+"pausa.txt"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal+"compacta.bat",mzlinha)
          MEMOWRIT(weblocal+"pausa.txt",MZLINHA)
          WAITRUN(weblocal+"compacta.bat",0)
          Do While FILE(weblocal+"pausa.txt")
            SysRefresh()
          ENDDO
          wtam1=DIRECTORY(cAnexo)
          wtam1=wtam1[1][2]

          wtam2=DIRECTORY(wcAnexo)
          wtam2=wtam2[1][2]
          If wtam1<wtam2
            Erase &wcAnexo
          Else
            Erase &cAnexo
            FRENAME(wcAnexo,cAnexo)
          Endif
          cAnexo=weblocal+"compacta.bat"
          Erase &cAnexo
          cAnexo=weblocal+"pausa.txt"
          Erase &cAnexo
          ZTCONTR1=ctod(ztcontr)
          WNRC='SELECT MD5("'+DTOC(ZTCONTR1)+STRZERO(ZEQUENC,10)+'4ss1n4")'
          MZ=SQLArrayAssoc(WNRC)
          WCHASH=MZ[1,1]
          arqQR=WEBLOCAL1+"qr.BMP"
          Erase &arqQR
          arqQR=WEBLOCAL1+"qr.PNG"
          Erase &arqQR
          texto="https://assinanet.exatus.net/validaassina.asp?dados="+WCHASH+"|"+STRZERO(ZEQUENC,10)+"|"+DTOC(ZTCONTR1)
          qrDLL:=LoadLibrary("QRCodelib.Dll" )
          #Define DC_CALL_STD 0x0020
          nResp:=DllCall(qrDLL, DC_CALL_STD, "FastQRCode", texto, arqQR)
          FreeLibrary(qrDLL)
          oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
          Erase &ZARQH
          oReg:Set("PDFFilename",ZARQH)
          oReg:Close()
          PRINT oPrn NAME "HASH" TO TEMPDF1
          DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
          nRowStep = oPrn:nVertRes() / 66
          nColStep = oPrn:nHorzRes() / 80
          PAGE
          oprn:cmsay(28.5,1.0,"Controle AssinaNet: "+STRZERO(ZEQUENC,10)+" "+WCHASH,oFont,,nRGB(0,0,0),,0)
          If AT("CLASSIFICACAO E FINALIDADE",zbs)#0
            oprn:cmsay(28.8,1.0,"Declaro estar de acordo com todas informações deste documento",oFont,nRGB(0,0,0),,0)
          Else
            oprn:cmsay(28.8,1.0,"Documento Assinado Eletronicamente",oFont,nRGB(0,0,0),,0)
          Endif
          If file(arqQr)
            sayBitmap(27.7,12,1.8,1.8,arqQR,oPrn)
          Endif
          ENDPAGE
          ENDPRINT
          tini=time()
          Do While .T.
            CURSORWAIT()
            DIF=TIMEASSECONDS(TIMEDIFF(tini,TIME()))
            If dif>60
              EXIT
            Endif
            If (nHandle:=FOpen(ZARQH,FO_READ+FO_EXCLUSIVE))!=-1
              fClose(nHandle)
              EXIT
            Else
              fClose(nHandle)
            Endif
          ENDDO
          Erase &ZARQH1
          comando=WCORRENTE+"\PDFTK "+NARQ+" stamp "+ZARQH+" output "+ZARQH1+chr(13)+chr(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","HASH"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          Do While FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &NARQ
          Erase &ZARQH
          FRENAME(ZARQH1,NARQ)
          comando=WCORRENTE+"\PDFTK "+NARQ+" background c:\exatus\eletro.pdf output "+zrqpdf+chr(13)+chr(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina contracapa"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          Do While FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &NARQ
          BEGIN TRANSACTION
          ADIREG(0)
          REPLACE DATA WITH CTOD(ZTCONTR)
          If W01TPDOCTO#0
            REPLACE TIPODOCUMENTO WITH W01TPDOCTO
          Else
            REPLACE TIPODOCUMENTO WITH 3
          Endif
          REPLACE CAMINHO WITH ZRQPDF
          REPLACE CAMINHOTXT WITH ZRQPDF
          REPLACE CODIf WITH CORRETOR->SQL_ROWID
          REPLACE CODCLIIf WITH WCODCLIIF
          REPLACE CODCORR WITH WCODINT
          REPLACE VALOR WITH ZEAIS
          REPLACE OBS WITH ZBS
          REPLACE INCLUIDO WITH DTOC(DATE())+"-"+TIME()
          REPLACE CHASH WITH wchash
          REPLACE USERINC WITH "-AssinaNet-n"
          REPLACE NOVO WITH "S"
          REPLACE NR_CAMBIO WITH ZCHAVEFICHA
          REPLACE HIDDEN WITH ZHIDDEN
          If ZCOMPLETO="S"
            REPLACE COMPLETO WITH "S"
          Endif
        END TRANSACTION
          COMMIT
          ZPDOCTO=TIPODOCUMENTO
          DECLARE TX1[1]
          TX1[1]="Documento nr: "+ZBS+" - Incluido no sistema"
          LOGFILE(WDADOS,TX1)
          WSQLNRDOCTO=sql_rowid
          ASSINACAPA()
          If WCODCLIIF#0 .And. EMPTY(ZPFASSINOU)
            If VAL(WCODIF)=24
              // criar campo na tabela instituição (se envia email para assinantes) SIGN (no cadastro de usuarios)
              // Ler cadastro de usuarios da wcodif=codIf quem tiver SIGN enviar email
              ENVIAEMAIL("elaine.santos@moneycorp.com",WEBASSUNTO,WEBCORPO,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
            Endif
            SELE 31
            wfiltass='CODCLIIf LIKE "%'+WIF+'%"'
            SQL FILTER TO wfiltass
            GO TOP
            Do While !EOF()
              WNRC1="SELECT sql_rowid,nome from assinantes where cnpjcpf="+cnpjcpf
              mz=SQLArrayAssoc(WNRC1)
              If len(mz)>0
                WNRC1="SELECT sql_rowid,vencto,valormax,tipodocumento FROM vencimentos where sql_deleted='F' and codassinante="+mz[1,1]+" and codcliif="+WIF
                mznome=mz[1,2]
                mzcnpjcpf=cnpjcpf
                mz=SQLArrayAssoc(WNRC1)
                If LEN(MZ)>0
                  MWSQL_ROWID=MZ[1,1]
                  MWVENCTO=ALLTRIM(MZ[1,2])
                  MWVENCTO=CTOD(RIGHT(MWVENCTO,2)+"/"+SUBSTR(MWVENCTO,6,2)+"/"+LEFT(MWVENCTO,4))
                  MWVALORMAX=VAL(MZ[1,3])
                  MWTIPODOCUMENTO=MZ[1,4]
                  If MWVENCTO=CTOD("")
                    //                wnrc1="update vencimentos set vencto="+dtom(ctod("31/12/2100"))+" where sql_rowid="+mwsql_rowid
                    //                MWVENCTO=CTOD("31/12/2100")
                    //                sql execute wnrc1
                  Endif
                  MWOBS="Cliente- "+ZPFCNPJ+" Assinante: "+mzcnpjcpf+"-"+mznome
                  If MWVENCTO<ctod(ZTCONTR)
                    If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                      ENVIAEMAIL(alltrim(emailquemenviou),"Procuração Vencida","Procuração Vencida - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                    Endif
                    SELE 38
                    ADIREG(0)
                    REPLACE CODIf WITH CORRETOR->SQL_ROWID
                    REPLACE USUARIO WITH quemenviou
                    REPLACE REFERENCIA WITH MWOBS
                    REPLACE MOTIVO WITH "Procuração Vencida"
                    SELE 31
                  Endif
                  If AT("@",EMAIL)=0
                    If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                      ENVIAEMAIL(alltrim(emailquemenviou),"Email Inválido","Email Inválido - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                    Endif
                    SELE 38
                    ADIREG(0)
                    REPLACE CODIf WITH CORRETOR->SQL_ROWID
                    REPLACE USUARIO WITH quemenviou
                    REPLACE REFERENCIA WITH MWOBS
                    REPLACE MOTIVO WITH "Email Inválido"
                    SELE 31
                  Endif
                  If SEENVIAEMAIL="S" .And. ZEAIS<=MWVALORMAX .And. MWVENCTO>=ctod(ZTCONTR) .And. AT(STRZERO(ZPDOCTO,10),MWTIPODOCUMENTO)#0 .And. at("@",EMAIL)#0
                    webcorpo1=webcorpo
                    If W01TPDOCTO=1 .And. !EMPTY(WEBCORPODOCTO1)
                      webcorpo1=webcorpodocto1
                    Endif
                    If W01TPDOCTO=2 .And. !EMPTY(WEBCORPODOCTO2)
                      webcorpo1=webcorpodocto2
                    Endif
                    If W01TPDOCTO=3 .And. !EMPTY(WEBCORPODOCTO3)
                      webcorpo1=webcorpodocto3
                    Endif
                    If W01TPDOCTO=4 .And. !EMPTY(WEBCORPODOCTO4)
                      webcorpo1=webcorpodocto4
                    Endif
                    If W01TPDOCTO=5 .And. !EMPTY(WEBCORPODOCTO5)
                      webcorpo1=webcorpodocto5
                    Endif
                    If W01TPDOCTO=6 .And. !EMPTY(WEBCORPODOCTO6)
                      webcorpo1=webcorpodocto6
                    Endif
                    If W01TPDOCTO=7 .And. !EMPTY(WEBCORPODOCTO7)
                      webcorpo1=webcorpodocto7
                    Endif
                    If W01TPDOCTO=8 .And. !EMPTY(WEBCORPODOCTO8)
                      webcorpo1=webcorpodocto8
                    Endif
                    webcorpo1=Strtran(webcorpo1,";;a02nome;:",mznome)
                    //          ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,ZRQPDF,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                    ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                    enviaaviso("Há documento(s) a ser(em) assinado(s)",cnpjcpf)
                    DECLARE TX1[1]
                    TX1[1]="Documento nr: "+STRZERO(ZRCONTR,19)+" DE: "+ZTCONTR+" enviado aviso para: "+alltrim(NOME)+" email: "+alltrim(EMAIL)
                    LOGFILE(WDADOS,TX1)
                  Endif
                Endif
              Endif
              sele 31
              SKIP
            ENDDO
          Endif
          // AQUI INTERMEDIARIO
          If WCODINT#0
            SELE 31
            SQL FILTER TO 'CODCLIIf LIKE "%'+STRZERO(WCODINT,10)+'%"'
            GO TOP
            Do While !EOF()
              WNRC1="SELECT sql_rowid,nome from assinantes where cnpjcpf="+cnpjcpf
              mz=SQLArrayAssoc(WNRC1)
              If len(mz)>0 .And. VAL(WIF)>0
                WNRC1="SELECT sql_rowid,vencto,valormax,tipodocumento FROM vencimentos where sql_deleted='F' and codassinante="+mz[1,1]+" and codcliif="+WIF
                mznome=mz[1,2]
                mzcnpjcpf=cnpjcpf
                mz=SQLArrayAssoc(WNRC1)
                If LEN(MZ)>0
                  MWSQL_ROWID=MZ[1,1]
                  MWVENCTO=ALLTRIM(MZ[1,2])
                  MWVENCTO=CTOD(RIGHT(MWVENCTO,2)+"/"+SUBSTR(MWVENCTO,6,2)+"/"+LEFT(MWVENCTO,4))
                  MWVALORMAX=VAL(MZ[1,3])
                  MWTIPODOCUMENTO=MZ[1,4]
                  If MWVENCTO=CTOD("")
                    //                wnrc1="update vencimentos set vencto="+dtom(ctod("31/12/2100"))+" where sql_rowid="+mwsql_rowid
                    //                MWVENCTO=CTOD("31/12/2100")
                    //                sql execute wnrc1
                  Endif
                  MWOBS="Intermediario- "+ZPFCNPJ+" Assinante: "+mzcnpjcpf+"-"+mznome
                  If MWVENCTO<=ctod(ZTCONTR)
                    If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                      ENVIAEMAIL(alltrim(emailquemenviou),"Procuração Vencida","Procuração Vencida - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                    Endif
                    SELE 38
                    ADIREG(0)
                    REPLACE CODIf WITH CORRETOR->SQL_ROWID
                    REPLACE USUARIO WITH quemenviou
                    REPLACE REFERENCIA WITH MWOBS
                    REPLACE MOTIVO WITH "Procuração Vencida"
                    SELE 31
                  Endif
                  If AT("@",EMAIL)=0
                    If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                      ENVIAEMAIL(alltrim(emailquemenviou),"Email Inválido","Email Inválido - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                    Endif
                    SELE 38
                    ADIREG(0)
                    REPLACE CODIf WITH CORRETOR->SQL_ROWID
                    REPLACE USUARIO WITH quemenviou
                    REPLACE REFERENCIA WITH MWOBS
                    REPLACE MOTIVO WITH "Email Inválido"
                    SELE 31
                  Endif
                  If SEENVIAEMAIL="S" .And. ZEAIS<=MWVALORMAX .And. MWVENCTO>=ctod(ZTCONTR) .And. AT(STRZERO(ZPDOCTO,10),MWTIPODOCUMENTO)#0 .And. at("@",EMAIL)#0
                    //                ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO,ZRQPDF,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                    webcorpo1=Strtran(webcorpo,";;a02nome;:",mznome)
                    ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                    enviaaviso("Há documento(s) a ser(em) assinado(s)",cnpjcpf)
                    DECLARE TX1[1]
                    TX1[1]="Documento nr: "+STRZERO(ZRCONTR,19)+" DE: "+ZTCONTR+" enviado aviso para: "+alltrim(NOME)+" email: "+alltrim(EMAIL)
                    LOGFILE(WDADOS,TX1)
                  Endif
                Endif
              Endif
              sele 31
              SKIP
            ENDDO
          Endif
        Endif
        NEXT OI

      Else

        // leitura PDF
        FOR OI=1 TO LEN(VLABEL)
          ZNOMECLI=""
          compltipo=""
          ZPFCNPJ=""
          zrcontr=0
          ztcontr=ctod("")
          MWAC10K=.F.
          tipodocumento=0
          ZNPJINT=""
          wcodint=0
          wcodcliif=0
          zeais=0
          ZOEDA=""
          ZOESTRAN=0
          ZBS=""
          E_HARD=.F.
          SELE 32
          GO BOTT
          SKIP
          DECLARE _VAR1[FCOUNT()]
          AFIELDS(_VAR1)
          FOR t = 1 TO FCOUNT()
            SysRefresh()
            _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
            _var5 = _var1 [t]
            &_var4 = &_var5
          NEXT t
          RELEASE _VAR1
          NARQ=WEBLOCAL1+MAISNOME+ALLTRIM(VLABEL[OI])
          NARQ1=LEFT(NARQ,LEN(NARQ)-3)+"TXT"
          If !FILE(NARQ1)
          LOOP
        Endif
        SELE 33
        ZAP
        APPEND FROM &NARQ1 SDF
        ZNCLUIDO=DTOC(DATE())+" "+TIME()
        USERINC="WORKER ARQUIVO "+ALLTRIM(VLABEL[OI])
        GO 1
        If VAL(ZPFCNPJ)=0
          ZPFCNPJ="0"
        Endif
        If AT("Contrato de Cambio",LINHA)#0 .OR. AT("Contrato de Câmbio",LINHA)#0 .OR. AT("Contrato de câmbio",LINHA)#0
          SKIP 2
          If AT("Tipo de Contrato",LINHA)#0
            SKIP 4
          Endif
          If EMPTY(LINHA)
            SKIP
          Endif
          If AT("Compra",LINHA)#0
            TPDOCTO="COMPRA"
          Else
            TPDOCTO="VENDA"
          Endif
          If UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\"
            GO 4
            If AT("Compra",LINHA)#0
              TPDOCTO="COMPRA"
            Else
              TPDOCTO="VENDA"
            Endif
          Endif
          If AT("Número do Contrato de Câmbio",LINHA)#0
            SKIP 3
            If AT("Compra",LINHA)#0
              TPDOCTO="COMPRA"
            Else
              TPDOCTO="VENDA"
            Endif
          Endif
          If AT("Cancelamento",LINHA)>0
            compltipo="-Cancela"
          Endif
          If AT("Altera",LINHA)>0
            compltipo="-Altera"
          Endif
          If UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA" .OR. UPPER(M1CAMINHO)="C:\NOTAS\TREVISO\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. lower(m1caminho)="c:\notas\bconfide\bmf\assina\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\" .or. upper(M1CAMINHO)="C:\NOTAS\BT\BMF\ASSINA\"
            go 16
          Endif
          If UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA"
            go 21
          Endif
          ANLIN=AT("CNPJ",LINHA)
          ZIPO="J"
          If ANLIN=0
            ANLIN=AT("CPF",LINHA)
            ZIPO="F"
          Endif
          If ANLIN=0
            ANLIN=1
          Endif
          SKIP
          NOMEG=.F.
          If UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\"  .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. lower(m1caminho)="c:\notas\bconfide\bmf\assina\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            GO 19
            If AT("Endereço",LINHA)#0
              GO 18
              If !EMPTY(LINHA)
                GO 19
                DELE
              Else
                GO 18
                DELE
              Endif
              PACK
            Endif
            GO 17
          Endif
          ZPFCNPJ=SUBSTR(LINHA,ANLIN,20)
          If UPPER(M1CAMINHO)="C:\NOTAS\FAIR\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONEXION\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\TURISCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONSEG\BMF\ASSINA\"
            go 18
            wli1=Strtran(linha," ","              ")
            ZPFCNPJ=RIGHT(ALLTRIM(wli1),14)
          Endif
          ZPFCNPJ=Strtran(ZPFCNPJ,".","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
          ZPFCNPJ=Strtran(ZPFCNPJ," ","")
          GO 3
          If upper(M1CAMINHO)="C:\NOTAS\BT\BMF\ASSINA\"
            go 3
            ZRCONTR=VAL(SUBSTR(LINHA,52,16))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\"  .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            GO 2
            WAT=AT("Número do contrato",LINHA)
            GO 4
            ZRCONTR=VAL(SUBSTR(LINHA,WAT,9))
            wat=at("/",LINHA)-2
            //        If LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\"
            //          WA2=AT("Contratação",LINHA)+12
            //          ZRCONTR=VAL(SUBSTR(LINHA,WA2,50))
            //        Endif
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
            //        If at("Cancelamento",linha)#0 .And. (UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\"  .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\") .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            //          ZRCONTR=val(SUBSTR(LINHA,70,11))
            //        Endif
          ElseIf M1CAMINHO="C:\NOTAS\TREVISO\BMF\ASSINA\"
            GO 4
            ZRCONTR=VAL(SUBSTR(LINHA,59,12))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA\"
            GO 4
            ZRCONTR=VAL(SUBSTR(LINHA,55,18))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
          Else
            GO 3
            ZRCONTR=VAL(SUBSTR(LINHA,56,14))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
            If AT("CAM010",LINHA)#0 .OR. AT("CAM009",LINHA)#0
              ZPFCNPJ=""
            Endif
            If AT("Cancelamento",LINHA)>0
              compltipo="-Cancela"
            Endif
            If AT("Altera",LINHA)>0
              compltipo="-Altera"
            Endif
          Endif
          If M1CAMINHO="C:\NOTAS\TREVISO\BMF\ASSINA\"
            GO 29
            If AT("Código",LINHA)=0
              GO 26
            Endif
            SKIP
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA\"
            GO 37
            If EMPTY(LINHA)
              GO 36
            Endif
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\"  .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            GO 27
            If NOMEG
              SKIP
            Endif
          Else
            GO 28
          Endif
          ZOEDA=ALLTRIM(LEFT(LINHA,4))
          If UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA\"  .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            GO 27
            ZOESTRAN=SUBSTR(LINHA,19,20)
            ZOESTRAN=LEFT(ZOESTRAN,AT("(",ZOESTRAN)-1)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 31
            If SUBSTR(LINHA,19,1)=":"
              ZEAIS=SUBSTR(LINHA,20,20)
            Else
              ZEAIS=SUBSTR(LINHA,19,20)
            Endif
            ZEAIS=Strtran(ZEAIS,"R$:","")
            ZEAIS=LEFT(ZEAIS,AT("(",ZEAIS)-1)
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
          ElseIf upper(M1CAMINHO)="C:\NOTAS\BT\BMF\ASSINA\"
            GO 28
            ZOEDA=ALLTRIM(LEFT(LINHA,4))
            GO 27
            ZOESTRAN=SUBSTR(LINHA,25,26)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 30
            ZEAIS=SUBSTR(LINHA,25,26)
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
            GO 17
            ZNOMECLI=LEFT(LINHA,68)
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA\"
            GO 37
            ZOEDA=ALLTRIM(LEFT(LINHA,4))
            ZOESTRAN=SUBSTR(LINHA,22,26)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            If ZOESTRAN=0
              GO 36
              ZOESTRAN=SUBSTR(LINHA,22,26)
              ZOESTRAN=Strtran(ZOESTRAN,".","")
              ZOESTRAN=Strtran(ZOESTRAN," ","")
              ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            Endif
            GO 42
            ZEAIS=SUBSTR(LINHA,22,26)
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
            If ZEAIS=0
              GO 41
              ZEAIS=SUBSTR(LINHA,22,26)
              ZEAIS=Strtran(ZEAIS,".","")
              ZEAIS=Strtran(ZEAIS," ","")
              ZEAIS=VAL(Strtran(ZEAIS,",","."))
            Endif
          ElseIf UPPER(M1CAMINHO)="C:\NOTAS\TREVISO\BMF\ASSINA\"
            GO 29
            If AT("Código",LINHA)=0
              GO 26
            Endif
            SKIP
            FND=AT("(",LINHA)
            FND=FND-20
            If FND<1
              FND=1
            Endif
            ZOESTRAN=SUBSTR(LINHA,20,FND)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 29
            If AT("Código",LINHA)=0
              GO 31
            Else
              GO 34
            Endif
            FND=AT("(",LINHA)
            mrea=at("R$:",LINHA)+3
            If MREA=0
              MREA=1
            Endif
            ZEAIS=SUBSTR(LINHA,MREA,FND)
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
          Else
            GO 28
            ZOESTRAN=SUBSTR(LINHA,29,26)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 32
            zlinha=Strtran(linha,",",".")
            If VAL(ZLINHA)=0
              GO 31
              zlinha=Strtran(linha,",",".")
            Endif
            If VAL(ZLINHA)=0
              GO 30
              zlinha=Strtran(linha,",",".")
            Endif
            ZEAIS=SUBSTR(LINHA,29,26)
            ZEAIS=Strtran(ZEAIS,"R","")
            ZEAIS=Strtran(ZEAIS,"$","")
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
          Endif
          ZPFCNPJ=ALLTRIM(ZPFCNPJ)
          If LEN(ZPFCNPJ)>14
            ZPFCNPJ=RIGHT(ZPFCNPJ,14)
          Endif
          // Aqui é o do intermediario (exatus)
          GO 21
          ANLIN=AT("CNPJ",LINHA)
          GO 22
          ZNPJINT=SUBSTR(LINHA,ANLIN,20)
          ZNPJINT=Strtran(ZNPJINT,".","")
          ZNPJINT=Strtran(ZNPJINT,"/","")
          ZNPJINT=Strtran(ZNPJINT,"-","")
          ZNPJINT=Strtran(ZNPJINT," ","")
          ZNPJINT=STRZERO(VAL(ZNPJINT),14)
          If UPPER(M1CAMINHO)="C:\NOTAS\AGK\BMF\ASSINA\" .OR. upper(M1CAMINHO)="C:\NOTAS\BT\BMF\ASSINA\" .OR. upper(M1CAMINHO)="C:\NOTAS\LEVYCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\INTERCAM\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\CONECTA\BMF\ASSINA\"
            ZNPJINT=""
          Endif
          If UPPER(M1CAMINHO)="C:\NOTAS\MSBANK\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\NOVOMUNDO\BMF\ASSINA\" .OR. LOWER(m1caminho)="c:\notas\bconfide\bmf\assina\"
            GO 23
            ANLIN=AT("CNPJ",LINHA)
            GO 24
            ZNPJINT=SUBSTR(LINHA,ANLIN,20)
            ZNPJINT=Strtran(ZNPJINT,".","")
            ZNPJINT=Strtran(ZNPJINT,"/","")
            ZNPJINT=Strtran(ZNPJINT,"-","")
            ZNPJINT=Strtran(ZNPJINT," ","")
            ZNPJINT=STRZERO(VAL(ZNPJINT),14)
          Endif
          If VAL(ZNPJINT)#0
            SELE 34
            SQL FILTER TO "CNPJCPF="+ZNPJINT
            GO TOP
            WCODINT=0
            Do While !EOF()
              If wcodif<>CODIF
                SKIP
              Else
                EXIT
              Endif
            ENDDO
            If val(CNPJCPF)=val(ZNPJINT)
              WCODINT=CLIIF->SQL_ROWID
            Endif
          Endif
        Endif
        sele 33
        If VAL(ZPFCNPJ)=0 .And. AT("Registro da Operação no Banco Central",LINHA)#0
          ZRCONTR=VAL(SUBSTR(LINHA,74,9))
          WCODINT=0
          ZNPJINT=STRZERO(0,14)
          GO 5
          ZTCONTR=CTOD(RIGHT(ALLTRIM(LINHA),10))
          GO 2
          If AT("X ] venda",LINHA)#0
            TPDOCTO="VENDA"
          Else
            TPDOCTO="COMPRA"
          Endif
          go 17
          zpfcnpj=right(trim(linha),20)
          ZPFCNPJ=Strtran(ZPFCNPJ,".","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
          ZPFCNPJ=Strtran(ZPFCNPJ," ","")
          go 25
          ZEAIS=SUBSTR(LINHA,111,28)
          ZEAIS=Strtran(ZEAIS,"R","")
          ZEAIS=Strtran(ZEAIS,"$","")
          ZEAIS=Strtran(ZEAIS,".","")
          ZEAIS=Strtran(ZEAIS," ","")
          ZEAIS=VAL(Strtran(ZEAIS,",","."))
          GO 24
          ZOESTRAN=SUBSTR(LINHA,111,28)
          ZOESTRAN=Strtran(ZOESTRAN,".","")
          ZOESTRAN=Strtran(ZOESTRAN," ","")
          ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
          ZOEDA=SUBSTR(LINHA,27,3)
        Endif
        If UPPER(M1CAMINHO)="C:\NOTAS\LEVYCAM\BMF\ASSINA\"
          ZLAYOUT=0
          GO 3
          If EMPTY(LINHA)
            ZLAYOUT=1
          Else
            ZLAYOUT=2
          Endif
          If ZLAYOUT=1
            ZNPJINT=""
            GO 4
            If AT("Venda",LINHA)#0
              TPDOCTO="VENDA"
            Else
              TPDOCTO="COMPRA"
            Endif
            ZRCONTR=VAL(SUBSTR(LINHA,63,14))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
            go 37
            ACRESC=0
            If EMPTY(LINHA)
              ACRESC=1
              skip
            Endif
            ZOEDA=LEFT(alltrim(LINHA),3)
            If ACRESC=0
              skip -1
            Endif
            ZOESTRAN=SUBSTR(LINHA,32,18)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 41
            If AT("Taxa Cambial",LINHA)#0
              GO 43
            Endif
            ZEAIS=SUBSTR(LINHA,31,19)
            ZEAIS=Strtran(ZEAIS,"R","")
            ZEAIS=Strtran(ZEAIS,"$","")
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
            GO 22
            zpfcnpj=right(trim(linha),20)
            ZPFCNPJ=Strtran(ZPFCNPJ,".","")
            ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
            ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
            ZPFCNPJ=Strtran(ZPFCNPJ," ","")
          Else
            ZNPJINT=""
            GO 4
            If AT("Venda",LINHA)#0
              TPDOCTO="VENDA"
            Else
              TPDOCTO="COMPRA"
            Endif
            ZRCONTR=VAL(SUBSTR(LINHA,55,14))
            wat=at("/",LINHA)-2
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
            go 33
            ZOEDA=LEFT(alltrim(LINHA),3)
            skip -1
            ZOESTRAN=ALLTRIM(LINHA)
            ZOESTRAN=LEFT(ZOESTRAN,AT("(",ZOESTRAN)-1)
            ZOESTRAN=Strtran(ZOESTRAN,".","")
            ZOESTRAN=Strtran(ZOESTRAN," ","")
            ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            GO 37
            ZEAIS=ALLTRIM(SUBSTR(LINHA,24,LEN(LINHA)))
            ZEAIS=LEFT(ZEAIS,AT("(",ZEAIS)-1)
            ZEAIS=Strtran(ZEAIS,"R","")
            ZEAIS=Strtran(ZEAIS,"$","")
            ZEAIS=Strtran(ZEAIS,".","")
            ZEAIS=Strtran(ZEAIS," ","")
            ZEAIS=VAL(Strtran(ZEAIS,",","."))
            GO 20
            zpfcnpj=right(trim(linha),20)
            ZPFCNPJ=Strtran(ZPFCNPJ,".","")
            ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
            ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
            ZPFCNPJ=Strtran(ZPFCNPJ," ","")
          Endif
        Endif
        If UPPER(M1CAMINHO)="C:\NOTAS\CORREP\BMF\ASSINA\" .or. UPPER(M1CAMINHO)="C:\NOTAS\SADOC\BMF\ASSINA\" .OR. UPPER(M1CAMINHO)="C:\NOTAS\EXIM\BMF\ASSINA\"
          ZNPJINT=""
          GO 3
          If AT("Venda",LINHA)#0
            TPDOCTO="VENDA"
          Else
            TPDOCTO="COMPRA"
          Endif
          ZRCONTR=VAL(SUBSTR(LINHA,54,14))
          wat=at("/",LINHA)-2
          ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
          go 32
          ZOEDA=LEFT(alltrim(LINHA),3)
          ZOESTRAN=SUBSTR(LINHA,38,18)
          ZOESTRAN=Strtran(ZOESTRAN,".","")
          ZOESTRAN=Strtran(ZOESTRAN," ","")
          ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
          GO 35
          ZEAIS=SUBSTR(LINHA,38,18)
          ZEAIS=Strtran(ZEAIS,"R","")
          ZEAIS=Strtran(ZEAIS,"$","")
          ZEAIS=Strtran(ZEAIS,".","")
          ZEAIS=Strtran(ZEAIS," ","")
          ZEAIS=VAL(Strtran(ZEAIS,",","."))
          GO 18
          zpfcnpj=right(trim(linha),20)
          ZPFCNPJ=Strtran(ZPFCNPJ,".","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
          ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
          ZPFCNPJ=Strtran(ZPFCNPJ," ","")
        Endif
        // Hardcopy
        If VAL(ZPFCNPJ)=0 .And. LASTREC()>30
          SELE 33
          GO 3
          If AT("Relatório de Conferência",LINHA)<>0
            LOCATE FOR AT("Valor reais:",LINHA)<>0
            If !EOF()
              ZEAIS=SUBSTR(LINHA,17,60)
              ZEAIS=Strtran(ZEAIS," ","")
              ZEAIS=Strtran(ZEAIS,".","")
              ZEAIS=VAL(Strtran(ZEAIS,",","."))
            Endif
            GO 1
            LOCATE FOR AT("Valor moeda: ",LINHA)<>0
            If !EOF()
              ZOESTRAN=SUBSTR(LINHA,17,60)
              ZOESTRAN=Strtran(ZOESTRAN," ","")
              ZOESTRAN=Strtran(ZOESTRAN,".","")
              ZOESTRAN=VAL(Strtran(ZOESTRAN,",","."))
            Endif
            GO 1
            LOCATE FOR AT("Moeda estrangeira:",LINHA)<>0
            If !EOF()
              SKIP
              ZOEDA=SUBSTR(LINHA,12,3)
            Endif
            GO 1
            LOCATE FOR AT("Cliente:  ",LINHA)<>0
            If !EOF()
              wat=at("CNPJ / CPF:",LINHA)
              SKIP
              ZPFCNPJ=alltrim(SUBSTR(LINHA,wat,23))
              ZPFCNPJ=Strtran(ZPFCNPJ," ","")
              ZPFCNPJ=Strtran(ZPFCNPJ,".","")
              ZPFCNPJ=Strtran(ZPFCNPJ,"-","")
              ZPFCNPJ=Strtran(ZPFCNPJ,"/","")
            Endif
            GO 4
            wat=at("Fechamento:",LINHA)
            SKIP
            ZTCONTR=CTOD(SUBSTR(LINHA,WAT,10))
            If ZTCONTR=CTOD("")
              GO 2
              ZTCONTR=CTOD(RIGHT(ALLTRIM(LINHA),10))
            Endif
            E_HARD=.T.
            GO 5
            zrcontr=VAL(SUBSTR(LINHA,17,20))
            If zrcontr=0
              GO 5
              zrcontr=val(LEFT(LINHA,8))
            Endif
          Endif
        Endif
        WCODCLIIF=0



        //MSGINFO(ZNPJINT,"intermediação")
        //MSGINFO(ZEAIS,"REAIS")
        //mSGINFO(ZOESTRAN,"ESTRANGEIRA")
        //MSGINFO(ZOEDA,"MOEDA")
        //MSGINFO(ZPFCNPJ,"CLIENTE")
        //MSGINFO(TPDOCTO,"TIPO")
        //MSGINFO(ZTCONTR,"DATA")
        //MSGINFO(zrcontr,"NR.CONTRATO")

        //LOOP
        //QUIT

        If UPPER(M1CAMINHO)="C:\NOTAS\BT\BMF\ASSINA\"
          ZNPJINT=""
        Endif
        janaoenco=.f.
        zbs=strzero(zrcontr,10)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+compltipo
        If E_HARD
          zbs="CF"+strzero(zrcontr,9)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))
        Endif
        If VAL(ZPFCNPJ)=0
          SELE 38
          ADIREG(0)
          REPLACE CODIf WITH CORRETOR->SQL_ROWID
          REPLACE USUARIO WITH quemenviou
          REPLACE REFERENCIA WITH ZBS
          REPLACE MOTIVO WITH "Documento não Reconhecido"
          janaoenco=.t.
          SELE 32
        Endif
        WIF="0"
        If VAL(ZPFCNPJ)#0
          SELE 34
          WNRC="codif="+wcodif+" and CNPJCPF="+ZPFCNPJ
          SQL FILTER TO WNRC
          GO TOP
          MWAC10K=.F.
          If val(CNPJCPF)=val(ZPFCNPJ)
            WCODCLIIF=CLIIF->SQL_ROWID
            MWNOVO=CLIIF->NOVO
            WIF=STRZERO(WCODCLIIF,10)
            MWAC10K=CLIIF->AC10K
            If MWAC10K="T"
              MWAC10K=.T.
            Else
              MWAC10K=.F.
            Endif
          Endif
        Endif
        zbs=strzero(zrcontr,10)+" - "+ZPFCNPJ+" - "+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+compltipo
        If E_HARD
          zbs="CF"+strzero(zrcontr,9)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))
        Endif
        //CONTRATOS NENORES QUE 10.000 USD
        WBPAR=.F.
        If UPPER(ALLTRIM(ZOEDA))="USD" .OR. VAL(ZOEDA)=220
          If MWAC10K=.F. .And. ZOESTRAN<10000 .And. E_HARD=.F.
            WBPAR=.T.
          Endif
        Else
          WNRC="SELECT venda from paridade where moeda=220 and data<"+dtom(ZTCONTR)+" ORDER BY DATA DESC LIMIT 0,1"
          mz=SQLArrayAssoc(WNRC)
          If LEN(mz)=0
            WBPAR=.T.
          Else
            If (ZEAIS/VAL(MZ[1,1]) < 10000) .And. MWAC10K=.F. .And. E_HARD=.F.
              WBPAR=.T.
            Endif
          Endif
        Endif
        If WBPAR=.T. .And. !E_HARD
          If !EMPTY(emailquemenviou) .And. SEENVIAEMAIL="S"
            ENVIAEMAIL(alltrim(emailquemenviou),"Contrato menor que 10.000 USD","Contrato menor que 10.000 USD",NARQ,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
          Endif
          SELE 38
          ADIREG(0)
          REPLACE CODIf WITH CORRETOR->SQL_ROWID
          REPLACE USUARIO WITH quemenviou
          REPLACE REFERENCIA WITH ZBS
          REPLACE MOTIVO WITH "Contrato menor que 10.000 USD"
          janaoenco=.t.
          WCODCLIIF=0
          SELE 32
        Endif
        If WCODCLIIF#0 .And. !janaoenco
          SELE 32
          If compltipo#"-Altera" .OR. E_HARD
            WNRC="SELECT sql_rowid FROM documentos WHERE sql_deleted='F' and data="+dtom(ZTCONTR)+" and obs='"+strzero(zrcontr,10)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+compltipo+"'"
            If E_HARD
              WNRC="SELECT sql_rowid FROM documentos WHERE sql_deleted='F' and data="+dtom(ZTCONTR)+" and obs='CF"+strzero(zrcontr,9)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+"'"
            Endif
            mz=SQLArrayAssoc(WNRC)
            If LEN(mz)#0
              WCODCLIIF=0
              SELE 38
              ADIREG(0)
              REPLACE CODIf WITH CORRETOR->SQL_ROWID
              REPLACE USUARIO WITH quemenviou
              REPLACE REFERENCIA WITH ZBS
              REPLACE MOTIVO WITH "Já Cadastrado"
              janaoenco=.t.
              SELE 32
            Endif
          Endif
        Endif
        If WCODCLIIF=0 .And. SEENVIAEMAIL="S" .And. !janaoenco
          If !EMPTY(emailquemenviou)
            ENVIAEMAIL(alltrim(emailquemenviou),"Documento não processado","Documento não Processado",NARQ,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
          Endif
        Endif
        If wcodcliif=0 .And. !janaoenco
          SELE 38
          ADIREG(0)
          REPLACE CODIf WITH CORRETOR->SQL_ROWID
          REPLACE USUARIO WITH quemenviou
          REPLACE REFERENCIA WITH ZBS
          REPLACE MOTIVO WITH "Assinante não cadastrado"
          janaoenco=.t.
          SELE 32
        Endif
        If WCODCLIIF#0
          SELE 35
          ZEQUENC=SEQUENCIA+1
          If REGLOCK(10)
            REPLACE SEQUENCIA WITH ZEQUENC
          Endif
          SELE 36
          BEGIN TRANSACTION
          ADIREG(0)
          REPLACE NOMEARQ WITH VLABEL[OI]
          REPLACE ENVIADO WITH DTOC(DATE())+"-"+TIME()
          REPLACE USUARIO WITH QUEMENVIOU
          REPLACE NOMEIf WITH WEBCOR
        END TRANSACTION
        COMMIT
        SELE 32
        NARQ=WEBLOCAL1+MAISNOME+ALLTRIM(VLABEL[OI])
        ZRQPDF=WEBLOCAL3+STRZERO(ZEQUENC,10)+".PDF"
        ZRQTXT=WEBLOCAL3+STRZERO(ZEQUENC,10)+".TXT"
        ZRQPDF1=WEBLOCAL3+STRZERO(ZEQUENC,10)+".PDF1"
        ZARQH=WEBLOCAL1+STRZERO(ZEQUENC,10)+".PDF"
        ZARQH1=WEBLOCAL1+STRZERO(ZEQUENC,10)+".PDF1"
        cAnexo=LOWER(NARQ)
        wcAnexo=Strtran(cAnexo,".pdf","_C.pdf")
        comando=WCORRENTE+"\gswin32c -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/prepress -dNOPAUSE -dQUIET -dBATCH -sOutputFile="+wcAnexo+" "+cAnexo
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+weblocal+"pausa.txt"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal+"compacta.bat",mzlinha)
        MEMOWRIT(weblocal+"pausa.txt",MZLINHA)
        WAITRUN(weblocal+"compacta.bat",0)
        Do While FILE(weblocal+"pausa.txt")
          SysRefresh()
        ENDDO
        wtam1=DIRECTORY(cAnexo)
        wtam1=wtam1[1][2]
        wtam2=DIRECTORY(wcAnexo)
        wtam2=wtam2[1][2]
        If wtam1<wtam2
          Erase &wcAnexo
        Else
          Erase &cAnexo
          FRENAME(wcAnexo,cAnexo)
        Endif
        cAnexo=weblocal+"compacta.bat"
        Erase &cAnexo
        cAnexo=weblocal+"pausa.txt"
        Erase &cAnexo
        WNRC='SELECT MD5("'+DTOC(ZTCONTR)+STRZERO(ZEQUENC,10)+'4ss1n4")'
        MZ=SQLArrayAssoc(WNRC)
        WCHASH=MZ[1,1]
        arqQR=WEBLOCAL1+"qr.BMP"
        Erase &arqQR
        arqQR=WEBLOCAL1+"qr.PNG"
        Erase &arqQR
        texto="https://assinanet.exatus.net/validaassina.asp?dados="+WCHASH+"|"+STRZERO(ZEQUENC,10)+"|"+DTOC(ZTCONTR)
        qrDLL:=LoadLibrary("QRCodelib.Dll" )
        #Define DC_CALL_STD 0x0020
        nResp:=DllCall(qrDLL, DC_CALL_STD, "FastQRCode", texto, arqQR)
        FreeLibrary(qrDLL)

        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        Erase &ZARQH
        oReg:Set("PDFFilename",ZARQH)
        oReg:Close()
        PRINT oPrn NAME "HASH" TO TEMPDF1
        DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
        nRowStep = oPrn:nVertRes() / 66
        nColStep = oPrn:nHorzRes() / 80
        PAGE
        oprn:cmsay(28.5,1.0,"Controle AssinaNet: "+STRZERO(ZEQUENC,10)+" "+WCHASH,oFont,,nRGB(0,0,0),,0)
        If !E_HARD
          oprn:cmsay(28.8,1.0,"Documento Assinado Eletronicamente",oFont,nRGB(0,0,0),,0)
        Else
          oprn:cmsay(28.8,1.0,"Declaro estar de acordo com todas informações deste documento",oFont,nRGB(0,0,0),,0)
        Endif
        If file(arqQR)
          sayBitmap(27.7,12,1.8,1.8,arqqr,oPrn)
        Endif
        ENDPAGE
        ENDPRINT
        tini=time()
        Do While .T.
          CURSORWAIT()
          DIF=TIMEASSECONDS(TIMEDIFF(tini,TIME()))
          If dif>60
            EXIT
          Endif
          If (nHandle:=FOpen(ZARQH,FO_READ+FO_EXCLUSIVE))!=-1
            fClose(nHandle)
            EXIT
          Else
            fClose(nHandle)
          Endif
        ENDDO
        comando=WCORRENTE+"\PDFTK "+NARQ+" stamp "+ZARQH+"  output "+ZARQH1+chr(13)+chr(10)
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
        MEMOWRIT(weblocal1+"AGUARDE.TXT","HASH"+MZLINHA)
        CURSORWAIT()
        WAITRUN(weblocal1+"FZ.BAT",0)
        CURSORWAIT()
        Do While FILE(weblocal1+"AGUARDE.TXT")
          SysRefresh()
        ENDDO
        Erase &NARQ
        Erase &ZARQH
        FRENAME(ZARQH1,NARQ)
        comando=WCORRENTE+"\PDFTK "+NARQ+" background c:\exatus\eletro.pdf output "+zrqpdf+chr(13)+chr(10)
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
        MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina contracapa"+MZLINHA)
        CURSORWAIT()
        WAITRUN(weblocal1+"FZ.BAT",0)
        CURSORWAIT()
        Do While FILE(weblocal1+"AGUARDE.TXT")
          SysRefresh()
        ENDDO
        Erase &NARQ
        Erase &NARQ1
        ZPDOCTO=4
        If !E_HARD
          If TPDOCTO="COMPRA"
            ZPDOCTO=1
          ElseIf TPDOCTO="VENDA"
            ZPDOCTO=2
          Endif
        Else
          TPDOCTO="CF"
          ZPDOCTO=3
        Endif
        sele 32
        BEGIN TRANSACTION
        ADIREG(0)
        REPLACE DATA WITH ZTCONTR
        REPLACE TIPODOCUMENTO WITH ZPDOCTO
        REPLACE CAMINHO WITH ZRQPDF
        REPLACE CAMINHOTXT WITH ZRQPDF
        REPLACE CODIf WITH CORRETOR->SQL_ROWID
        REPLACE CODCORR WITH WCODINT
        REPLACE CODCLIIf WITH WCODCLIIF
        REPLACE VALOR WITH ZEAIS
        If !E_HARD
          REPLACE OBS WITH strzero(zrcontr,10)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+compltipo
        Else
          REPLACE OBS WITH "CF"+strzero(zrcontr,9)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))
        Endif
        REPLACE INCLUIDO WITH DTOC(DATE())+"-"+TIME()
        REPLACE CHASH WITH WCHASH
        REPLACE USERINC WITH quemenviou
        REPLACE NOVO WITH "S"
      END TRANSACTION
        COMMIT
        DECLARE TX1[1]
        If !E_HARD
          TX1[1]="Documento nr: "+strzero(zrcontr,10)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+compltipo+" - Incluido no Assinanet"
        Else
          WSQLNRDOCTO=sql_rowid
          FOR qtd_ass=1 to wqtd_ass_if
            SQL EXECUTE "INSERT INTO assinou (nome,cnpjcpf,documento,origem,back1,back2) VALUES ( 'ASSINATURA','99999999999',"+STRZERO(WSQLNRDOCTO,11)+",'I','S','S');"
            SQL EXECUTE "COMMIT"
          NEXT qtd_ass
          TX1[1]="CONFORMIDADE E FINALIDADE nr: "+strzero(zrcontr,9)+"-"+ZOEDA+" "+alltrim(TRANSFORM(ZOESTRAN,"@E 9,999,999,999.99"))+" - Incluido no Assinanet"
        Endif
        LOGFILE(WDADOS,TX1)
        If WCODCLIIF#0
          SELE 31
          SQL FILTER TO 'CODCLIIf LIKE "%'+WIF+'%"'
          GO TOP
          Do While !EOF()
            WNRC1="SELECT sql_rowid,nome from assinantes where cnpjcpf="+cnpjcpf
            mz=SQLArrayAssoc(WNRC1)
            If len(mz)>0
              WNRC1="SELECT sql_rowid,vencto,valormax,tipodocumento FROM vencimentos where sql_deleted='F' and codassinante="+mz[1,1]+" and codcliif="+WIF
              mznome=mz[1,2]
              mzcnpjcpf=cnpjcpf
              mz=SQLArrayAssoc(WNRC1)
              If LEN(MZ)>0
                MWSQL_ROWID=MZ[1,1]
                MWVENCTO=ALLTRIM(MZ[1,2])
                MWVENCTO=CTOD(RIGHT(MWVENCTO,2)+"/"+SUBSTR(MWVENCTO,6,2)+"/"+LEFT(MWVENCTO,4))
                MWVALORMAX=VAL(MZ[1,3])
                MWTIPODOCUMENTO=MZ[1,4]
                If MWVENCTO=CTOD("")
                  //                wnrc1="update vencimentos set vencto="+dtom(ctod("31/12/2100"))+" where sql_rowid="+mwsql_rowid
                  //                MWVENCTO=CTOD("31/12/2100")
                  //                sql execute wnrc1
                Endif
                MWOBS="Cliente: "+ZPFCNPJ+" Assinante: "+mzcnpjcpf+"-"+mznome
                If MWVENCTO<=ZTCONTR
                  If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                    ENVIAEMAIL(alltrim(emailquemenviou),"Procuração Vencida","Procuração Vencida - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                  Endif
                  SELE 38
                  ADIREG(0)
                  REPLACE CODIf WITH CORRETOR->SQL_ROWID
                  REPLACE USUARIO WITH quemenviou
                  REPLACE REFERENCIA WITH MWOBS
                  REPLACE MOTIVO WITH "Procuração Vencida"
                  SELE 31
                Endif
                If AT("@",EMAIL)=0
                  If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                    ENVIAEMAIL(alltrim(emailquemenviou),"Email Inválido","Email Inválido - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                  Endif
                  SELE 38
                  ADIREG(0)
                  REPLACE CODIf WITH CORRETOR->SQL_ROWID
                  REPLACE USUARIO WITH quemenviou
                  REPLACE REFERENCIA WITH MWOBS
                  REPLACE MOTIVO WITH "Email Inválido"
                  SELE 31
                Endif
                If SEENVIAEMAIL="S" .And. ZEAIS<=MWVALORMAX .And. MWVENCTO>=ZTCONTR .And. AT(STRZERO(ZPDOCTO,10),MWTIPODOCUMENTO)#0 .And. at("@",EMAIL)#0
                  webcorpo1=webcorpo
                  If ZPDOCTO=1 .And. !EMPTY(WEBCORPODOCTO1)
                    webcorpo1=webcorpodocto1
                  Endif
                  If ZPDOCTO=2 .And. !EMPTY(WEBCORPODOCTO2)
                    webcorpo1=webcorpodocto2
                  Endif
                  If ZPDOCTO=3 .And. !EMPTY(WEBCORPODOCTO3)
                    webcorpo1=webcorpodocto3
                  Endif
                  If ZPDOCTO=4 .And. !EMPTY(WEBCORPODOCTO4)
                    webcorpo1=webcorpodocto4
                  Endif
                  If ZPDOCTO=5 .And. !EMPTY(WEBCORPODOCTO5)
                    webcorpo1=webcorpodocto5
                  Endif
                  If ZPDOCTO=6 .And. !EMPTY(WEBCORPODOCTO6)
                    webcorpo1=webcorpodocto6
                  Endif
                  If ZPDOCTO=7 .And. !EMPTY(WEBCORPODOCTO7)
                    webcorpo1=webcorpodocto7
                  Endif
                  If ZPDOCTO=8 .And. !EMPTY(WEBCORPODOCTO8)
                    webcorpo1=webcorpodocto8
                  Endif
                  webcorpo1=Strtran(webcorpo1,";;a02nome;:",mznome)
                  //            ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,ZRQPDF,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR) // enviava o pdf junto, por seguranca, agora o cliente terá que acessar o sistema para verificar
                  ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                  enviaaviso("Há documento(s) a ser(em) assinado(s)",cnpjcpf)
                  DECLARE TX1[1]
                  TX1[1]="Documento nr: "+STRZERO(ZRCONTR,19)+" DE: "+DTOC(ZTCONTR)+" enviado aviso para: "+alltrim(NOME)+" email: "+alltrim(EMAIL)
                  LOGFILE(WDADOS,TX1)
                Endif
              Endif
            Endif
            SELE 31
            SKIP
          ENDDO
        Endif
        // AQUI INTERMEDIARIO
        If WCODINT#0
          SELE 31
          SQL FILTER TO 'CODCLIIf LIKE "%'+STRZERO(WCODINT,10)+'%"'
          GO TOP
          Do While !EOF()
            WNRC1="SELECT sql_rowid,nome from assinantes where cnpjcpf="+cnpjcpf
            mz=SQLArrayAssoc(WNRC1)
            If len(mz)>0
              WNRC1="SELECT sql_rowid,vencto,valormax,tipodocumento FROM vencimentos where sql_deleted='F' and codassinante="+mz[1,1]+" and codcliif="+WIF
              mznome=mz[1,2]
              mzcnpjcpf=cnpjcpf
              mz=SQLArrayAssoc(WNRC1)
              If LEN(MZ)>0
                MWSQL_ROWID=MZ[1,1]
                MWVENCTO=ALLTRIM(MZ[1,2])
                MWVENCTO=CTOD(RIGHT(MWVENCTO)+"/"+SUBSTR(MWVENCTO,6,2)+"/"+LEFT(MWVENCTO,4))
                MWVALORMAX=VAL(MZ[1,3])
                MWTIPODOCUMENTO=MZ[1,4]
                If MWVENCTO=CTOD("")
                  //                wnrc1="update vencimentos set vencto="+dtom(ctod("31/12/2100"))+" where sql_rowid="+mwsql_rowid
                  //                MWVENCTO=CTOD("31/12/2100")
                  //                sql execute wnrc1
                Endif
                MWOBS="Intermediario: "+ZPFCNPJ+" Assinante: "+mzcnpjcpf+"-"+mznome
                If MWVENCTO<=ZTCONTR
                  If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                    ENVIAEMAIL(alltrim(emailquemenviou),"Procuração Vencida","Procuração Vencida - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                  Endif
                  SELE 38
                  ADIREG(0)
                  REPLACE CODIf WITH CORRETOR->SQL_ROWID
                  REPLACE USUARIO WITH quemenviou
                  REPLACE REFERENCIA WITH MWOBS
                  REPLACE MOTIVO WITH "Procuração Vencida"
                  SELE 31
                Endif
                If AT("@",EMAIL)=0
                  If SEENVIAEMAIL="S" .And. !EMPTY(emailquemenviou)
                    ENVIAEMAIL(alltrim(emailquemenviou),"Email Inválido","Email Inválido - "+MWOBS,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WEBCOR,WEBCOR)
                  Endif
                  SELE 38
                  ADIREG(0)
                  REPLACE CODIf WITH CORRETOR->SQL_ROWID
                  REPLACE USUARIO WITH quemenviou
                  REPLACE REFERENCIA WITH MWOBS
                  REPLACE MOTIVO WITH "Email Inválido"
                  SELE 31
                Endif
                If SEENVIAEMAIL="S" .And. ZEAIS<=MWVALORMAX .And. MWVENCTO>=(ZTCONTR) .And. AT(STRZERO(ZPDOCTO,10),MWTIPODOCUMENTO)#0 .And. at("@",EMAIL)#0
                  webcorpo1=Strtran(webcorpo,";;a02nome;:",mznome)
                  //                ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,ZRQPDF,"",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                  ENVIAEMAIL(alltrim(EMAIL),WEBASSUNTO,WEBCORPO1,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                  enviaaviso("Há documento(s) a ser(em) assinado(s)",cnpjcpf)
                  DECLARE TX1[1]
                  TX1[1]="Documento nr: "+STRZERO(ZRCONTR,19)+" DE: "+DTOC(ZTCONTR)+" enviado aviso para: "+alltrim(NOME)+" email: "+alltrim(EMAIL)
                  LOGFILE(WDADOS,TX1)
                Endif
              Endif
            Endif
            SKIP
          ENDDO
        Endif
      Endif
      Erase &NARQ
      Erase &NARQ1
      NEXT OI
    Endif
    EMAILRETORNO=emailquemenviou // WEBEMAIL
    If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
      ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"AssinaNet - Log dos Documentos enviados","Em anexo relatorio dos documentos para assinatura digital",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    Endif
    SELE 36
    USE
    SELE 34
    USE
    SELE 33
    USE
    SELE 31
    USE
    SELE 38
    USE
    SELE 39
    USE
    SELE 32
    USE
    SELE 35
  Return

  Procedure COBRADOC(PARAM)
    SEENVIAEMAIL=PARAM
    If !FILE("CONEX.DAD")
      Return
    Endif
    HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
    If (DOW(DATE())#6 .OR. HORAS<2000) .And. SEENVIAEMAIL="S"
      Return
    Endif
    If SEENVIAEMAIL="S"
      If FAZENDO
        Return
      Endif
    Endif
    fazendo=.T.
    RESTORE FROM CONEX.DAD ADDITIVE
    CONEXON=trim(Decrypt(CONEXON,"leopardo"))
    CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
    CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
    CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
    CONEXINTO=90
    CURSORWAIT()
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTO

    If SQL_ErrorNO() > 0
      Return
    Endif
    SQLSETCONNECTION(CONEXINTO)
    SET PACKETSIZE TO 25
    SELE 35
    If !USEREDE("instituicao",.F.,10,"CORRETOR","M")
      SELE 30
      Return
    Endif
    SELE 31
    If !USEREDE("assinantes",.F.,10,"ASSINA","M")
      SELE 35
      USE
      SQLDISCONNECT(CONEXINTO)
      SELE 30
      CURSORARROW()
    Endif
    SELE 35
    GO TOP
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    Do While .NOT. EOF()
      SELE 35
      If COBRAASS<>"S"
        SKIP
        LOOP
      Endif
      DECLARE TX1[1]
      TX1[1]="Cobra Docto"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      wcpfcnpj=cnpjcpf
      WNRC="SELECT sql_rowid FROM INSTITUICAO WHERE CNPJCPF="+WCPFCNPJ
      mz=SQLArrayAssoc(WNRC)
      wcodif=MZ[1,1]
      WCODIF=STRZERO(VAL(WCODIF),10)
      SysRefresh()
      M->LOCAL21=TRIM("/"+CURDIR())
      SELE 35
      WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
      VERDIRETORIO(WEBLOCAL1)
      WEBLOCAL2=ALLTRIM(CAMINHO)+"ARQF\"
      VERDIRETORIO(WEBLOCAL2)
      WEBLOCAL3=alltrim(CAMINHO)+"DOCTOS\"
      VERDIRETORIO(WEBLOCAL3)
      DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
      ADIR(WEBLOCAL1+"*.*",VLABEL1)
      FOR O=1 TO LEN(VLABEL1)
        SysRefresh()
        ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
        Erase &ARQDE
      NEXT O
      DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
      ADIR(WEBLOCAL1+"*.*",VLABEL1)
      If LEN(VLABEL1)>0
        MSGINFO("Verifique pasta temporaria, não apagou todos arquivos",wsist)
        EXIT
      Endif
      ARQF3=WEBLOCAL2+"COBRADO.TXT"
      ULTCOBR=CTOD("")
      If FILE(ARQF3)
        ULTCOBR=CTOD(MEMOREAD(ARQF3))
        If ULTCOBR=DATE()
          SELE 35
          SKIP
          LOOP
        Endif
      Endif
      MEMOWRIT(ARQF3,DTOC(DATE()))
      SELE 35
      WEBLOCAL=ALLTRIM(CAMINHO)
      WEBCOR=NOME
      WEBASSUNTO="Assinanet - Documento(s) para ser(em) assinado(s) eletronicamente"
      WDADOS=WEBLOCAL1+"ENVIADO.TXT"
      Erase &WDADOS
      WEBCORPO=OEMTOANSI(CORPOEMAIL)
      WEBPROCE=OEMTOANSI(CORPOPROCESSO)
      WEBSENHA=SENHAEMAIL
      WEBLOGIN=USUARIOEMAIL
      WEBSENHA=SENHAEMAIL
      WEBSERV=ALLTRIM(SMTP)
      WEBEMAIL=ALLTRIM(EMAIL)
      DECLARE TX1[1]
      TX1[1]="Assinanet Cobrança: "+WEBCOR
      LOGFILE(WEBLOCAL2+"ROTINAS.TXT",TX1)
      COBRARDOCS()
      If SEENVIAEMAIL="S" .And. !EMPTY(WEBEMAIL)
        ENVIAEMAIL(ALLTRIM(WEBEMAIL),"AssinaNet - Cobranças enviadas","Em anexo cobrança enviados",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      If SEENVIAEMAIL="S"
        ENVIAEMAIL("bene@exatus.net","AssinaNet - Cobranças enviadas "+WEBCOR,"Em anexo cobrança enviadas",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      If SEENVIAEMAIL="S"
        ENVIAEMAIL("wander@exatus.net","AssinaNet "+WEBCOR,"Cobranças enviadas",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      LCHDIR(M->LOCAL21)
      SELE 35
      SKIP
    ENDDO
    cursorarrow()
    SELE 31
    USE
    SELE 35
    USE
    SQLDISCONNECT(CONEXINTO)
    SELE 30
    CURSORARROW()
    FAZENDO=.F.
  Return

  Procedure COBRARDOCS()
    oConn := TOleAuto():New("ADODB.Connection")
    wconec="driver=MySQL ODBC 5.1 Driver;server="+conexon+";uid="+conexuser+";Pwd="+conexsenha+";database=assinanet"
    oConn:Open(WCONEC)
    oRs1 := TOleAuto():New("ADODB.Recordset")
    oRS1:CursorLocation:=3
    oRS1:CursorType:=0
    oRS1:LockType:=1
    oRS1:ActiveConnection(oConn)
    SINT1="SELECT documentos.codif,clientesif.nome,documentos.codcliif,COUNT(documentos.sql_rowid) quantos FROM documentos LEFT JOIN clientesIf ON documentos.codcliif=clientesif.`sql_rowid` WHERE completo='N' AND documentos.sql_deleted='F' AND documentos.codif="+WCODIF+" AND documentos.data>"+dtom(date()-1095)+" GROUP BY documentos.codif,documentos.codcliif"
    oRS1:Open(SINT1)
    oConn:Execute(sint1)
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    Do While !(oRs1:Eof)
      WIF=STRZERO(RETNUL1("CODCLIIF","N"),10)
      wmpresa=retnul1("nome","C")
      If VAL(WIF)=0
        oRs1:MoveNext()
        LOOP
      Endif
      WEBASSUNTO="Assinanet - "+STRZERO(RETNUL1("QUANTOS","N"),4)+" Documento(s) para ser(em) assinado(s) eletronicamente"
      SELE 31
      MFILTER='CODCLIIf LIKE "%'+WIF+'%"'
      SQL FILTER TO MFILTER
      GO TOP
      Do While !EOF()
        WNRC1="SELECT sql_rowid,nome from assinantes where sql_deleted='F' and cnpjcpf="+cnpjcpf
        mz1=SQLArrayAssoc(WNRC1)
        If len(mz1)>0
          WNRC1="SELECT sql_rowid,vencto,valormax,tipodocumento FROM vencimentos where sql_deleted='F' and codassinante="+mz1[1,1]+" and codcliif="+WIF+" and codif="+WCODIF
          mznome=mz1[1,2]
          mz1=SQLArrayAssoc(WNRC1)
          If LEN(MZ1)>0
            MWSQL_ROWID=MZ1[1,1]
            MWVENCTO=ALLTRIM(MZ1[1,2])
            MWVENCTO=CTOD(RIGHT(MWVENCTO,2)+"/"+SUBSTR(MWVENCTO,6,2)+"/"+LEFT(MWVENCTO,4))
            MWVALORMAX=VAL(MZ1[1,3])
            MWTIPODOCUMENTO=MZ1[1,4]
            If at("@",EMAIL)#0
              ZMAIL1=EMAIL
              //         zmail1="WANDERLEI@EXATUS.NET"
              If SEENVIAEMAIL="S"
                webcorpo1=Strtran(webcorpo,";;a02nome;:",mznome)
                ENVIAEMAIL(alltrim(ZMAIL1),WEBASSUNTO,WEBCORPO1,"","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,NOME,WEBCOR)
                enviaaviso("Há documento(s) a ser(em) assinado(s)",cnpjcpf)
              Endif
              DECLARE TX1[1]
              TX1[1]="Cobrança empresa: "+wmpresa+" documentos "+webassunto+" enviado aviso para: "+alltrim(NOME)+" email: "+alltrim(ZMAIL1)
              LOGFILE(WDADOS,TX1)
            Endif
          Endif
        Endif
        SKIP
      ENDDO
      oRs1:MoveNext()
      SysRefresh()
    ENDDO
    If oConn:State <> 0
      oConn:Close()
    Endif
    SELE 35
  Return

  Procedure ENVCOMPLETO(PARAM)
    SEENVIAEMAIL=PARAM
    If !FILE("CONEX.DAD")
      Return
    Endif
    HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
    If HORAS<2130 .OR. DATE()=LOCALDASSINA
      Return
    Endif
    LOCALDASSINA=DATE()
    If SEENVIAEMAIL="S"
      If FAZENDO
        Return
      Endif
    Endif
    fazendo=.T.
    RESTORE FROM CONEX.DAD ADDITIVE
    CONEXON=trim(Decrypt(CONEXON,"leopardo"))
    CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
    CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
    CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
    CONEXINTO=90
    CURSORWAIT()
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTO

    If SQL_ErrorNO() > 0
      Return
    Endif
    SQLSETCONNECTION(CONEXINTO)
    SET PACKETSIZE TO 25
    SELE 35
    If !USEREDE("instituicao",.F.,10,"CORRETOR","M")
      SELE 30
      Return
    Endif
    SELE 31
    If !USEREDE("documentos",.F.,10,"DOCUMENTOS","M")
      SELE 35
      USE
      SQLDISCONNECT(CONEXINTO)
      SELE 30
      CURSORARROW()
    Endif
    SELE 35
    GO TOP
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    Do While .NOT. EOF()
      SELE 35
      wcpfcnpj=cnpjcpf
      WNRC="SELECT sql_rowid FROM INSTITUICAO WHERE CNPJCPF="+WCPFCNPJ
      mz=SQLArrayAssoc(WNRC)
      wcodif=MZ[1,1]
      WCODIF=STRZERO(VAL(WCODIF),10)
      SysRefresh()
      M->LOCAL21=TRIM("/"+CURDIR())
      SELE 35
      WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
      VERDIRETORIO(WEBLOCAL1)
      WEBLOCAL2=ALLTRIM(CAMINHO)+"ARQF\"
      VERDIRETORIO(WEBLOCAL2)
      WEBLOCAL3=alltrim(CAMINHO)+"DOCTOS\"
      VERDIRETORIO(WEBLOCAL3)
      DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
      ADIR(WEBLOCAL1+"*.*",VLABEL1)
      FOR O=1 TO LEN(VLABEL1)
        SysRefresh()
        ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
        Erase &ARQDE
      NEXT O
      DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
      ADIR(WEBLOCAL1+"*.*",VLABEL1)
      If LEN(VLABEL1)>0
        MSGINFO("Verifique pasta temporaria, não apagou todos arquivos",wsist)
        EXIT
      Endif
      SELE 35
      WEBLOCAL=ALLTRIM(CAMINHO)
      WEBPROCE=OEMTOANSI(CORPOPROCESSO)
      WEBCOR=NOME
      WEBASSUNTO="Assinanet - Documento(s) assinado(s) eletronicamente (Processo Completo)"
      If EMPTY(WEBPROCE)
        WEBCORPO="Prezados, seguem documento(s) da "+WEBCOR+" assinado(s) eletronicamente"
      Else
        WEBCORPO=WEBPROCE
      Endif
      WDADOS=WEBLOCAL1+"ENVIADO.TXT"
      Erase &WDADOS
      WEBSENHA=SENHAEMAIL
      WEBLOGIN=USUARIOEMAIL
      WEBSENHA=SENHAEMAIL
      WEBSERV=ALLTRIM(SMTP)
      WEBEMAIL=ALLTRIM(EMAIL)
      DECLARE TX1[1]
      TX1[1]="Assinanet Documentos (processo Completo): "+WEBCOR
      LOGFILE(WEBLOCAL2+"ROTINAS.TXT",TX1)
      ALGUM=.F.
      ENVIAFINAL()
      If SEENVIAEMAIL="S" .And. !EMPTY(WEBEMAIL) .And. ALGUM=.t.
        ENVIAEMAIL(ALLTRIM(WEBEMAIL),"AssinaNet - Envio documentos assinados","Em anexo envios completos enviados",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      If SEENVIAEMAIL="S" .And. ALGUM=.t.
        ENVIAEMAIL("bene@exatus.net","AssinaNet - Envio documentos assinados "+WEBCOR,"Em anexo envio completos enviadas",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      If SEENVIAEMAIL="S" .And. ALGUM=.t.
        ENVIAEMAIL("wander@exatus.net","AssinaNet "+WEBCOR,"Envio Documentos assinados "+webcor,WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      Endif
      LCHDIR(M->LOCAL21)
      SELE 35
      SKIP
    ENDDO
    cursorarrow()
    SELE 31
    wnrc1='UPDATE documentos SET emailenviado="S" WHERE completo="S" AND emailenviado="N"'
    SQL EXECUTE WNRC1
    USE
    SELE 35
    USE
    SQLDISCONNECT(CONEXINTO)
    SELE 30
    CURSORARROW()
    FAZENDO=.F.
  Return

  Procedure ENVIAFINAL()
    oConn := TOleAuto():New("ADODB.Connection")
    wconec="driver=MySQL ODBC 5.1 Driver;server="+conexon+";uid="+conexuser+";Pwd="+conexsenha+";database=assinanet"
    oConn:Open(WCONEC)
    oRs1 := TOleAuto():New("ADODB.Recordset")
    oRS1:CursorLocation:=3
    oRS1:CursorType:=0
    oRS1:LockType:=1
    oRS1:ActiveConnection(oConn)
    SINT1="SELECT documentos.codcliif,documentos.caminho,clientesif.nome,clientesif.email,clientesif.cnpjcpf,documentos.sql_rowid FROM documentos LEFT JOIN clientesIf ON documentos.codcliif=clientesif.`sql_rowid` WHERE documentos.completo='S' AND documentos.sql_deleted='F' and documentos.emailenviado='N' AND documentos.codif="+WCODIF+" AND LENGTH(clientesif.email)>4 order by documentos.codcliif"
    oRS1:Open(SINT1)
    oConn:Execute(sint1)
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    If !oRs1:eof
      antes=retnul1("codcliif","N")
      EMANT=RETNUL1("email","C")
      WNOME=RETNUL1("NOME","C")
      WCPFANT=RETNUL1("CNPJCPF","C")
    Endif
    ALGUMAQUI=.F.
    Do While !(oRs1:Eof)
      If ANTES<>RETNUL1("CODCLIIF","N")
        DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
        ADIR(WEBLOCAL1+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
          If UPPER(VLABEL1[O])<>"ENVIADO.TXT"
            Erase &ARQDE
          Endif
        NEXT O
        DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
        ADIR(WEBLOCAL1+"*.*",VLABEL1)
        If LEN(VLABEL1)>1
          MSGINFO("Verifique pasta temporaria, não apagou todos arquivos",wsist)
          EXIT
        Endif
        ANTES=RETNUL1("CODCLIIF","N")
        EMANT=RETNUL1("EMAIL","C")
        WNOME=RETNUL1("NOME","C")
        WCPFANT=RETNUL1("CNPJCPF","C")
        ALGUMAQUI=.F.
      Endif
      If AT("@",EMAIL)<>0
        If FILE(ALLTRIM(RETNUL1("CAMINHO","C")))
          NRDOCTO=RETNUL1("sql_rowid","N")
          DECLARE VLABEL2[ADIR(alltrim(retnul1("CAMINHO","C")))]
          ADIR(ALLTRIM(RETNUL1("CAMINHO","C")),VLABEL2)
          ARQP=WEBLOCAL1+VLABEL2[1]
          LZCOPYFILE(ALLTRIM(RETNUL1("CAMINHO","C")),ARQP)
          oPgW1:=nil
          oPgW1 := CreateObject("Microsoft.XMLHTTP")
          xComando := "https://assinanet.exatus.net/administrativo/server/wsdownloadcertificado.asp?documento="+strzero(nrdocto,9)
          oPgW1:Open("GET",xComando,.F.)
          oPgW1:Send()
          WCERTIF=opgW1:responsebody
          ARQCERT=alltrim(lower(WEBLOCAL1+LEFT(VLABEL2[1],LEN(VLABEL2[1])-4)+"_CERTIFICADO.PDF"))
          Erase &arqcert
          memowrit(ARQCERT,WCERTIF)
          oPgW1:=nil
          ALGUM=.T.
          ALGUMAQUI=.T.
        Endif
      Endif
      oRs1:MoveNext()
      If (oRs1:Eof) .OR. ANTES<>RETNUL1("CODCLIIF","N")
        SINT2="update documentos set emailenviado='S' WHERE COMPLETO='S' AND SQL_DELETED='F' AND EMAILENVIADO='N' AND codif="+WCODIF+" AND CODCLIIF="+STRZERO(ANTES,9)
        If AT("@",EMANT)=0 .OR. ALGUMAQUI=.F.
          oConn:Execute(sint2)
        Else
          comando=WCORRENTE+"\BRAZIP.EXE AH "+WEBLOCAL1+STRZERO(ANTES,9)+".ZIP "+WEBLOCAL1+"*.PDF"+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina ENVIO COMPLETO"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          Do While FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          MTEZ=WEBLOCAL1+"FZ.BAT"
          Erase &MTEZ
          DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.PDF")]
          ADIR(WEBLOCAL1+"*.PDF",VLABEL1)
          FOR O=1 TO LEN(VLABEL1)
            SysRefresh()
            ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
            Erase &ARQDE
          NEXT O
          If SEENVIAEMAIL="S"
            DECLARE TX1[1]
            TX1[1]=webassunto+" enviado documentos completos para: "+alltrim(WNOME)+" email: "+alltrim(EMANT)
            LOGFILE(WDADOS,TX1)
            webcorpo1=Strtran(webcorpo,";;a02nome;:","")
            ENVIAEMAIL(alltrim(EMANT),WEBASSUNTO,WEBCORPO1,WEBLOCAL1+STRZERO(ANTES,9)+".ZIP","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WNOME,WEBCOR)
            //    ENVIAEMAIL("wanderlei.cruz@starmoney.com.br",WEBASSUNTO,WEBCORPO1,WEBLOCAL1+STRZERO(ANTES,9)+".ZIP","",WEBEMAIL,WEBSENHA,WEBSERV,WEBLOGIN,.F.,WNOME,WEBCOR)
            enviaaviso("Documento(s) completo(s) enviado(s) para seu E-mail",WCPFANT)
            oConn:Execute(sint2)
          Endif
        Endif
      Endif
      SysRefresh()
    ENDDO
    If oConn:State <> 0
      oConn:Close()
    Endif
    SELE 35
  Return

  Procedure ATCADASS()
    SELE 37
    mat_stru={}
    aAdd(mat_stru,{"LINHA","C",300,0})
    ATASS=WEBLOCAL1+"AT_ASSINA.DAT"
    ATASSO=WEBLOCAL+"AT_ASSINA.CSV"
    DBCREATE(ATASS,MAT_STRU,"DBFCDX")
    RELEASE mat_stru
    USE &ATASS ALIAS TEMP VIA "DBFCDX"
    If NETERR()
      DECLARE TX1[1]
      TX1[1]="Problema geração arquivo temporario"
      LOGFILE("ErroSql.txt",TX1)
      SELE 37
      USE
      Erase &ATASS
      Return
    Endif
    APPEND FROM &ATASSO SDF
    M->LOCAL3=ALLTRIM(WEBLOCAL2)+MAISNOME+"AT_ASSINA.CSV"
    LZCOPYFILE(ATASSO,M->LOCAL3)
    GO TOP
    contad=0
    SELE 37
    Do While !EOF()
      contad=contad+1
      If CONTAD=1
        SKIP
      Endif
      SysRefresh()
      WLI=ALLTRIM(LINHA)+"; ; ; ; ; ;"+SPACE(200)
      COLU=AT(";",WLI)
      WTIPO=SUBSTR(WLI,1,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WQTASS=SUBSTR(WLI,1,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WCNPJ=SUBSTR(WLI,1,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WCPF=SUBSTR(WLI,1,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WNOME=SINAL(SUBSTR(WLI,1,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WCARGO=SINAL(SUBSTR(WLI,1,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WEMAIL=LOWER(SUBSTR(WLI,1,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WVCTO=SUBSTR(WLI,1,COLU-1)
      wvcto=left(alltrim(WVCTO),10)
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      COLU=AT(";",WLI)
      WVALOR=SUBSTR(WLI,1,COLU-1)
      WVALOR=Strtran(WVALOR,".","")
      WVALOR=Strtran(WVALOR,",","")
      WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
      WCPF=Strtran(WCPF,".","")
      WCPF=Strtran(wcpf,"-","")
      WCPF=Strtran(wcpf,"/","")
      WCPF=VAL(WCPF)
      WCNPJ=Strtran(WCNPJ,".","")
      WCNPJ=Strtran(WCNPJ,"-","")
      WCNPJ=Strtran(WCNPJ,"/","")
      WCNPJ=VAL(WCNPJ)


      // MSGINFO(WTIPO,"WTIPO")
      // MSGINFO(WQTASS,"WQTASS")
      // MSGINFO(WCNPJ,"WCNPJ")
      // MSGINFO(WCPF,"WCPF")
      // MSGINFO(WNOME,"WNOME")
      // MSGINFO(WCARGO,"WCARGO")
      // MSGINFO(WEMAIL,"WEMAIL")
      // MSGINFO(WVCTO,"WVCTO")
      // MSGINFO(WVALOR,"WVALOR")

      If WCPF=11111111111 .OR. WCPF=22222222222 .OR. WCPF=33333333333 .OR. WCPF=44444444444 .OR. WCPF=55555555555 .OR. WCPF=66666666666 .OR. WCPF=77777777777 .OR. WCPF=88888888888 .OR. WCPF=99999999999
        SKIP
        LOOP
      Endif
      If WCPF>0 .And. !CIC(STRZERO(WCPF,11))
        SKIP
        LOOP
      Endif
      If WCNPJ>0 .And. !CGC(STRZERO(WCNPJ,14))
        SKIP
      LOOP
    Endif
    If WTIPO="F" .OR. WTIPO="J"
      If WTIPO="F"
        WWCPFCNPJ=STRZERO(WCPF,11)
      Else
        WWCPFCNPJ=STRZERO(WCNPJ,14)
      Endif
      SELE 34
      WFILT='CNPJCPF='+WWCPFCNPJ+' AND CODIf LIKE "%'+STRZERO(VAL(WCODIF),10)+'%"'
      SQL FILTER TO WFILT
      GO TOP
      BEGIN TRANSACTION
      If VAL(CNPJCPF)=0
        ADIREG(0)
        REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
        REPLACE USERINC WITH quemenviou
      Else
        REGLOCK(10)
        REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
        REPLACE USERALT WITH quemenviou
      Endif
      REPLACE NOME WITH WNOME
      If WTIPO="F"
        REPLACE CNPJCPF WITH STRZERO(WCPF,11)
      Else
        REPLACE CNPJCPF WITH STRZERO(WCNPJ,14)
      Endif
      REPLACE CODIf WITH STRZERO(VAL(WCODIF),10)
      REPLACE CORRETORA WITH "N"
      REPLACE QTDASS WITH VAL(WQTASS)
      If QTDASS=0
        REPLACE QTDASS WITH 1
      Endif
      If WTIPO="F"
        WCNPJ=WCPF
        WTIPO="A"
      Endif
    END TRANSACTION
    COMMIT
  Endif
    If WTIPO="A"
      If WCNPJ=0
        WWCPFCNPJ=STRZERO(WCPF,11)
      Else
        WWCPFCNPJ=STRZERO(WCNPJ,14)
      Endif
      WNRC='SELECT sql_rowid FROM CLIENTESIf WHERE CNPJCPF='+WWCPFCNPJ+' AND CODIF='+STRZERO(VAL(WCODIF),10)
      mz=SQLArrayAssoc(WNRC)
      If LEN(MZ)>0
        WCODCLI=MZ[1,1]
        WCODCLI=STRZERO(VAL(WCODCLI),10)
      Else
        WCODCLI=STRZERO(0,10)
      Endif
      If VAL(WCODCLI)#0
        WNRC='SELECT sql_rowid,codcliif,email FROM assinantes WHERE sql_deleted="F" AND CNPJCPF='+strzero(wcpf,11)
        mz=SQLArrayAssoc(WNRC)
        If LEN(MZ)=0
          wnrc1="insert into assinantes (`trocarsenha`,`incluido`,`userinc`,`deacordo`,`datadeacordo`,`email`,`nome`,`cnpjcpf`,`cargo`,`valormax`,`venctoprocuracao`,`tipodocumento`,`codcliif`) values (0,'"+DTOC(DATE())+" "+TIME()+"','"+quemenviou+"',0,'0001-01-01 00:00:00','"+wemail+"','"+wnome+"','"+strzero(wcpf,11)+"','"+wcargo+"',"+wvalor+","+DTOM(CTOD(WVCTO))+",'0000000001,0000000002,0000000003,0000000004,0000000005,0000000006','"+STRZERO(val(wcodcli),10)+"');"
          SQL EXECUTE WNRC1
        Else
          If AT(STRZERO(val(wcodcli),10),mz[1,2])=0
            If LEN(mz[1,2])<9
              WEMPRES=STRZERO(val(wcodcli),10)
            Else
              WEMPRES=mz[1,2]+","+STRZERO(val(wcodcli),10)
            Endif
            WNRC1="UPDATE assinantes set CARGO='"+WCARGO+"',VALORMAX="+WVALOR+",venctoprocuracao="+dtom(ctod(wvcto))+",alterado='"+dtoc(date())+" "+time()+"',useralt='"+quemenviou+"',codcliif='"+wempres+"' where sql_deleted='F' and cnpjcpf="+strzero(wcpf,11)
            SQL EXECUTE WNRC1
            If EMPTY(mz[1,3]) .OR. AT("@",mz[1,3])=0
              WNRC1="UPDATE assinantes set email='"+wemail+"' where sql_deleted='F' and cnpjcpf="+strzero(wcpf,11)
              SQL EXECUTE WNRC1
            Endif
          Endif
        Endif
        WNRC1="SELECT sql_rowid FROM assinantes where sql_deleted='F' and cnpjcpf="+strzero(wcpf,11)
        mz=SQLArrayAssoc(WNRC1)
        If LEN(MZ)>0
          MWCLIENTE=ALLTRIM(MZ[1,1])
          MWCODCLIIF=WCODCLI
          MWCODIF=WCODIF
          MWVENCTO=DTOM(CTOD(WVCTO))
          MWCARGO=WCARGO
          MWVALORMAX=WVALOR
          MWTIPODOCTO="0000000001,0000000002,0000000003,0000000004,0000000005,0000000006"
          MWISOLADAMENTE="F"
          WNRC1="SELECT sql_rowid FROM vencimentos where sql_deleted='F' and codassinante="+mwcliente+" and codcliif="+mwcodcliif+" and codif="+STRZERO(VAL(WCODIF),10)
          mz=SQLArrayAssoc(WNRC1)
          If LEN(MZ)=0
            wnrc1="insert into vencimentos (`codassinante`,`codcliif`,`codif`,`vencto`,`cargo`,`valormax`,`tipodocumento`,`isoladamente`,`userinc`,`incluido`) values ("+mwcliente+","+mwcodcliif+","+mwcodif+","+MWVENCTO+",'"+trim(MWCARGO)+"',"+trim(MWVALORMAX)+",'"+MWTIPODOCTO+"','"+MWISOLADAMENTE+"','"+QUEMENVIOU+"',now());"
            SQL EXECUTE WNRC1
          Endif
        Endif
      Endif
    Endif
    COMMIT
    SELE 37
    SKIP
    ENDDO
    SELE 37
    USE
    Erase &ATASS
    Erase &ATASSO
  Return

  Procedure MALAD()
    //SELECT 57506_documentos.arquivo,57506_documentos.cliente,57506_clientes.nome,57506_clientes.cep,57506_clientes.logradouro,57506_clientes.numero,57506_clientes.compl,57506_clientes.bairro,57506_clientes.cidade,57506_clientes.uf,57506_clientes.cnpjcpf,57506_clientes.ddd,57506_clientes.telefone,57506_clientes.dddcel,57506_clientes.celular,57506_clientes.documento,57506_clientes.dtdoc,57506_clientes.dtnasc FROM 57506_clientes LEFT JOIN 57506_DOCUMENTOS ON 57506_clientes.CODIGO=57506_DOCUMENTOS.CLIENTE WHERE 57506_clientes.sql_deleted="F" AND 57506_clientes.bloqueado="N" AND 57506_clientes.inativo="N" AND 57506_clientes.Tipo="F" AND 57506_clientes.uf="SP" AND 57506_DOCUMENTOS.ARQUIVO LIKE "%RGRNE%" ORDER BY 57506_clientes.CODIGO LIMIT  5951,1800 (proximo..)
    // SALVAR EM EXCEL CSV, informar antes que os campos serão separados por ;
    If !FILE("C:\T0\MALAD.DBF")
      mat_stru={}
      AADD(MAT_STRU,{"Arquivo","C",100,0})
      AADD(MAT_STRU,{"PASTA","C",11,0})
      AADD(MAT_STRU,{"NOME","C",120,0})
      AADD(MAT_STRU,{"CEP","C",12,0})
      AADD(MAT_STRU,{"LOGRAD","C",150,0})
      AADD(MAT_STRU,{"NUMERO","C",12,0})
      AADD(MAT_STRU,{"COMPL","C",18,0})
      AADD(MAT_STRU,{"BAIRRO","C",100,0})
      AADD(MAT_STRU,{"CIDADE","C",50,0})
      AADD(MAT_STRU,{"UF","C",4,0})
      AADD(MAT_STRU,{"CPF","C",16,0})
      AADD(MAT_STRU,{"DDD","C",8,0})
      AADD(MAT_STRU,{"FONE","C",20,0})
      AADD(MAT_STRU,{"DDDC","C",8,0})
      AADD(MAT_STRU,{"CELUL","C",20,0})
      AADD(MAT_STRU,{"DOCTO","C",30,0})
      AADD(MAT_STRU,{"DTDOCTO","C",12,0})
      AADD(MAT_STRU,{"DTNASC","C",12,0})
      AADD(MAT_STRU,{"RESULT","C",600,0})
      AADD(MAT_STRU,{"SITUA","C",20,0})
      DBCREATE("C:\T0\MALAD.DBF",mat_stru)
    Endif
    SELE 36
    USE C:\T0\MALAD.DBF EXCLUSIVE
    YARQ="C:\T0\MALAD.CSV"
    If FILE(YARQ)
      //  ZAP
      SELE 37
      mat_stru={}
      aAdd(mat_stru,{"LINHA","C",300,0})
      ATASS="C:\T0\TEMP.DAT"
      DBCREATE(ATASS,MAT_STRU,"DBFCDX")
      RELEASE mat_stru
      USE &ATASS ALIAS TEMP VIA "DBFCDX"
      If NETERR()
        MSGINFO("Problema geração arquivo temporario para Taxas",WSIST)
        SELE 37
        USE
        Erase &ATASS
        Return
      Endif
      APPEND FROM &YARQ SDF
      GO TOP
      contad=0
      SELE 37
      Do While !EOF()
        contad=contad+1
        If CONTAD=1
          SKIP
        Endif
        SysRefresh()
        WLI=ALLTRIM(LINHA)+"; ; ; ; ; ;"+SPACE(200)
        COLU=AT(";",WLI)
        ZRQUIVO=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZASTA=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZOME=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZEP=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZOGRAD=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZUMERO=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZOMPL=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZAIRRO=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZIDADE=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZF=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZPF=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZDD=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZONE=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZDDC=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZELUL=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZOCTO=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZTDOCTO=SUBSTR(WLI,1,COLU-1)
        WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
        COLU=AT(";",WLI)
        ZTNASC=SUBSTR(WLI,1,COLU-1)
        //    ZRQUIVO=STRZERO(VAL(ZPF),11)+SUBSTR(ZRQUIVO,AT(".",ZRQUIVO),LEN(ZRQUIVO))
        ZESULT=""
        ZDE="C:\T0\57506\"+STRZERO(VAL(ZASTA),8)+"\RGR*.*"
        COMANDO="COPY "+ZDE+" C:\T0\"+CHR(13)+CHR(10)
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL C:\T0\AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT("C:\T0\FZ.BAT",MZLINHA)
        MEMOWRIT("C:\T0\AGUARDE.TXT","HASH"+MZLINHA)
        CURSORWAIT()
        WAITRUN("C:\T0\FZ.BAT",0)
        CURSORWAIT()
        Do While FILE("C:\T0\AGUARDE.TXT")
          SysRefresh()
        ENDDO
        zdtnasc=substr(ztnasc,9,2)+"/"+substr(ztnasc,6,2)+"/"+substr(ztnasc,1,4)
        zitua=""
        //   Cmacro=chr(13)+chr(10)+'wget -T120 -o tt.txt -Oc:\t0\ret.txt "https://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf='+strzero(val(zpf),11)+'&data='+zdtnasc+'&token=8216570RjmqHVALJo14834768&ignore_db" --no-check-certificate  -U --user-agent=MOZILLA'+chr(13)+chr(10)+"DEL c:\t0\AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        //   MEMOWRIT("c:\t0\AGUARDE.TXT"," ")
        //   MEMOWRIT("c:\t0\FZ.BAT",CMACRO)
        //   WAITRUN("c:\t0\FZ.BAT",0)
        //   Do While FILE("c:\t0\AGUARDE.TXT")
        //   ENDDO
        //   zesult=memoread("c:\t0\ret.txt")
        //   If AT("REGULAR",ZESULT)#0
        //     ZITUA="REGULAR"
        //     Else
        //     ZITUA="IRREGULAR"
        //   Endif
        ZITUA="REGULAR"
        If ZITUA="REGULAR"
          SELE 36
          ADIREG(0)
          PUTREC()
        Endif
        SELE 37
        SKIP
      ENDDO
      SELE 37
      USE
      Erase &YARQ
      Erase &ATASS
    Endif
    SELE 36
    ARQ_XLS="c:\t0\maladi.xls"
    GERAXLS(3,"Relatório Mala direta")
    close database
    MSGINFO("Terminou !",WSIST)
  Return

  Procedure ABREASSINA()
    SELE 31
    If !USEREDE("ASSINANTES",.F.,10,"ASSINA","M")
      SELE 33
      USE
      Erase &NARQ3
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
    SELE 32
    If !USEREDE("DOCUMENTOS",.F.,10,"DOCTOS","M")
      SELE 33
      USE
      Erase &NARQ3
      SELE 31
      USE
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
    SELE 34
    If !USEREDE("CLIENTESIF",.F.,10,"CLIIF","M")
      SELE 32
      USE
      SELE 33
      USE
      Erase &NARQ3
      SELE 31
      USE
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
    SELE 36
    If !USEREDE("LOGENV",.F.,10,"LOGS","M")
      SELE 34
      USE
      SELE 32
      USE
      SELE 33
      USE
      Erase &NARQ3
      SELE 31
      USE
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
    SELE 38
    If !USEREDE("NAOINTEGRADOS",.F.,10,"NAOINT","M")
      SELE 34
      USE
      SELE 32
      USE
      SELE 33
      USE
      Erase &NARQ3
      SELE 31
      USE
      SELE 36
      USE
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
    SELE 39
    If !USEREDE("PARIDADE",.F.,10,"PARIDADE","M")
      SELE 34
      USE
      SELE 32
      USE
      SELE 33
      USE
      Erase &NARQ3
      SELE 31
      USE
      SELE 36
      USE
      SELE 38
      USE
      SELE 35
      TEVEERRO=.T.
      Return
    Endif
  Return

  Procedure GERAASSINA(PARAM)
    HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
    If DOW(DATE())=1 .OR. DOW(DATE())=7 .OR. HORAS>2000 .OR. HORAS<0900
      Return
    Endif
    If !FILE("CONEX.DAD")
      Return
    Endif
    RESTORE FROM CONEX.DAD ADDITIVE
    CONEX1ON=trim(Decrypt(CONEXON,"leopardo"))
    CONEX1DTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
    CONEX1DTBAS="assinanet"
    CONEX1USER=trim(decrypt(CONEXUSER,"leopardo"))
    CONEX1SENHA=trim(decrypt(CONEXSENHA,"leopardo"))
    CONEX1PORT=CONEXPORT
    CONEXEX=40

    SQL CONNECT ON CONEX1ON;
      PORT CONEX1PORT;
      DATABASE CONEX1DTBAS;
      USER CONEX1USER;
      PASSWORD CONEX1SENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXEX

    If SQL_ErrorNO() > 0
      Return
    Endif
    SET PACKETSIZE TO 25
    SQLSETCONNECTION(CONEXEX)
    SELE 35
    If !USEREDE("instituicao",.F.,10,"GCORRETOR","M")
      SQLDISCONNECT(CONEXEX)
      Return
    Endif
    GO TOP
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    Do While .NOT. EOF()
      SELE 35
      If arquivarem#0
        wwcodif=strzero(sql_rowid,10)
        wfunc="UPDATE documentos SET arquivado='T' WHERE codif="+WWCODIF+" AND sql_deleted='F' AND arquivado='F' AND backnaif<>'0001-01-01 00:00:00' AND TIMESTAMPDIFF(DAY,backnaif,NOW()) > "+STRZERO(arquivarem,5)
        SQL EXECUTE WFUNC
        SQL EXECUTE "COMMIT"
      Endif
      If NOVO="N"
        SKIP
        LOOP
      Endif
      GMCAMINHO=alltrim(CAMINHO)
      GWMCORRET=ALLTRIM(QCOR)
      If SUBSTR(GMCAMINHO,2,1)<>":"
        GMCAMINHO=ALLTRIM(CURDRIVE()+":"+caminho)
      Endif
      WEBLOCAL2=GMCAMINHO+"FEITO.ASS"
      If FILE(WEBLOCAL2)
        SKIP
        LOOP
      Endif
      CONTT=0
      MEMOWRIT(weblocal2,"FEITO.ASS")
      BGERAASSINA()
      If CONTT=0
        WEBLOCAL2=GMCAMINHO+"FEITO.ASS"
        Erase &WEBLOCAL2
      Endif
      SELE 35
      SKIP
    ENDDO
    cursorarrow()
    SELE 35
    COMMIT
    USE
    SQL DISCONNECT ALL
    CLOSE DATABASE
    CURSORARROW()
  Return

  Procedure BGERAASSINA()
    CONEXON="mysqlcorretoras.exatus.net"
    CONEXUSER="WorkerAccess"
    CONEXSENHA="W0rk3r@cc355"
    CONEXDTBAS="db_turismo"
    CONEXPORT=3306
    CONEXINTOO=50
    CURSORWAIT()

    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTOO

    SQLSETCONNECTION(CONEXINTOO)
    If SQL_ErrorNO() > 0
      Return
    Endif
    SET PACKETSIZE TO 25
    SELE 31
    If !USEREDE("EMPRESAS",.F.,10,"EMPRESAS","M")
      SQLDISCONNECT(CONEXINTOO)
      Return
    Endif
    LOCATE FOR ALLTRIM(QCOR)=ALLTRIM(GWMCORRET)
    If ALLTRIM(QCOR)#ALLTRIM(GWMCORRET)
      SELE 31
      USE
      SQLDISCONNECT(CONEXINTOO)
      SQLSETCONNECTION(CONEXEX)
      Return
    Endif
    wcodbacen=alltrim(codbacen)
    NTABEL=STRZERO(VAL(CODBACEN),5)+"_movimento"
    CONEXDTBAS=ALLTRIM(BANCODADOS)
    CONEXUSER=ALLTRIM(USUARIODB)
    CONEXSENHA=ALLTRIM(SENHADB)
    CONEXON=ALLTRIM(SERVERDB)
    SELE 31
    USE
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
    CONEXINTOO=50
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTOO

    If SQL_ErrorNO() > 0
      Return
    Endif
    SQLSETCONNECTION(CONEXINTOO)
    SET PACKETSIZE TO 2
    SELE 31
    If !USEREDE(NTABEL,.F.,10,"BACEN","M")
      SQLDISCONNECT(CONEXINTOO)
      Return
    Endif
    GWFILT="DATA='"+STRZERO(YEAR(DATE()),4)+"-"+STRZERO(MONTH(DATE()),2)+"-"+STRZERO(DAY(DATE()),2)+"' AND ADIG='S' AND GEROADIG='N' AND (ACIC='S' OR NR_BACEN>0) AND (NAT1<>32009 AND NAT1<>32016 AND NAT1<>32023)"
    SQL FILTER TO GWFILT
    GO TOP
    Do While !EOF()
      Gwde="link.dat"
      Erase &Gwde
      Gwnrint=strzero(codigo,15)
      Gcmacro='wget.exe -T10 -olog.txt -Olink.dat "https://sictur.exatus.net/ws_contratos.asp?codbacen='+wcodbacen+'&nrinterno='+Gwnrint+'" --no-check-certificate  -U --user-agent=MOZILLA'
      waitrun(GCMACRO,0)
      Gwde=Strtran(memoread("link.dat"),"/","\")
      If FILE(Gwde)
        NOMAARQUIVO:=DIRECTORY(gwde)
        If NOMAARQUIVO[1][2]>100
          If REGLOCK(10)
            CONTT=CONTT+1
            Gwpara=Gmcaminho+strzero(contt,8)+".zip"
            LZCOPYFILE(Gwde,Gwpara)
            NOMAARQUIVO:=DIRECTORY(Gwpara)
            If NOMAARQUIVO[1][2]>100
              REPLACE GEROADIG WITH "S"
            Else
              Erase &gwde
              Erase &Gwpara
              CONTT=CONTT-1
            Endif
          Else
            Erase &gwde
            Erase &Gwpara
          Endif
        Else
          Erase &gwde
          Erase &Gwpara
        Endif
      Else
        Erase &gwde
      Endif
      COMMIT
      SKIP
    ENDDO
    SELE 31
    USE
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
  Return

  Procedure CADASTROCAPA()
    If AT("01EMPRESA:",MZLINHA)=0 .OR. AT("01CNPJCPF:",MZLINHA)=0 .OR. AT("01QTDASSI:",MZLINHA)=0 .OR. AT("01NOME:",MZLINHA)=0 .OR. AT("01CPF:",MZLINHA)=0 .OR. AT("01EMAIL:",MZLINHA)=0 .OR. AT("01CARGO:",MZLINHA)=0 .OR. AT("01VENCTO:",MZLINHA)=0 .OR. AT("01VALOR:",MZLINHA)=0 .OR. AT("01TPDOCTO:",MZLINHA)=0 .OR. AT("01TPPESS:",MZLINHA)=0
      Return
    Endif
    // Cadastro PF (area empresa)
    WNOMEEMPR=TRIM(LECAPA("01EMPRESA:"))
    WWCPFCNPJ=TRIM(LECAPA("01CNPJCPF:"))
    WQTASS=VAL(TRIM(LECAPA("01QTDASSI:")))
    WTIPO=TRIM(LECAPA("01TPPESS:"))
    W01NOME=TRIM(LECAPA("01NOME:"))
    W01CPF=TRIM(LECAPA("01CPF:"))
    W01EMAIL=TRIM(LECAPA("01EMAIL:"))
    W01CARGO=TRIM(LECAPA("01CARGO:"))
    W01VENCTO=TRIM(LECAPA("01VENCTO:"))
    W01VALOR=TRIM(LECAPA("01VALOR:"))
    W01TIPODOCTO=TRIM(LECAPA("01TPDOCTO:"))
    W01CELULAR=TRIM(LECAPA("01CELULAR:"))
    SELE 34
    WFILT='CNPJCPF='+WWCPFCNPJ+' AND CODIf LIKE "%'+STRZERO(VAL(WCODIF),10)+'%"'
    SQL FILTER TO WFILT
    GO TOP
    BEGIN TRANSACTION
    If VAL(CNPJCPF)=0 .or. eof()
      ADIREG(0)
      REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
      REPLACE USERINC WITH "Web-Ficha Cadastral"
    Else
      REGLOCK(10)
      REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
      REPLACE USERALT WITH "Web-Ficha Cadastral"
    Endif
    REPLACE NOME WITH WNOMEEMPR
    If WTIPO="F"
      REPLACE CNPJCPF WITH STRZERO(VAL(WWCPFCNPJ),11)
      REPLACE EMAIL WITH TRIM(W01EMAIL)
    Else
      REPLACE CNPJCPF WITH STRZERO(VAL(WWCPFCNPJ),14)
    Endif
    REPLACE CODIf WITH STRZERO(VAL(WCODIF),10)
    REPLACE CORRETORA WITH "N"
    REPLACE QTDASS WITH WQTASS
    If QTDASS=0
      REPLACE QTDASS WITH 1
    Endif
  END TRANSACTION
  COMMIT
  WCODEMPRESA=sql_rowid
  //Cadastro assinante (PF)
  Do While .T.
    SELE 31
    WFILT='CNPJCPF='+W01CPF
    SQL FILTER TO WFILT
    GO TOP
    If VAL(CNPJCPF)=0
      ADIREG(0)
      REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
      REPLACE USERINC WITH "Web-Ficha Cadastral"
      REPLACE DEACORDO WITH 0
      REPLACE DATADEACORDO WITH "0001-01-01 00:00:00"
    Else
      REGLOCK(10)
      REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
      REPLACE USERALT WITH "AT_ASSINA"
    Endif
    REPLACE NOME WITH W01NOME
    REPLACE CNPJCPF WITH W01CPF
    REPLACE CELULAR WITH W01CELULAR
    REPLACE CARGO WITH W01CARGO
    REPLACE VALORMAX WITH VAL(W01VALOR)
    REPLACE VENCTOPROCURACAO WITH CTOD(W01VENCTO)
    REPLACE TIPODOCUMENTO WITH W01TIPODOCTO
    If AT(STRZERO(WCODEMPRESA,10),CODCLIIF)=0
      If LEN(CODCLIIF)<9
        REPLACE CODCLIIf WITH STRZERO(WCODEMPRESA,10)
      Else
        REPLACE CODCLIIf WITH CODCLIIF+","+STRZERO(WCODEMPRESA,10)
      Endif
    Endif
    REPLACE EMAIL WITH W01EMAIL
    REPLACE TROCARSENHA WITH 0
    COMMIT
    SQL EXECUTE "COMMIT"
    WNRC1="SELECT sql_rowid FROM assinantes where sql_deleted='F' and cnpjcpf="+w01cpf
    mz=SQLArrayAssoc(WNRC1)
    If LEN(MZ)>0
      MWCLIENTE=ALLTRIM(MZ[1,1])
      MWCODCLIIF=STRZERO(WCODEMPRESA,10)
      MWCODIF=WCODIF
      MWVENCTO=DTOM(CTOD(W01VENCTO))
      MWCARGO=W01CARGO
      MWVALORMAX=W01VALOR
      MWTIPODOCTO=W01TIPODOCTO
      MWISOLADAMENTE="F"
      WNRC1="SELECT sql_rowid FROM vencimentos where sql_deleted='F' and codassinante="+mwcliente+" and codcliif="+mwcodcliif
      mz=SQLArrayAssoc(WNRC1)
      If LEN(MZ)=0
        wnrc1="insert into vencimentos (`codassinante`,`codcliif`,`codif`,`vencto`,`cargo`,`valormax`,`tipodocumento`,`isoladamente`,`userinc`,`incluido`) values ("+mwcliente+","+mwcodcliif+","+mwcodif+","+MWVENCTO+",'"+trim(MWCARGO)+"',"+trim(MWVALORMAX)+",'"+MWTIPODOCTO+"','"+MWISOLADAMENTE+"','"+quemenviou+"',now());"
        SQL EXECUTE WNRC1
      Endif
    Endif
    W01NOME=TRIM(LECAPA("01NOME:"))
    W01CPF=TRIM(LECAPA("01CPF:"))
    W01EMAIL=TRIM(LECAPA("01EMAIL:"))
    W01CARGO=TRIM(LECAPA("01CARGO:"))
    If EMPTY(W01CARGO) .And. !EMPTY(W01CPF)
      W01CARGO="NAO INFORMADO"
    Endif
    W01VENCTO=TRIM(LECAPA("01VENCTO:"))
    W01VALOR=TRIM(LECAPA("01VALOR:"))
    W01TIPODOCTO=TRIM(LECAPA("01TPDOCTO:"))
    W01CELULAR=TRIM(LECAPA("01CELULAR:"))
    If empty(w01nome) .or. empty(w01cpf) .or. empty(w01email) .or. empty(w01cargo) .or. empty(w01vencto) .or. empty(w01valor) .or. empty(w01tipodocto)
      EXIT
    Endif
  ENDDO
Return

Procedure LECAPA(QUAL)
  If AT(QUAL,MZLINHA)=0
    Return ""
  Else
    WCONTEM=SUBSTR(MZLINHA,AT(QUAL,MZLINHA)+LEN(QUAL),AT(CHR(13),MZLINHA)-1-len(qual))
    If QUAL="01HIDDEN:"
      WCONTEM=SUBSTR(MZLINHA,AT(QUAL,MZLINHA)+LEN(QUAL),36)
    Endif
    WCONTEM=SINAL(WCONTEM)
    //  MSGINFO(WCONTEM,qual)
    //  If AT(CHR(13),WCONTEM)#0
    //    WCONTEM=LEFT(WCONTEM,AT(CHR(13),WCONTEM)-1)
    //  Endif
    MZLINHA=SUBSTR(MZLINHA,AT(CHR(13),MZLINHA)+1,LEN(MZLINHA))
    //  msginfo(mzlinha)
    Return WCONTEM
  Endif

Procedure ASSINACAPA()
  If AT("CLASSIFICACAO E FINALIDADE",ZBS)#0 // Concordancia com os dados do Contrato
    FOR qtd_ass=1 to wqtd_ass_if
      SQL EXECUTE "INSERT INTO assinou (nome,cnpjcpf,documento,origem,back1,back2) VALUES ( 'ASSINATURA','99999999999',"+STRZERO(WSQLNRDOCTO,11)+",'I','S','S');"
      SQL EXECUTE "COMMIT"
    NEXT qtd_ass
  Endif
  If !EMPTY(ZPFASSINOU) // Cliente
    wfunc="INSERT into `assinou` (`nome`,`cnpjcpf`,`documento`,`origem`,`validade`,`ipa`,`incluido`) VALUES ('"+WNOMEEMPR+"','"+WWCPFCNPJ+"',"+STRZERO(WSQLNRDOCTO,11)+",'C','2099-12-31 00:00:00','"+ZIPPF+"','"+ZPFASSINOU+"')"
    SQL EXECUTE wfunc
    SQL EXECUTE "COMMIT"
  Endif
  If ZCORRETASSINA="S"     // Cadastro
    wfunc="INSERT into `assinou` (`nome`,`cnpjcpf`,`documento`,`origem`,`validade`,`ipa`,`incluido`) VALUES ('"+ZACEITEPOR+"','"+ZACEITEPORCPF+"',"+STRZERO(WSQLNRDOCTO,11)+",'I','2099-12-31 00:00:00','"+ZIPCORRETORA+"',"+zaceiteem+")"
    SQL EXECUTE wfunc
    SQL EXECUTE "COMMIT"
  Endif
  If VAL(wmcpfoperador)#0 // Operador
    wfunc="INSERT into `assinou` (`nome`,`cnpjcpf`,`documento`,`origem`,`validade`,`ipa`) VALUES ('"+WMNOMEOPER+"','"+WMCPFOPERADOR+"',"+STRZERO(WSQLNRDOCTO,11)+",'I','2099-12-31 00:00:00','"+ZIPOPERADOR+"')"
    SQL EXECUTE wfunc
    SQL EXECUTE "COMMIT"
  Endif
  If VAL(ZNPJINT)>0
    If VAL(ZNPJINT)=VAL(ZACEITEPORCPF) .And. ZCORRETASSINA="S" // Moneycorp
      SQL EXECUTE "INSERT into `assinou` (`nome`,`cnpjcpf`,`documento`,`origem`,`validade`,`ipa`,`incluido`) VALUES ('"+ZACEITEPOR+"','"+ZACEITEPORCPF+"',"+STRZERO(WSQLNRDOCTO,11)+",'I','2099-12-31 00:00:00','"+ZIPCORRETORA+"',"+zaceiteem+")"
      SQL EXECUTE "COMMIT"
    Endif
  Endif
Return

Procedure ASSINABACKUP()
  // Roda no servidor EXATUS (PRINCIPAL)
  If !FILE("CONEXBACKUP.DAD")
    Return
  Endif
  RESTORE FROM CONEXBACKUP.DAD ADDITIVE
  CONEX1ON=trim(Decrypt(CONEXON,"leopardo"))
  CONEX1DTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
  CONEX1USER=trim(decrypt(CONEXUSER,"leopardo"))
  CONEX1SENHA=trim(decrypt(CONEXSENHA,"leopardo"))
  CONEX1TP=CONEXTP
  CONEXEX=40
  CONEX1PORT=3306
  If conex1tp=0 .OR. conex1tp=1
    conex1tp="back1"
  Else
    conex1tp="back2"
  Endif
  CURSORWAIT()
  If conex1tp="back1"
    oConn := TOleAuto():New("ADODB.Connection")
    wWconec="driver=MySQL ODBC 5.1 Driver;server="+conex1on+";uid="+conex1user+";Pwd="+conex1senha+";database=assinanet"
    oConn:Open(WWCONEC)
    SINT1="UPDATE assinou SET fezgeoloc='T' WHERE ISNULL(ipa) OR sql_deleted='T'"
    oConn:Execute(sint1)
    oRs1 := TOleAuto():New("ADODB.Recordset")
    oRS1:CursorLocation:=3
    oRS1:CursorType:=0
    oRS1:LockType:=1
    oRS1:ActiveConnection(oConn)
    SINT1="SELECT assinou.sql_rowid,assinou.ipa FROM assinou where sql_deleted='F' and fezgeoloc='F' and not ISNULL(ipa) LIMIT 0,2500"
    oRS1:Open(SINT1)
    CURSORWAIT()
    Do While !(oRs1:Eof)
      wwsql_rowid=strzero(RETNUL1("sql_rowid","N"),10)
      wwipa=RETNUL1("ipa","C")
      If empty(wwipa)
        oConn:Execute("UPDATE assinou set fezgeoloc='T' where sql_rowid="+wwsql_rowid)
        oRs1:MoveNext()
        LOOP
      Endif
      oPg := CreateObject("Microsoft.XMLHTTP")
      xComando := "https://www.ip-tracker.org/blacklist-check.php?ip="+ALLTRIM(WWIPA)
      oPg:Open("GET",xComando,.f.)
      oPg:Send()
      wResult=opg:responsebody
      oPg:=nil
      wwlinha=""
      wwlinha=substr(wresult,at('Hostname:</th> <td>',wResult)+19,len(wResult))
      wwlinha="HostName: "+left(wwlinha,at("</td>",wwlinha)-4)
      geoloc=substr(wResult,at("IP Location:</th><td class='tracking'>",wResult)+38,len(wResult))
      geoloc=left(geoloc,at("<img",geoloc)-1)
      wwlinha=wwlinha+" Localização: "+geoloc
      geoloc=substr(wResult,at("ISP:</th><td class='tracking'>",wResult)+30,len(wResult))
      geoloc=left(geoloc,at("</td>",geoloc)-1)
      wwlinha=wwlinha+" ISP: "+geoloc
      geoloc=substr(wResult,at("Organization:</th><td>",wResult)+22,len(wResult))
      geoloc=left(geoloc,at("</td>",geoloc)-1)
      wwlinha=wwlinha+" Organização: "+geoloc
      geoloc=substr(wResult,at("class='trackingrond'>Blacklist Status:",wResult)+38,len(wResult))
      geoloc=left(geoloc,at("</td>",geoloc)-1)
      wwlinha=wwlinha+" Lista: "+geoloc
      WWLINHA=Strtran(WWLINHA,"'","")
      oConn:Execute("UPDATE assinou set fezgeoloc='T',geoloc='"+sinal(alltrim(left(wwlinha,200)))+"' where sql_rowid="+wwsql_rowid)
      oRs1:MoveNext()
    ENDDO
    oRs1:Close()
    SINT1="SELECT assinou.sql_rowid,vencimentos.cargo,documentos.codIf FROM assinou LEFT JOIN documentos ON documentos.sql_rowid=assinou.documento LEFT JOIN clientesIf ON clientesif.sql_rowid=documentos.codcliIf LEFT JOIN clientesIf AS intermediadora ON intermediadora.sql_rowid=documentos.codcorr LEFT JOIN assinantes ON assinou.cnpjcpf=assinantes.cnpjcpf AND INSTR(assinantes.codcliif,LPAD(documentos.codcliif,10,'0'))>0 LEFT JOIN vencimentos ON clientesif.sql_rowid=vencimentos.codcliIf AND assinantes.sql_rowid=vencimentos.codassinante WHERE assinou.cnpjcpf<>'99999999999' AND documentos.codif=24 AND documentos.completo='S' AND vencimentos.cargo LIKE '%MONEYCORP%' AND AJUSTEIF='F'"
    oRs1:Open(SINT1)
    CURSORWAIT()
    Do While !(oRs1:Eof)
      wwsql_rowid=strzero(RETNUL1("sql_rowid","N"),10)
      oConn:Execute("UPDATE assinou set origem='I',ajusteif='T' where sql_rowid="+wwsql_rowid)
      oRs1:MoveNext()
    ENDDO
    If oConn:State <> 0
      oConn:Close()
    Endif
  Endif
  SQL CONNECT ON CONEX1ON;
    PORT CONEX1PORT;
    DATABASE CONEX1DTBAS;
    USER CONEX1USER;
    PASSWORD CONEX1SENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXEX
  If SQL_ErrorNO() > 0
    Return
  Endif
  SQLSETCONNECTION(CONEXEX)
  WNRC='SELECT assinou.documento,assinou.sql_rowid,documentos.caminho,assinou.origem,clientesif.qtdass,assinantes.isoladamente,assinantes.sql_rowid as sql_assinantes,documentos.codcliIf FROM assinou LEFT JOIN assinantes ON assinou.cnpjcpf=assinantes.cnpjcpf LEFT JOIN documentos ON assinou.documento=documentos.sql_rowid LEFT JOIN clientesIf ON documentos.codcliif=clientesif.sql_rowid WHERE assinou.'+conex1tp+'="N" AND assinou.sql_deleted="F" AND documentos.sql_deleted="F" ORDER BY assinou.sql_rowid'
  mz=SQLArrayAssoc(WNRC)
  wtotal=LEN(MZ)
  If wtotal>500
    wtotal=500
  Endif
  FOR IZ=1 TO wtotal
    wcaminho=upper(mz[iz,3])
    WDOCTO=MZ[IZ,1]
    WLINHA=MZ[IZ,2]
    wassinante=MZ[IZ,7]
    wcliassina=MZ[IZ,8]
    If CONEXTP=2
      WCAMINHO=Strtran(WCAMINHO,"C:","G:")
    Endif
    Erase &wcaminho
    CMACRO='wget.exe -T30 -olog.txt -O'+wcaminho+' "https://assinanet.exatus.net/administrativo/wsdownloadassinado.asp?documento='+wdocto+'" --no-check-certificate'
    WAITRUN(CMACRO,0)
    CURSORWAIT()
    //  WNOMAARQUIVO:=DIRECTORY(wcaminho)
    //  If file(wcaminho) .And. WNOMAARQUIVO[1][2]>500
    SQL EXECUTE 'UPDATE assinou set '+conex1tp+'="S" where sql_rowid='+wlinha
    SQL EXECUTE "COMMIT"
    //  Endif
    // documentos assinados isoladamente
    WNRC1='SELECT isoladamente FROM vencimentos WHERE sql_deleted="F" and codassinante='+wassinante+" and codcliif="+wcliassina


    If VAL(WASSINANTE)=0 .OR. VAL(WCLIASSINA)=0
      //   WNRC=OUTROS("",WNRC,"")
      //   MSGINFO(IZ)
    Endif

    If VAL(wassinante)#0 .And. VAL(wcliassina)#0
      mz1=SQLArrayAssoc(WNRC1)
      If LEN(mz1)>0
        WORIGEM=MZ[IZ,4]
        WQTDASS=MZ[IZ,5]
        WISOLADO=MZ1[1,1]
        // documentos que precisam de apenas 1 assinatura
        WNRC1='SELECT tipodocumento FROM documentos WHERE tipodocumento=7 and sql_deleted="F" and sql_rowid='+WDOCTO
        mz1=SQLArrayAssoc(WNRC1)
        TIPO7=.F.
        If LEN(MZ1)#0
          TIPO7=.T.
        Endif
        If CONEXTP=1 .And. WORIGEM="C" .And. (WISOLADO="T" .OR. TIPO7)
          FOR IIZ=1 TO VAL(WQTDASS-1)
            SQL EXECUTE "INSERT INTO assinou (nome,cnpjcpf,documento,origem,back1,back2) VALUES ( 'ASSINATURA','99999999999',"+WDOCTO+",'C','S','S');"
          NEXT IIZ
          SQL EXECUTE "COMMIT"
        Endif
      Endif
    Endif
  NEXT IZ
  SQLSETCONNECTION(CONEXEX)
  SQLDISCONNECT(CONEXEX)
  SQL DISCONNECT ALL
  close database
  abrearqs()
  sele 30
  CURSORARROW()
Return

Procedure BACKDOCSICTUR()
  If !FILE("CONEXBACKUP.DAD")
    Return
  Endif
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If HORAS>2100 .OR. HORAS<0900
    Return
  Endif
  RESTORE FROM CONEXBACKUP.DAD ADDITIVE
  CONEX1TP=CONEXTP
  If conex1tp=0 .OR. conex1tp=1
    conex1tp="back1"
  Else
    conex1tp="back2"
  Endif
  CONEXON="mysqlcorretoras.exatus.net"
  CONEXDTBAS="bacen"
  CONEXUSER="WorkerAccess"
  CONEXSENHA="W0rk3r@cc355"
  CONEXDTBAS="db_turismo"
  CONEXPORT=3306
  CONEXINTOO=50
  CURSORWAIT()

  ULT=""
  Do While .T.
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTOO

    If SQL_ErrorNO() > 0
      Return
    Endif

    SET PACKETSIZE TO 25
    SQLSETCONNECTION(CONEXINTOO)
    SELE 35
    If !USEREDE("empresas",.F.,10,"GCORRETOR","M")
      SQLDISCONNECT(CONEXINTOO)
      Return
    Endif
    GO TOP
    If VAL(ULT)=0
      ULT=CODBACEN
    Else
      LOCATE FOR ULT=CODBACEN
      SKIP
    Endif
    If EOF()
      SQLDISCONNECT(CONEXINTOO)
      EXIT
    Endif
    ULT=CODBACEN
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
    VERDIRETORIO(WEBLOCAL1)
    If fazbackup="N"
      SQLDISCONNECT(CONEXINTOO)
      LOOP
    Endif
    CONEX1ON=alltrim(serverdb)
    CONEX1DTBAS=alltrim(bancodados)
    CONEX1USER=alltrim(usuariodb)
    CONEX1SENHA=alltrim(senhadb)
    CONEX1PORT=3306
    CONEX1INTOO=60
    CONEX1CORRE=codbacen
    WGUARDALISTAS=guardalistas
    SQLDISCONNECT(CONEXINTOO)
    SQL DISCONNECT ALL
    BGERABACKUSICTUR()
  ENDDO
  SQLDISCONNECT(CONEXINTOO)
  SQL DISCONNECT ALL
  CLOSE DATABASE
  abrearqs()
  CURSORARROW()
Return

Procedure BGERABACKUSICTUR()
  CURSORWAIT()
  SQL CONNECT ON CONEX1ON;
    PORT CONEX1PORT;
    DATABASE CONEX1DTBAS;
    USER CONEX1USER;
    PASSWORD CONEX1SENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEX1INTOO

  SQLSETCONNECTION(CONEX1INTOO)
  If SQL_ErrorNO() > 0
    Return
  Endif
  SET PACKETSIZE TO 25
  WNRC="SELECT codigo,arquivo,cliente,documento FROM "+conex1corre+"_documentos  WHERE "+conex1tp+"='N' AND sql_deleted='F'"
  mz=SQLArrayAssoc(WNRC)
  wtotal=LEN(MZ)
  If wtotal>1000
    wtotal=1000
  Endif
  FOR IZ=1 TO wtotal
    wcodigo=MZ[IZ,1]
    wcaminho="C:\DOCUMENTOS\"
    If CONEXTP=2
      WCAMINHO=Strtran(WCAMINHO,"C:","G:")
    Endif
    If !VERDIRETORIO(WCAMINHO)
      loop
    Endif
    wcaminho=trim(wcaminho)+trim(conex1corre)+"\"
    If !VERDIRETORIO(WCAMINHO)
      loop
    Endif
    wcaminho=wcaminho+STRZERO(val(MZ[IZ,3]),8)+"\"
    If !VERDIRETORIO(WCAMINHO)
      loop
    Endif
    wcaminho=wcaminho+trim(MZ[IZ,2])
    Erase &wcaminho
    wcliente=MZ[IZ,3]
    CMACRO='wget.exe -T30 -olog.txt -O"'+wcaminho+'" "https://sictur.exatus.net/wsdownloaddocumento.asp?codbacen='+conex1corre+'&documento='+wcodigo+'" --no-check-certificate'
    WAITRUN(CMACRO,0)
    CURSORWAIT()
    WNOMAARQUIVO:=DIRECTORY(wcaminho)
    If WGUARDALISTAS="S" .And. MZ[IZ,4]="Comprovante Receita" .And. at(".",wcaminho)#0 .And. conex1tp="back1"
      WNRC1="SELECT cnpjcpf,nome,dtnasc from "+conex1corre+"_clientes  WHERE codigo="+wcliente
      mz1=SQLArrayAssoc(WNRC1)
      If LEN(mz1)>0
        wpastapdf="c:\sites\turismo\documentos\"+conex1corre+"\"+strzero(val(wcliente),8)+"\"
        WBNASC=stod(Strtran(MZ1[1,3],"-",""))
        buscalista(conex1corre,MZ1[1,2],MZ1[1,1],WBNASC,strzero(val(wcliente),9),wpastapdf)
      Endif
    Endif
    wfunc='UPDATE '+conex1corre+'_documentos set '+conex1tp+'="S" where codigo='+wcodigo
    If conex1tp="back2"
      declare tx1[1]
      tx1[1]=wfunc
      LOGFILE("c:\exatus\copia.txt",tx1)
    Endif
    SQL EXECUTE wfunc
    SQL EXECUTE "COMMIT"
  NEXT IZ
  SQLDISCONNECT(CONEX1INTOO)
  SQLSETCONNECTION(CONEXINTOO)
Return

Procedure buscalista(wCDBACEN,MWNOME,WBCPF,WBNASC,qualcli,wpastapdf)
  wxmlB2='<?xml version="1.0" encoding="UTF-8"?> <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://microsoft.com/webservices/"> <soapenv:Header> <web:Autenticacao> <web:Usuario>'+wCDBACEN+'</web:Usuario> <web:Senha>SEMPREAMESMASENHA1.</web:Senha> </web:Autenticacao> </soapenv:Header> <soapenv:Body> <web:BuscaCompleta> <web:tipo>NOME</web:tipo> <web:corretora>'+wCDBACEN+'</web:corretora> <web:nome>'+MWNOME+'</web:nome>  <web:cnpjCpf>'+WBCPF+'</web:cnpjCpf> <web:dataNascimento>'+DTOC(WBNASC)+'</web:dataNascimento> </web:BuscaCompleta> </soapenv:Body></soapenv:Envelope>'
  objHTML:= TOleAuto():New("MSXML2.ServerXMLHTTP.6.0")
  objHTML:setTimeouts(250000,250000,250000,250000) // Resolve(link), Connect (conectar no lugar), Send (envio do xml), Receive (aguardar para receber)
  objHTML:open("POST","https://sictur.exatus.net/lr/ListasRestritivas.asmx")
  objHTML:setRequestHeader("Host","https://sictur.exatus.net/lr/ListasRestritivas.asmx")
  objHTML:setRequestHeader("Content-Type","text/xml; charset=utf-8")
  objHTML:setRequestHeader("SOAPAction","http://microsoft.com/webservices/BuscaCompleta")
  objHtml:setRequestHeader("Content-Length:",strzero(len(wxmlb2),4))
  objHTML:send(wxmlB2)
  objretorno=objHTML:status
  wxmlB3=wxmlB2
  wxmlB2=objHTML:Responsebody
  objHTML:=nil
  wevidencias=lexml('ArquivoEvidencia',wxmlb2)
  wlink="https://sictur.exatus.net/lr/evidencias/"+wevidencias
  If objretorno=200
    If AT(".PDF",wlink)#0
      oPg1 := CreateObject("Microsoft.XMLHTTP")
      oPg1:Open("GET",wlink,.f.)
      oPg1:Send()
      marq2=Strtran(ALLTRIM(wevidencias),"/","\")
      memowrit(WLOCTEMP+marq2,opg1:responsebody)
      uploadexatussvb(WLOCTEMP,marq2,wpastapdf)
      oPg1:=nil
      WFUNC="INSERT INTO "+wCDBACEN+"_documentos (`documento`,`arquivo`,`cliente`,`userinc`,`incluido`) VALUES ('PLD ','"+marq2+"',"+qualcli+",'Worker','"+dtoc(date())+" "+time()+"')"
      SQL EXECUTE wfunc
      SQL EXECUTE "COMMIT"
      WXML1=""
    Endif
  Else
    MEMOWRIT("C:\EXATUS\FLAG.TXT",wxmlB3)
    //  ENVIAEMAIL("bene@exatus.net;wanderlei.cruz@starmoney.com.br",wxmlb3,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
  Endif
  objHTML:=nil
Return

Procedure uploadexatus(wlocal,warq,pastaweb)
  LINK="https://sictur.exatus.net/uploadexatus.php"
  nomearq=wlocal+warq
  wpastaweb=pastaweb
  wcontxml=preenchexml(MEMOREAD("c:\exatus\upload.bas"))
  MEMOWRIT("ENVIA.VBS",WCONTXML)
  RUNFILE("ENVIA.VBS")
  DECLARE TX1[1]
  TX1[1]=pastaweb
  LOGFILE("c:\exatus\upload.con",tx1)
Return

Procedure uploadexatussvb(WORIGEM,WARQORIGEM,WGRAVAEM)
  wLINK="https://doctosnet.exatus.net/cadastro/uploadexatus.php"
  oXMLHTTP:=CreateObject("MSXML2.XMLHTTP")
  boundary = "12345678"
  objStream:=CreateObject("ADODB.Stream")
  objStream:Type = 1 //' binary
  objStream:Open()
  objStream:LoadFromFile(worigem+warqorigem)
  sOut:=CreateObject("ADODB.Stream")
  sOut:Charset="us-ascii"
  sOut:Type=2 // Text!
  sOut:Open()
  sOut:WriteText("--" + boundary + CRLF)
  sOut:WriteText("Content-Disposition: form-data; name=caminho" + CRLF + CRLF)
  sOut:WriteText(WGRAVAEM + CRLF)
  sOut:WriteText("--" + boundary + CRLF)
  sOut:WriteText("Content-Disposition: form-data; name=churrasco; filename="+warqorigem+CRLF)
  sOut:WriteText("Content-Type: application/pdf" + CRLF + CRLF)
  sOut2:=CreateObject("ADODB.Stream")
  sOut2:Charset="us-ascii"
  sOut2:Type = 2 // Text
  sOut2:Open()
  sOut2:WriteText(CRLF + "--" + boundary + "--" + CRLF + CRLF)
  sAll:=CreateObject("ADODB.Stream")
  sAll:Type = 1 // binary
  sAll:Open()
  sOut:Position = 0
  sOut:CopyTo(sAll)
  objStream:CopyTo(sAll)
  sOut2:Position = 0
  sOut2:CopyTo(sAll)
  oXMLHTTP:Open("POST",wlink, .f.)
  oXMLHTTP:setRequestHeader("Content-Type", "multipart/form-data; boundary="+boundary)
  oXMLHTTP:setRequestHeader("Content-length", sAll:Size)
  sAll:Position = 0
  oXMLHTTP:Send(sAll:Read())
  objStrem:=nil
  sOut:=nil
  sOut2:=nil
  sAll:=nil
  oXMLHTTP:=nil
Return

Procedure GERAFICHAPF(PARAM)
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If (DOW(DATE())=1 .OR. DOW(DATE())=7 .OR. HORAS>2000 .OR. HORAS<0600) .And. PARAM<>"N"
    Return
  Endif
  If !FILE("CONEX.DAD")
    Return
  Endif
  RESTORE FROM CONEX.DAD ADDITIVE
  CONEX1ON=trim(Decrypt(CONEXON,"leopardo"))
  CONEX1DTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
  CONEX1DTBAS="assinanet"
  CONEX1USER=trim(decrypt(CONEXUSER,"leopardo"))
  CONEX1SENHA=trim(decrypt(CONEXSENHA,"leopardo"))
  CONEX1PORT=CONEXPORT
  CONEXEX=40

  SQL CONNECT ON CONEX1ON;
    PORT CONEX1PORT;
    DATABASE CONEX1DTBAS;
    USER CONEX1USER;
    PASSWORD CONEX1SENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXEX

  If SQL_ErrorNO() > 0
    Return
  Endif
  SET PACKETSIZE TO 25
  SQLSETCONNECTION(CONEXEX)
  SELE 35
  If !USEREDE("instituicao",.F.,10,"GCORRETOR","M")
    SQLDISCONNECT(CONEXEX)
    Return
  Endif
  GO TOP
  CURSORWAIT()
  WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\ASSINANET\"
  VERDIRETORIO(WEBLOCAL1)
  Do While .NOT. EOF()
    SELE 35
    If EMPTY(GCORRETOR->cserverdb) .OR. EMPTY(GCORRETOR->cbancodados) .OR. EMPTY(GCORRETOR->cusuariodb) .OR. EMPTY(GCORRETOR->csenhadb)
      SKIP
      LOOP
    Endif
    GMCAMINHO=alltrim(CAMINHO)
    If SUBSTR(GMCAMINHO,2,1)<>":"
      GMCAMINHO=ALLTRIM(CURDRIVE()+":"+caminho)
    Endif
    WEBLOCAL2=GMCAMINHO+"FEITO.ASS"
    If FILE(WEBLOCAL2)
      SKIP
      LOOP
    Endif
    CONTT=0
    MEMOWRIT(weblocal2,"FEITO.ASS")
    BGERAFICHAPF()
    If CONTT=0
      WEBLOCAL2=GMCAMINHO+"FEITO.ASS"
      Erase &WEBLOCAL2
    Endif
    SELE 35
    SKIP
  ENDDO
  cursorarrow()
  SELE 35
  COMMIT
  USE
  SQL DISCONNECT ALL
  CLOSE DATABASE
  CURSORARROW()
  If PARAM="N"
    MSGINFO("Terminou as Fichas PF e PJ",wsist)
  Endif
Return

Procedure BGERAFICHAPF()
  CONEXON=TRIM(GCORRETOR->cserverdb)
  CONEXDTBAS=TRIM(GCORRETOR->cbancodados)
  CONEXUSER=TRIM(GCORRETOR->cusuariodb)
  CONEXSENHA=TRIM(GCORRETOR->csenhadb)
  CONEXPORT=3306
  CONEXINTOO=50
  CURSORWAIT()
  SQL CONNECT ON CONEXON;
    PORT CONEXPORT;
    DATABASE CONEXDTBAS;
    USER CONEXUSER;
    PASSWORD CONEXSENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXINTOO

  SQLSETCONNECTION(CONEXINTOO)
  If SQL_ErrorNO() > 0
    Return
  Endif
  SET PACKETSIZE TO 25
  //**** PESSOA FISICA
  WNRC="SELECT versao2,pastaid,pastapdf FROM doctosnet WHERE sql_rowid=1"
  mz=SQLArrayAssoc(WNRC)
  wversao=MZ[1,1]
  wpastaid=MZ[1,2]
  wpastapdf=MZ[1,3]
  SELE 31
  If !USEREDE("fisica",.F.,10,"fisica","M")
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
    Return
  Endif
  WFILT="enviaassinanet='S'"
  SQL FILTER TO WFILT
  GO TOP
  Do While !EOF()
    CONTT=CONTT+1
    cBuf=""
    If wversao="N"
      xComando=TRIM(GCORRETOR->caminho_pdf)+TRIM(CAMINHO_PDF)
    Else
      xComando=TRIM(wpastapdf+wpastaid+"/"+caminho_pdf)
    Endif
    If ISINTERNET() .And. !EMPTY(xComando)
      oPg := CreateObject("Microsoft.XMLHTTP")
      oPg:Open("GET",xComando,.f.)
      oPg:Send()
      cBuf :=oPg:responseBody
      oPg:=nil
    Endif
    If len(cBuf)<100
      contt=contt-1
      skip
      loop
    Endif
    WNRC="SELECT cliente FROM users WHERE id="+strzero(iduser,11)
    // se cliente="S" significa que quem ativou o cadastro foi o Dept.de cadastro, então não pode ir assinado.
    mz=SQLArrayAssoc(WNRC)
    WASSINACLI="S"
    If len(mz)>0
      If MZ[1,1]="N"
        WASSINACLI="N"
      Endif
    Else
      WASSINACLI="N"
    Endif
    arqpdf=lower(GMCAMINHO+Strtran(UPPER(caminho_pdf),"PDF/",""))
    arqtxt=Strtran(arqpdf,".pdf",".txt")
    MEMOWRIT(arqpdf,cBuf)
    WCPF=Strtran(CPF,".","")
    WCPF=ALLTRIM(Strtran(WCPF,"-",""))
    If VAL(WCPF)=0
      WCPF=Strtran(cpf_estrangeiro,".","")
      WCPF=Strtran(WCPF,"-","")
    Endif
    wcpf=val(WCPF)
    cBuf="CNPJ/CPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
    cBuf=cBuf+"DATA:"+dtoc(date())+chr(13)+chr(10)
    cBuf=cBuf+"R$:0,00"+chr(13)+chr(10)
    cBuf=cBuf+"OBS:Web-Ficha Cadastral"+chr(13)+chr(10)
    If val(cpfoperador)#0  // Assinatura Operador
      cBuf=cBuf+"INTERMEDIARIO:"+strzero(val(cpfoperador),11)+chr(13)+chr(10)
    Endif
    cBuf=cbuf+"IPCORRETORA:"+sinal(ipcorretora)+chr(13)+chr(10)
    If WASSINACLI="N"
      cBuf=cbuf+"PFASSINOU:"+chr(13)+chr(10)
    Else
      cBuf=cbuf+"PFASSINOU:"+sinal(assinadoem)+chr(13)+chr(10) // broker está vindo sem.
    Endif
    cBuf=cBuf+"01IPPF:"+sinal(ALLTRIM(IP))+CHR(13)+CHR(10)
    cBuf=cBuf+"01NRDOCTO:0000"+CHR(13)+CHR(10)
    If ASSINACOR="S"
      cBuf=cBuf+"01INSTASSINA:S"+CHR(13)+CHR(10)     // inclui a assinatura de quem Aceitou o Cadastro
    Else
      cBuf=cBuf+"01INSTASSINA:N"+CHR(13)+CHR(10)     // não inclui a assinatura de quem Aceitou o Cadastro
    Endif
    cBuf=cBuf+"01ACEITEPOR:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(ACEITEPOR)))))+CHR(13)+CHR(10)
    cBuf=cBuf+"01ACEITEPORCPF:"+sinal(aceiteporcpf)+CHR(13)+CHR(10)
    cBuf=cBuf+"01TIPODOCUMENTO:5"+CHR(13)+CHR(10)
    cBuf=cBuf+"01EMPRESA:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(NOME_COMPLETO)))))+CHR(13)+CHR(10)
    cBuf=cBuf+"01CNPJCPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
    cBuf=cBuf+"01QTDASSI:1"+CHR(13)+CHR(10)
    cBuf=cBuf+"01TPPESS:F"+CHR(13)+CHR(10)
    cBuf=cBuf+"01NOME:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(NOME_COMPLETO)))))+CHR(13)+CHR(10)
    cBuf=cBuf+"01CPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
    cBuf=cBuf+"01EMAIL:"+sinal(TRIM(email))+chr(13)+chr(10)
    cBuf=cBuf+"01CARGO:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(OBS_EMPRESA)))))+chr(13)+chr(10)
    cBuf=cBuf+"01VENCTO:01/01/2099"+chr(13)+chr(10)
    cBuf=cBuf+"01VALOR:9999999999999.99"+chr(13)+chr(10)
    cBuf=cBuf+"01TPDOCTO:0000000001,0000000002,0000000003,0000000004,0000000005,0000000006"+chr(13)+chr(10)+CHR(13)+CHR(10)
    cBuf=cBuf+"01CELULAR:"+Strtran(Strtran(Strtran(Strtran(Strtran(celphone,"-","")," ",""),"(",""),")",""),"+","")+chr(13)+chr(10)
    MEMOWRIT(arqtxt,cBuf)
    If REGLOCK(10)
      REPLACE enviaassinanet with "N"
      REPLACE enviadoassinanet with "S"
    Endif
    COMMIT
    SKIP
  ENDDO

  //*** PESSOA FISICA DENTRO DA JURIDICA ***
  SELE 31
  If !USEREDE("juridica",.F.,10,"juridica","M")
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
    Return
  Endif
  SELE 32
  If !USEREDE("fisicapj",.F.,10,"fisicapj","M")
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
    Return
  Endif
  SELE 33
  If !USEREDE("pjacproc",.F.,10,"pjacproc","M")
    SQLDISCONNECT(CONEXINTOO)
    SQLSETCONNECTION(CONEXEX)
    Return
  Endif
  SELE 31
  WFILT="enviaassinanet='S'"
  SQL FILTER TO WFILT
  GO TOP
  Do While !EOF()
    WNRC="SELECT sql_rowid FROM juridica WHERE CNPJ='"+alltrim(cnpj)+"' and sql_deleted='F'"
    mz=SQLArrayAssoc(WNRC)
    wcodif=MZ[1,1]
    WCODIF=STRZERO(VAL(WCODIF),13)
    If VAL(WCODIF)=0
      SKIP
      LOOP
    Endif
    SELE 33
    SQL FILTER TO "id_pj="+wcodif+" and tipo_cad=2 and sql_deleted='F'"
    //  SELECT * FROM pjacproc WHERE tipo_cad=2 AND id_pj=257 UNION ALL SELECT * FROM pjacproc WHERE tipo_cad=1 AND id_pj=257 AND cpf NOT IN (SELECT cpf FROM pjacproc WHERE tipo_cad=2 AND id_pj=257)
    GO TOP
    Do While !EOF()
      SQL EXECUTE "UPDATE fisicapj set idpj="+wcodif+" where sql_rowid="+strzero(pjacproc->id_pf,13)
      WNRC="SELECT CPF FROM fisicapj WHERE sql_rowid="+strzero(pjacproc->id_pf,13)
      mz=SQLArrayAssoc(WNRC)
      If LEN(MZ)=0 .OR. VAL(MZ[1,1])=0
        SKIP
        LOOP
      Endif
      wcpf=MZ[1,1]
      sele 32
      SQL FILTER TO "CPF='"+WCPF+"' and idpj="+wcodif
      GO TOP
      If EOF()
        SELE 33
        SKIP
        LOOP
      Endif
      If fisicapj->ASSINA="F"
        SELE 33
        SKIP
        LOOP
      Endif
      cBuf=""
      xComando=TRIM(GCORRETOR->caminho_pdf)+TRIM(CAMINHO_PDF)
      If ISINTERNET() .And. !EMPTY(xComando)
        oPg := CreateObject("Microsoft.XMLHTTP")
        oPg:Open("GET",xComando,.f.)
        oPg:Send()
        cBuf :=oPg:responseBody
        oPg:=nil
      Endif
      If len(cBuf)<100
        skip
        loop
      Endif
      CONTT=CONTT+1
      arqpdf=GMCAMINHO+Strtran(caminho_pdf,"PDF/","")
      arqtxt=Strtran(arqpdf,".pdf",".txt")
      MEMOWRIT(arqpdf,cBuf)
      WCPF=Strtran(CPF,".","")
      WCPF=ALLTRIM(Strtran(WCPF,"-",""))
      If VAL(WCPF)=0
        WCPF=Strtran(cpf_estrangeiro,".","")
        WCPF=Strtran(WCPF,"-","")
      Endif
      wcpf=val(WCPF)
      cBuf="CNPJ/CPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
      cBuf=cBuf+"DATA:"+dtoc(date())+chr(13)+chr(10)
      cBuf=cBuf+"R$:0,00"+chr(13)+chr(10)
      cBuf=cBuf+"OBS:Web-Ficha Cadastral"+chr(13)+chr(10)
      If val(juridica->cpfoperador)#0  // Assinatura Operador
        cBuf=cBuf+"INTERMEDIARIO:"+strzero(val(juridica->cpfoperador),11)+chr(13)+chr(10)
      Endif
      cBuf=cbuf+"IPCORRETORA:"+sinal(juridica->ipcorretora)+chr(13)+chr(10)
      cBuf=cBuf+"01NRDOCTO:0000"+CHR(13)+CHR(10)
      cBuf=cBuf+"01INSTASSINA:N"+CHR(13)+CHR(10)
      cBuf=cBuf+"01ACEITEPOR:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(JURIDICA->ACEITEPOR)))))+CHR(13)+CHR(10)
      cBuf=cBuf+"01ACEITEPORCPF:"+sinal(juridica->aceiteporcpf)+CHR(13)+CHR(10)
      cBuf=cBuf+"01TIPODOCUMENTO:5"+CHR(13)+CHR(10)
      cBuf=cBuf+"01EMPRESA:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(NOME_COMPLETO)))))+CHR(13)+CHR(10)
      cBuf=cBuf+"01CNPJCPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
      cBuf=cBuf+"01QTDASSI:1"+CHR(13)+CHR(10)
      cBuf=cBuf+"01TPPESS:F"+CHR(13)+CHR(10)
      cBuf=cBuf+"01NOME:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(NOME_COMPLETO)))))+CHR(13)+CHR(10)
      cBuf=cBuf+"01CPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
      cBuf=cBuf+"01EMAIL:"+sinal(TRIM(email))+chr(13)+chr(10)
      cBuf=cBuf+"01CARGO:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(OBS_EMPRESA)))))+chr(13)+chr(10)
      If ASSINAATE=CTOD("")
        cBuf=cBuf+"01VENCTO:01/01/2099"+chr(13)+chr(10)
      Else
        cBuf=cBuf+"01VENCTO:"+dtoc(assinaate)+chr(13)+chr(10)
      Endif
      If vlmaximoassina=0
        cBuf=cBuf+"01VALOR:9999999999999.99"+chr(13)+chr(10)
      Else
        cBuf=cBuf+"01VALOR:"+alltrim(TRANSFORM(vlmaximoassina,"9999999999999.99"))+chr(13)+chr(10)
      Endif
      cBuf=cBuf+"01TPDOCTO:0000000001,0000000002,0000000003,0000000004,0000000005,0000000006"+chr(13)+chr(10)+chr(13)+chr(10)
      cBuf=cBuf+"01CELULAR:"+Strtran(Strtran(Strtran(Strtran(Strtran(celphone,"-","")," ",""),"(",""),")",""),"+","")+chr(13)+chr(10)
      MEMOWRIT(arqtxt,cBuf)
      SELE 33
      SKIP
    ENDDO
    //** CADASTRO PJ E ASSINANTES
    SELE 31
    cBuf=""
    xComando=TRIM(GCORRETOR->caminho_pdf)+Strtran(TRIM(doc_pdf),"../","")
    If ISINTERNET() .And. !EMPTY(xComando)
      oPg := CreateObject("Microsoft.XMLHTTP")
      oPg:Open("GET",xComando,.f.)
      oPg:Send()
      cBuf :=oPg:responseBody
      oPg:=nil
    Endif
    If len(cBuf)<100
      skip
      loop
    Endif
    CONTT=CONTT+1
    arqpdf=trim(GMCAMINHO+Strtran(doc_pdf,"PDF_pj/",""))
    arqtxt=trim(Strtran(arqpdf,".pdf",".txt"))
    MEMOWRIT(arqpdf,cBuf)
    wcnpj=Strtran(cnpj,"-","")
    wcnpj=Strtran(wcnpj,".","")
    wcnpj=val(Strtran(wcnpj,"/",""))
    cBuf="CNPJ/CPF:"+STRZERO(wcnpj,14)+chr(13)+chr(10)
    cBuf=cBuf+"DATA:"+dtoc(date())+chr(13)+chr(10)
    cBuf=cBuf+"R$:0,00"+chr(13)+chr(10)
    cBuf=cBuf+"OBS:Web-Ficha Cadastral"+chr(13)+chr(10)
    If val(juridica->cpfoperador)#0  // Assinatura Operador
      cBuf=cBuf+"INTERMEDIARIO:"+strzero(val(juridica->cpfoperador),11)+chr(13)+chr(10)
    Endif
    cBuf=cbuf+"IPCORRETORA:"+sinal(juridica->ipcorretora)+chr(13)+chr(10)
    cBuf=cBuf+"01NRDOCTO:0000"+CHR(13)+CHR(10)
    cBuf=cBuf+"01INSTASSINA:N"+CHR(13)+CHR(10)
    cBuf=cBuf+"01ACEITEPOR:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(JURIDICA->ACEITEPOR)))))+CHR(13)+CHR(10)
    cBuf=cBuf+"01ACEITEPORCPF:"+sinal(juridica->aceiteporcpf)+CHR(13)+CHR(10)
    cBuf=cBuf+"01TIPODOCUMENTO:5"+CHR(13)+CHR(10)
    cBuf=cBuf+"01EMPRESA:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(RAZAO_SOCIAL)))))+CHR(13)+CHR(10)
    cBuf=cBuf+"01CNPJCPF:"+STRZERO(WCNPJ,14)+chr(13)+chr(10)
    cBuf=cBuf+"01QTDASSI:"+STRZERO(qtdassina,2)+CHR(13)+CHR(10)
    cBuf=cBuf+"01TPPESS:J"+CHR(13)+CHR(10)
    SELE 33
    SQL FILTER TO "id_pj="+wcodif+" and sql_deleted='F'"
    //  SELECT * FROM pjacproc WHERE tipo_cad=2 AND id_pj=257 UNION ALL SELECT * FROM pjacproc WHERE tipo_cad=1 AND id_pj=257 AND cpf NOT IN (SELECT cpf FROM pjacproc WHERE tipo_cad=2 AND id_pj=257)
    GO TOP
    jafez=""
    Do While !EOF()
      WNRC="SELECT CPF FROM fisicapj WHERE not isnull(fisicapj.idpj) and sql_rowid="+strzero(pjacproc->id_pf,13)
      mz=SQLArrayAssoc(WNRC)
      If LEN(MZ)=0 .OR. VAL(MZ[1,1])=0
        SKIP
        LOOP
      Endif
      wcpf=MZ[1,1]
      If at(wcpf,jafez)#0
        skip
        loop
      Endif
      jafez=jafez+wcpf+","
      sele 32
      SQL FILTER TO "CPF='"+WCPF+"'"
      GO TOP
      If EOF()
        SELE 33
        SKIP
        LOOP
      Endif
      If fisicapj->ASSINA="F"
        SELE 33
        SKIP
        LOOP
      Endif
      WCPF=Strtran(CPF,".","")
      WCPF=ALLTRIM(Strtran(WCPF,"-",""))
      If VAL(WCPF)=0
        WCPF=Strtran(cpf_estrangeiro,".","")
        WCPF=Strtran(WCPF,"-","")
      Endif
      wcpf=val(WCPF)
      cBuf=cBuf+"01NOME:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(NOME_COMPLETO)))))+CHR(13)+CHR(10)
      cBuf=cBuf+"01CPF:"+STRZERO(wcpf,11)+chr(13)+chr(10)
      cBuf=cBuf+"01EMAIL:"+sinal(TRIM(email))+chr(13)+chr(10)
      cBuf=cBuf+"01CARGO:"+TRIM(UPPER(sinal(oemtoansi(HB_UTF8TOSTR(OBS_EMPRESA)))))+chr(13)+chr(10)
      If ASSINAATE=CTOD("")
        cBuf=cBuf+"01VENCTO:01/01/2099"+chr(13)+chr(10)
      Else
        cBuf=cBuf+"01VENCTO:"+dtoc(assinaate)+chr(13)+chr(10)
      Endif
      If vlmaximoassina=0
        cBuf=cBuf+"01VALOR:9999999999999.99"+chr(13)+chr(10)
      Else
        cBuf=cBuf+"01VALOR:"+alltrim(TRANSFORM(vlmaximoassina,"9999999999999.99"))+chr(13)+chr(10)
      Endif
      cBuf=cBuf+"01TPDOCTO:0000000001,0000000002,0000000003,0000000004,0000000005,0000000006"+chr(13)+chr(10)
      cBuf=cBuf+"01CELULAR:"+Strtran(Strtran(Strtran(Strtran(Strtran(celphone,"-","")," ",""),"(",""),")",""),"+","")+chr(13)+chr(10)
      SELE 33
      SKIP
    ENDDO
    MEMOWRIT(arqtxt,cBuf)
    SELE 31
    If REGLOCK(10)
      REPLACE enviaassinanet with "N"
      REPLACE enviadoassinanet with "S"
    Endif
    COMMIT
    SKIP
  ENDDO
  COMMIT
  SELE 31
  USE
  SELE 32
  USE
  SELE 33
  USE
  SELE 35
  SQLDISCONNECT(CONEXINTOO)
  SQLSETCONNECTION(CONEXEX)
Return

Procedure DTOM(T_DATA)
Return "'"+STRZERO(YEAR(T_DATA),4)+"-"+STRZERO(MONTH(T_DATA),2)+"-"+STRZERO(DAY(T_DATA),2)+"'"

Procedure BACKDOCNET()
  If !FILE("CONEXBACKUP.DAD")
    Return
  Endif
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If HORAS>2200 .OR. HORAS<0600
    Return
  Endif
  RESTORE FROM CONEXBACKUP.DAD ADDITIVE
  CONEX1TP=CONEXTP
  If conex1tp=0 .OR. conex1tp=1
    conex1tp="back1"
  Else
    conex1tp="back2"
  Endif
  CONEXON="mysqlcorretoras.exatus.net"
  CONEXDTBAS="bacen"
  CONEXUSER="WorkerAccess"
  CONEXSENHA="W0rk3r@cc355"
  CONEXDTBAS="db_turismo"
  CONEXPORT=3306
  CONEXINTOO=50
  CURSORWAIT()
  ULT=""
  Do While .T.
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTOO

    If SQL_ErrorNO() > 0
      Return
    Endif

    SET PACKETSIZE TO 25
    SQLSETCONNECTION(CONEXINTOO)
    SELE 35
    If !USEREDE("empfichas",.F.,10,"GCORRETOR","M")
      SQLDISCONNECT(CONEXINTOO)
      Return
    Endif
    GO TOP
    If VAL(ULT)=0
      ULT=CODBACEN
    Else
      LOCATE FOR ULT=CODBACEN
      SKIP
    Endif
    If EOF()
      SQLDISCONNECT(CONEXINTOO)
      EXIT
    Endif
    ULT=CODBACEN
    CURSORWAIT()
    WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\DOCTOSNET\"
    VERDIRETORIO(WEBLOCAL1)
    If backup="N"
      SQLDISCONNECT(CONEXINTOO)
      LOOP
    Endif
    CONEX1ON=alltrim(serverdb)
    CONEX1DTBAS=alltrim(bancodados)
    CONEX1USER=alltrim(usuariodb)
    CONEX1SENHA=alltrim(senhadb)
    CONEX1PORT=3306
    CONEX1INTOO=60
    CONEX1CORRE=codbacen
    SQLDISCONNECT(CONEXINTOO)
    SQL DISCONNECT ALL
    BDOCTOSNET()
  ENDDO
  SQLDISCONNECT(CONEXINTOO)
  SQL DISCONNECT ALL
  CLOSE DATABASE
  abrearqs()
  CURSORARROW()
Return

Procedure BDOCTOSNET()
  CURSORWAIT()
  SQL CONNECT ON CONEX1ON;
    PORT CONEX1PORT;
    DATABASE CONEX1DTBAS;
    USER CONEX1USER;
    PASSWORD CONEX1SENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEX1INTOO

  SQLSETCONNECTION(CONEX1INTOO)
  If SQL_ErrorNO() > 0
    Return
  Endif
  SET PACKETSIZE TO 25
  WNRC="SELECT versao2,pastaid,pastapdf FROM doctosnet  WHERE sql_rowid=1"
  mz=SQLArrayAssoc(WNRC)
  wversao=MZ[1,1]
  wpastaid=MZ[1,2]
  wpastapdf=MZ[1,3]
  WNRC="SELECT sql_rowid,sql_rowid_cliente,CONVERT(CAST(CONVERT(arquivo USING LATIN1) AS BINARY) USING UTF8) arquivo FROM documentos  WHERE "+conex1tp+"='N' AND sql_deleted='F'"
  mz=SQLArrayAssoc(WNRC)
  wtotal=LEN(MZ)
  If wtotal>1000
    wtotal=1000
  Endif
  FOR IZ=1 TO wtotal
    wcodigo=MZ[IZ,1]
    wcaminho="C:\DOCTOSNET\"
    If CONEXTP=2
      WCAMINHO=Strtran(WCAMINHO,"C:","G:")
    Endif
    If !VERDIRETORIO(WCAMINHO)
      loop
    Endif
    wcaminho=trim(wcaminho)+trim(conex1corre)+"\"
    If !VERDIRETORIO(WCAMINHO)
      loop
    Endif
    wcaminho=wcaminho+alltrim(MZ[IZ,3])
    Erase &wcaminho
    If wversao="N"
      WARQ:=ALLTRIM(wpastapdf+"documentos/")+alltrim(MZ[IZ,3])
    Else
      WARQ:=ALLTRIM(wpastapdf)+wpastaid+"/documentos/"+alltrim(MZ[IZ,3])
    Endif
    oPg := CreateObject("Microsoft.XMLHTTP")
    oPg:Open("GET",warq,.f.)
    oPg:Send()
    wResult=opg:responsebody
    wstats=opg:status
    oPg:=nil
    If wstats<>200
      wfunc='UPDATE documentos set SQL_DELETED="T",useralt="Robo Backup",alterado=now() where sql_rowid='+wcodigo
      SQL EXECUTE WFUNC
      LOOP
    Endif
    memowrit(wcaminho,wResult)
    wfunc='UPDATE documentos set '+conex1tp+'="S" where sql_rowid='+wcodigo
    SQL EXECUTE wfunc
    SQL EXECUTE "COMMIT"
  NEXT IZ
  SQLDISCONNECT(CONEX1INTOO)
  SQLSETCONNECTION(CONEXINTOO)
Return

Procedure ATUALIZAASSINA(SEENVIAEMAIL)
  DECLARE TX1[1]
  NOMEARQ=""
  If (dow(date())=7 .or. dow(date())=1) .And. SEENVIAEMAIL="S"
    Return
  Endif
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If (HORAS<1400 .OR. DATE()=LOCALDPARASS) .And. SEENVIAEMAIL="S"
    Return
  Endif
  LCHDIR(WCORRENTE)
  If !FILE("CONEX.DAD")
    Return
  Endif
  RESTORE FROM CONEX.DAD ADDITIVE
  CONEXON=trim(Decrypt(CONEXON,"leopardo"))
  CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
  CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
  CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
  CONEXINTO=90
  If dow(date())=7
    CDT=DTOS(DATE()-1)+".CSV"
    ElseIf dow(date())=1
    CDT=DTOS(DATE()-2)+".CSV"
    Else
    CDT=DTOS(DATE())+".CSV"
  EndIf
  Erase &cdt
  oPg := CreateObject("MSXML2.ServerXMLHTTP.6.0")
  xComando := "http://www4.bcb.gov.br/Download/fechamento/"+CDT
  oPg:Open("GET",xComando,.f.)
  oPg:Send()
  wresult=opg:responsebody
  oPg:=nil
  marq1=WLOCTEMP+"COMPL.TXT"
  marq3=WLOCTEMP+"NAOFEITO.TXT"
  marq4=WLOCTEMP+"LINHAS.DBF"
  Erase &MARQ1
  memowrit(marq1,wresult)
  strzero(day(date()),2)
  deucerto=strzero(day(date()),2)+"/"+strzero(month(date()),2)+"/"+strzero(year(date()),4)
  If (AT(deucerto,wresult)=0 .OR. LEN(wresult)<7000) .and. SEENVIAEMAIL="S"
    Erase log.txt
    Erase &marq1
    Return
  Endif
  MAT_STRU={}
  AADD(MAT_STRU,{"LINHA","C",254,0})
  DBCREATE(marq4,MAT_STRU,"DBFCDX")
  SELE 33
  If !USEREDE(marq4,.T.,10,"TMP")
    SELE 35
    Return
  Endif
  APPEND FROM &marq1 SDF
  ULTMOE="000"
  SQL CONNECT ON CONEXON;
    PORT CONEXPORT;
    DATABASE CONEXDTBAS;
    USER CONEXUSER;
    PASSWORD CONEXSENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXINTO

  If SQL_ErrorNO() > 0
    Erase &marq1
    Return
  Endif
  SQLSETCONNECTION(CONEXINTO)
  SET PACKETSIZE TO 25
  SELE 35
  If !USEREDE("paridade",.F.,10,"PARIDADE","M")
    SELE 30
    Erase &marq1
    Return
  Endif
  WZII=1
  WDT=CTOD("")
  SELE 33
  GO TOP
  do while !EOF()
    WLI=ALLTRIM(LINHA)+";"+SPACE(50)
    If SUBSTR(WLI,3,1)<>"/"
      TX1[1]=WLI
      LOGFILE(marq3,TX1)
      SKIP
      LOOP
    Endif
    WDT=CTOD(LEFT(WLI,10))
    ZMOEDA=SUBSTR(WLI,12,3)
    ULTMOE=ZMOEDA
    ZSIMBOL=SUBSTR(WLI,18,3)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    COLU=AT(";",WLI)
    ZCOMP=SUBSTR(WLI,1,COLU-1)
    ZCOMP=Strtran(Strtran(ZCOMP,",","."),";","")
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    ZVEND=SUBSTR(WLI,1,COLU-1)
    ZVEND=Strtran(Strtran(ZVEND,",","."),";","")
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    ZPARC=SUBSTR(WLI,1,COLU-1)
    ZPARC=Strtran(Strtran(ZPARC,",","."),";","")
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,LEN(WLI))
    ZPARV=SUBSTR(WLI,1,COLU-1)
    ZPARV=Strtran(Strtran(ZPARV,",","."),";","")
    sele 35
    SINT1="moeda='"+zmoeda+"' and data="+dtom(wdt)
    SQL FILTER TO SINT1
    GO TOP
    If EOF() .or. BOF()
      ADIREG(0)
      REPLACE DATA WITH wdt
      REPLACE MOEDA WITH val(ZMOEDA)
      REPLACE SIMBOLO WITH ZSIMBOL
      REPLACE PARC WITH val(ZPARC)
      REPLACE PARV WITH val(ZPARV)
      REPLACE COMPRA WITH val(ZCOMP)
      REPLACE VENDA WITH val(ZVEND)
    Endif
    SELE 33
    SKIP
  ENDDO
  SELE 33
  USE
  //Erase &marq1
  SQLSETCONNECTION(CONEXINTO)
  sele 35
  use
  sele 30
  marq2=WLOCTEMP+dtos(date())+".ULT"
  If ULTMOE="000"
    If SEENVIAEMAIL="S" .and. !File(marq2)
      ENVIAEMAIL("wanderlei@exatus.net","BolsasNet "+ULTMOE+dtoc(wdt),"NÃO ATUALIZOU","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      MEMOWRIT(marq2,"enviado email de aviso")
    elseif SEENVIAEMAIL="N"
      MSGINFO("NÃO ATUALIZOU",wsist)
    Endif
    Return
  Endif
  LOCALDPARASS=DATE()
  Erase &marq2
  If SEENVIAEMAIL="S"
    ENVIAEMAIL("wanderlei@exatus.net","BolsasNet "+ULTMOE+dtoc(wdt),"Paridades Assinanet atualizadas com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
  Else
    MSGINFO("Paridades Assinanet atualizadas",wsist)
  Endif
Return

Procedure BAIXACONTRATOMENOS10K()
  DECLARE TX1[1]
  If dow(date())=7 .or. dow(date())=1
    Return
  Endif
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If !FILE("CONEX.DAD")
    Return
  Endif
  RESTORE FROM CONEX.DAD ADDITIVE
  CONEXON=trim(Decrypt(CONEXON,"leopardo"))
  CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
  CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
  CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
  CONEXINTO=90

  If SEENVIAEMAIL="N" .OR. (HORAS>2000 .And. DATE()#LOCALDPARID)
    LOCALDPARID=date()
    oPg1 := CreateObject("Microsoft.XMLHTTP")
    xComando := "https://www.exatus.net/grafico/taxasc.csv"
    oPg1:Open("GET",xComando,.f.)
    oPg1:Send()
    WRESULT=opg1:responsebody
    oPg1:=nil
    CURSORWAIT()
    wdt=ctod(left(alltrim(right(wresult,17)),10))
    marq1=WLOCTEMP+"COMPLETO.TXT"
    memowrit(marq1,wresult)
    oText:=TTxtfile():NEW(marq1)
    CONTAD=oText:RecCount()
    If VALTYPE(CONTAD)<>"N"
      CONTAD=1
    Endif
    ULTMOE="000"
    If CONTAD<20
      oText:Close()
      Erase &marq1
    Else
      CURSORWAIT()
      SQL CONNECT ON CONEXON;
        PORT CONEXPORT;
        DATABASE CONEXDTBAS;
        USER CONEXUSER;
        PASSWORD CONEXSENHA;
        OPTIONS SQL_NO_WARNING;
        LIB 'MySQL';
        INTO CONEXINTO


      If SQL_ErrorNO() > 0
        Return
      Endif
      SQLSETCONNECTION(CONEXINTO)
      SET PACKETSIZE TO 25
      SELE 35
      If !USEREDE("paridade",.F.,10,"PARIDADE","M")
        SELE 30
        Return
      Endif
      FOR WZII=1 TO CONTAD
        WLI=ALLTRIM(oText:ReadLine())+";"+SPACE(50)
        If SUBSTR(WLI,4,1)<>"-"
          oText:Skip()
          LOOP
        Endif
        ZMOEDA=LEFT(WLI,3)
        ULTMOE=ZMOEDA
        ZSIMBOL=SUBSTR(WLI,5,3)
        ZPARC=Strtran(Strtran(SUBSTR(WLI,9,14),",","."),";","")
        ZPARV=Strtran(Strtran(SUBSTR(WLI,24,14),",","."),";","")
        ZCOMP=Strtran(Strtran(SUBSTR(WLI,39,14),",","."),";","")
        ZVEND=Strtran(Strtran(SUBSTR(WLI,54,14),",","."),";","")
        SINT1="moeda="+zmoeda+" and data="+dtom(wdt)
        SQL FILTER TO SINT1
        GO TOP
        If EOF() .or. BOF()
          ADIREG(0)
          REPLACE DATA WITH wdt
          REPLACE MOEDA WITH val(ZMOEDA)
          REPLACE SIMBOLO WITH ZSIMBOL
          REPLACE PARC WITH val(ZPARC)
          REPLACE PARV WITH val(ZPARV)
          REPLACE COMPRA WITH val(ZCOMP)
          REPLACE VENDA WITH val(ZVEND)
        Endif
        oText:Skip()
      NEXT
    Endif
    USE
    oText:Close()
    Erase &marq1
    If SEENVIAEMAIL="S"
      ENVIAEMAIL("wanderlei@exatus.net","BolsasNet "+ULTMOE,"Paridades Assinanet atualizadas com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    Else
      MSGINFO("Paridades Assinanet atualizadas",wsist)
    Endif
  Endif
Return

