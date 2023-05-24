//**************************************************************************************************************
// Programa.: WORKER                                                                                           *
// Autor....: Wanderlei Souza Cruz                                                                             *
// Data.....: 21/01/2003                                                                                       *
// Assinanet: 14/08/2011                                                                                       *
// BBI......: 16/11/2011                                                                                       *
// ASSINADOW: 18/08/2014                                                                                       *
// TAXAS AUTO 30/04/2015                                                                                       *
// LISTAS...: 11/02/2016 -CONEX.LIS - RODA JUNTO COM O ASSINANET APOS O DIA 23 DE CADA MES AOS SABADOS         *
// COBRA DOC: 28/09/2016                                                                                       *
// ASSINANET-ENVIO AUTOMATICO DE CONTRATOS DO SICTUR: 30/11/2018                                               *
// Aviso....: Copyright (c) 2015, Exatus Com.Proc.Dados Ltda, All Rights Reserved                              *
//                                                                                                             *
// Observações:                                                                                                *
//                                                                                                             *
// Sistema Local, arquivo (Bolsanet.dbf) se for incluir campos, colocar no final para haver o reconhecimento do campo validade E NUNCA FAZER PELO WORKER *
// ARQUIVO 14203.TXT = CONFIDENCE                                                                              *
// https://assinanet.exatus.net/administrativo/wsdownloadassinado.asp?documento=64215                          *                                                                                  *
//**************************************************************************************************************
#include "FiveWin.ch"
#include "TSButton.ch"
#include "Report.ch"
#include "FILEIO.CH"
#include "InKey.ch"
#include "winapi.ch"
#include "BMPGET.CH"
#include "FILEXLS.CH"
#include "Struct.Ch"
#include "barcode.ch"
#include "sqllib.ch"
#include "tip.ch"
#define  HKEY_CURRENT_USER 2147483649

REQUEST MySql
REQUEST DBFFPT

Function Main()
  Local oBrush, oBar, oIcon,oChildWnd
  PUBLIC oWnd,oTimer,oTimer1,oFont4,tx1[1],tx2[2],oActiveX
  PUBLIC oCT
  cSvSetEOL := SET(_SET_EOL) //Guardando na váriavel o caracter de finalização, para ser usado no APPEND FROM...SDF
  SET DATE FRENCH
  SET CENTURY ON
  SET MULTIPLE  ON
  DECLARE TX1[1]
  DECLARE TX2[2]
  TX2[1]=""
  TX2[2]=.T.
  TX1[1]=""
  wcorpt='<html><head> <title>Travelex Bank - Assinatura de Câmbio</title>    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head><body><table width="700" cellspacing="0" bgcolor="#ffffff" cellpadding="0" border="0" align="center">    <tr>        <td border="0" width="700" align="center">            <a href="https://www.travelexbank.com.br/" target="_blank" rel="noopener noreferrer"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_01.png"></a>                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_02.png">                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_03.png">                                               <a href="http://exatus.net/downloads/setupassinanetcliente.exe"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_04.png"></a>                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_05.png">                                               <a href="https://assinanet.exatus.net/"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_06.png"></a>                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_07.png">                                               <a href="https://assinanet.exatus.net/manualcliente.pdf"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_08.png"></a>                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_09.png">                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_10.png">                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_11.png">                                               <a href="https://api.whatsapp.com/send/?phone=551130040490&text&app_absent=0"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_12.png"></a>                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_13.png">        </td>                </tr></table><table width="700" cellspacing="0" bgcolor="#ffffff" cellpadding="0" border="0" align="center">    <tr>        <td border="0" align="center">                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_14.png">        </td>        <td border="0" align="center">                                               <a href="https://www.instagram.com/travelexbank/"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_15.png"></a>        </td>        <td border="0" align="center">                                               <a href="https://twitter.com/travelexbank"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_16.png"></a>        </td>        <td border="0" align="center">                                               <a href="https://www.linkedin.com/company/travelex-confidence?trk=tyah"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_17.png"></a>        </td>        <td border="0" align="center">                                               <img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_18.png">        </td>                </tr></table><table width="700" cellspacing="0" bgcolor="#ffffff" cellpadding="0" border="0" align="center">    <tr>        <td border="0" align="center">                                               <a href="https://www.travelexbank.com.br/"><img src="https://respositorio-bi.s3.amazonaws.com/tvx-bank/21.09.15-assinatura-de-cambio/assinatura-de-cambio_19.png"></a>        </td>                </tr></table></body></html>'

  INIOUVI="00:00:00"
  INITROC="00:00:00"
  VALIDOATE=CTOD("30/06/2050")
  confidence=.f.
  PARAMETERS sen
  WSEN=SEN
  QPAR=PCOUNT()
  FAZENDO=.F.
  EXTERN DBFCDX
  RDDSETDEFAULT("DBFCDX")
  If FILE("14203.TXT")
    CONFIDENCE=.T.
  EndIf
  MARQ=CURDRIVE()+TRIM(":\"+CURDIR())+"\LIBMYSQL.DLL"
  If !FILE(MARQ)
    MSGINFO("Falta "+MARQ+", não será possivel continuar !","")
    QUIT
  EndIf
  If !FILE("CONEX.DAD")
    msgSTOP("Utilize o programa automa.exe C para configurar conexão !","")
    return 
  EndIf
  RESTORE FROM CONEX.DAD ADDITIVE
  TEMPDF=.F.
  WDADOS=""
  ECORRETORA=.F.
  WSIST=""
  TEMPDF1=""
  VERIMPRESSORAS()
  If !TEMPDF
    MSGINFO("Não existe a impressora WIN2PDF instalada, não é possivel continuar !","")
    return 
  EndIf
  WCORRENTE1=ALLTRIM(CURDIR())
  WCORRENTE=CURDRIVE()+":"
  If LEN(WCORRENTE1)>1
    WCORRENTE=WCORRENTE+"\"+ALLTRIM(CURDIR())
  EndIf
  LOCALWEB=wcorrente+"\dados\"
  If !VERDIRETORIO(LOCALWEB)
    MSGINFO("Não foi possivel criar o diretorio "+localweb,WSIST)
    QUIT
  EndIf
  WLOCTEMP :="C:\TEMPO\"
  If !VERDIRETORIO(WLOCTEMP)
    MSGSTOP("Diretorio: "+WLOCTEMP+" Inexistente !, deve ser criado para poder continuar o sistema",wsist)
    return 
  EndIf
  LOCALNOT1=WCORRENTE+"\TEMP\"
  If !VERDIRETORIO(LOCALNOT1)
    MSGINFO("Nao foi possivel criar "+LOCALNOT1,"")
    return 
  EndIf
  LOCALNOT1=WCORRENTE+"\NOTTEMP\"
  LOCALNOT2=WCORRENTE+"\TAXAS\"
  LOCALNOT3=WCORRENTE+"\NOTDADOS\"
  LOCALNOT5=WCORRENTE+"\NEGADOS\"
  LOCALDTAXAS=DATE()-1
  LOCALDPARASS=DATE()-1
  LOCALDASSINA=DATE()-1
  LOCALDPARID=DATE()-1
  LOCALTAXAS=WCORRENTE+"\INETPUB\WWWROOT\"
  LOCALCART=WCORRENTE+"\WINCART\"
  LOCALCOTNOT=3
  LOCALINTERVALO=120
  LOCALdiasnot=3
  LOCALDIASTAXAS=3
  LOCALFROM="exatus@exatus.net"
  LOCALSERVER="smtp.exatus.net"
  LOCALGRAFIC=WCORRENTE+"\INETPUB\WWWROOT\GRAFICOS\"
  LOCALGNEW=WCORRENTE+"\INETPUB\WWWROOT\GRAFICO\"
  LOCALBC=ALLTRIM(LOCALGRAFIC)
  LOCALINICIO="0000"
  LOCALFIM="2358"
  WWDATADOL=""
  ULTBUSCA="00:00:00"
  WWHORADOL=""
  WWDATANOT=""
  WWHORANOT=""
  If FILE("VERHORA.MEM")
    RESTORE FROM VERHORA.MEM ADDITIVE
  EndIf
  If FILE("ARQPAR.DAT")
    RESTORE FROM ARQPAR.DAT ADDITIVE
  EndIf
  LOCNOT1=ALLTRIM(LOCALNOT1)
  LOCNOT2=ALLTRIM(LOCALNOT2)
  LOCNOT3=ALLTRIM(LOCALNOT3)
  LOCALBC=ALLTRIM(LOCALBC)
  LOCALTAXAS=ALLTRIM(LOCALTAXAS)
  LOCALCART=ALLTRIM(LOCALCART)
  LOCNOT5=ALLTRIM(LOCALNOT5)
  M->FROM=ALLTRIM(LOCALFROM)
  CIPSERVER=LOCALSERVER
  PREV=.T.
  WA="$ÁÉÍÓÚáéíóúçÇãõÃÕàèìòùÀÈÌÒÙÂÊÎÔÛâêîôûäëïöüÄËÏÖÜºª"
  FOR WI=1 TO 31
    SysRefresh()
    WA=WA+CHR(WI)
  NEXT
  FOR WI=128 TO 255
    SysRefresh()
    WA=WA+CHR(WI)
  NEXT
  WB="sAEIOUaeioucCaoAOaeiouAEIOUAEIOUaeiouaeiouAEIOU"+SPACE(300)
  EMP3="Exatus.Net"
  NAOCRIOU=""
  If !VERDIRETORIO(LOCNOT1) .OR. !VERDIRETORIO(LOCNOT2) .OR. !VERDIRETORIO(LOCNOT3) .OR. !VERDIRETORIO(LOCNOT5) .OR. !VERDIRETORIO(LOCALCART)
    MSGINFO("Nao foi possivel criar "+NAOCRIOU,"")
    return 
  EndIf
  ULTIMAATU="00:00:00"
  ULTATU1="00:00:00"
  PRIMEIRA=0
  PRIMEIRA1=CTOD("")
  PRIMEIRA2=0
  If FILE("LOCAIS.DAT")
    RESTORE FROM LOCAIS.DAT ADDITIVE
  EndIf
  PUBLIC oFontSay
  PUBLIC oFontSayB
  PUBLIC oFontGet
  PUBLIC oFontBut
  PUBLIC oFontTit
  PUBLIC oFontChk
  PUBLIC oFontLis
  PUBLIC oBrushTela
  MFONTES()
  SET TALK OFF
  SET DATE FRENCH
  SET BELL OFF
  SET STAT OFF
  SET SCORE OFF
  SET WRAP ON
  SET CURSOR OFF
  SET DELETED ON
  SET EPOCH TO 2000
  SET DECIMAL TO 2
  SET CENTURY ON
  SET EXCLUSIVE OFF
  SETHANDLECOUNT(250)
  If !FILE("PREV32.DLL")
    MSGINFO("Falta PREV32.DLL, sem o qual não poderá ser emitido qualquer relatório ","")
    return  .F.
  EndIf
  If !FILE("qrcodelib.dll")
    MSGINFO("Falta QRCODELIB.DLL, sem o qual não poderá ser gerado o QRCODE ","")
    return  .F.
  EndIf
  COD1="000000"
  M->CALCATE=DATE()
  _DATA=""
  EMP1=""
  _LOCAL=""
  NRAA="000"
  NRAD=SPACE(40)
  NRAPREV=.T.
  _DATA=STRZERO(VAL(_DATA),4)
  arq01 = LOCALCART+"EMPRESAS.DBF"   // Cadastro de Empresas
  arq02 = LOCALCART+"OPERADOR.DBF"   // Cadastro de Operadores
  arq03 = _LOCAL+"PAPE2006.DBF"      // Cadastro de Papeis da Carteira
  arq04 = _LOCAL+"MERC2006.DBF"      // Cadastro de Mercados
  ARQ05 = _LOCAL+"RDAR2006.DBF"      // Resumo de Auração de Ganhos Renda Variavel
  ARQ06 = LOCALCART+"BDIS.DBF"       // Arquivo de BDI completo
  ARQ07 = _Local                     // Arquivos de Notas em XML
  ARQ08 = LOCALCART+"PARCMV.DBF"     // Parametrizaçào do arquivo CMVD
  arq10 = _LOCAL+"MOVT2006.DBF"      // Movimento Diario
  arq11 = _LOCAL+"MOVN2006.DBF"      // Movimento Notas
  arq12 = _LOCAL+"NOTA2006.DBF"      // Dados das Notas
  arq13 = _LOCAL+"INTE2006.DBF"      // Integração dos Lançamentos
  arq14 = LOCALCART+"PAPEIS.DBF"     // Cadastro de Papeis p/Movimento
  arq15 = _LOCAL+"BOLSA.DBF"         // Arquivo de Notas de Bolsa
  arq16 = _LOCAL+"BMF.DBF"           // Arquivo de Notas de BMF
  ind0101 = LOCALCART+"EMPRESA1.IND" // Indice de Empresas
  ind0201 = LOCALCART+"OPERADOR.IND" // Indice de Operadores
  ind0301 = _LOCAL+"PAPE2006.IND"    // Indice de Papeis da Carteira
  ind0401 = _LOCAL+"MERC2006.IND"    // Indice do Cadastro de Mercados
  ind0501 = _LOCAL+"RDAR2006.IND"    // Indice Resumo de Auração de Ganhos Renda Variavel
  ind0601 = LOCALCART+"BDIS.IND"     // Indice do Arquivo de BDI
  ind0801 = LOCALCART+"PARCMV.IND"   // Indice da Parametrizaçào do arquivo CMVD
  ind1001 = _LOCAL+"MVTO2006.IND"    // Indice do Movimento Diairo
  ind1101 = _LOCAL+"MVTC2006.IND"    // Indice do Movimento Notas
  ind1201 = _LOCAL+"NOTA2006.IND"    // Indice Dados da Nota
  ind1301 = _LOCAL+"INTE2006.IND"    // Indice Integração Contabil
  ind1401 = LOCALCART+"PAPEIS.IND"   // Indice Cadastro de Papeis p/Movimento
  ind1501 = _LOCAL+"BOLSA.IND"       // Indice Arquivo de Notas Bolsa
  ind1601 = _LOCAL+"BMF.IND"         // Indice Arquivo de Notas BMF
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
  chv0601 = "CODPAPEL+DTOS(DATA)"
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
  _CGCCPF=""
  AUTOMATICO:=.F.
  If LOCALQTIPO="W"
    If QPAR=0
      AUTOMATICO=.F.
    Else
      AUTOMATICO=.T.
    EndIf
  EndIf
  A02NOME=""
  If AUTOMATICO
    cod_oper=999
    A02NOME="MASTER - Exatus"
    FOR I=1 TO 20
      SysRefresh()
      _var4="OP"+ALLTRIM(STR(I))
      &_var4="S"
    NEXT I
  EndIf
  CRIARQ()
  CRIAMOV()
  INDICES()
  ABREARQS()
  EMP2="(0xx11) 2227.3078 - www.exatus.net "
  M->DAT_HOJE=DATE()
  OP1:=OP2:=OP3:=OP4:=OP5:=OP6:=OP7:=OP8:=OP9:=OP10:=" "
  OP11:=OP12:=OP13:=OP14:=OP15:=OP16:=OP17:=OP18:=OP19:=OP20:=op21:=op22:=op23:=op24:=op25:=op26:=op27:=op28:=op29:=op30:=" "
  CONTINUA=.F.
  SNDPLAYSOUND("WELCOME.WAV",0)

  If QPAR=0
    If !QUEM()
      QUIT
    EndIf
  EndIf

  If QPAR<>0
    FOR I=1 TO 20
      SysRefresh()
      _var4="OP"+ALLTRIM(STR(I))
      &_var4="S"
    NEXT I
    COD_OPER:=999
    A02NOME="MASTER - Exatus"
  EndIf
  If M->DAT_HOJE>VALIDOATE
    QUIT
  EndIf
  SOUADM=.F.   //  PARA FAZER ATUALIZAÇÃO AUTOMATICA
  If LOCALQTIPO<>"W"
    MMMM(1)
    CRIARQC()
    INDICESC()
  EndIf
  If LOCALQTIPO="C"
    CLOSE DATABASE
    WSIST="Sistema de Ganho Renda Variavel em Ações"
    DEFINE ICON oIcon FILE "CARTEIRA.ICO"
  Else
    WSIST="Worker"
    DEFINE ICON oIcon FILE "WORKER.ICO"
  EndIf
  If AUTOMATICO
    If WSEN="C"
      WSIST="Carteira de Ações"
    ElseIf WSEN="B"
      WSIST="CambioNet"
    ElseIf WSEN="A"
      WSIST="Assinanet"
    ElseIf WSEN="W"
      WSIST="BolsasNet"
    EndIf
    DEFINE ICON oIcon FILE "WORKER.ICO"
  EndIf

  If LOCALQTIPO="W" .And. COD_OPER<997
    If !VERDIRETORIO("C:\BOLSANET\")
      MSGINFO("Nao foi possivel criar C:\BOLSANET",WSIST)
      CLOSE DATABASE
      QUIT
    EndIf
    SELE 30
    GO 1
    WPASTA=ALLTRIM(PASTA)
    WSIST="Bolsas Net - Desktop"
    DEFINE ICON oIcon FILE "CARTEIRA.ICO"
    SetKey( VK_F2, { || (PRINTERSETUP(),oitem2:settext(PRNGETNAME()),oitem2:refresh()) } )
    DEFINE FONT oFonttela NAME "verdana" SIZE 6,12
    DEFINE BITMAP oBmp RESOURCE "TELA"
    DEFINE WINDOW oWnd FROM 0,0 TO 31,97.4 STYLE nOr( WS_POPUP ) TITLE "BolsasNet"

    // CADASTRO
    SX=90
    If OP17="S"
      SX=SX+40
      @ SX,25  BTNBMP RESOURCE "mcli" ADJUST OF oWnd PIXELS SIZE 200,31 //ACTION  NCADCLI()  // cadastro de clientes
    EndIf
    If OP19="S"
      SX=SX+40
      @ SX,25 BTNBMP RESOURCE "USER"  ADJUST OF oWnd PIXELS SIZE 200,31 ACTION CADUSER()
    EndIf
    // RELATORIOS
    If OP18="S"
      SX=90
      WARQ1W=WPASTA+"IMPRIME\ETIQAT10.PDF"
      WARQ2W=WPASTA+"IMPRIME\ETIQM10.PDF"
      WARQ3W=WPASTA+"SEPARADO\IMPRIME.ZIP"
      WARQ4W=WPASTA+"NAOENCO\IMPRIME.ZIP"
      WARQ5W=WPASTA+"IMPRIME\IMPRIME.ZIP"
      WARQ6W=WPASTA+"FRENTEV\IMPRIME.ZIP"
      WARQ7W=WPASTA+"IMPRIME\RELAT.PDF"
      WARQ8W=WPASTA+"IMPRIME\RELAT.XLS"
      If FILE(WARQ1W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL1" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ1W) // Etiquetas até 10
      EndIf
      If FILE(WARQ2W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL2" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ2W) // Etiquetas acima de 10
      EndIf
      If FILE(WARQ3W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL3" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ3W) // separados
      EndIf
      If FILE(WARQ4W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL4" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ4W) // Não encontrados
      EndIf
      If FILE(WARQ5W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL5" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ5W) // mais de 1 para envelopar
      EndIf
      If FILE(WARQ6W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL6" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ6W) // Frente e verso
      EndIf
      If FILE(WARQ7W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL7" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ7W) // Resumo PDF
      EndIf
      If FILE(WARQ8W)
        SX=SX+40
        @ SX,265 BTNBMP RESOURCE "REL8" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION RUNFILE(WARQ8W) // Resumo XLS
      EndIf
    EndIf
    // MANUTENÇÃO
    SX=90
    If OP21="S"
      SX=SX+40
      @ SX,505 BTNBMP RESOURCE "PARAM" ADJUST OF oWnd PIXELS SIZE 200,31 // ACTION PARA1()
    EndIf
    SX=SX+40
    @ SX,505 BTNBMP RESOURCE "IMPRESS" ADJUST OF oWnd PIXELS SIZE 200,31 ACTION (PRINTERSETUP())
    SX=SX+40
    @ SX,505 BTNBMP RESOURCE "LOG" ADJUST OF oWnd PIXELS SIZE 200,31 // ACTION (LOGS())
    If OP20="S"
      SX=SX+80
      @ SX,505 BTNBMP RESOURCE "GERAR" ADJUST OF oWnd PIXELS SIZE 200,31 // ACTION (EXECUTA())
    EndIf
    @ 3,752  BTNBMP RESOURCE "mfechar" bdt1 NOBORDER OF oWnd PIXELS SIZE 24,21 ACTION (SNDPLAYSOUND("GOODBYE.WAV",0),oWnd:End()) // Sair do Sistema

    DEFINE TIMER oTimer INTERVAL (30*1000) ACTION ROT() OF oWnd
    ACTIVATE TIMER oTimer
    oWnd:Center()
    ACTIVATE WINDOW oWnd NORMAL     ON PAINT PalBmpDraw( oWnd:hDC,0,0,oBmp:hBitmap, oBmp:hPalette)
  Else
    DEFINE BRUSH oBrush COLOR nRGB(0,0,0)
    WFIGC="CONTR01.fig"
    If localqtipo="C"
      DEFINE BRUSH oBrush COLOR nRGB(255,255,255)
      WFIGC="TELAACOES"
    EndIf
    DEFINE WINDOW oWnd FROM 0,40 to 40,120 TITLE WSIST BRUSH oBrush ICON oIcon MENU BuildMenu()
    DEFINE BITMAP oBmps RESOURCE WFIGC ADJUST
    If LOCALQTIPO="W"
      emp2=space(40)
      SET MESSAGE OF oWnd to emp2 CENTERED
      oItem := TMsgItem():New( oWnd:oMsgBar, dtoc(date()), 80,oFontsay,"R/W",,.T. )
      oItem := TMsgItem():New( oWnd:oMsgBar, trim(A02NOME), 80,oFontsay,"R/W",,.T. )
      oItem := TMsgItem():New( oWnd:oMsgBar, "Versão: 22/05/2023", 140,oFontsay,"R/W",,.T. )
    Else
      WEMP2=EMP2+TRIM(EMP1)+" ref:"+_DATA+" Local: "+_LOCAL
      SET MESSAGE OF oWnd TO WEMP2 CENTERED
      oItem2 := TMSGItem():New( oWnd:oMSGBar, PRNGETNAME(), 140,oFontsay,"R/W",,.T. )
      oItem := TMsgItem():New( oWnd:oMsgBar, dtoc(date()), 80,oFontsay,"R/W",,.T. )
      oItem := TMsgItem():New( oWnd:oMsgBar, "Versão: 22/05/2023", 140,oFontsay,"R/W",,.T. )
    EndIf
    If AUTOMATICO
      If WSEN="C"
        oItem:= TMsgItem():New( oWnd:oMsgBar, "CarteiraNet", 140,oFontsay,"R/W",,.T. )
      ElseIf WSEN="W"
        oItem:= TMsgItem():New( oWnd:oMsgBar, "CambioNet", 140,oFontsay,"R/W",,.T. )
      ElseIf WSEN="B"
        oItem:= TMsgItem():New( oWnd:oMsgBar, "BolsasNet", 140,oFontsay,"R/W",,.T. )
      ElseIf WSEN="A"
        oItem:= TMsgItem():New( oWnd:oMsgBar, "AssinaNet", 140,oFontsay,"R/W",,.T. )
      EndIf
    EndIf
    If AUTOMATICO
      ROTINAS()
      DEFINE TIMER oTimer INTERVAL (LOCALINTERVALO*1000) ACTION (ROTINAS()) OF oWnd
      ACTIVATE TIMER oTimer
    EndIf
    If !AUTOMATICO
      ACTIVATE WINDOW oWnd MAXIMIZED ON PAINT PalBmpDraw( oWnd:hDC,(oWnd:nHeight()/2)-(oBmps:nHeight()/2)-50,(oWnd:nWidth() /2)-(oBmps:nWidth() /2),oBmps:hBitmap, oBmps:hPalette)
    Else
      ACTIVATE WINDOW oWnd ICONIZED  ON PAINT PalBmpDraw( oWnd:hDC,(oWnd:nHeight()/2)-(oBmps:nHeight()/2)-50,(oWnd:nWidth() /2)-(oBmps:nWidth() /2),oBmps:hBitmap, oBmps:hPalette)
    EndIf
  EndIf
  sysRefresh()
  CLOSE DATABASE
  If LOCALQTIPO="C"
    M->LOCAL=LOCALCART
    DECLARE VLABEL[ADIR(LOCALCART+"TMP*.*")]
    ADIR(LOCALCART+"TMP*.*",vlabel)
    VLABEL:=ASORT(VLABEL)
    I=0
    FOR I=1 TO LEN(VLABEL)
      SysRefresh()
      NARQ=LOCALCART+VLABEL[I]
      Erase &NARQ
    NEXT I
  Else
    DECLARE VLABEL[ADIR(LOCALNOT1+"*.*")]
    ADIR(LOCALNOT1+"*.*",vlabel)
    I=0
    FOR I=1 TO LEN(VLABEL)
      Sysrefresh()
      NARQ=LOCALNOT1+VLABEL[I]
      Erase &NARQ
    NEXT I
    DECLARE VLABEL[ADIR("TMP*.*")]
    ADIR("TMP*.*",vlabel)
    I=0
    FOR I=1 TO LEN(VLABEL)
      SysRefresh()
      NARQ=VLABEL[I]
      Erase &NARQ
    NEXT I
    DECLARE VLABEL[ADIR("*.TMP")]
    ADIR("*.TMP",vlabel)
    I=0
    FOR I=1 TO LEN(VLABEL)
      SysRefresh()
      NARQ=VLABEL[I]
      Erase &NARQ
    NEXT I
  EndIf
return ( NIL )

Procedure LIMPAMEM()
  oWnd:Normal()
  oWnd:Minimize()
return 

Procedure MFONTES()
  DEFINE FONT oFontsay  NAME "Verdana" SIZE 0,-11
  DEFINE FONT oFontsayB NAME "Verdana" SIZE 0,-10 BOLD
  DEFINE FONT oFontGet  NAME "Verdana" SIZE 0,-9
  DEFINE FONT oFontBut  NAME "Verdana" SIZE 0,-11
  DEFINE FONT oFontTit  NAME "Verdana" SIZE 0,-11
  DEFINE FONT oFontChk  NAME "Verdana" SIZE 0,-11
  DEFINE FONT oFontLis  NAME "Verdana" SIZE 0,-9
  DEFINE BRUSH oBrushTela COLOR nRGB(236,233,216)
return 

Function BuildMenu()
  Local oMenu
  MENU oMenu
  If LOCALQTIPO="W" .And. COD_OPER<997
    MENUITEM "&Cadastros"
    MENU
    MENUITEM "Clientes"
    ENDMENU
  ElseIf LOCALQTIPO="W" .And. !AUTOMATICO
    MENUITEM "&Cadastros"
    MENU
    SEPARATOR
    MENUITEM "&Comercial"         ACTION COMERCIAL() ACCELERATOR ACC_ALT,ASC("C")
    separator
    //      MENUITEM "&Noticias-EXATUS"   ACTION CADNOT(1)
    //      MENUITEM "Ta&xas"             ACTION CADTAXA(1)
    //      MENUITEM "&Usuarios-Boletins" ACTION CADBOL()
    MENUITEM "Empresas BolsasNet"  ACTION EMPVAL(3)
    //      MENUITEM "Empresas CambioNet"  ACTION EMPVAL(2)
    //      MENUITEM "Usuários CambioNet"  ACTION CADWINCC()
    SEPARATOR
    MENUITEM "Todas Taxas" ACTION MSGINFO("Salva as taxas do dia que voce quer no diretório \exatus com o nome de Taxas.csv !",WSIST),TODASTAXAS("N")
    menuitem "Todas Taxas baixando do BACEN" action TODASTAXAS("S")
    SEPARATOR
    MENUITEM "BolsasNet" ACTION RECPDF("N")
    MENUITEM "Integra Cambio-Sictur/Viewer/CSNU" ACTION INTCAM("N") // planilhas enviadas ao servidor integra no SICTUR
    MENUITEM "Arq.Finald"  ACTION TXFINAUD("N")
    SEPARATOR
    MENUITEM "Assina Net"  ACTION ASSINANET("N")
    MENUITEM "AssinaNet-Envia Completos" ACTION ENVCOMPLETO("N")
    MENUITEM "Cobrar Documentos" ACTION COBRADOC("N")
    MENUITEM "Backup Assinanet" ACTION ASSINABACKUP()
    MENUITEM "Backup Sictur" ACTION BACKDOCSICTUR()
    MENUITEM "Backup Doctosnet" ACTION BACKDOCNET()
    MENUITEM "Gera SICTUR Assinanet" ACTION GERAASSINA("N")
    MENUITEM "Gera Ficha Assinanet-PF/PJ" ACTION GERAFICHAPF("N")
    MENUITEM "Atualiza Paridade Assinanet"  ACTION ATUALIZAASSINA("N")
    separator
    //      MENUITEM "Puxa servidores" ACTION PUXASERVIDORES()
    MENUITEM "Ajusta Base" ACTION AJUSTABASE()
    If CONFIDENCE
      menuitem "Tem - Confidence"  ACTION ATUACONFIDENCE()
    Else
      menuitem "Nao Tem - Confidence"  ACTION ATUACONFIDENCE()
    EndIf
    SEPARATOR
    MENUITEM "&Cad.Listas Restritivas" ACTION CADTROCA()
    MENUITEM "Busca e Processa Listas"  ACTION LISTA("N")
    MENUITEM "TESTE PushNotification" ACTION (enviaaviso("teste de envio Push Notification","06666901894"),enviaaviso("Teste de envio Push Notification ","19861349898"))
    SEPARATOR
    MENUITEM "Calculo Get Money"  ACTION (DT:=DATE(),DT:=OUTROS("Data:",DT,"@D"),CALCULAGET(DT))
    SEPARATOR
    MENUITEM "CAM021" ACTION CAM0021()
    MENUITEM "msbank-Continuar de onde parou" ACTION QB()
    separator

    MENUITEM "Envelopes Diversos" ACTION ETIQU()
    SEPARATOR
    MENUITEM "Altera Senha SICTUR" ACTION senhasictur()
    MENUITEM "Gera Senha" action msginfo(GEraseNHA(12),wsist)
    MENUITEM "Email Exatus" ACTION  ENVIAEMAIL("Wanderlei.cruz@starmoney.com.br","Assunto: Teste Email Exatus",wcorpt,"c:\t4\cpf.pdf","c:\t4\comp.pdf","exatus@exatus.net","ckutzfiu","smtp.exatus.net","exatus@exatus.net",.F.)
    MENUITEM "Planner" action limpaplanner()
    MENUITEM "Baixar Contratos menores que 10k" action BAIXACONTRATOMENOS10K()
    MENUITEM "Email Amazon EWS-SES" ACTION (ENVIAEMAIL("Wanderlei.cruz@starmoney.com.br","Assunto: Teste Email EWS-SES","<html><head><title>Teste exemplo pela AMAZON-EWS</title></head><body><h1>Teste Envio pela AMAZON-EWS</h1><p>","c:\t4\w1.pdf","c:\t4\w2.pdf","AKIAVEWB2BCCYC2EY3PI","BD0+qOl39Uy+iUrHlzNBfGCGg+eJ3S/+RoG6KbPQ/MWL","email-smtp.sa-east-1.amazonaws.com","custodia@simpaul.com.br",.F.),msginfo("enviado"))
    ENDMENU
    MENUITEM "Manutenção"
    MENU
    MENUITEM "&Usuarios-WORKER" ACTION CADUSER()
    SEPARATOR
    MENUITEM "&Manut.Arquivos"  ACTION IIF( MSGYesno( "Faz manutenção dos Arquivos ?",WSIST ),INDICES(1),)
    MENUITEM "&Limpeza nos Arquivos" ACTION ROTINASDIA()
    MENUITEM "&Zera Ultima Busca" ACTION (ULTBUSCA:="00:00:00",ROTINAS(),MSGINFO("Zerado !",""))
    SEPARATOR
    MENUITEM "&Log Utilização" ACTION CADLOG() MESSAGE "Consulta as Rotinas utilizadas pelos Usuários"
    SEPARATOR
    MENUITEM "&Parametros" ACTION PARAMETROS() MESSAGE "Cadastro dos Parametros"
    SEPARATOR
    MENUITEM "&Setup da Impressora" ACTION PRINTERSETUP()
    SEPARATOR
    If A02NOME="MASTER - Exatus"
      SEPARATOR
      MENUITEM "Muda para Carteira" ACTION (MUDAQTIPO("C"),oWnd:End())
    EndIf
    MENUITEM "Mala direta"  action MALAD()
    MENUITEM "CONTROLE" ACTION ACERTACONTROL()
    MENUITEM "CONSULTA LISTA INDIVIDUAL" ACTION CONSIND()
    ENDMENU
    MENUITEM "&Sair" MESSAGE "Finaliza o Sistema" ACTION sair()
  ElseIf LOCALQTIPO="C" .And. !AUTOMATICO
    MENUITEM "&Arquivos"
    MENU
    MENUITEM "&Papeis"
    MENU
    MENUITEM "&Papeis" ACTION CADPAP()
    If COD_OPER>990
      MENUITEM "BDI &Automático" ACTION AUTOPU(0)
    EndIf
    ENDMENU
    MENUITEM "&Carteira" ACTION CADSALD()
    SEPARATOR
    If COD_OPER>990
      MENUITEM "&Base e Imposto Anterior" ACTION CADIMP()
      SEPARATOR
    EndIf
    MENUITEM "&Classificação Contabil" ACTION CADCONT()
    ENDMENU
    MENUITEM "&Movimentação"
    MENU
    MENUITEM "&Movimento Notas" ACTION CADNOTAS()
    MENUITEM "Movimento &Extra" ACTION CADDIV()
    MENUITEM "Importa Notas (Layout Worker)" ACTION INTNOTA()
    If COD_OPER>990
      SEPARATOR
      MENUITEM "Calcular todas Carteiras" ACTION recalall()
    EndIf
    ENDMENU
    MENUITEM "&Relatórios"
    MENU
    MENUITEM "Demonstrativo &Carteira" ACTION IMPCART(0)
    MENUITEM "&Movimento da Carteira" ACTION IMPMOV(1,0)
    MENUITEM "&Demonstrativo Contabil" ACTION DCONT(0)
    If COD_OPER>990
      MENUITEM "&Resumo Apuração I.R." ACTION IMPRES(1)
    EndIf
    ENDMENU
    MENUITEM "&Integração"
    MENU
    MENUITEM "Int.Sinacor" ACTION INTSINA(0)
    ENDMENU
    MENUITEM "Manutenção"
    MENU
    MENUITEM "&Usuários" ACTION CADUSER()
    MENUITEM "&Arquivos" ACTION IIF( MsgYesno( "Manutenção dos Arquivos ?","Manutenção" ),INDICESC(1),)
    MENUITEM "Arqv.&Carteira" ACTION IIF( MsgYesno( "Manutenção dos Arquivos ?","Manutenção" ),INDI1(),)
    If COD_OPER>990
      MENUITEM "Marca todos para Recalcular" ACTION IIF( MsgYesno( "Marcar para recalcular todos ?","Manutenção" ),INDI21(1),)
      SEPARATOR
    EndIf
    MENUITEM "&Impressora" ACTION (PRINTERSETUP(),oitem2:settext(PRNGETNAME()),oitem2:refresh())
    SEPARATOR
    MENUITEM "Troca Investidor" ACTION (MMMM(2),CRIARQC(),INDICESC())
    If A02NOME="MASTER - Exatus"
      SEPARATOR
      MENUITEM "Muda para Worker" ACTION (MUDAQTIPO("W"),oWnd:End())
    EndIf
    ENDMENU
    MENUITEM "&Sair" MESSAGE "Finaliza o Sistema" ACTION (oWnd:End(), SNDPLAYSOUND("GOODBYE.WAV",0)) ACCELERATOR ACC_ALT,ASC("S")
  EndIf
  If AUTOMATICO
    MENUITEM "&Sair" MESSAGE "Finaliza o Sistema" ACTION SAIR()
  EndIf
  ENDMENU
return ( oMenu )

Procedure sair()
  If OP7#"S"
    MSGINFO("Estou Trabalhando ",WSIST)

    //  MSGRUN("Estou trabalhando, aguarde fechamento",wsist,{ ||FECHAROBO()})
  EndIf
  oWnd:End()
  SNDPLAYSOUND("GOODBYE.WAV",0)
return 

Procedure FECHAROBO()
  DO WHILE OP7#"S"
  ENDDO
return 

Procedure RECALALL()
  CLOSE DATABASE
  SELE 1
  If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"EMP")
    MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
    CLOSE DATABASE
    return 
  EndIf
  ARQ=LOCALCART+"EMPRESA1.IND"
  SET INDEX TO &ARQ
  SET ORDER TO 1
  SEEK COD1
  If EOF()
    CLOSE DATABASE
    return 
  EndIf
  If A01CODNAC<>"S"
    CLOSE DATABASE
    MSGINFO("Não é Corretora, utilize a opção Manutenção->Troca Investidor, e escolha o codigo da corretora !",WSIST)
    return 
  EndIf
  CLOSE DATABASE
  INDI21(2)
  LOCALQTIPO:="W"
  ABREARQS()
  CARTEIRANET()
  VERMC(16)
  LOCALQTIPO:="C"
  MSGINFO("Calculos Efetuados !",WSIST)
  CLOSE DATABASE
return 

Procedure MUDAQTIPO(QUAL)
  LOCALQTIPO=QUAL
  SAVE ALL LIKE LOCAL* TO ARQPAR.DAT
return 

Function ADIREG
  PARA TEMPO
  PRIVATE SEMPRE
  APPEND BLANK
  If .NOT. NETERR()
    return  .T.
  EndIf
  M->SEMPRE=(M->TEMPO=0)
  DO WHILE (M->SEMPRE .OR. M->TEMPO>0)
    SysRefresh()
    APPEND BLANK
    If .NOT. NETERR()
      return  .T.
    EndIf
    M->TEC=INKEY(1)
    M->TEMPO=M->TEMPO - 0.5
    If M->TEC=27
      EXIT
    EndIf
  ENDDO
return  .F.

Function USEREDE
  PARA ARQU,EXUSE,TEMPO,ARQ2,QZ
  PRIVATE SEMPRE
  DEFAULT QZ:="N"
  ARQU=UPPER(TRIM(ARQU))
  If AT("NOTAEX",ARQU)#0 .OR. QZ="S"
    AB=.T.
  Else
    AB=.F.
  EndIf
  If (ARQ2=Nil)
    ARQ2:=ARQU
  EndIf
  If AT(".",ARQ2)>0
    ARQ2=LEFT(ARQ2,5)
  EndIf
  If AT(":",ARQ2)>0
    ARQ2="ARQS"
  EndIf
  M->SEMPRE = (M->TEMPO=0)
  If AT(".",ARQU)>0
    ARQIND=TRIM(LEFT(ARQU,AT(".",ARQU)-1))+".KEY"
  Else
    ARQIND=TRIM(ARQU)+".KEY"
  EndIf
  If QZ="M"
    USE &arqu ALIAS &arq2 VIA "mysql"
    return  .T.
  EndIf
  DO WHILE (M->SEMPRE .OR. M->TEMPO>0)
    SysRefresh()
    If AB
      If EXUSE
        USE &ARQU EXCLUSIVE ALIAS (ARQ2) VIA "DBFNTX"
      Else
        USE &ARQU ALIAS (ARQ2) VIA "DBFNTX"
      EndIf
    Else
      If EXUSE
        USE &ARQU EXCLUSIVE ALIAS (ARQ2)
      Else
        USE &ARQU ALIAS (ARQ2)
      EndIf
    EndIf
    If .NOT. NETERR()
      If FILE(ARQIND)
        If AT("NOTAEX",ARQU)#0
          ARQIN1=LOCNOT3+"NOTAEX1.KEY"
          SET INDEX TO &ARQIND,&ARQIN1
        Else
          SET INDEX TO &ARQIND
        EndIf
        SET ORDER TO 1
      EndIf
      return  .T.
    EndIf
    M->TEC=INKEY(1)
    M->TEMPO=M->TEMPO-1
    If M->TEC=27
      EXIT
    EndIf
  ENDDO
return  .F.

Function REGLOCK
  PARA TEMPO
  PRIVATE SEMPRE
  If RLOCK()
    return  .T.
  EndIf
  M->SEMPRE=(M->TEMPO=0)
  while (M->SEMPRE .OR. M->TEMPO>0)
    SysRefresh()
    If RLOCK()
      return  .T.
    EndIf
    M->TEC=INKEY(1)
    M->TEMPO=M->TEMPO - 0.5
    If M->TEC=27
      EXIT
    EndIf
  ENDDO
return  .F.

Procedure VERDIRETORIO(DIRET)
  M->LOCAL2=curdrive()+TRIM(":/"+CURDIR())
  M->LOCAL1=ALLTRIM(DIRET)
  M->LOCAL1=LEFT(M->LOCAL1,LEN(M->LOCAL1)-1)
  If !lCHDIR(M->LOCAL1)
    If !lMKDIR(M->LOCAL1)
      LCHDIR(M->LOCAL2)
      NAOCRIOU=M->LOCAL1
      return  .F.
    EndIf
  EndIf
  LCHDIR(M->LOCAL2)
return  .T.

Function PUTREC()
  If !REGLOCK()
    MSGINFO("Registro sendo utilizado por outro Usuário","Tente Novamente")
    return  .T.
  EndIf
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    If VALTYPE(&_var5)="C"
      REPLACE &_var5 WITH TRIM(&_var4)
    Else
      REPLACE &_var5 WITH &_var4
    EndIf
  NEXT t
  UNLOCK
return (.T.)

Procedure CADLOG()
  If !USUARIO(6)
    return 
  EndIf
  SELE 21
  If !USEREDE(LOCNOT3+"LOG.DBF",.F.,10,"LOG1")
    return 
  EndIf
  Browse(1,"Log. de Usuários","Log. de Usuários")
  SELE 21
  USE
return 

Procedure CADUSER()
  SELE 21
  If !USEREDE(LOCALCART+"USER.DBF",.F.,10,"USUARIO1")
    return 
  EndIf
  Browse(2,"Cadastro de Usuários","Usuários", { || buscauser() },{ ||NCADuser("I") },{ ||NCADuser("A") })
  SELE 21
  USE
return 

Function BUSCAUSER()
  TEXTO="Código"
  VARIAVEL=0
  MASCARA="999"
  REG=RECNO()
  DO WHILE .T.
    SysRefresh()
    If !PROCURA()
      EXIT
    EndIf
    CHAVE=VARIAVEL
    If EMPTY(CHAVE)
      EXIT
    EndIf
    SEEK chave
    If EOF() .OR. DELE()
      MSGALERT("Não Encontrado !","Alerta")
      GO REG
    Else
      EXIT
    EndIf
  ENDDO
return 

Procedure NCADuser(FS)
  Local oDlg,oBrush
  PUBLIC F
  F=FS
  DECLARE _VAR1[FCOUNT()]
  AFIELDS(_VAR1)
  If F="I"
    GO BOTT
    SKIP
  EndIf
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  RELEASE _VAR1
  TROCOP(1)
  MFONTES()
  DEFINE DIALOG oDlg FROM 0,0 TO 26,57 TITLE "Dados do Usuário" BRUSH oBrushTela STYLE WS_CAPTION
  @   0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 219,165
  @   0.87,02 SAY "Código" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   1.74,02 SAY "Nome" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @   2.61,02 SAY "Senha" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  A="N"
  If F="I"
    @ 01,10 GET ZODIGO PICTURE "999" VALID (VSTCON(ZODIGO)) .And. ZODIGO<900 FONT oFontget SIZE 15,11
  Else
    @ 01,10 GET ZODIGO PICTURE "999" WHEN A="S" FONT oFontget SIZE 15,11
  EndIf
  @ 2,10 GET ZOME PICTURE "@!" FONT oFontget SIZE 139,11
  @ 3,10 GET ZEN OF oDlg PASSWORD FONT ofontget SIZE 20,11
  If LOCALQTIPO="W" .And. COD_OPER=999
    @ 04,1  CHECKBOX ZC1  PROMPT "Senhas-Clientes" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 05,1  CHECKBOX ZC2  PROMPT "Usuarios-WORKER" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 06,1  CHECKBOX ZC3  PROMPT "Listas Restritivas" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 07,1  CHECKBOX ZC4  PROMPT "Usuarios-Boletins" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 08,1  CHECKBOX ZC5  PROMPT "Manut.Arquivos" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 09,1  CHECKBOX ZC6  PROMPT "Log de Utilizacao" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 10,1  CHECKBOX ZC7  PROMPT "Gerar Noticias/Cotacoes" size 110,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 11,1  CHECKBOX ZC8  PROMPT "Limpeza nos Arquivos" size 106,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 05,12 CHECKBOX ZC11 PROMPT "Ativos" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 06,12 CHECKBOX ZC12 PROMPT "Usuarios x Ativos" SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 07,12 CHECKBOX ZC13 PROMPT "Comercial" SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 08,12 CHECKBOX ZC14 PROMPT "Usuarios BolsasNet" SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 09,12 CHECKBOX ZC15 PROMPT "Empresas BolsasNet/CambioNet" SIZE 135,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)

    @ 04,20  CHECKBOX ZC17  PROMPT "Cad.Clientes/Grupos" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 05,20  CHECKBOX ZC18  PROMPT "Imprimir relatorios" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 06,20  CHECKBOX ZC19  PROMPT "Cad.Usuários" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 07,20  CHECKBOX ZC20  PROMPT "Gerar Relatórios" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 08,20  CHECKBOX ZC21  PROMPT "Parametros" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)

  ElseIf LOCALQTIPO="W" .And. COD_OPER<999
    @ 04,2  CHECKBOX ZC17  PROMPT "Cad.Clientes" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 05,2  CHECKBOX ZC18  PROMPT "Imprimir Relatórios" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 06,2  CHECKBOX ZC19  PROMPT "Cad.Usuários" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 07,2  CHECKBOX ZC20  PROMPT "Gerar Relatórios" size 120,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
    @ 08,2  CHECKBOX ZC21  PROMPT "Parametros" size 100,11  of oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
  EndIf
  @ 178,065 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (TROCOP(2),PUTREC(),oDlg:end())
  @ 178,125 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
  oDlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
  COMMIT
return 

Function TROCOP(X)
  If X=2
    If F="I"
      ADIREG(0)
    EndIf
  EndIf
  If X=1
    FOR I=1 TO 30
      SysRefresh()
      _var4="ZC"+ALLTRIM(STR(I))
      If &_VAR4="S"
        &_VAR4=.T.
      Else
        &_VAR4=.F.
      EndIf
    NEXT
  ElseIf X=2
    FOR I=1 TO 30
      SysRefresh()
      _var4="ZC"+ALLTRIM(STR(I))
      If &_VAR4=.T.
        &_VAR4="S"
      Else
        &_VAR4="N"
      EndIf
    NEXT
  EndIf
return 

Function Browse(aq, cTitle, cListName, bSearch, bNew, bModify, bDelete, bList, bedit )
  Local oDlg, oFont, oBrush
  Local btnNew, btnModify, btnDelete, btnSearch, btnList, btnEnd
  Local n,BORDEM
  Local CONTINUA:=.T.
  PUBLIC AQS
  PUBLIC oLbx
  AQS=AQ

  MFONTES()
  ORDLIST=0

  DEFAULT cTitle  := "Browse", cListName := "Registros",;
    bDelete := { || RecDelete( oLbx ),oLbx:Setfocus() },;
    bList   := { || wReport(aq, oLbx,clistname,oLbx:Setfocus() )},BEDIT:=.F.

  DEFINE DIALOG oDlg FROM 1, 5 TO 27, 79 TITLE cTitle FONT oFontTit BRUSH oBrushTela STYLE WS_CAPTION

  FOR I=1 TO 15
    SysRefresh()
    If !BOF()
      SKIP -1
    EndIf
  NEXT I

  If AQ=1
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(LOG1->DATA,"@D"),LOG1->HORA,LOG1->USUARIO,LOG1->ROTN HEADERS "Data","Hora","Usuário","Rotina Utilizada" SIZE 275,137 OF oDlg
  ElseIf AQ=2
    @ 1,1 LISTBOX oLbx FIELDS TRANSFORM(USUARIO1->CODIGO,"9999"),USUARIO1->NOME HEADERS "Cod","Nome" SIZE 275, 137  OF oDlg
  ElseIf AQ=3
    @ 1,1 LISTBOX oLbx FIELDS STRZERO(SENHAS1->CODIGO,4),SENHAS1->CLIENTE,SENHAS1->LOGIN,SENHAS1->NOTICIAS,SENHAS1->COTACOES HEADERS "Codigo","Cliente","Login","Noticias","Cotacoes" SIZE 275,137 OF oDlg
  ElseIf AQ=4
    @ 1,1 LISTBOX oLbx FIELDS TROCA1->PARA,TROCA1->DE HEADERS "Lista","Endereço" SIZE 275,137 OF oDlg
  ElseIf AQ=5
    @ 1,1 LISTBOX oLbx FIELDS PESQ1->SMTP,PESQ1->NOME,PESQ1->EMAIL,PESQ1->PALAVRA1 HEADERS "Cliente","Usuario","Email","Chave 01" SIZE 275,137 OF oDlg
  ElseIf AQ=7
    @ 1,1 LISTBOX oLbx FIELDS STRZERO(AGENCIA1->CODIGO),OEMTOANSI(AGENCIA1->DESC) HEADERS "Codigo","Descr." SIZE 275,137 OF oDlg
  ElseIf AQ=8
    @ 1,1 LISTBOX oLbx FIELDS ATIVOS1->CODIGO,ATIVOS1->GRAFICO,OEMTOANSI(ATIVOS1->DESC),TRANSFORM(ATIVOS1->DTULTNEG,"@D") HEADERS "Codigo","Grafico","Descr.","Data Ult." SIZE 275,137 OF oDlg
  ElseIf AQ=9
    @ 1,1 LISTBOX oLbx FIELDS STRZERO(NOTAEX1->IDNOT,10),STRZERO(NOTAEX1->AGENCIA,5),STRZERO(NOTAEX1->ASSUNTO,5),TRANSFORM(NOTAEX1->DTNOT,"@D"),NOTAEX1->HRNOT,(NOTAEX1->MANCHETE) HEADERS "ID","Agencia","Assunto","Data","Hora","Manchete" SIZE 275,137 OF oDlg
  ElseIf AQ=10
    @ 1,1 LISTBOX oLbx FIELDS TAXAS1->CODATI,OEMTOANSI(PESQ(TAXAS1->CODATI,"10",1,"DESC")),TAXAS1->HORAULTNEG,TRANSFORM(TAXAS1->DTULTNEG,"@D"),TRANSFORM(TAXAS1->PROFCOMPRA,"@E 999,999.99999999"),TRANSFORM(TAXAS1->PROFVENDA,"@E 999,999.99999999") HEADERS "Codigo","Nome","Hora","Data","Compra","Venda" SIZE 275,137 OF oDlg
  ElseIf AQ=11
    @ 1,1 LISTBOX oLbx FIELDS STRZERO(SENHAS1->CODIGO,4),SENHAS1->CLIENTE HEADERS "Codigo","Cliente" SIZE 275,137 OF oDlg
  ElseIf AQ=12
    filtro=1
    @ 0.5,12.5 COMBOBOX oCbx VAR Filtro items {"Todos","Bolsa","Cambio","BMF"} OF oDlg FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 70,60 ON CHANGE (FILTRO(FILTRO),oLbx:Refresh(),oLbx:SetFocus())
    GO TOP
    @ 2,1 LISTBOX oLbx FIELDS TRANSFORM(CLIEUS->ULTCONT,"@D"),LEFT(CLIEUS->NOME,40),CLIEUS->TELEFONE HEADERS "Ult.Cont","Nome","Telefone" SIZE 279, 127  OF oDlg
    oLbx:bLDblClick = { | nRow, nCol | (MWCLI:=CLIEUS->CODIGO,OBSERVA(STRZERO(CLIEUS->CODIGO,6)+"000000000-","Anotações "+CLIEUS->NOME)) }
    oLbx:NclrText = { || SelCor(CLIEUS->INATIVO)}
  ElseIf AQ=13 .OR. AQ=14 .OR. AQ=15 .OR. AQ=16 .OR. AQ=17 .OR. AQ=18
    @ 1,1 LISTBOX oLbx FIELDS SIZE 279, 137  OF oDlg

  EndIf
  oLbx:nClrPane:={|| IIF((oLbx:cAlias)->(ordlist())%2==0,nRGB(240,255,255),nRGB(230,230,230))}
  @ 156,010 SBUTTON PIXELS btnNew RESOURCE "INCLUI" COLOR 0,nRGB(255,255,255) PROMPT "&Inclui" FONT oFontbut  TEXT ON_BOTTOM SIZE 35,35 OF oDlg
  @ 156,050 SBUTTON PIXELS TEXT ON_BOTTOM FONT oFontbut COLOR 0,nRGB(255,255,255) PROMPT "&Altera" btnModify  RESOURCE "ALTERA"  SIZE 35,35 OF oDlg
  @ 156,090 SBUTTON PIXELS TEXT ON_BOTTOM FONT oFontbut COLOR 0,nRGB(255,255,255) PROMPT "&Exclui" btnDelete  RESOURCE "EXCLUI"  SIZE 35,35 OF oDlg
  @ 156,130 SBUTTON PIXELS TEXT ON_BOTTOM FONT oFontbut COLOR 0,nRGB(255,255,255) PROMPT "&Procura" btnSearch RESOURCE "PROC"    SIZE 35,35 OF oDlg
  @ 156,170 SBUTTON PIXELS TEXT ON_BOTTOM FONT oFontbut COLOR 0,nRGB(255,255,255) PROMPT "I&mprimir" btnList  RESOURCE "IMPRIME" SIZE 35,35 OF oDlg
  btnlist:ctooltip="Imprime dados"
  @ 156,210 SBUTTON PIXELS TEXT ON_BOTTOM FONT oFontbut COLOR 0,nRGB(255,255,255) PROMPT "&Sair" btnEnd       RESOURCE "SAIR2"   SIZE 35,35 OF oDlg
  btnEnd:ctooltip="Finaliza e volta para o menu principal"

  If AQ=12
    MWCLI=0
    @ 156,250 SBUTTON PIXELS TOOLTIP "Anotações" PROMPT "&Notas" TEXT ON_BOTTOM COLOR 0,nRGB(255,255,255) FONT oFontbut RESOURCE "CHAMADA" OF oDlg SIZE 35,35 ACTION (MWCLI:=CLIEUS->CODIGO,OBSERVA(STRZERO(CLIEUS->CODIGO,6)+"000000000-","Anotações "+CLIEUS->NOME),oLbx:Refresh(),oLbx:Setfocus())
  EndIf


  btnNew:bAction    = { || Eval( bNew ), oLbx:UpStable(),oLbx:Refresh(), oLbx:SetFocus() }

  btnModify:bAction = If( bModify != nil,;
    { || Eval( bModify ),;
    oLbx:UpStable(),oLbx:Refresh(), oLbx:SetFocus() },)

  btnDelete:bAction = If( bDelete != nil,{ || Eval( bDelete ),oLbx:UpStable(),oLbx:Refresh(), oLbx:SetFocus() },)

  btnSearch:bAction = If( bSearch != nil,;
    { || Eval( bSearch ),;
    oLbx:UpStable(),oLbx:Refresh(), oLbx:SetFocus() },)
  btnList:bAction   = { || Eval( bList ),oLbx:UpStable(), oLbx:Refresh() }
  btnEnd:bAction    = { || oDlg:End() }


  If AQ=1
    btnNew:Disable()
    btnModify:Disable()
    btnDelete:Disable()
    btnSearch:Disable()
    btnList:Disable()
  EndIf
  If AQ=9 .or. AQ=10
    btnNew:Disable()
    btnList:Disable()
  EndIf
  If AQ=11
    btnNew:Disable()
    btnList:Disable()
    btnDelete:Disable()
  EndIf
  odlg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlg CENTERED
  COMMIT
return  nil

Procedure SelCoR(tpz)
  Local nColor := CLR_BLACK
  If TPZ
    ncolor = CLR_RED
  EndIf
return  nColor

Procedure FILTRO(X)
  If X=1
    set filter to
  ElseIf X=2
    SET FILTER TO EBOLSA
  ElseIf X=3
    SET FILTER TO ECAMBIO
  ElseIf X=4
    SET FILTER TO EBMF
  EndIf
  GO TOP
return 

Function WREPORT( AQ,oLbx,ctitle)
  Local n
  Local cAlias := If( oLbx != nil, oLbx:cAlias, Alias() )
  Local oFont1, oFont2, oFont3
  PUBLIC oRpt
  DEFINE FONT oFont1 NAME "ARIAL" SIZE 0,-10
  DEFINE FONT oFont2 NAME "ARIAL" SIZE 0,-7
  DEFINE FONT oFont3 NAME "ARIAL" SIZE 0,-6
  reg=recno()

  If AQ=12
    MWSREG=RECNO()
    RELATCOM()
    SELE 14
    SET ORDER TO 1
    SELE 13
    GO MWSREG
    return 
  EndIf

  If EOF()
    MSGINFO("Não Existe Registros",WSIST)
    return 
  EndIf
  GO TOP
  If PREV

    REPORT oRpt TITLE " "," "," "," ",EMP3,CTitle FONT oFont1,oFont2,oFont3;
      PREVIEW;
      FOOTER " ","Emitido em "+DTOC(DAT_HOJE)+"                  Folha: "+strzero(oRpt:nPage,3)," ",'"Documento interno de Propriedade da empresa."','"Proibida sua reprodução, difusão e conservação (apropriação)."' ;
      CAPTION ctitle;

    Else

    REPORT oRpt TITLE " "," "," "," ",EMP3,CTitle FONT oFont1,oFont2,oFont3;
      FOOTER " ","Emitido em "+DTOC(DAT_HOJE)+"                  Folha: "+strzero(oRpt:nPage,3)," ",'"Documento interno de Propriedade da empresa."','"Proibida sua reprodução, difusão e conservação (apropriação)."' ;
      CAPTION ctitle;

    EndIf

    If AQ=2
      COLUMN TITLE "Cod." DATA STRZERO(USUARIO1->CODIGO,4)
      COLUMN TITLE "Nome" DATA USUARIO1->NOME
    ElseIf AQ=3
      COLUMN TITLE "Cliente" DATA SENHAS1->CLIENTE
      COLUMN TITLE "Login" DATA SENHAS1->LOGIN
      COLUMN TITLE "Senha" DATA SENHAS1->SENHA
      COLUMN TITLE "NOT" DATA SENHAS1->NOTICIAS
      COLUMN TITLE "COT" DATA SENHAS1->COTACOES
    ElseIf AQ=4
      COLUMN TITLE "Descrição" DATA TROCA1->PARA FONT 3
      COLUMN TITLE "Endereço" DATA TROCA1->DE FONT 3
    ElseIf AQ=5
      COLUMN TITLE "SMPT-Cliente" DATA PESQ1->SMTP FONT 3
      COLUMN TITLE "Usuario" DATA PESQ1->NOME
      COLUMN TITLE "Email" DATA PESQ1->EMAIL
      COLUMN TITLE "Chave 01" DATA PESQ1->PALAVRA1
    ElseIf AQ=8
      COLUMN TITLE "CODIGO" DATA ATIVOS1->CODIGO
      COLUMN TITLE "Descr." DATA OEMTOANSI(ATIVOS1->DESC)
    ElseIf AQ=9
      COLUMN TITLE "Data","Follow-up" DATA DTOC(WOBS->DATA),DTOC(WOBS->AVISAREM) FONT 2 GRID
      COLUMN TITLE "Resumo da Descrição"," " DATA WOBS->HISTO,LEFT(WOBS->ANOTA,100) FONT 2 GRID
    ElseIf AQ=13
      COLUMN TITLE "Usuário" DATA USUARIOS->USUARIO FONT 2 GRID
      COLUMN TITLE "Senha" DATA LEFT(USUARIOS->SENHA,25) FONT 2 GRID
      COLUMN TITLE "Frase" DATA LEFT(USUARIOS->FRASE,50) FONT2 GRID
    ElseIf AQ=14
      If QTP=1
        COLUMN TITLE "Corretora" DATA STRZERO(EMPVAL->CORRETORA,3) FONT 2 GRID
        COLUMN TITLE "Nome" DATA EMPVAL->NOMECORR FONT 2 GRID
      ElseIf QTP=2
        COLUMN TITLE "Corretora" DATA STRZERO(EMPWINCC->CORRETORA,3) FONT 2 GRID
        COLUMN TITLE "Nome" DATA EMPWINCC->NOMECORR FONT 2 GRID
      ElseIf QTP=3
        COLUMN TITLE "Corretora" DATA BOLSANET->CORRETORA FONT 2 GRID
        COLUMN TITLE "Pasta" DATA BOLSANET->PASTA FONT 2 GRID
      EndIf
    ElseIf AQ=15
      COLUMN TITLE "Usuário" DATA USUAWIN->USUARIO FONT 2 GRID
      COLUMN TITLE "Senha" DATA LEFT(USUAWIN->SENHA,25) FONT 2 GRID
      COLUMN TITLE "Frase" DATA LEFT(USUAWIN->FRASE,50) FONT2 GRID
    EndIf

    oRpt:CELLVIEW()

    ENDREPORT

    ACTIVATE REPORT oRpt

    oFont1:End()
    oFont2:End()
    GO reg
    oFont1:End()
    oFont2:End()
  return  nil
  //----------------------------------------------------------------------------//
  Function PORQUEEXCLUI()
    Local oPar1,oBrush51
    If LOCALQTIPO="C"
      return  .T.
    EndIf
    varia=space(40)
    EXCLUIR=.F.
    MFONTES()
    DEFINE DIALOG oPar1 FROM 1, 1 TO 13.5 , 55 TITLE WSIST BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,7.5 say "" box raised size 135,50 OF oPar1
    @   0.87,09.5 say "Motivo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,11
    @ 3.0,09.5 get VARIA picture "@!" size 75,10 OF oPar1
    @ 0.7,0.5 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER size 46,65 OF oPar1
    @ 70,60 SBUTTON oBtn PIXELS PROMPT "&Confirma"  TEXT ON_CENTER NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST FONT oFontbut OF oPar1 size 40,12 action (oPar1:end(),APAGAMESMO()) DEFAULT
    @ 70,120 SBUTTON oBtn PIXELS PROMPT "C&ancela"  TEXT ON_CENTER NOBOX RESOURCE "OFF","ON","DISABLE" ADJUST FONT oFontbut OF oPar1 size 40,12 action (oPar1:end()) DEFAULT
    oPar1:lhelpicon:=.f.
    ACTIVATE DIALOG oPar1 CENTERED
  return (EXCLUIR)

  Procedure APAGAMESMO()
    M->U_AREATUA = ALLTRIM(STR(SELECT(),3))
    If EMPTY(VARIA)
      MSGINFO("Nao pode ser excluido sem motivo !",WSIST)
      return 
    EndIf
    EXCLUIR=.T.
    DECLARE _var1 [FCOUNT()],_VAR2[FCOUNT()],_VAR3[FCOUNT()]
    aFIELDS(_VAR1,_VAR2,_VAR3)
    TX=""
    FOR I=1 TO LEN(_VAR1)
      SysRefresh()
      a=_var1[i]
      A=&A
      If _VAR2[I]="N"
        A=STR(A)
      ElseIf _VAR2[I]="M"
        A=" "
      ElseIf _VAR2[I]="D"
        A=DTOC(A)
      ElseIf _VAR2[I]="L"
        A=" "
      EndIf
      TX=TX+","+A
    NEXT I
    DECLARE TX1[1]
    TX1[1]="USUARIO: "+A02NOME+" Apagou: "+TX
    LOGFILE(LOCNOT3+"LOGEXC.TXT",TX1)
    SELE 11
    PORQUE=VARIA
    If ADIREG(0)
      REPLACE DATA WITH DATE()
      REPLACE HORA WITH TIME()
      REPLACE USUARIO WITH A02NOME
      REPLACE ROTN WITH PORQUE
      UNLOCK
    EndIf
    SELE &U_AREATUA
  return 

  Function RecDelete( oLbx )
    If MSGYesNo( "Exclui este Registro ?", "Wincam")
      If !PORQUEEXCLUI()
        return 
      EndIf
      If REGLOCK(10)
        DELETE
      Else
        MSGInfo("Registro Bloqueado por outro Usuário","Wincam")
      EndIf
      oLbx:UpStable()
      oLbx:Refresh()
    EndIf
  return  nil
  //----------------------------------------------------------------------------//
  Procedure QUEM()
    PUBLIC oDlg
    PUBLIC cPassword := "    "
    PUBLIC nTries    := 0
    PUBLIC COD_OPER:=0
    SELE 12
    MFONTES()
    DEFINE FONT oFontgtela  NAME "Verdana" SIZE 6,13 BOLD
    DEFINE FONT oFonttela  NAME "Verdana" SIZE 6,13 BOLD
    DEFINE FONT oFontbat  NAME "Verdana" SIZE 6,13
    DEFINE BITMAP oBmpl RESOURCE "lgn"
    DEFINE DIALOG oDlg FROM 1,1 TO 16,51 TITLE "LOGIN DO SISTEMA"
    @ 2.61,8 SAY "Usuário"  FONT oFonttela COLOR nRGB(100,100,100),nRGB(255,255,255) SIZE 25,10
    @ 3.48,8 SAY "Senha"    FONT oFonttela COLOR nRGB(100,100,100),nRGB(255,255,255) SIZE 25,10
    @ 4.35,8 SAY "Data"     FONT oFonttela COLOR nRGB(100,100,100),nRGB(255,255,255) SIZE 25,10
    @ 3,11 GET ocod var cod_oper picture "999" FONT oFontgtela size 45,11 VALID COD_OPER>0 RIGHT
    @ 4,11 GET cPassword OF oDlg PASSWORD FONT oFontgtela size 45,11
    @ 5,11 GET M->dat_hoje FONT oFontgtela size 45,11 VALID M->DAT_HOJE>DATE()-1
    @ 93,45  SBUTTON PIXELS RESOURCE "OFF","ON","DISABLE" NOBOX  TITLE "&Entrar"  TEXT ON_CENTER ADJUST FONT oFontbat OF oDlg SIZE 40,12 ACTION (VERS())
    @ 93,115 SBUTTON PIXELS RESOURCE "OFF","ON","DISABLE" NOBOX  TITLE "&Sair"    TEXT ON_CENTER ADJUST FONT oFontbat OF oDlg SIZE 40,12 ACTION (oDlg:End())
    oDlg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlg CENTERED   ON INIT oCod:Setfocus() ON PAINT PalBmpDraw( oDlg:hDC,(oDlg:nHeight()/2)-(oBmpl:nHeight()/2)-10,(oDlg:nWidth() /2)-(oBmpl:nWidth() /2),oBmpl:hBitmap, oBmpl:hPalette)
    SELE 12
    USE
  return  continua

  Function VERS()
    ntries+=1
    If NTRIES=4
      MSGSTOP("Saindo do Sistema !","")
      CLOSE DATABASE
      QUIT
    EndIf
    SELE 12
    SEEK COD_OPER
    If EOF() .And. COD_OPER#999
      MSGINFO("Senha Incorreta !",WSIST)
      return 
    EndIf
    If CPASSWORD=SEN .OR. (COD_OPER=999 .And. CPASSWORD="4372")
      A02NOME=NOME
      OP1=FC1;OP2=FC2;OP3=FC3;OP4=FC4;OP5=FC5;OP6=FC6;OP7=FC7;OP8=FC8;OP9=FC9;OP10=FC10
      OP11=FC11;OP12=FC12;OP13=FC13;OP14=FC14;OP15=FC15;OP16=FC16;OP17=FC17;OP18=FC18;OP19=FC19;OP20=FC20;OP21=FC21;OP22=FC22;OP23=FC23;OP24=FC24;OP25=FC25;OP26=FC26;OP27=FC27;OP28=FC28;OP29=FC29;OP30=FC30
      If COD_OPER=999
        A02NOME="MASTER - Exatus"
        FOR I=1 TO 20
          SysRefresh()
          _var4="OP"+ALLTRIM(STR(I))
          &_var4="S"
        NEXT I
      EndIf
      CONTINUA=.T.
      odlg:end()
      return 
    EndIf
    MSGINFO("Senha Incorreta !",WSIST)
  return 

  Function USUARIO(X)
    OPC="OP"+ALLTRIM(STR(X))
    If X<=50
      If &OPC#"S"
        MSGINFO("Usuário Não Autorizado !!","Rotina")
        return  .F.
      EndIf
    EndIf
    SELE 11
    DECLARE ROTN [50]
    ROTN[1]="Senhas-Clientes"
    ROTN[2]="Usuarios-WORKER"
    ROTN[3]="Listas Restritivas"
    ROTN[4]="Usuarios-Boletins"
    ROTN[5]="Manut.Arquivos"
    ROTN[6]="Log de Utilizacao"
    ROTN[7]="Arquivo de Noticias/Cotacoes"
    ROTN[8]="Limpeza nos Arquivos"
    ROTN[9]="Cad.Assuntos"
    ROTN[10]="Cad.Agencias"
    ROTN[11]="Ativos"
    ROTN[12]="Usuarios x Ativos"
    ROTN[13]="Comercial"
    ROTN[14]="Usuários BolsasNet "
    ROTN[15]="Empresas BolsasNet"
    ROTN[16]="Relatórios Especiais-CambioNet"
    If ADIREG(0)
      REPLACE DATA WITH DATE()
      REPLACE HORA WITH TIME()
      REPLACE USUARIO WITH A02NOME
      REPLACE ROTN WITH ROTN[X]
      UNLOCK
    EndIf
  return  .T.

  Function CRIARQ
    If !FILE(ARQ06)
      mat_stru={}
      Aadd(mat_stru,{"DATA","D",8,0})
      Aadd(mat_stru,{"CODPAPEL","C",12,0})
      Aadd(mat_stru,{"PRECOM","N",19,5})
      AADD(MAT_STRU,{"POP","C",1,0})
      DBCREATE(arq06,mat_stru)
      USE &arq06
      INDEX ON &chv0601 TAG CHAVE1 DESCENDING TO &ind0601
      CLOSE DATABASE
    EndIf
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
      HIST[1]="COMPRA OPÇÕES"
      HIST[2]="COMPRA VISTA"
      HIST[3]="COMPRA TERMO"
      HIST[4]="VENDA OPÇÕES"
      HIST[5]="VENDA VISTA"
      HIST[6]="VENDA TERMO"
      HIST[7]="ENTRADA SUBSCRIÇÃO"
      HIST[8]="ENTRADA OUTROS"
      HIST[9]="SAIDA OUTROS"
      HIST[10]="BONIFICAÇÃO"
      HIST[11]="ENTRADA REF.TRANSFERENCIA"
      HIST[12]="SAIDA TRANSFERENCIA"
      HIST[13]="IMPLIT"
      HIST[14]="LUCRO OPCOES NORMAIS"
      HIST[15]="LUCRO VISTA NORMAIS"
      HIST[16]="LUCRO TERMO NORMAIS"
      HIST[17]="LUCRO OPÇÕES DAY-TRADE"
      HIST[18]="LUCRO VISTA DAY-TRADE"
      HIST[19]="LUCRO TERMO DAY-TRADE"
      HIST[20]="PREJUIZO OPÇÕES NORMAIS"
      HIST[21]="PREJUIZO VISTA NORMAIS"
      HIST[22]="PREJUIZO TERMO NORMAIS"
      HIST[23]="PREJUIZO OPÇÕES DAY-TRADE"
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
        EndIf
      NEXT ZS
      CLOSE DATABASE
    EndIf
    If !FILE(ARQ14)
      // Arquivo de Papeis BDI *
      mat_stru={}
      AADD(MAT_STRU,{"a03codp","C",12,0})  // Codigo do papel
      AADD(MAT_STRU,{"a03desc","C",40,0})  // Descricao
      AADD(MAT_STRU,{"a03tipo","C",01,0})  // Tipo de papel (1-A Vista 2-Opcao 3-A Termo)
      AADD(MAT_STRU,{"a03codt","C",02,0})  // Codigo da tabela
      AADD(MAT_STRU,{"a03lmil","C",01,0})  // Lote de mil (S/N)
      AADD(MAT_STRU,{"a03pm","N",19,2})    // Preco de Mercado
      AADD(MAT_STRU,{"A03DTBD","D",8,0})   // Data Ult.Movimento
      AADD(MAT_STRU,{"A03DTOP","D",8,0})   // Data de Validade da Opção
      AADD(MAT_STRU,{"A03CODPO","C",12,0}) // Código do Papel a Vista
      AADD(MAT_STRU,{"A03POP","C",1,0}) //    S=POP
      DBCREATE(ARQ14,MAT_STRU)
      USE &ARQ14
      INDEX ON &CHV1401 TAG CHAVE1 TO &IND1401
      INDEX ON &CHV1402 TAG CHAVE2 TO &IND1401
      INDEX ON &CHV1403 TAG CHAVE3 TO &IND1401
      CLOSE DATABASE
    EndIf
    ARQ=LOCALCART+"EMPRESAS.DBF"
    If !FILE(ARQ)
      MAT_STRU={}
      AADD(mat_stru,{"a01code","C",06,0})   // Código na Exatus
      AADD(mat_stru,{"a01nome","C",60,0})   // Nome
      AADD(MAT_STRU,{"A01ENDER","C",60,0})  // Endereço
      AADD(MAT_STRU,{"A01CEPCID","C",60,0}) // Cep/Cidade
      AADD(mat_stru,{"a01data","C",04,0})   // Data (utilizar 4 para o ano)
      AADD(mat_stru,{"a01Tipo","N",1,0})    // 1-Fisica 2-Juridica
      AADD(mat_stru,{"a01CGCCPF","C",14,0}) // CGCCPF
      AADD(MAT_Stru,{"a01SENHA","C",20,0})  // Senha de Acesso Internet
      AADD(MAT_STRU,{"a01Frase","C",60,0})  // Frase Secreta Acesso Internet
      AADD(MAT_STRU,{"A01VALID","D",8,0})   // Data Validade do Pagamento
      AADD(MAT_STRU,{"A01INATI","L",1,0})   // Se verdadeiro = Inativo
      AADD(MAT_STRU,{"A01DIAVENC","N",2,0}) // Dia de Vencimento da Fatura
      AADD(MAT_STRU,{"A01EMITIDO","D",8,0}) // Data do Vencimento da Fatura
      AADD(MAT_STRU,{"A01PAGO","D",8,0})    // Data do Pagamento da Fatura
      AADD(MAT_STRU,{"A01EMAIL","C",253,0}) // Email de Recuperação de Senha
      AADD(MAT_STRU,{"A01VALMES","N",19,2}) // Valor da Mensalidade
      AADD(MAT_STRU,{"A01CORRET","N",3,0})  // Se deve atualizar a carteira esta 1
      AADD(MAT_STRU,{"A01CODNAC","C",10,0}) // Se for Administrador terá "S" aqui
      AADD(MAT_STRU,{"A01CONFER","L",1,0})  // Se Verdadeiro foi verificado e pode utilizar as notas
      AADD(MAT_STRU,{"A01CORP","N",3,0}) // Codigo da Corretora Administradora ou da Corretora no caso do campo a01codnac = "S"
      AADD(MAT_STRU,{"A01COD1","N",8,0}) // Codigo do Cliente na Corretora Principal
      AADD(MAT_STRU,{"A01COD1I","N",8,0})// Codigo do Cliente na Corretora Principal Investimento
      AADD(MAT_STRU,{"A01COR2","N",3,0}) // Codigo da Segunda Corretora
      AADD(MAT_STRU,{"A01COD2","N",8,0}) // Codigo do Cliente na Segunda Corretora
      AADD(MAT_STRU,{"A01COD2I","N",8,0})// Codigo do Cliente na Segunda Corretora Investimento
      AADD(MAT_STRU,{"A01COR3","N",3,0}) // Codigo da Terceira Corretora
      AADD(MAT_STRU,{"A01COD3I","N",8,0})// Codigo do Cliente na Terceira Corretora Investimento
      AADD(MAT_STRU,{"A01COD3","N",8,0}) // Codigo do Cliente na Terceira Corretora
      AADD(MAT_STRU,{"A01COR4","N",3,0}) // Codigo da Quarta Corretora
      AADD(MAT_STRU,{"A01COD4","N",8,0}) // Codigo do Cliente na Quarta Corretora
      AADD(MAT_STRU,{"A01COD4I","N",8,0})// Codigo do Cliente na Quarta Corretora Investimento
      AADD(MAT_STRU,{"A01COR5","N",3,0}) // Codigo da Quinta Corretora
      AADD(MAT_STRU,{"A01COD5","N",8,0}) // Codigo do Cliente na Quinta Corretora
      AADD(MAT_STRU,{"A01COD5I","N",8,0})// Codigo do Cliente na Quinta Corretora Investimento
      AADD(MAT_STRU,{"A01CALCATE","D",8,0}) // Calcular até
      DBCREATE(ARQ,MAT_STRU)
    EndIf
    ARQ=LOCALCART+"USER.DBF"
    If !FILE(ARQ)
      mat_stru={}
      AADD(mat_stru,{"CODIGO","N",3,0})
      AADD(mat_stru,{"NOME","C",40,0})
      AADD(mat_stru,{"SEN","C",4,0})
      AADD(mat_stru,{"FC1","C",1,0})
      AADD(mat_stru,{"FC2","C",1,0})
      AADD(mat_stru,{"FC3","C",1,0})
      AADD(mat_stru,{"FC4","C",1,0})
      AADD(mat_stru,{"FC5","C",1,0})
      AADD(mat_stru,{"FC6","C",1,0})
      AADD(mat_stru,{"FC7","C",1,0})
      AADD(mat_stru,{"FC8","C",1,0})
      AADD(mat_stru,{"FC9","C",1,0})
      AADD(mat_stru,{"FC10","C",1,0})
      AADD(mat_stru,{"FC11","C",1,0})
      AADD(mat_stru,{"FC12","C",1,0})
      AADD(mat_stru,{"FC13","C",1,0})
      AADD(mat_stru,{"FC14","C",1,0})
      AADD(mat_stru,{"FC15","C",1,0})
      AADD(mat_stru,{"FC16","C",1,0})
      AADD(mat_stru,{"FC17","C",1,0})
      AADD(mat_stru,{"FC18","C",1,0})
      AADD(mat_stru,{"FC19","C",1,0})
      AADD(mat_stru,{"FC20","C",1,0})
      AADD(mat_stru,{"FC21","C",1,0})
      AADD(mat_stru,{"FC22","C",1,0})
      AADD(mat_stru,{"FC23","C",1,0})
      AADD(mat_stru,{"FC24","C",1,0})
      AADD(mat_stru,{"FC25","C",1,0})
      AADD(mat_stru,{"FC26","C",1,0})
      AADD(mat_stru,{"FC27","C",1,0})
      AADD(mat_stru,{"FC28","C",1,0})
      AADD(mat_stru,{"FC29","C",1,0})
      AADD(mat_stru,{"FC30","C",1,0})
      DBCREATE(ARQ,mat_stru)
    EndIf
    CLOSE DATABASE
    If LOCALQTIPO#"W"
      return 
    EndIf
    ARQ=LOCALWEB+"USUARIO.DBF"
    If .NOT. FILE(ARQ)
      aStructure={}
      aAdd( aStructure, { "USUARIO"   , "C",  50, 0 } )
      aAdd( aStructure, { "SENHA"     , "C",  50, 0 } )
      aAdd( aStructure, { "PASTA"     , "C", 254, 0 } )
      aAdd( aStructure, { "ISADMIN"   , "L",   1, 0 } )
      aAdd( aStructure, { "BLOQUEADO" , "L",   1, 0 } )
      aAdd( aStructure, { "CORRETORA" , "C",  50, 0 } )
      aAdd( aStructure, { "ACESSO"    , "N",  20, 5 } )
      aAdd( aStructure, { "ULTACESSO" , "D",   8, 0 } )
      aAdd( aStructure, { "FRASE"     , "C", 100, 0 } )
      aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
      aAdd( aStructure, { "ENVEXTR"   , "L",   1, 0 } )
      DBCREATE(ARQ,aStructure)
    EndIf
    ARQ=LOCNOT3+"BANCOML.DBF"
    If .NOT. FILE(ARQ)
      MAT_STRU={}
      AADD(MAT_STRU,{"CODIGO","N",  6,  0})
      AADD(MAT_STRU,{"CODCLI","N",6,0})
      AADD(MAT_STRU,{"CLIENTE","C",1,0})
      AADD(MAT_STRU,{"NOME","C", 50,  0})
      AADD(MAT_STRU,{"ENDERECO","C", 50,  0})
      AADD(MAT_STRU,{"CIDADE","C", 20,  0})
      AADD(MAT_STRU,{"ESTADO","C",  2,  0})
      AADD(MAT_STRU,{"CEP","C",  8,  0})
      AADD(MAT_STRU,{"TELEFONE","C", 20,  0})
      AADD(MAT_STRU,{"FAX","C",11,0})
      AADD(MAT_STRU,{"CONTATO","C", 20,  0})
      AADD(MAT_STRU,{"TIPO","C",  1,  0})
      AADD(MAT_STRU,{"CGCCPF","C", 15,  0})
      AADD(MAT_STRU,{"OPERADOR","N",  4,  0})
      AADD(MAT_STRU,{"EMAIL","C",254,0})
      AADD(MAT_STRU,{"EMBAILBM","C",254,0})
      AADD(MAT_STRU,{"EMAILBOV","C",254,0})
      AADD(MAT_STRU,{"EMAILCONT","C",254,0})
      AADD(MAT_STRU,{"ULTCONT","D",8,0})
      AADD(MAT_STRU,{"PROXCONT","D",8,0})
      AADD(MAT_STRU,{"ULTEMAIL","D",8,0})
      AADD(MAT_STRU,{"PROXEMAI","D",8,0})
      AADD(MAT_STRU,{"INATIVO","L",1,0})
      AADD(MAT_STRU,{"EBOLSA","L",1,0})
      AADD(MAT_STRU,{"ECAMBIO","L",1,0})
      AADD(MAT_STRU,{"EBMF","L",1,0})
      AADD(MAT_STRU,{"ECONTAB","L",1,0})
      AADD(MAT_STRU,{"EOUTROS","L",1,0})
      AADD(MAT_STRU,{"ASEC","L",1,0})
      DBCREATE(ARQ,MAT_STRU)
    EndIf
    ARQ=LOCALWEB+"USUARIOS.DBF"
    If .NOT. FILE(ARQ)
      aStructure={}
      aAdd( aStructure, { "USUARIO"   , "C",  10, 0 } )
      aAdd( aStructure, { "SENHA"     , "C",  20, 0 } )
      aAdd( aStructure, { "PASTA"     , "C",  50, 0 } )
      aAdd( aStructure, { "ISADMIN"   , "L",   1, 0 } )
      aAdd( aStructure, { "BLOQUEADO" , "L",   1, 0 } )
      aAdd( aStructure, { "CORRETORA" , "C",  50, 0 } )
      aAdd( aStructure, { "ACESSO"    , "N",   5, 0 } )
      aAdd( aStructure, { "ULTACESSO" , "D",   8, 0 } )
      aAdd( aStructure, { "FRASE"     , "C", 100, 0 } )
      aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
      aAdd( aStructure, { "MBOLSA"    , "C", 100, 0 } )
      aAdd( aStructure, { "MBMF"      , "C", 100, 0 } )
      aAdd( aStructure, { "MCONTAB"   , "C", 100, 0 } )
      aAdd( aStructure, { "CORREIO"   , "L",   1, 0 } )
      aAdd( aStructure, { "ENVMAIL"   , "L",   1, 0 } )
      aAdd( aStructure, { "DADOS"     , "L",   1, 0 } )
      aAdd( aStructure, { "ASSBOV"    , "N",   6, 0 } )
      aAdd( aStructure, { "ASSBMF"    , "N",   6, 0 } )
      aAdd( aStructure, { "ASSCON"    , "N",   6, 0 } )
      DBCREATE(ARQ,aStructure)
    EndIf
    ARQ=LOCALWEB+"EMPWINCC.DBF"
    If .NOT. FILE(ARQ)
      aStructure={}
      aAdd( aStructure, { "Corretora" , "N",   3, 0 } )
      aAdd( aStructure, { "NOMECORR"  , "C", 100, 0 } )
      aAdd( aStructure, { "SMTPCORR"  , "C", 100, 0 } )
      aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
      aAdd( aStructure, { "LOCALCOR"  , "C", 100, 0 } )
      aAdd( aStructure, { "ENVMAIL"   , "L",   1, 0 } )
      aAdd( aStructure, { "CORREIO"   , "L",   1, 0 } )
      aAdd( aStructure, { "DADOS"     , "L",   1, 0 } )
      aAdd( aStructure, { "INSTALA"   , "L",   1, 0 } )
      aAdd( aStructure, { "PASTALOG"  , "C", 150, 0 } )
      aAdd( aStructure, { "WINIMAGE"  , "L",   1, 0 } )
      aAdd( aStructure, { "ULTEXTR"   , "D",   8, 0 } )
      aAdd( aStructure, { "INST"      , "N",   5, 0 } )
      DBCREATE(ARQ,aStructure)
    EndIf
    ARQ=LOCNOT3+"ANOTAC.DBF"
    If .NOT. FILE(ARQ)
      //* NUNCA CRIAR CRIAMOV PARA ESTE ARQUIVO
      MAT_STRU={}
      AADD(MAT_STRU,{"CHAVE","C",30,0})
      AADD(MAT_STRU,{"DATA","D",8,0})
      AADD(MAT_STRU,{"USUARIO","C",30,0})
      AADD(MAT_STRU,{"AVISAREM","D",8,0})
      AADD(MAT_STRU,{"HORA","C",10,0})
      AADD(MAT_STRU,{"HISTO","C",60,0})
      AADD(MAT_STRU,{"ANOTA","M",10,0})
      DBCREATE(ARQ,MAT_STRU)
    EndIf
    ARQ=LOCNOT3+"REGISTRO.DBF"
    If !FILE(ARQ)
      MAT_STRU={}
      AADD(MAT_STRU,{"USUARIO","C",50,0})
      AADD(MAT_STRU,{"SENHA","C",50,0})
      AADD(MAT_STRU,{"DATA","D",8,0})
      DBCREATE(ARQ,MAT_STRU)
    EndIf
    ARQ=LOCNOT3+"TAXAS.DBF"
    If !FILE(ARQ)
      mat_stru={}
      AADD(mat_stru,{"CODATI","C",12,0})
      AADD(mat_stru,{"HORAULTNEG","C",6,0})
      AADD(mat_stru,{"DTULTNEG","D",8,0})
      AADD(mat_stru,{"ULTNEG","N",19,8})
      AADD(mat_stru,{"OSCILA","N",19,8})
      AADD(mat_stru,{"PROFCOMPRA","N",19,8})
      AADD(mat_stru,{"QUOFCOMPRA","N",19,8})
      AADD(mat_stru,{"PROFVENDA","N",19,8})
      AADD(mat_stru,{"QUOFVENDA","N",19,8})
      AADD(mat_stru,{"PRMAXIMO","N",19,8})
      AADD(mat_stru,{"PRMINIMO","N",19,8})
      AADD(mat_stru,{"PRABERTURA","N",19,8})
      AADD(mat_stru,{"PRMEDIO","N",19,8})
      AADD(mat_stru,{"PRFECANT","N",19,8})
      AADD(mat_stru,{"NRNEG","N",19,0})
      AADD(mat_stru,{"QTDULTNEG","N",19,8})
      AADD(mat_stru,{"QTDTOTAL","N",19,8})
      AADD(mat_stru,{"VOLUME","N",19,8})
      AADD(mat_stru,{"VOLUMETOT","N",19,8})
      AADD(mat_stru,{"PARV","N",19,8})
      AADD(mat_stru,{"PARC","N",19,8})
      DBCREATE(ARQ,mat_stru)
    EndIf
    ARQ=LOCNOT3+"PESQ.DBF"
    If !FILE(ARQ)
      mat_stru:={}
      AADD(mat_stru,{"EMAIL","C",254,0})
      AADD(mat_stru,{"DTBOL","D",8,0})
      AADD(mat_stru,{"NOME","C",40,0})
      AADD(mat_stru,{"MOEDA","C",10,0})
      AADD(mat_stru,{"ADMIN","C",1,0})
      AADD(mat_stru,{"MAIORMENOR","C",1,0})
      AADD(mat_stru,{"VALOR","N", 19,8})
      AADD(mat_stru,{"PALAVRA1","C",40,0})
      AADD(mat_stru,{"CORRETORA","C",40,0})
      AADD(mat_stru,{"USUARIO","C",40,0})
      AADD(mat_stru,{"SENHA","C",40,0})
      AADD(mat_stru,{"PALAVRA5","C",40,0})
      AADD(mat_stru,{"CONSULTANR","N",8,0})
      AADD(mat_stru,{"SMTP","C",100,0})
      aadd(mat_stru,{"senhae","C",50,0})
      AADD(mat_stru,{"CONTA","C",100,0})
      DBCREATE(ARQ,mat_stru)
    EndIf
    ARQ=LOCNOT3+"TROCA.DBF"
    If !FILE(ARQ)
      mat_stru={}
      //** NAO FAZER CRIAMOV DESTE ARQUIVO TEM MEMO
      AADD(mat_stru,{"CLIENTE","C",50,0})
      AADD(mat_stru,{"DE","C",100,0}) // Endereçco Lista
      AADD(mat_stru,{"PARA","C",100,0}) // Resumo nome Lista
      AADD(mat_stru,{"ENVIAFTP","C",1,0})
      AADD(mat_stru,{"FTP","C",100,0})
      AADD(mat_stru,{"SENHAFTP","C",50,0})
      AADD(mat_stru,{"BOLETIM","C",1,0})
      AADD(mat_stru,{"SMTP","C",100,0})
      AADD(mat_stru,{"ATUA","N",3,0}) // nr. do layout
      AADD(mat_stru,{"ULTATU","C",10,0})
      AADD(mat_stru,{"DTULT","D",8,0})
      AADD(MAT_STRU,{"RODAPE","M",10,0}) // Lista..
      AADD(MAT_STRU,{"SITE","M",10,0}) // Descrição Lista
      DBCREATE(ARQ,mat_stru)
    EndIf
    ARQ=LOCNOT3+"LOG.DBF"
    If !FILE(ARQ)
      mat_stru={}
      AADD(mat_stru,{"DATA","D",8,0})
      AADD(mat_stru,{"HORA","C",8,0})
      AADD(mat_stru,{"USUARIO","C",20,0})
      AADD(mat_stru,{"ROTN","C",40,0})
      DBCREATE(ARQ,mat_stru)
    EndIf
    ARQTEMP=LOCNOT3+"NOTAEX.DBF"
    ARQTEMPM=LOCNOT3+"NOTAEX.DBT"
    //** NAO FAZER CRIAMOV DESTE ARQUIVO ... TEM MEMO
    If !FILE(ARQTEMP)
      Erase &ARQTEMPM
      mat_stru={}
      AADD(mat_stru,{"AGENCIA","N",5,0})
      AADD(mat_stru,{"DESCAG","C",30,0})
      AADD(mat_stru,{"DESCASS","C",30,0})
      AADD(mat_stru,{"ASSUNTO","N",5,0})
      AADD(mat_stru,{"DTNOT","D",8,0})
      AADD(mat_stru,{"HRNOT","C",8,0})
      AADD(mat_stru,{"PRIORID","N",5,0})
      AADD(mat_stru,{"MANCHETE","C",150,0})
      AADD(mat_stru,{"IDNOT","N",8,0})
      AADD(mat_stru,{"TEXTO","M",10,0})
      AADD(MAT_STRU,{"LINK","C",250,0})
      AADD(MAT_STRU,{"NRDIAS","N",3,0})
      DBCREATE(ARQTEMP,mat_stru,"DBFNTX")
    EndIf
    CLOSE DATABASE
  return 

  Function INDICES()
    If PCOUNT()<>0
      If !USUARIO(5)
        return 
      EndIf
    EndIf
    CLOSE DATABASE
    SELE 1
    If !FILE(IND0601) .OR. PCOUNT()<>0
      Erase &IND0601
      If !USEREDE(ARQ06,.T.,10,"TEMP")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MsgStop("Não foi possível acesso ao arquivo "+ARQ14,WSIST)
        QUIT
      EndIf
      PACK
      INDEX ON &chv0601 TAG CHAVE1 DESCENDING TO &ind0601
      CLOSE DATABASE
    EndIf
    CLOSE DATABASE
    SELE 1
    If .NOT. FILE(IND1401) .OR. PCOUNT()<>0
      Erase &IND1401
      If .NOT. USEREDE(ARQ14,.T.,10,"TEMP")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MsgStop("Não foi possível acesso ao arquivo "+ARQ14,WSIST)
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDICC( oMeter, oText, oDlg, @lEnd,2 )}              , ;
        IND1401,"Indexação")
    EndIf
    CLOSE DATABASE
    ARQ=LOCALCART+"EMPRESA1.IND"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCALCART+"EMPRESAS.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","PLANILHA")
        QUIT
      EndIf
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,14 )}, ;
        "Empresas Carteira", "Manutenção")
    EndIf
    CLOSE DATABASE
    ARQ=LOCALCART+"USER.KEY"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCALCART+"USER.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","OPER.DBF")
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,1 )}, ;
        "Usuários", "Manutenção")
    EndIf
    If LOCALQTIPO#"W"
      return 
    EndIf
    ARQ=LOCNOT3+"NOTAEX.KEY"
    ARQ1S=LOCNOT3+"NOTAEX1.KEY"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCNOT3+"NOTAEX.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","OPER.DBF")
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,5 )}, ;
        "Noticias", "Manutenção")
    EndIf
    ARQ=LOCNOT3+"TAXAS.KEY"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCNOT3+"TAXAS.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","OPER.DBF")
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,7 )}, ;
        "Taxas", "Manutenção")
    EndIf
    CLOSE DATABASE
    ARQ=LOCNOT3+"BANCOML.KEY"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCNOT3+"BANCOML.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","SENHAS")
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,11 )}, ;
        "Clientes", "Manutenção")
    EndIf
    CLOSE DATABASE
    ARQ=LOCNOT3+"ANOTAC.KEY"
    If .NOT. FILE(ARQ) .OR. PCOUNT()<>0
      If .NOT. USEREDE(LOCNOT3+"ANOTAC.DBF",.T.,10,"TIPO01")
        MESSAGEBEEP(-1);MESSAGEBEEP(-1);MESSAGEBEEP(-1)
        MSGStop("Não foi possível acesso ao arquivo","SENHAS")
        QUIT
      EndIf
      PACK
      MsgMeter( { | oMeter, oText, oDlg, lEnd |                    ;
        INDIC( oMeter, oText, oDlg, @lEnd,12 )}, ;
        "Anotações", "Manutenção")
    EndIf
    CLOSE DATABASE
    If PCOUNT()<>0
      ABREARQS()
    EndIf
  return 

  Function INDIC( oMeter, oText, oDlg, lEnd,CHAVE )
    oMeter:nTotal := lastrec()
    If CHAVE=1
      INDEX ON CODIGO TAG CODIGO TO &ARQ EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
    ElseIf CHAVE=5
      INDEX ON DTOS(DTNOT)+HRNOT+STRZERO(AGENCIA,3) TAG DTHRAG TO &ARQ EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
      INDEX ON LINK TAG LINK TO &ARQ1S EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
    ElseIf CHAVE=7
      INDEX ON CODATI+DTOS(DTULTNEG)+HORAULTNEG TAG COD DESCEND TO &ARQ EVAL( oMeter:Set( recno() ), SysRefresh(), !lEnd )
    ElseIf CHAVE=10
      INDEX ON HORAULTNEG TAG TES to &arq eval ( OmETER:sET( RECNO() ), sYSrEFRESH(), !LeND )
    ElseIf CHAVE=11
      INDEX ON CODIGO TAG WCOD TO &ARQ
      INDEX ON NOME TAG WNOME TO &ARQ EVAL ( oMeter:Set( recno() ), SysRefresh(), !lEnd )
    ElseIf CHAVE=12
      INDEX ON CHAVE TAG ANJ TO &ARQ
      INDEX ON DTOS(AVISAREM) TAG AVISO TO &ARQ EVAL( oMeter:Set( recno() ), SysRefresh(), !lEnd )
    ElseIf CHAVE=14
      INDEX ON A01CODE TAG CODIGO TO &IND0101
      INDEX ON VAL(A01CGCCPF) TAG CGCCPF TO &IND0101
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
    ElseIf CHAVE=15
      INDEX ON NRSOLIC+CODASS+GRUPO+STRZERO(CGCCPF,14) TAG CHAVE1 TO &ARQ
      INDEX ON DTOS(DATA)+NRSOLIC+CODASS+GRUPO+STRZERO(CGCCPF,14) TAG CHAVE2 TO &ARQ
    EndIf
  return  (NIL)

  Procedure PROCURA()
    Local oD1lg
    PUBLIC RETORNO
    MFONTES()
    RETORNO=.T.
    DEFINE DIALOG oD1lg FROM 1, 1 TO 13,55 TITLE "Pesquisa" BRUSH oBrushTela STYLE WS_CAPTION
    @ 0.44,9 say "" box raised COLOR CLR_BLACK,nRGB(248,248,248) size 155,60 OF oD1lg
    @ 0.87,10 say Texto FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,11
    @ 3.0,7.4 get VARIAVEL picture mascara  size 145,10 OF oD1lg
    @ 0.5,0.5 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER size 46,65 OF oD1lg
    @ 73,70  SBUTTON PIXELS PROMPT "&Confirma" NOBOX RESOURCE "OFF","ON","DISABLE"  ADJUST TEXT ON_CENTER OF oD1lg SIZE 40,12 ACTION oD1lg:end() FONT oFontbut
    @ 73,120 SBUTTON PIXELS PROMPT "C&ancela" NOBOX RESOURCE "OFF","ON","DISABLE"  ADJUST TEXT ON_CENTER OF oD1lg SIZE 40,12 ACTION (VERM(5),oD1lg:End()) FONT oFontbut
    od1lg:lhelpicon:=.f.
    ACTIVATE DIALOG oD1lg CENTERED
  return  RETORNO

  Function VERM(X)
    MFONTES()
    If X=1
      If ZOTIC
        ZOTICIAS="S"
      Else
        ZOTICIAS="N"
      EndIf
      If ZOTAC
        ZOTACOES="S"
      Else
        ZOTACOES="N"
      EndIf
    EndIf
    If X=2
      If !EMPTY(ZTP)
        ZNVIAFTP="S"
      Else
        ZNVIAFTP="N"
      EndIf
      If ZOLET
        ZOLETIM="S"
      Else
        ZOLETIM="N"
      EndIf
    EndIf
    If X=3
      If ZOLET
        MSGINFO("Na linha FTP entre com a Palavra chave do Boletim ex: BOLETIM98",WSIST)
      EndIf
    EndIf
    If X=4
      ZGENCIAS=""
      ZSSUNTOS=""
      FOR WPI=1 TO 100
        SysRefresh()
        VARIAS="AG"+ALLTRIM(STR(WPI))
        If &VARIAS
          ZGENCIAS=ZGENCIAS+"X"
        Else
          ZGENCIAS=ZGENCIAS+"_"
        EndIf
        VARIAS="ASS"+ALLTRIM(STR(WPI))
        If &VARIAS
          ZSSUNTOS=ZSSUNTOS+"X"
        Else
          ZSSUNTOS=ZSSUNTOS+"_"
        EndIf
      NEXT WPI
    EndIf
    If X=5
      RETORNO = .F.
    EndIf
    If X=6
      If EMPTY(Z01CALCATE)
        Z01DATA=STRZERO(YEAR(DATE()),4)+" "
      Else
        Z01DATA=STRZERO(YEAR(Z01CALCATE),4)+" "
      EndIf
      o01data:refresh()
    EndIf
    If X=7
      If val(ZOD)=0
        return  .T.
      EndIf
      SELE 37
      SET ORDER TO 1
      SEEK val(ZOD)
      If !FOUND()
        ZOD=space(15)
        MSGBEEP()
        MSGINFO("Não encontrado !",WSIST)
        return  .f.
      EndIf
      ZOME=NOME
      onome1:refresh()
    EndIf
    If X=8
      WAITRUN(ARQOBS3,0)
      DO WHILE FILE(ARQOBS2)
      ENDDO
    EndIf
  return  .T.

  Function VSTCON(CHAVE)
    SET ORDER TO 1
    SEEK CHAVE
    If EOF()
      return  .T.
    Else
      MSGINFO("Já Existente !",WSIST)
      return  .F.
    EndIf
  return 

  Procedure CADTROCA()
    If !USUARIO(3)
      return 
    EndIf
    SELE 21
    If !USEREDE(LOCNOT3+"TROCA.DBF",.F.,10,"TROCA1")
      return 
    EndIf
    Browse(4,"Listas Restritivas","Listas Restritivas", { || buscatroca() },{ ||NCADTROCA("I") },{ ||NCADTROCA("A") })
    SELE 21
    USE
  return 

  Function BUSCATROCA()
    TEXTO="Lista"
    VARIAVEL=SPACE(50)
    MASCARA="@!"
    REG=RECNO()
    DO WHILE .T.
      SysRefresh()
      If !PROCURA()
        EXIT
      EndIf
      CHAVE=VARIAVEL
      If EMPTY(CHAVE)
        EXIT
      EndIf
      LOCATE FOR LOWER(PARA)=LOWER(ALLTRIM(CHAVE))
      If EOF() .OR. DELE()
        MSGALERT("Não Encontrado !","Alerta")
        GO REG
      Else
        EXIT
      EndIf
    ENDDO
  return 

  Procedure NCADTROCA(FS)
    Local oDlg,oBrush
    PUBLIC F
    F=FS
    DECLARE _VAR1[FCOUNT()]
    AFIELDS(_VAR1)
    If F="I"
      GO BOTT
      SKIP
    EndIf
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t
    RELEASE _VAR1
    MFONTES()
    If ZOLETIM="S"
      ZOLET=.T.
    Else
      ZOLET=.F.
    EndIf
    DEFINE FONT oFont NAME "courier new" SIZE 0,-12
    DEFINE DIALOG oDlg FROM 0,0 TO 23,62 TITLE "Lista Restritiva" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 239,139
    @   0.87,02 SAY "Descrição:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   1.74,02 SAY "Endereço:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,02 SAY "Ultima Atualizacao" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,20 SAY "Pag/Rec:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

    @   3.48,02 SAY "Layout:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,20 SAY "Tp Arquivo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)


    @ 01,10 GET ZARA PICTURE "@!" FONT oFontget SIZE 150,11
    @ 02,10 GET ZE  FONT oFontget SIZE 150,11
    @ 03,10 GET ZTULT PICTURE "@D" FONT oFontget SIZE 40,11
    @ 03,20 GET ZNVIAFTP PICTURE "!" FONT oFontget SIZE 10,11 VALID ZNVIAFTP$"SN"
    @ 04,10 GET ZTUA PICTURE "99" RIGHT  FONT oFontget SIZE 20,11
    @ 04,20 GET ZLIENTE PICTURE "!!!" FONT oFontget SIZE 20,11

    @ 05,1 GET ZODAPE FONT oFontget MEMO SIZE 200,75

    @ 154,035 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"    ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (IIF(F="I",ADIREG(0),""),PUTREC(),oDlg:end())
    @ 154,100 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"     ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
    oDlg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlg CENTERED
  return 

  Procedure CADBOL()
    If !USUARIO(4)
      return 
    EndIf
    SELE 21
    If !USEREDE(LOCNOT3+"PESQ.DBF",.F.,10,"PESQ1")
      return 
    EndIf
    Browse(5,"Usuarios de Boletins/Alertas","Usuarios de Boletins/Alertas", { || buscabol() },{ ||NCADBOL("I") },{ ||NCADBOL("A") })
    SELE 21
    USE
  return 

  Function BUSCABOL()
    TEXTO="Email"
    VARIAVEL=SPACE(100)
    MASCARA=""
    REG=RECNO()
    DO WHILE .T.
      SysRefresh()
      If !PROCURA()
        EXIT
      EndIf
      CHAVE=VARIAVEL
      If EMPTY(CHAVE)
        EXIT
      EndIf
      LOCATE FOR lower(EMAIL)=LOWER(ALLTRIM(CHAVE))
      If EOF() .OR. DELE()
        MSGALERT("Não Encontrado !","Alerta")
        GO REG
      Else
        EXIT
      EndIf
    ENDDO
  return 

  Procedure NCADBOL(FS)
    Local oDlg,oBrush
    PUBLIC F
    F=FS
    DECLARE _VAR1[FCOUNT()]
    AFIELDS(_VAR1)
    If F="I"
      GO BOTT
      SKIP
    EndIf
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t
    RELEASE _VAR1
    MFONTES()
    DEFINE DIALOG oDlg FROM 0,0 TO 34,63 TITLE "Boletim/Avisos" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 243,230
    @   0.87,02 SAY "SMTP - Cliente" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   1.74,02 SAY "Conta-Cliente" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,02 SAY "Nome" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,02 SAY "Email" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,02 SAY "Ult.Bol.Enviado" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,02 SAY "Ativo" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,02 SAY "Admin" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.96,02 SAY "Maior,Menor" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   7.83,02 SAY "Valor" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   9.57,02 SAY "Chave 01" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  10.44,02 SAY "Corretora" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  11.31,02 SAY "Usuario" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  12.18,02 SAY "Senha" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  13.05,02 SAY "Chave 02" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  14.79,02 SAY "Consulta NR" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

    @ 01,10 GET ZMTP FONT oFontget SIZE 150,11
    @ 02,10 GET ZONTA FONT oFontget SIZE 150,11
    @ 03,10 GET ZOME FONT oFontget SIZE 80,11
    @ 04,10 GET ZMAIL FONT oFontget SIZE 150,11
    @ 05,10 GET ZTBOL PICTURE "@D" FONT oFontget SIZE 45,11
    @ 06,10 GET ZOEDA FONT oFontget SIZE 45,11
    @ 07,10 GET ZDMIN PICTURE "!" FONT oFontget SIZE 8,11
    @ 08,10 GET ZAIORMENOR PICTURE "!" FONT oFontget SIZE 8,11
    @ 09,10 GET ZALOR PICTURE "@E 999,999.9999" FONT oFontget SIZE 60,11
    @ 11,10 GET ZALAVRA1 FONT oFontget SIZE 150,11
    @ 12,10 GET ZORRETORA FONT oFontget SIZE 150,11
    @ 13,10 GET ZSUARIO FONT oFontget SIZE 150,11
    @ 14,10 GET ZENHA FONT oFontget SIZE 150,11
    @ 15,10 GET ZALAVRA5 FONT oFontget SIZE 150,11
    @ 17,10 GET ZONSULTANR PICTURE "@E 99,999,999" RIGHT FONT oFontget SIZE 60,11

    @ 242,080 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (IIF(F="I",ADIREG(0),""),PUTREC(),oDlg:end())
    @ 242,140 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
    oDlg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlg CENTERED
  return 

  Procedure ROTINASDIA()
    If OP5#"S"
      return 
    EndIf
    CURSORWAIT()
    If PRIMEIRA1=DATE()
      return 
    EndIf
    TMP_="c:\exatus\ExRotina.TXT"
    Erase &TMP_
    TMP_="c:\exatus\rotbolsa.txt"
    Erase &TMP_
    CLOSE DATABASE
    PRIMEIRA1=DATE()
    SAVE ALL LIKE PRIMEIRA* TO LOCAIS.DAT
    SELE 1
    If !USEREDE(LOCNOT3+"NOTAEX.DBF",.T.,10,"NOTAEX")
      CLOSE DATABASE
      ABREARQS()
      return 
    EndIf
    GO TOP
    DO WHILE .NOT. EOF()
      SysRefresh()
      If DTNOT<DATE()-NRDIAS
        DELE
      EndIf
      SKIP
    ENDDO
    PACK
    //Sx_MemoPack( 32, {|| QQOut(".")}, LastRec() / 10 )
    GO TOP
    DO WHILE .NOT. EOF()
      SysRefresh()
      REPLACE IDNOT WITH RECNO()
      SKIP
    ENDDO
    CLOSE DATABASE
    SELE 6
    If !USEREDE(LOCNOT3+"TAXAS.DBF",.T.,10,"TAXAS")
      CLOSE DATABASE
      ABREARQS()
      return 
    EndIf
    GO TOP
    DO WHILE .NOT. EOF()
      SysRefresh()
      If AT(";",CODATI)#0
        SKIP
        LOOP
      EndIf
      If DTULTNEG<DATE()-LOCALDIASTAXAS
        DELE
      EndIf
      SKIP
    ENDDO
    PACK
    CLOSE DATABASE
    ABREARQS()
    COMMIT
  return 

  Function PESQ
    PARA M_RELACAO,M_AREA,M_ORDEM,M_CAMPO,M_SAIDA,M_VAZIO
    M->U_AREATUA = ALLTRIM(STR(SELECT(),3))
    M->U_ORDEM = INDEXORD()
    SELE &M_AREA
    SET ORDER TO M->M_ORDEM
    SEEK M->M_RELACAO
    M->RE_TORNO = &M_CAMPO
    SELE &U_AREATUA
    SET ORDER TO M->U_ORDEM
    If PCOUNT()>4
      &M_SAIDA=M->RE_TORNO
      If PCOUNT()=6
        return  ""
      EndIf
    EndIf
  return  M->RE_TORNO

  Procedure ROTINAS()
    M->DAT_HOJE=DATE()
    If OP7#"S"
      return 
    EndIf
    LIMPAMEM()
    HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
    If HORAS<VAL(LOCALINICIO) .OR. HORAS>VAL(LOCALFIM)
      return 
    EndIf
    OP7="N"
    U1_AREATUA = ALLTRIM(STR(SELECT(),3))
    U1_ORDEM = INDEXORD()
    CURSORWAIT()
    ARQWOR=ALLTRIM(LOCALNOT3)+DTOS(DATE())+".LOG"
    If WSEN="C"
      MARQ=LOCALBC+"HORARIC.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      MARQ=LOCALCART+"FEITO.ACO"
      If FILE(MARQ)
        DECLARE TX1[1]
        TX1[1]="Carteira Net"
        LOGFILE(ARQWOR,TX1)
        CARTEIRANET()
        MARQ=LOCALCART+"FEITO.ACO"
        Erase &MARQ
      EndIf
      SELE &U1_AREATUA
      If U1_ORDEM#0
        SET ORDER TO U1_ORDEM
      EndIf
      CURSORARROW()
      OP7="S"
      return 
    EndIf
    If WSEN="A"
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
      DECLARE TX1[1]
      TX1[1]="Gera Assina"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      GERAASSINA("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
      DECLARE TX1[1]
      TX1[1]="GeraFichaPF"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      GERAFICHAPF("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Assinanet"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      ASSINANET("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Completo"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      ENVCOMPLETO("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Cobra Docto"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      COBRADOC("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="TodasTaxas"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      TODASTAXAS("S")
      DECLARE TX1[1]
      TX1[1]="TodasTaxas"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      ATUALIZAASSINA("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      MARQ=LOCALGNEW+DTOS(DATE())+".HUB"
      HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
      If HORAS>0915 .And. !FILE(MARQ)
        MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
        DECLARE TX1[1]
        TX1[1]="Verifica Hub"
        LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
        VERHUB()
      EndIf
      MARQ=LOCALGNEW+DTOS(DATE())+".HU1"
      If HORAS>0915 .And. !FILE(MARQ)
        MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
        DECLARE TX1[1]
        TX1[1]="VerHUb1"
        LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
        VERHUB1()
      EndIf
      MARQ=LOCALBC+DTOS(DATE())+".FEZ"
      If DOW(DATE())=7 .And. day(date())>23 .And. !FILE(MARQ)
        MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
        DECLARE TX1[1]
        TX1[1]="Listas"
        LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
        LISTA("S")
      EndIf
    EndIf
    If WSEN="B"
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      FEZCSV=.F.
      //    CADWEB1("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      FEZCSV=.F.
      DECLARE TX1[1]
      TX1[1]="Rotinas Dia"
      LOGFILE(WCORRENTE+"\rotbolsa.TXT",TX1)
      ROTINASDIA()
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Ouvidoria"
      LOGFILE(WCORRENTE+"\rotbolsa.TXT",TX1)
      OUVIDORIA()
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    BUSCAXML() // noticias
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    BUSCANOVAS() // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      ATUOUCOTAS=.F.
      //    ATUACOTAS()  // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    tuacota1()  // RETIRADO EM 25/07/2020
      FAZGRAFMES1=.F.
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    GRAFMES()  // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    GRAFMES2()  // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    GRAFDIA1()  // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    GRAFMES1()   // RETIRADO EM 25/07/2020
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Intcam"
      LOGFILE(WCORRENTE+"\rotbolsa.TXT",TX1)
      INTCAM("S") // integra Sictur
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Todas Taxas"
      LOGFILE(WCORRENTE+"\rotbolsa.TXT",TX1)
      TODASTAXAS("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //    ATUAWEBWINCC()
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="RecPdf"
      LOGFILE(WCORRENTE+"\rotbolsa.TXT",TX1)
      RECPDF("S")
      MARQ=LOCALBC+"HORARIO.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //  CALCULAGET(DATE())
      //  MARQ=LOCALBC+"HORARIO.CSV"
      //  MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      //  CALCCART("S")
      If CONFIDENCE
        MARQ=LOCALBC+"HORARIO.CSV"
        MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
        DECLARE TX1[1]
        TX1[1]="Confidence"
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        ATUACONFIDENCE()
        MARQ=LOCALBC+"HORARIO.CSV"
        MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      EndIf
    ElseIf WSEN="W"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Backup Assinanet"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      ASSINABACKUP()
      MARQ=LOCALBC+"HORARIW.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Backup Documentos"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      BACKDOCSICTUR()
      MARQ=LOCALBC+"HORARIW.CSV"
      MEMOWRIT(MARQ,TIME()+DTOC(DATE()))
      DECLARE TX1[1]
      TX1[1]="Backup Doctosnet"
      LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
      BACKDOCNET()
    EndIf
    COMMIT
    //* PARA RETORNAR A ROTINA QUE ESTAVA FAZENDO ANTERIORMENTE **
    SELE &U1_AREATUA
    If U1_ORDEM#0
      SET ORDER TO U1_ORDEM
    EndIf
    CURSORARROW()
    DECLARE VLABEL[ADIR(LOCNOT1+"*.TMP")]
    ADIR(LOCNOT1+"*.TMP",vlabel)
    VLABEL:=ASORT(VLABEL)
    I=0
    FOR I=1 TO LEN(VLABEL)
      SysRefresh()
      NARQ=LOCNOT1+VLABEL[I]
      Erase &NARQ
    NEXT I
    DECLARE TX1[1]
    TX1[1]="Final dos Serviços automaticos"+chr(13)+chr(10)
    If WSEN="B"
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
    Else
      LOGFILE(WCORRENTE+"\ExRotina.TXT",TX1)
    EndIf
    OP7="S"
  return 

  Procedure SINAL(WMTEXTO)
    If VALTYPE(WMTEXTO)<>"C"
      return  WMTEXTO
    EndIf
    WA="#$Á&ÉÍÓÚáéíóúçÇãõÃÕàèìòùÀÈÌÒÙÂÊÎÔÛâêîôûäëïöüÄËÏÖÜºªÑ"
    FOR I=1 TO 31
      WA=WA+CHR(I)
    NEXT I
    FOR I=128 TO 255
      WA=WA+CHR(I)
    NEXT I
    WB=" DAEEIOUaeioucCaoAOaeiouAEIOUAEIOUaeiouaeiouAEIOU"+SPACE(300)
    FOR I=1 TO LEN(WA)
      DO WHILE .T.
        LETRA=SUBSTR(WA,I,1)
        TROCA=SUBSTR(WB,I,1)
        ACHOU=AT(LETRA,WMTEXTO)
        If ACHOU>0
          WMTEXTO=LEFT(WMTEXTO,ACHOU-1)+TROCA+RIGHT(WMTEXTO,LEN(WMTEXTO)-ACHOU)
        Else
          EXIT
        EndIf
      ENDDO
    NEXT I
  return  WMTEXTO

  Function TimeDiff( cStartTime, cEndTime )
    return  TimeAsString(IF(cEndTime < cStartTime, 86400 , 0) +;
      TimeAsSeconds(cEndTime) - TimeAsSeconds(cStartTime))

  Function TimeAsString( nSeconds )
    return  StrZero(INT(Mod(nSeconds / 3600, 24)), 2, 0) + " " +;
      StrZero(INT(Mod(nSeconds / 60, 60)), 2, 0) + " " +;
      StrZero(INT(Mod(nSeconds, 60)), 2, 0)

  Function TimeAsSeconds( cTime )
    return  VAL(cTime) * 3600 + VAL(SUBSTR(cTime, 4)) * 60 +;
      VAL(SUBSTR(cTime, 7))

  Procedure CADNOT(QYZ)
    QY1=QYZ
    SELE 21
    If !USEREDE(LOCNOT3+"NOTAEX.DBF",.F.,10,"NOTAEX1")
      return 
    EndIf
    SET ORDER TO 1
    If QY1=1
      Browse(9,"Noticias-Exatus"," Noticias-Exatus", { || BUSCANOTI() },{ ||NCADNOTI("I") },{ ||NCADNOTI("A") })
    EndIf
    SELE 21
    USE
  return 

  Procedure BUSCANOTI()
    TEXTO="Código"
    VARIAVEL=0
    MASCARA="99999999"
    REG=RECNO()
    DO WHILE .T.
      SysRefresh()
      If !PROCURA()
        EXIT
      EndIf
      CHAVE=VARIAVEL
      If EMPTY(CHAVE)
        EXIT
      EndIf
      SEEK chave
      If EOF() .OR. DELE()
        MSGALERT("Não Encontrado !","Alerta")
        GO REG
      Else
        EXIT
      EndIf
    ENDDO
  return 

  Procedure NCADNOTI()
    Local oDlg,oBrush
    PUBLIC F
    DECLARE _var1 [FCOUNT()]
    AFIELDS(_var1)
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t
    MFONTES()
    DEFINE DIALOG oDlg FROM 1,1 TO 26,78 TITLE "Noticia" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 SAY " " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 300,160 box raised
    @   0.87,1 SAY "Agencia:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   1.74,1 SAY "Assunto:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,1 SAY "Data:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,1 SAY "Hora:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,1 SAY "Prioridade:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,1 SAY "Manchete:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,1 SAY "LINK:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    A="N"
    ZEXTO=OEMTOANSI(ZEXTO)
    ZANCHETE=OEMTOANSI(ZANCHETE)
    ZESCAG=OEMTOANSI(ZESCAG)
    ZESCASS=OEMTOANSI(ZESCASS)

    @ 1,8 GET ZGENCIA FONT oFontget SIZE 20,11 RIGHT WHEN A="S"
    @ 1,12 GET ZESCAG FONT oFontget SIZE 203,11 WHEN A="S"
    @ 2,8 GET ZSSUNTO FONT oFontget SIZE 20,11 RIGHT WHEN A="S"
    @ 2,12 GET ZESCASS FONT oFontget SIZE 203,11 WHEN A="S"
    @ 3,8 GET ZTNOT FONT oFontget SIZE 45,11 WHEN A="S"
    @ 4,8 GET ZRNOT FONT oFontget SIZE 35,11 WHEN A="S"
    @ 5,8 GET ZRIORID FONT oFontget SIZE 15,11 RIGHT WHEN A="S"
    @ 6,8 GET ZANCHETE FONT oFontget SIZE 234,11 WHEN A="S"
    @ 7,8 GET ZINK FONT oFontget SIZE 234,11 WHEN A="S"
    @ 8,1 GET ZEXTO FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) OF oDlg MEMO SIZE 290,60
    @ 13.3,28 SBUTTON NOBOX RESOURCE "OFF","ON","DISABLE"  ADJUST TEXT ON_CENTER  PROMPT "&OK" SIZE 35,12 ACTION (oDlg:End()) FONT oFontbut
    odlg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlG CENTERED
  return 

  Procedure ATUACOTAS()
    Local NOMAARQUIVO:=DIRECTORY(LOCNOT1+"C*.TXT")
    CURSORWAIT()
    FAZERATE=LEN(NOMAARQUIVO)
    If FAZERATE>50
      FAZERATE=50
    EndIf
    FOR ZII=1 TO FAZERATE
      SysRefresh()
      ATUOUCOTAS=.T.
      YARQ=ALLTRIM(LOCNOT1+NOMAARQUIVO[ZII][1])
      If NOMAARQUIVO[zii][3]=date()
        DIF=TIMEDIFF(NOMAARQUIVO[1][4],TIME())
        DIF1=TIMEASSECONDS(DIF)
        If DIF1<15
          LOOP
        EndIf
      EndIf
      oText:=TTxtfile():NEW(YARQ)
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)<>"N"
        CONTAD=1
      EndIf
      nrbol=VAL(oText:ReadLine())+1
      If nrbol=1 .And. date()>CTOD("05/01/"+STRZERO(YEAR(DATE()),4))
      Else
        LINHA='<%'+CHR(13)+CHR(10)+'NrBol = "'+RIGHT(STRZERO(YEAR(M->DAT_HOJE),4),2)+STRZERO(NRBOL,4)+'"'+CHR(13)+CHR(10)+'%>'+CHR(13)+CHR(10)
        ARQ_NOM=LOCALGNEW+"CONFIG.INC"
        Erase &arq_nom
        If !file(arq_nom)
          ARQS=FCREATE(ARQ_NOM,0)
          FWRITE(arqS,LINHA)
          FCLOSE(ARQS)
        EndIf
      EndIf
      oText:Skip()
      horau=oText:ReadLine()
      qualatual=substr(horau,8,10)
      horau=left(horau,5)
      oText:Skip()
      If qualatual="Fech. Ptax"
        PORRA=ALLTRIM(YARQ)
        PORRA=LEFT(YARQ,LEN(YARQ)-4)+".CS1"
        LZCOPYFILE(YARQ,PORRA)
        lzcopyfile(YARQ,LOCNOT2+"TAXAS.CSV")
        LZCOPYFILE(YARQ,"\SITES\EXATUS\GRAFICO\TAXAS.CSV")
      EndIf
      If QUALATUAL="Abertura"
        LZCOPYFILE(YARQ,"\SITES\EXATUS\GRAFICO\ABERTURA.CSV")
      EndIf
      FOR WZI=3 TO CONTAD
        SysRefresh()
        WLI1=ALLTRIM(oText:ReadLine())
        WLI=ALLTRIM(SUBSTR(WLI1,2,LEN(WLI1)))+"; ; ;"
        WLI=Strtran(WLI,",",".")
        COLU=AT(";",WLI)
        ZODATI=LEFT(WLI,COLU-1)
        ZODATI=ALLTRIM("EX"+SUBSTR(ZODATI,4,3))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZARC=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZARV=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZROFCOMPRA=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZROFVENDA=VAL(LEFT(WLI,COLU-1))
        ZORAULTNEG=HORAU
        ZLTNEG=ZROFCOMPRA
        ZSCILA=0
        ZRMAXIMO=0
        ZRMINIMO=0
        ZRABERTURA=0
        ZRMEDIO=0
        ZRFECANT=0
        ZRNEG=0
        ZTDULTNEG=0
        ZTDTOTAL=0
        ZOLUME=0
        ZOLUMETOT=0
        ZTULTNEG=DATE()
        ZUOFCOMPRA=0
        ZUOFVENDA=0
        ZODATI1=ALLTRIM(ZODATI)
        ZODATI=ZODATI1+SPACE(12-LEN(ZODATI1))
        ZODATI1=ZODATI
        ZRAFICO="S"
        SELE 6
        If ZRAFICO="S"
          SEEK ZODATI+DTOS(ZTULTNEG)+LEFT(ZORAULTNEG,5)
          If EOF()
            ADIREG(0)
          Else
            REGLOCK(10)
          EndIf
          PUTREC()
          COMMIT
        EndIf
        If PRIMEIRA2=0
          GO BOTT
          PRIMEIRA2=RECNO()
        EndIf
        oText:Skip()
      NEXT WZI
      oText:Close()
      Erase &YARQ
    NEXT ZII
  return 

  Procedure CRIAMOV()
    FAZER=.F.
    If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"EMPVAL")
      MSGStop("O arquivo EMPRESAS não está  disponível",WSIST)
      return 
    EndIf
    If FCOUNT()<35
      CLOSE DATABASE
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
        ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
        ARQ_PRN1:=ARQ_PRN+".DAT"
        ARQ_PRN2:=ARQ_PRN+".DBT"
        If !FILE(ARQ_PRN1)
          EXIT
        EndIf
      ENDDO
      MAT_STRU={}
      AADD(mat_stru,{"a01code","C",06,0})   // Código na Exatus
      AADD(mat_stru,{"a01nome","C",60,0})   // Nome
      AADD(MAT_STRU,{"A01ENDER","C",60,0})  // Endereço
      AADD(MAT_STRU,{"A01CEPCID","C",60,0}) // Cep/Cidade
      AADD(mat_stru,{"a01data","C",04,0})   // Data (utilizar 4 para o ano)
      AADD(mat_stru,{"a01Tipo","N",1,0})    // 1-Fisica 2-Juridica
      AADD(mat_stru,{"a01CGCCPF","C",14,0}) // CGCCPF
      AADD(MAT_Stru,{"a01SENHA","C",20,0})  // Senha de Acesso Internet
      AADD(MAT_STRU,{"a01Frase","C",60,0})  // Frase Secreta Acesso Internet
      AADD(MAT_STRU,{"A01VALID","D",8,0})   // Data Validade do Pagamento
      AADD(MAT_STRU,{"A01INATI","L",1,0})   // Se verdadeiro = Inativo
      AADD(MAT_STRU,{"A01DIAVENC","N",2,0}) // Dia de Vencimento da Fatura
      AADD(MAT_STRU,{"A01EMITIDO","D",8,0}) // Data do Vencimento da Fatura
      AADD(MAT_STRU,{"A01PAGO","D",8,0})    // Data do Pagamento da Fatura
      AADD(MAT_STRU,{"A01EMAIL","C",253,0}) // Email de Recuperação de Senha
      AADD(MAT_STRU,{"A01VALMES","N",19,2}) // Valor da Mensalidade
      AADD(MAT_STRU,{"A01CORRET","N",3,0})  // Se deve atualizar a carteira esta 1
      AADD(MAT_STRU,{"A01CODNAC","C",10,0}) // Se for Administrador terá "S" aqui
      AADD(MAT_STRU,{"A01CONFER","L",1,0})  // Se Verdadeiro foi verificado e pode utilizar as notas
      AADD(MAT_STRU,{"A01CORP","N",3,0}) // Codigo da Corretora Administradora ou da Corretora no caso do campo a01codnac = "S"
      AADD(MAT_STRU,{"A01COD1","N",8,0}) // Codigo do Cliente na Corretora Principal
      AADD(MAT_STRU,{"A01COD1I","N",8,0})// Codigo do Cliente na Corretora Principal Investimento
      AADD(MAT_STRU,{"A01COR2","N",3,0}) // Codigo da Segunda Corretora
      AADD(MAT_STRU,{"A01COD2","N",8,0}) // Codigo do Cliente na Segunda Corretora
      AADD(MAT_STRU,{"A01COD2I","N",8,0})// Codigo do Cliente na Segunda Corretora Investimento
      AADD(MAT_STRU,{"A01COR3","N",3,0}) // Codigo da Terceira Corretora
      AADD(MAT_STRU,{"A01COD3I","N",8,0})// Codigo do Cliente na Terceira Corretora Investimento
      AADD(MAT_STRU,{"A01COD3","N",8,0}) // Codigo do Cliente na Terceira Corretora
      AADD(MAT_STRU,{"A01COR4","N",3,0}) // Codigo da Quarta Corretora
      AADD(MAT_STRU,{"A01COD4","N",8,0}) // Codigo do Cliente na Quarta Corretora
      AADD(MAT_STRU,{"A01COD4I","N",8,0})// Codigo do Cliente na Quarta Corretora Investimento
      AADD(MAT_STRU,{"A01COR5","N",3,0}) // Codigo da Quinta Corretora
      AADD(MAT_STRU,{"A01COD5","N",8,0}) // Codigo do Cliente na Quinta Corretora
      AADD(MAT_STRU,{"A01COD5I","N",8,0})// Codigo do Cliente na Quinta Corretora Investimento
      AADD(MAT_STRU,{"A01CALCATE","D",8,0}) // Calcular até
      DBCREATE(ARQ_PRN1,MAT_STRU)
      FAZER=.T.
      SELE 1
      If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
        MSGINFO("Não foi possivel continuar !","Tente mais tarde")
        CLOSE DATABASE
        QUIT
      EndIf
      ARQ=LOCALCART+"EMPRESAS.DBF"
      APPEND FROM &ARQ
      CLOSE DATABASE
      Erase &ARQ
      FRENAME(ARQ_PRN1,ARQ)
    EndIf
    CLOSE DATABASE
    If LOCALQTIPO#"W"
      return 
    EndIf
    SELE 1
    If !USEREDE(LOCNOT3+"BANCOML.DBF",.F.,10,"EMPVAL")
      MSGStop("O arquivo EMPVAL não está  disponível",WSIST)
      return 
    EndIf
    If FCOUNT()<29
      CLOSE DATABASE
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
        ARQ_PRN=LOCNOT1+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
        ARQ_PRN1:=ARQ_PRN+".DAT"
        ARQ_PRN2:=ARQ_PRN+".DBT"
        If !FILE(ARQ_PRN1)
          EXIT
        EndIf
      ENDDO
      SELE 1
      MAT_STRU={}
      AADD(MAT_STRU,{"CODIGO","N",  6,  0})
      AADD(MAT_STRU,{"CODCLI","N",6,0})
      AADD(MAT_STRU,{"CLIENTE","C",1,0})
      AADD(MAT_STRU,{"NOME","C", 50,  0})
      AADD(MAT_STRU,{"ENDERECO","C", 50,  0})
      AADD(MAT_STRU,{"CIDADE","C", 20,  0})
      AADD(MAT_STRU,{"ESTADO","C",  2,  0})
      AADD(MAT_STRU,{"CEP","C",  8,  0})
      AADD(MAT_STRU,{"TELEFONE","C", 20,  0})
      AADD(MAT_STRU,{"FAX","C",11,0})
      AADD(MAT_STRU,{"CONTATO","C", 20,  0})
      AADD(MAT_STRU,{"TIPO","C",  1,  0})
      AADD(MAT_STRU,{"CGCCPF","C", 15,  0})
      AADD(MAT_STRU,{"OPERADOR","N",  4,  0})
      AADD(MAT_STRU,{"EMAIL","C",254,0})
      AADD(MAT_STRU,{"EMBAILBM","C",254,0})
      AADD(MAT_STRU,{"EMAILBOV","C",254,0})
      AADD(MAT_STRU,{"EMAILCONT","C",254,0})
      AADD(MAT_STRU,{"ULTCONT","D",8,0})
      AADD(MAT_STRU,{"PROXCONT","D",8,0})
      AADD(MAT_STRU,{"ULTEMAIL","D",8,0})
      AADD(MAT_STRU,{"PROXEMAI","D",8,0})
      AADD(MAT_STRU,{"INATIVO","L",1,0})
      AADD(MAT_STRU,{"EBOLSA","L",1,0})
      AADD(MAT_STRU,{"ECAMBIO","L",1,0})
      AADD(MAT_STRU,{"EBMF","L",1,0})
      AADD(MAT_STRU,{"ECONTAB","L",1,0})
      AADD(MAT_STRU,{"EOUTROS","L",1,0})
      AADD(MAT_STRU,{"ASEC","L",1,0})
      DBCREATE(ARQ_PRN1,MAT_STRU)
      FAZER=.T.
      SELE 1
      If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
        MSGINFO("Não foi possivel continuar !","Tente mais tarde")
        CLOSE DATABASE
        QUIT
      EndIf
      ARQ=LOCNOT3+"BANCOML.DBF"
      APPEND FROM &ARQ
      CLOSE DATABASE
      Erase &ARQ
      FRENAME(ARQ_PRN1,ARQ)
    EndIf
    CLOSE DATABASE
    /*
**** NAO FAZER ALTERAÇÕES NA BASE POR AQUI, FAZER MANUALMENTE E COLOCAR O NOVO CAMPO NO FINAL DO ARQUIVO
**** EXISTE CAMPO MEMO, E PODE OCORRER PROBLEMAS.
*/
    CLOSE DATABASE
    SELE 1
    If !USEREDE(LOCALWEB+"EMPWINCC.DBF",.F.,10,"EMPWINCC")
      msginfo(localweb+"empwincc.dbf",wsist)
      MSGStop("O arquivo EMPWINCC não está  disponível",WSIST)
      return 
    EndIf
    If FCOUNT()<13
      CLOSE DATABASE
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
        ARQ_PRN=LOCNOT1+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
        ARQ_PRN1:=ARQ_PRN+".DAT"
        ARQ_PRN2:=ARQ_PRN+".DBT"
        If !FILE(ARQ_PRN1)
          EXIT
        EndIf
      ENDDO
      SELE 1
      aStructure={}
      aAdd( aStructure, { "Corretora" , "N",   3, 0 } )
      aAdd( aStructure, { "NOMECORR"  , "C", 100, 0 } )
      aAdd( aStructure, { "SMTPCORR"  , "C", 100, 0 } )
      aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
      aAdd( aStructure, { "LOCALCOR"  , "C", 254, 0 } )
      aAdd( aStructure, { "ENVMAIL"   , "L",   1, 0 } )
      aAdd( aStructure, { "CORREIO"   , "L",   1, 0 } )
      aAdd( aStructure, { "DADOS"     , "L",   1, 0 } )
      aAdd( aStructure, { "PASTALOG"  , "C", 150, 0 } )
      aAdd( aStructure, { "INSTALA"   , "L",   1, 0 } )
      aAdd( aStructure, { "WINIMAGE"  , "L",   1, 0 } )
      aAdd( aStructure, { "ULTEXTR"   , "D",   8, 0 } )
      aAdd( aStructure, { "INST"      , "N",   5, 0 } )
      DBCREATE(ARQ_PRN1,aStructure)
      FAZER=.T.
      SELE 1
      If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
        MSGINFO("Não foi possivel continuar !","Tente mais tarde")
        CLOSE DATABASE
        QUIT
      EndIf
      ARQ=LOCALWEB+"EMPWINCC.DBF"
      APPEND FROM &ARQ
      CLOSE DATABASE
      Erase &ARQ
      FRENAME(ARQ_PRN1,ARQ)
    EndIf
    CLOSE DATABASE
    SELE 1
    If !USEREDE(LOCALWEB+"USUARIO.DBF",.F.,10,"USUARIO")
      MSGStop("O arquivo USUARIO não está  disponível",WSIST)
      return 
    EndIf
    If FCOUNT()<>13
      CLOSE DATABASE
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
        ARQ_PRN=LOCNOT1+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
        ARQ_PRN1:=ARQ_PRN+".DAT"
        ARQ_PRN2:=ARQ_PRN+".DBT"
        If !FILE(ARQ_PRN1)
          EXIT
        EndIf
      ENDDO
      SELE 1
      aStructure={}
      aAdd( aStructure, { "USUARIO"   , "C",  50, 0 } )
      aAdd( aStructure, { "SENHA"     , "C",  50, 0 } )
      aAdd( aStructure, { "CLIENTE"   , "C",  10, 0 } )
      aAdd( aStructure, { "DIREITOS"  , "C",  50, 0 } )
      aAdd( aStructure, { "PASTA"     , "C", 254, 0 } )
      aAdd( aStructure, { "ISADMIN"   , "C",   1, 0 } )
      aAdd( aStructure, { "BLOQUEADO" , "L",   1, 0 } )
      aAdd( aStructure, { "CORRETORA" , "C",  50, 0 } )
      aAdd( aStructure, { "ACESSO"    , "N",  20, 5 } )
      aAdd( aStructure, { "ULTACESSO" , "D",   8, 0 } )
      aAdd( aStructure, { "FRASE"     , "C", 100, 0 } )
      aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
      aAdd( aStructure, { "ENVEXTR"   , "L",   1, 0 } )
      DBCREATE(ARQ_PRN1,aStructure)
      FAZER=.T.
      SELE 1
      If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
        MSGINFO("Não foi possivel continuar !","Tente mais tarde")
        CLOSE DATABASE
        QUIT
      EndIf
      ARQ=LOCALWEB+"USUARIO.DBF"
      APPEND FROM &ARQ
      CLOSE DATABASE
      Erase &ARQ
      FRENAME(ARQ_PRN1,ARQ)
    EndIf
    CLOSE DATABASE
  return 

  Procedure CADTAXA(QUAL)
    SELE 21
    If !USEREDE(LOCNOT3+"TAXAS.DBF",.F.,10,"TAXAS1")
      return 
    EndIf
    Browse(10,"Taxas"," Taxas", { || BUSCATaxa() },{ ||NCADtaxa("I") },{ ||NCADtaxa("A") })
    SELE 21
    USE
  return 

  Procedure BUSCATAXA()
    TEXTO="Código"
    VARIAVEL=SPACE(12)
    MASCARA="@!"
    REG=RECNO()
    DO WHILE .T.
      SysRefresh()
      If !PROCURA()
        EXIT
      EndIf
      CHAVE=VARIAVEL
      If EMPTY(CHAVE)
        EXIT
      EndIf
      SEEK TRIM(chave)
      If EOF() .OR. DELE()
        MSGALERT("Não Encontrado !","Alerta")
        GO REG
      Else
        EXIT
      EndIf
    ENDDO
  return 

  Procedure NCADTAXA()
    Local oDlg,oBrush
    PUBLIC F
    DECLARE _var1 [FCOUNT()]
    AFIELDS(_var1)
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t

    MFONTES()
    DEFINE DIALOG oDlg FROM 1,1 TO 26,80 TITLE "Taxa" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 SAY " " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 305,155 box raised
    @   0.87,01 SAY "Codigo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   1.74,01 SAY "Hora:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   1.74,27 SAY "Data:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,01 SAY "Ult.Neg:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,27 SAY "Oscilacao:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,01 SAY "Compra:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,27 SAY "Quant.Compra:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,01 SAY "Venda:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,27 SAY "Quant.Venda:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,01 SAY "Maximo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,27 SAY "Minimo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,01 SAY "Abertura:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,27 SAY "Medio:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.96,01 SAY "Fech.Anterior:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.96,27 SAY "Nr. Negocios:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   7.83,01 SAY "Quant.Negocios:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   7.83,27 SAY "Quant.Total:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   8.70,01 SAY "Volume:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   8.70,27 SAY "Volume Total:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   9.57,01 SAY "Paridade Venda:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   9.57,27 SAY "Paridade Compra:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    A="N"
    @ 01,08 GET ZODATI FONT oFontget SIZE 50,11 WHEN A="S"
    @ 0.73,21 SAY OEMTOANSI(PESQ(ZODATI,"10",1,"DESC")) FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 170,11
    @ 02,08 GET ZORAULTNEG FONT oFontget SIZE 50,11 WHEN A="S"
    @ 02,28 GET ZTULTNEG FONT oFontget SIZE 50,11 WHEN A="S"
    @ 03,08 GET ZLTNEG FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 03,28 GET ZSCILA FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 04,08 GET ZROFCOMPRA FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 04,28 GET ZUOFCOMPRA FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 05,08 GET ZROFVENDA  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 05,28 GET ZUOFVENDA FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 06,08 GET ZRMAXIMO  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 06,28 GET ZRMINIMO  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 07,08 GET ZRABERTURA  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 07,28 GET ZRMEDIO  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 08,08 GET ZRFECANT  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 08,28 GET ZRNEG  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 09,08 GET ZTDULTNEG FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 09,28 GET ZTDTOTAL FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 10,08 GET ZOLUME  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 10,28 GET ZOLUMETOT FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 11,08 GET ZARV  FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 11,28 GET ZARC FONT oFontget SIZE 80,11 RIGHT WHEN A="S"
    @ 13,28 SBUTTON NOBOX RESOURCE "OFF","ON","DISABLE"  ADJUST TEXT ON_CENTER  PROMPT "&OK" SIZE 35,12 ACTION (oDlg:End()) FONT oFontbut
    odlg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlG CENTERED
  return 

  Procedure PARAMETROS()
    PUBLIC oPar
    If OP7#"S"
      MSGINFO("Aguarde término das rotinas","")
      return 
    EndIf
    MFONTES()
    If !USUARIO(6)
      return 
    EndIf

    DEFINE DIALOG oPar FROM 0,0 TO 36.5,78 TITLE "Parametros do Sistema" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 SAY " " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) BOX RAISED SIZE 302,255
    @   0.87,1 SAY "Guardar Noticias por                                       DIAS" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 220,11
    @   1.74,1 SAY "Conta Email" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,1 SAY "Servidor SMTP" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,1 SAY "Pasta Graficos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,1 SAY "Hora Início Servicos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,1 SAY "Hora Fim Servicos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,1 SAY "Guardar Taxas por                                          DIAS" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 220,11
    @   6.96,1 SAY "Arquivos Temporarios" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   7.83,1 say "Arquivo TAXA.CSV" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   8.70,1 SAY "Arquivos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   9.57,1 SAY "Log Email" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  10.44,1 SAY "Verificar a cada                                              SEGUNDOS" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 240,11
    @  11.31,1 SAY "FTP - Noticias" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  12.18,1 SAY "Diretório para Taxas" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  13.05,1 SAY "Pastas Graficos (Exatus)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  13.92,1 SAY "Usuários BolsasNet/CambioNet" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  14.79,1 SAY "Verificar Noticias/Cotaçoes a cada                    MINUTOS" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) size 220,11
    @  15.66,1 SAY "Pasta Cart.Ações" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  16.53,1 SAY "Logs Horarios" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

    LOCALNOT1=LOCNOT1+SPACE(50)
    LOCALNOT2=LOCNOT2+SPACE(50)
    LOCALNOT3=LOCNOT3+SPACE(50)
    LOCALNOT5=LOCNOT5+SPACE(50)
    LOCALFROM=LOCALFROM+SPACE(40)
    CIPSERVER=CIPSERVER+SPACE(40)
    LOCALGRAFIC=LOCALGRAFIC+SPACE(40)
    LOCALGNEW=LOCALGNEW+SPACE(40)
    LOCALTAXAS=LOCALTAXAS+SPACE(40)
    LOCALCART=LOCALCART+SPACE(40)
    LOCALWEB=LOCALWEB+SPACE(150)
    LOCALBC=LOCALBC+SPACE(150)

    @ 01,15 GET LOCALDIASNOT PICTURE "99999" FONT ofontget SIZE 20,11 MESSAGE "Guardar Noticias por nn dias"
    @ 02,15 GET LOCALFROM FONT oFontget SIZE 180,11 MESSAGE "Conta Email"
    @ 03,15 GET CIPSERVER FONT oFontget SIZE 180,11 MESSAGE "SMTP Email "
    @ 04,15 GET LOCALGRAFIC FONT oFontget SIZE 180,11 MESSAGE "Local dos Graficos"
    @ 05,15 GET LOCALINICIO FONT oFontget SIZE 20,11 MESSAGE "Hora Início dos Servicos"
    @ 06,15 GET LOCALFIM FONT oFontget SIZE 20,11 MESSAGE "Hora Fim dos Servicos"
    @ 07,15 GET LOCALDIASTAXAS PICTURE "99999" FONT oFontget SIZE 20,11 MESSAGE "Guardar Taxas por nn dias"
    @ 08,15 GET LOCALNOT1 FONT oFontget SIZE 180,11
    @ 09,15 GET LOCALNOT2 FONT oFontget SIZE 180,11
    @ 10,15 GET LOCALNOT3 FONT oFontget SIZE 180,11
    @ 11,15 GET LOCALNOT5 FONT oFontget SIZE 180,11
    @ 12,15 GET LOCALINTERVALO PICTURE "99999" FONT oFontget SIZE 20,11 MESSAGE "Intervalo para verificacao dos arquivos"
    @ 14,15 GET LOCALTAXAS FONT oFontget SIZE 180,11 MESSAGE "Local para solicitar taxas"
    @ 15,15 GET LOCALGNEW FONT oFontget SIZE 180,11 MESSAGE "Local Salvar Graficos (Exatus) "
    @ 16,15 GET LOCALWEB FONT oFontget SIZE 180,11 MESSAGE "Local BolsasNet"
    @ 17,15 GET LOCALCOTNOT PICTURE "99" RIGHT FONT oFontget SIZE 20,11 MESSAGE "A cada xx minutos "
    @ 18,15 GET LOCALCART FONT oFontget SIZE 180,11 MESSAGE "Local Carteiras"
    @ 19,15 GET LOCALBC FONT oFontget SIZE 180,11 MESSAGE "Local BC (logs)"
    @ 263,135 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE"  ADJUST TEXT ON_CENTER  PROMPT "&Ok" OF OPAR SIZE 40,12 ACTION (GRAVPAR(),OPAR:END()) FONT oFontbut
    oPar:lhelpicon:=.f.
    ACTIVATE DIALOG oPar CENTERED
  return 

  Procedure GRAVPAR()
    LOCALBC=ALLTRIM(LOCALBC)
    LOCALFROM=ALLTRIM(LOCALFROM)
    cIPServer=ALLTRIM(CIPSERVER)
    LOCALGRAFIC=ALLTRIM(LOCALGRAFIC)
    LOCALGNEW=ALLTRIM(LOCALGNEW)
    LOCALSERVER=CIPSERVER
    LOCALNOT1=ALLTRIM(LOCALNOT1)
    LOCALNOT2=ALLTRIM(LOCALNOT2)
    LOCALNOT3=ALLTRIM(LOCALNOT3)
    localtaxas=alltrim(localtaxas)
    LOCALWEB=ALLTRIM(LOCALWEB)
    LOCALNOT5=ALLTRIM(LOCALNOT5)
    LOCNOT1=LOCALNOT1
    LOCNOT2=LOCALNOT2
    LOCNOT3=LOCALNOT3
    LOCNOT5=LOCALNOT5
    SAVE ALL LIKE LOCAL* TO ARQPAR.DAT
  return 

  Function GERDOL(JARQn,CHAVE)
    JARQ_NOM=LOCALGRAFIC+JARQn
    JARQN=JARQ_NOM
    PRI=0
    WHILE .T.
      SysRefresh()
      If FILE(JARQN)
        Erase(JARQN)
      Else
        EXIT
      EndIf
    ENDDO
    JARQ=FCREATE(JARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    SELE 6
    LINHA="col ; Venda ; 16711680"+chr(13)+chr(10)+"col ; Compra ; 0000255"+chr(13)+chr(10)
    FWRITE(JARQ,LINHA)
    CVH=ALLTRIM(CHAVE)
    CHAVE=CVH+SPACE(12-LEN(CVH))
    SEEK CHAVE
    HANTES=" "
    while CODATI=CHAVE .And. !BOF()
      SysRefresh()
      If DATE()#DTULTNEG
        EXIT
      EndIf
      SKIP
    ENDDO
    SKIP -1
    DO WHILE CODATI=CHAVE .And. !BOF()
      SysRefresh()
      TMP=VAL(LEFT(HORAULTNEG,2)+SUBSTR(HORAULTNEG,3,2))
      TMP1=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
      If TMP>TMP1
        SKIP -1
        LOOP
      EndIf
      If PROFCOMPRA=0 .And. PROFVENDA=0
        SKIP -1
        LOOP
      EndIf
      If PRI=0
        PRI=1
        MENOR=PROFVENDA
      EndIf
      If HANTES=LEFT(HORAULTNEG,3)
        SKIP -1
        LOOP
      EndIf
      HANTES=LEFT(HORAULTNEG,3)
      FAZERISTO=.F.
      If SUBSTR(HORAULTNEG,3,1)$"1340"
        FAZERISTO=.T.
      EndIf
      If !FAZERISTO
        SKIP -1
        LOOP
      EndIf
      LINHA=LEFT(HORAULTNEG,2)+":"+SUBSTR(HORAULTNEG,3,2)+" ; "+TRANSFORM(PROFVENDA,"9,999.999999")+" ; "+TRANSFORM(PROFCOMPRA,"9,999.999999")+CHR(13)+CHR(10)
      If CODATI="ZDCO        " .OR. CODATI="EDOLCM  "
        DOLCOMC=PROFCOMPRA
        DOLCOMV=PROFVENDA
      EndIf
      FWRITE(jarq,linha)
      If PROFVENDA>MAIOR
        MAIOR=PROFVENDA
      EndIf
      If PROFVENDA<MENOR
        MENOR=PROFVENDA
      EndIf
      If PROFCOMPRA>MAIOR
        MAIOR=PROFCOMPRA
      EndIf
      If PROFCOMPRA<MENOR
        MENOR=PROFCOMPRA
      EndIf
      SKIP -1
    ENDDO
    MAIOR+=.005
    MENOR-=.005
    FCLOSE(JARQ)
  return 

  Procedure GRAFMES()
    Local oText
    DECLARE V1LABEL[ADIR(LOCNOT1+"*.CSV")]
    ADIR(LOCNOT1+"*.CSV",v1label)
    FOR JI=1 TO LEN(V1LABEL)
      SysRefresh()
      FEZCSV=.T.
      lzcopyfile(LOCNOT1+V1LABEL[JI],+"TAXAS.CSV")
      oText:=TTxtFile():New(LOCNOT1+V1LABEL[JI])
      FAZGRAFMES1=.T.
      FOR WZI=1 TO oText:RecCount()
        SysRefresh()
        LINHA=oText:ReadLine()
        DATA1=CTOD(LEFT(LINHA,2)+"/"+SUBSTR(LINHA,3,2)+"/"+SUBSTR(LINHA,5,4))
        CDMOE=SUBSTR(LINHA,16,4)
        COMPRA=SUBSTR(LINHA,20,14)
        VENDA=SUBSTR(LINHA,35,14)
        WPARC=SUBSTR(LINHA,50,14)
        WPARV=SUBSTR(LINHA,65,14)
        SELE 6
        If CDMOE="DKK;" .OR. CDMOE="NOK;" .OR. CDMOE="SEK;" .OR. CDMOE="AUD;" .OR. CDMOE="CAD;" .OR. CDMOE="USD;" .OR. CDMOE="ESC;" .OR. CDMOE="NLG;" .OR. CDMOE="BEF;" .OR. CDMOE="FRF;" .OR. CDMOE="CHF;" .OR. CDMOE="JPY;" .OR. CDMOE="GBP;" .OR. CDMOE="ITL;".OR. CDMOE="DEM;" .OR. CDMOE="FMK;" .OR. CDMOE="ESP;" .OR. CDMOE="ATS;" .OR. CDMOE="EUR;"
          SELE 6
          CDMOE=CDMOE+SPACE(8)
          SEEK CDMOE+DTOS(DATA1)+"2359"
          If EOF()
            ADIREG(0)
          Else
            REGLOCK(10)
          EndIf
          REPLACE CODATI WITH CDMOE
          REPLACE HORAULTNEG WITH "235900"
          REPLACE DTULTNEG WITH DATA1
          REPLACE PROFCOMPRA WITH VAL(TROCA(COMPRA))
          REPLACE PROFVENDA WITH VAL(TROCA(VENDA))
          REPLACE PARV WITH VAL(TROCA(WPARV))
          REPLACE PARC WITH VAL(TROCA(WPARC))
          COMMIT
          UNLOCK
        EndIf
        oText:skip()
      NEXT WZI
      oText:Close()
      WWARQ=LOCNOT1+V1LABEL[JI]
      Erase &WWARQ
    NEXT JI
  return 

  Procedure GRAFMES2()
    Local oText
    DECLARE V1LABEL[ADIR(LOCNOT1+"*.CS1")]
    ADIR(LOCNOT1+"*.CS1",v1label)
    FOR JI=1 TO LEN(V1LABEL)
      SysRefresh()
      FEZCSV=.T.
      oText:=TTxtFile():New(LOCNOT1+V1LABEL[JI])
      FAZGRAFMES1=.T.
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)<>"N"
        CONTAD=1
      EndIf
      nrbol=VAL(oText:ReadLine())+1
      oText:Skip()
      oText:Skip()
      FOR WZI=3 TO CONTAD
        SysRefresh()
        WLI1=ALLTRIM(oText:ReadLine())
        WLI=ALLTRIM(SUBSTR(WLI1,2,LEN(WLI1)))+"; ; ;"
        WLI=Strtran(WLI,",",".")
        COLU=AT(";",WLI)
        ZODATI=ALLTRIM(LEFT(WLI,COLU-1))+";"
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZARV=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZARC=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZROFCOMPRA=VAL(LEFT(WLI,COLU-1))
        WLI=SUBSTR(WLI,COLU+1,150)
        COLU=AT(";",WLI)
        ZROFVENDA=VAL(LEFT(WLI,COLU-1))
        ZLTNEG=ZROFCOMPRA
        CDMOE=SUBSTR(ZODATI,4,3)+";"
        DATA1=DATE()
        If CDMOE="DKK;" .OR. CDMOE="NOK;" .OR. CDMOE="SEK;" .OR. CDMOE="AUD;" .OR. CDMOE="CAD;" .OR. CDMOE="USD;" .OR. CDMOE="ESC;" .OR. CDMOE="NLG;" .OR. CDMOE="BEF;" .OR. CDMOE="FRF;" .OR. CDMOE="CHF;" .OR. CDMOE="JPY;" .OR. CDMOE="GBP;" .OR. CDMOE="ITL;".OR. CDMOE="DEM;" .OR. CDMOE="FMK;" .OR. CDMOE="ESP;" .OR. CDMOE="ATS;" .OR. CDMOE="EUR;"
          SELE 6
          CDMOE=CDMOE+SPACE(8)
          SEEK CDMOE+DTOS(DATA1)+"2359"
          If EOF()
            ADIREG(0)
          Else
            REGLOCK(10)
          EndIf
          REPLACE CODATI WITH CDMOE
          REPLACE HORAULTNEG WITH "235900"
          REPLACE DTULTNEG WITH DATA1
          REPLACE PROFCOMPRA WITH ZROFCOMPRA
          REPLACE PROFVENDA WITH ZROFVENDA
          REPLACE PARV WITH ZARV
          REPLACE PARC WITH ZARC
          UNLOCK
          COMMIT
        EndIf
        oText:skip()
      NEXT WZI
      oText:Close()
      WWARQ=LOCNOT1+V1LABEL[JI]
      Erase &WWARQ
    NEXT JI
  return 

  Function GERMES(JARQn,CHAVE)
    JARQ_NOM=LOCALGRAFIC+JARQn
    JARQN=JARQ_NOM
    PRI=0
    WHILE .T.
      SysRefresh()
      If FILE(JARQN)
        Erase(JARQN)
      Else
        EXIT
      EndIf
    ENDDO
    JARQ=FCREATE(JARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    ARQ=FCREATE(JARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    LINHA="col ; Venda ; 16711680"+chr(13)+chr(10)+"col ; Compra ; 0000255"+chr(13)+chr(10)
    FWRITE(arq,linha)
    SELE 6
    SEEK CHAVE
    MENOR=0
    If !EOF()
      MENOR=PROFVENDA
    EndIf
    DO WHILE CODATI=CHAVE .And. MONTH(DTULTNEG)=MONTH(DATE()) .And. !EOF()
      SysRefresh()
      SKIP
    ENDDO
    SKIP -1
    DO WHILE CODATI=CHAVE .And. !BOF() .And. MONTH(DATE())=MONTH(DTULTNEG)
      SysRefresh()
      If CODATI=CHAVE
        LINHA=LEFT(DTOC(DTULTNEG),5)+" ; "+ALLTRIM(TRANSFORM(PROFVENDA,"999999.999999"))+" ; "+ALLTRIM(TRANSFORM(PROFCOMPRA,"999999.999999"))+CHR(13)+CHR(10)
        FWRITE(arq,linha)
        If PROFVENDA>MAIOR
          MAIOR=PROFVENDA
        EndIf
        If PROFVENDA<MENOR
          MENOR=PROFVENDA
        EndIf
        If PROFCOMPRA>MAIOR
          MAIOR=PROFCOMPRA
        EndIf
        If PROFCOMPRA<MENOR
          MENOR=PROFCOMPRA
        EndIf
      EndIf
      SKIP -1
    ENDDO
    FCLOSE(ARQ)
    MAIOR+=.005
    MENOR-=.005
  return 

  Procedure TROCA(VARI)
    VARI1=""
    FOR I=1 TO LEN(VARI)
      SysRefresh()
      SU=SUBSTR(VARI,I,1)
      If SU="@"
        SU=" "
      EndIf
      If SU=","
        SU="."
      EndIf
      If SU="+"
        SU=""
      EndIf
      VARI1=VARI1+SU
    next
  return  VARI1

  Procedure ABREARQS()
    SELE 12
    If !USEREDE(LOCALCART+"USER.DBF",.F.,10,"USUARIO")
      return 
    EndIf
    If LOCALQTIPO<>"W"
      return 
    EndIf
    If WSEN="C"
      SELE 29
      If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
        return 
      Else
        arqind=localcart+"empresa1.ind"
        set index to &arqind
      EndIf
      return 
    EndIf
    SELE 1
    If !USEREDE(LOCNOT3+"NOTAEX.DBF",.F.,10,"NOTAEX")
      return 
    Else
      ARQ=LOCNOT3+"NOTAEX.KEY"
      ARQS1=LOCNOT3+"NOTAEX1.KEY"
      SET INDEX TO &ARQ,&ARQS1
    EndIf
    SELE 6
    If !USEREDE(LOCNOT3+"TAXAS.DBF",.F.,10,"TAXAS")
      return 
    EndIf
    SELE 7
    If !USEREDE(LOCNOT3+"TROCA.DBF",.F.,10,"TROCA")
      return 
    EndIf
    SELE 9
    If !USEREDE(LOCNOT3+"PESQ.DBF",.F.,10,"PESQ")
      return 
    EndIf
    SELE 11
    If !USEREDE(LOCNOT3+"LOG.DBF",.F.,10,"LOG")
      return 
    EndIf
    SELE 13
    If !USEREDE(LOCNOT3+"BANCOML.DBF",.F.,10,"CLIEUS")
      return 
    EndIf
    SELE 14
    If !USEREDE(LOCNOT3+"ANOTAC.DBF",.F.,10,"ANOTAC")
      return 
    EndIf
    SELE 22
    If !USEREDE(LOCALWEB+"EMPWINCC.DBF",.F.,10,"EMPWINCC")
      return 
    EndIf
    SELE 24
    If !USEREDE(LOCALWEB+"USUARIO.DBF",.F.,10,"USUAWIN")
      return 
    EndIf
    SELE 30
    If !USEREDE(LOCNOT3+"BOLSANET.DBF",.F.,10,"BOLSANET","S")
      return 
    EndIf
  return 

  Procedure GRAFDIA1()
    If !ATUOUCOTAS
      return 
    EndIf
    If FILE(LOCALGNEW+"\DIA\GRAFDIA.TXT")
      FErase(LOCALGNEW+"\DIA\GRAFDIA.TXT" )
    EndIf
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLCOM.TXT","EXUSD")
    GRAFICO="Dólar Comercial em R$ (PTAX)  ; DOLCOM.TXT ; "+TRANSFORM(MENOR,"9999.9999")+" ; "+TRANSFORM(MAIOR,"9999.9999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLCOM1.TXT","EDOLCM")
    GRAFICO=GRAFICO+"Dólar Comercial (Mercado) em R$ ; DOLCOM1.TXT ; "+TRANSFORM(MENOR,"9999.9999")+" ; "+TRANSFORM(MAIOR,"9999.9999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLTUR.TXT","EDOLTR")
    GRAFICO=GRAFICO+"Dólar Turismo em R$; DOLTUR.TXT ; "+TRANSFORM(MENOR,"9999.9999")+" ; "+TRANSFORM(MAIOR,"9999.9999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLPAP.TXT","EDOLPR")
    GRAFICO=GRAFICO+"Dólar Papel em R$; DOLPAP.TXT ; "+TRANSFORM(MENOR,"9999.9999")+" ; "+TRANSFORM(MAIOR,"9999.9999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLDKK.TXT","EXDKK")
    GRAFICO=GRAFICO+"Coroa Dinamarquesa em R$ ; DOLDKK.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLJPY.TXT","EXJPY")
    GRAFICO=GRAFICO+"Iene  em R$ ; DOLJPY.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLCHF.TXT","EXCHF")
    GRAFICO=GRAFICO+"Franco Suiço em R$ ; DOLCHF.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLCAD.TXT","EXCAD")
    GRAFICO=GRAFICO+"Dólar Canadense em R$ ; DOLCAD.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLGBP.TXT","EXGBP")
    GRAFICO=GRAFICO+"Libra Esterlina em R$; DOLGBP.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLEUR.TXT","EXEUR")
    GRAFICO=GRAFICO+"Euro em R$ ; DOLEUR.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLNOK.TXT","EXNOK")
    GRAFICO=GRAFICO+"Coroa Norueguesa em R$; DOLNOK.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLSEK.TXT","EXSEK")
    GRAFICO=GRAFICO+"Coroa Sueca em R$; DOLSEK.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERDOL1(LOCALGNEW+"\DIA\DOLAUD.TXT","EXAUD")
    GRAFICO=GRAFICO+"Dólar Australiano em R$; DOLAUD.TXT ; "+TRANSFORM(MENOR," 99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    ARQ_NOM=LOCALGNEW+"\DIA\GRAFDIA.TXT"
    ARQ=FCREATE(ARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    FWRITE(arq,GRAFICO)
    FCLOSE(ARQ)

  Function GERDOL1(JARQn,CHAVE)
    JARQ_NOM=JARQn
    JARQN=JARQ_NOM
    PRI=0
    WHILE .T.
      SysRefresh()
      If FILE(JARQN)
        Erase(JARQN)
      Else
        EXIT
      EndIf
    ENDDO
    JARQ=FCREATE(JARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    SELE 6
    LINHA="Hora ; Cotação"+CHR(13)+CHR(10)+"col ; Venda ; 16711680"+CHR(13)+CHR(10)+"col ; Compra ; 0000255"+CHR(13)+CHR(10)
    FWRITE(JARQ,LINHA)
    CVH=ALLTRIM(CHAVE)
    CHAVE=CVH+SPACE(12-LEN(CVH))
    SEEK CHAVE
    HANTES=" "
    while CODATI=CHAVE .And. !BOF()
      SysRefresh()
      If DATE()#DTULTNEG
        EXIT
      EndIf
      SKIP
    ENDDO
    SKIP -1
    DO WHILE CODATI=CHAVE .And. !BOF()
      SysRefresh()
      If PROFCOMPRA=0 .And. PROFVENDA=0
        SKIP -1
        LOOP
      EndIf
      If PRI=0
        PRI=1
        MENOR=PROFVENDA
      EndIf
      LINHA=LEFT(HORAULTNEG,5)+" ; "+TRANSFORM(PROFVENDA,"9,999.999999")+" ; "+TRANSFORM(PROFCOMPRA,"9,999.999999")+CHR(13)+CHR(10)
      If CODATI="EUSD        "
        DOLCOMC=PROFCOMPRA
        DOLCOMV=PROFVENDA
      EndIf
      FWRITE(jarq,linha)
      If PROFVENDA>MAIOR
        MAIOR=PROFVENDA
      EndIf
      If PROFVENDA<MENOR
        MENOR=PROFVENDA
      EndIf
      If PROFCOMPRA>MAIOR
        MAIOR=PROFCOMPRA
      EndIf
      If PROFCOMPRA<MENOR
        MENOR=PROFCOMPRA
      EndIf
      SKIP -1
    ENDDO
    MAIOR+=.005
    MENOR-=.005
    FCLOSE(JARQ)
  return 

  Procedure GRAFMES1()
    Local oText
    If !FAZGRAFMES1
      return 
    EndIf
    If FILE(LOCALGNEW+"\MES\GRAFMES.TXT")
      FErase(LOCALGNEW+"\MES\GRAFMES.TXT" )
    EndIf
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\DOLMCOM.TXT","USD;")
    GRAFICO="Dólar Comercial - 220 - USD ; DOLMCOM.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\DKK.TXT","DKK;")
    GRAFICO=GRAFICO+"Coroa Dinamarquesa - 055 - DKK ; DKK.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\NOK.TXT","NOK;")
    GRAFICO=GRAFICO+"Coroa Noroeguesa - 065 - NOK ; NOK.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\SEK.TXT","SEK;")
    GRAFICO=GRAFICO+"Coroa Sueca - 070 - SEK ; SEK.TXT ;"+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\AUD.TXT","AUD;")
    GRAFICO=GRAFICO+"Dólar Australiano - 150 - AUD ; AUD.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\CAD.TXT","CAD;")
    GRAFICO=GRAFICO+"Dólar Canadense - 165 - CAD ; CAD.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\CHF.TXT","CHF;")
    GRAFICO=GRAFICO+"Franco Suiço - 425 - CHF ; CHF.TXT ;"+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\JPY.TXT","JPY;")
    GRAFICO=GRAFICO+"Iene - 470 - JPY ; JPY.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\GBP.TXT","GBP;")
    GRAFICO=GRAFICO+"Libra Esterlina - 540 - GBP ; GBP.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    MENOR:=MAIOR:=0
    GERMES1(LOCALGNEW+"\MES\EUR.TXT","EUR;")
    GRAFICO=GRAFICO+"Euro - 978 - EUR ; EUR.TXT ; "+TRANSFORM(MENOR,"99999.9999999")+" ; "+TRANSFORM(MAIOR," 99999.9999999")+CHR(13)+CHR(10)
    ARQ_NOM=LOCALGNEW+"\MES\GRAFMES.TXT"
    ARQ=FCREATE(ARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    FWRITE(arq,GRAFICO)
    FCLOSE(ARQ)
  return 

  Function GERMES1(JARQN,CHAVE)
    JARQ_NOM=JARQN
    JARQN=JARQ_NOM
    PRI=0
    WHILE .T.
      SysRefresh()
      If FILE(JARQN)
        Erase(JARQN)
      Else
        EXIT
      EndIf
    ENDDO
    JARQ=FCREATE(JARQ_NOM,0)
    If FERROR()#0
      return 
    EndIf
    LINHA="Data ; Cotação"+CHR(13)+CHR(10)+"col ; Venda ; 16711680"+CHR(13)+CHR(10)+"col ; Compra ; 0000255"+CHR(13)+CHR(10)
    FWRITE(JARQ,linha)
    SELE 6
    SEEK CHAVE
    MENOR=0
    If !EOF()
      MENOR=PROFVENDA
    EndIf
    DO WHILE CODATI=CHAVE .And. MONTH(DTULTNEG)=MONTH(DATE()) .And. !EOF()
      SysRefresh()
      SKIP
    ENDDO
    SKIP -1
    DO WHILE CODATI=CHAVE .And. !BOF() .And. MONTH(DATE())=MONTH(DTULTNEG)
      SysRefresh()
      If CODATI=CHAVE
        LINHA=LEFT(DTOC(DTULTNEG),5)+" ; "+ALLTRIM(TRANSFORM(PROFVENDA,"999999.999999"))+" ; "+ALLTRIM(TRANSFORM(PROFCOMPRA,"999999.999999"))+CHR(13)+CHR(10)
        FWRITE(JARQ,linha)
        If PROFVENDA>MAIOR
          MAIOR=PROFVENDA
        EndIf
        If PROFVENDA<MENOR
          MENOR=PROFVENDA
        EndIf
        If PROFCOMPRA>MAIOR
          MAIOR=PROFCOMPRA
        EndIf
        If PROFCOMPRA<MENOR
          MENOR=PROFCOMPRA
        EndIf
      EndIf
      SKIP -1
    ENDDO
    FCLOSE(JARQ)
    MAIOR+=.005
    MENOR-=.005
  return 

  Procedure TUACOTA1()
    If !FILE(alltrim(LOCNOT1)+"DOLAR.TXT")
      return 
    EndIf
    YARQ=ALLTRIM(ALLTRIM(LOCNOT1)+"DOLAR.TXT")
    oText:=TTxtfile():NEW(YARQ)
    CONTAD=oText:RecCount()
    If VALTYPE(CONTAD)<>"N"
      CONTAD=1
    EndIf
    oText:Skip()
    FOR WZI=1 TO CONTAD
      SysRefresh()
      WLI1=ALLTRIM(oText:ReadLine())+"; ; ; ; ;"
      WLI=Strtran(WLI1,",",".")
      COLU=AT(";",WLI)
      ZODATI=LEFT(WLI,COLU-1)
      ZODATI=ALLTRIM("E"+ZODATI)
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      NOMEPAPEL=LEFT(WLI,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      ZARV=0
      ZARC=0
      ZROFCOMPRA=VAL(LEFT(WLI,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      ZROFVENDA=VAL(LEFT(WLI,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      ZSCILA=VAL(LEFT(WLI,COLU-1))
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      ZORAULTNEG=LEFT(WLI,COLU-1)
      WLI=SUBSTR(WLI,COLU+1,150)
      COLU=AT(";",WLI)
      ZTULTNEG=CTOD(LEFT(WLI,COLU-1))
      ZLTNEG=ZROFCOMPRA
      ZRMAXIMO=0
      ZRMINIMO=0
      ZRABERTURA=0
      ZRMEDIO=0
      ZRFECANT=0
      ZRNEG=0
      ZTDULTNEG=0
      ZTDTOTAL=0
      ZOLUME=0
      ZOLUMETOT=0
      ZUOFCOMPRA=0
      ZUOFVENDA=0
      ZODATI1=ALLTRIM(ZODATI)
      ZODATI=ZODATI1+SPACE(12-LEN(ZODATI1))
      ZODATI1=ZODATI
      ZRAFICO="S"
      SELE 6
      If ZRAFICO="S"
        SEEK ZODATI+DTOS(ZTULTNEG)+LEFT(ZORAULTNEG,5)
        If EOF()
          ADIREG(0)
        Else
          REGLOCK(10)
        EndIf
        PUTREC()
        COMMIT
      EndIf
      If PRIMEIRA2=0
        GO BOTT
        PRIMEIRA2=RECNO()
      EndIf
      oText:Skip()
    NEXT WZI
    oText:Close()
    Erase &YARQ
  return 

  Procedure COMERCIAL()
    If !USUARIO(13)
      return 
    EndIf
    SELE 13
    set order to 2
    Browse(12,"Manutenção Comercial","Comercial", { || BUSCACOM("oLbx") },{ ||NCADCOM("I") },{ ||NCADCOM("A") })
  return 

  Procedure NCADCOM(ZS)
    PUBLIC oDlg1,oBrush
    PUBLIC F
    F=ZS
    SELE 13
    set order to 1
    EREG=RECNO()
    DECLARE _VAR1[FCOUNT()]
    AFIELDS(_VAR1)
    If F="I"
      GO BOTT
      SKIP
    EndIf
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t
    RELEASE _VAR1
    MFONTES()
    A="N"
    DEFINE DIALOG oDlg1 FROM 0,0 TO 28,73 TITLE "Dados Comercial" BRUSH oBrushTela STYLE WS_CAPTION
    @   0.44,0.5 say "" box raised COLOR CLR_BLACK,nRGB(248,248,248) size 279,178 OF oDlg1
    @   1.74,1  SAY "Código:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,1  SAY "Nome:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   3.48,1  SAY "Endereço:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,1  SAY "Cidade/UF:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   4.35,28 SAY "Cep:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   5.22,1  SAY "CNPJ/CPF:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,1  SAY "CONTATO" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.96,1  SAY "Nome" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,28 SAY "FONE" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   6.09,38 SAY "FAX" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   7.83,1  SAY "Operador:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   8.70,1  SAY "Email(Cambio)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   9.57,1  SAY "Email(Bolsa)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  10.44,1  say "Email(BMF)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @  11.31,1  say "Email (Contab)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

    If F="I"
      @ 02,7 GET ozodigo VAR ZODIGO PICTURE "999999" VALID (VSTCON(ZODIGO)) FONT oFontget SIZE 30,11 RIGHT
    Else
      @ 02,7 GET ZODIGO PICTURE "999999"  FONT oFontget SIZE 30,11 RIGHT NO MODIFY
    EndIf

    @ 1,1   CHECKBOX ZBOLSA  PROMPT "Bolsa" SIZE 50,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,6   CHECKBOX ZCAMBIO PROMPT "Cambio" SIZE 55,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,12  CHECKBOX ZBMF    PROMPT "BMF" SIZE 50,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,17  CHECKBOX ZCONTAB PROMPT "Contabil" SIZE 60,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,23  CHECKBOX ZOUTROS PROMPT "Outros" SIZE 50,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,28  CHECKBOX ZSEC    PROMPT "Asec" SIZE 50,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 1,33  CHECKBOX ZNATIVO PROMPT "Inativo" SIZE 60,10 FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

    @ 3,7  GET ZOME PICTURE "@!" FONT oFontget SIZE 220,11
    @ 4,7  GET ZNDERECO PICTURE "@!" FONT oFontget SIZE 220,11
    @ 5,7  GET ZIDADE PICTURE "@!" FONT oFontget SIZE 90,11
    @ 5,19 GET ZSTADO PICTURE "!!" FONT oFontget SIZE 12,11
    @ 5,24 GET ZEP PICTURE "@R 99999-999" FONT oFontget SIZE 35,11
    @ 6,7  GET ZGCCPF PICTURE "@!" FONT oFontget SIZE 70,11
    @ 8,7  GET ZONTATO PICTURE "@!" FONT oFontget SIZE 80,11
    @ 8,18 GET ZELEFONE PICTURE "@!" FONT oFontget SIZE 60,11
    @ 8,26 GET ZAX PICTURE "@!" FONT oFontget SIZE 68,11
    @ 10,7    GET ZMAIL FONT oFontget SIZE 220,11
    @ 11,7 GET ZMBAILBM FONT oFontget SIZE 220,11
    @ 12,7 GET ZMAILBOV FONT oFontget SIZE 220,11
    @ 13,7 get zmailcont FONT oFontget SIZE 220,11

    @ 192,100 SBUTTON PIXELS TEXT ON_CENTER ADJUST  FONT oFontbut PROMPT "&Confirma" RESOURCE "OFF","ON","DISABLE" NOBOX  OF odlg1 SIZE 40,12 ACTION (IIF(F="I",ADIREG(0,2),REGLOCK(10)),PUTREC(2),oDlg1:END())
    @ 192,150 SBUTTON PIXELS TEXT ON_CENTER ADJUST  FONT oFontbut PROMPT "C&ancela" RESOURCE "OFF","ON","DISABLE" NOBOX  OF odlg1 SIZE 40,12 ACTION (oDlg1:END())
    oDlg1:lhelpicon:=.f.
    ACTIVATE DIALOG oDlg1 CENTERED
    SELE 13
    set order to 2
  return 

  Function BUSCACOM(OLBX)
    TEXTO="Código/Nome, +CNPJ/CPF"
    VARIAVEL=SPACE(40)
    MASCARA="@!"
    REG=RECNO()
    DO WHILE .T.
      SysRefresh()
      If !BPROCURA()
        EXIT
      EndIf
      CHAVE=VARIAVEL
      If EMPTY(CHAVE)
        EXIT
      EndIf
      If LEFT(CHAVE,1)="+"
        If LEN(CHAVE)>2
          CHAVE=ALLTRIM(RIGHT(CHAVE,LEN(CHAVE)-1))
        EndIf
        LOCATE FOR CGCCPF=CHAVE
      ElseIf VAL(CHAVE)>0
        SET ORDER TO 1
        SEEK VAL(chave)
      Else
        SET ORDER TO 2
        SET SOFTSEEK ON
        SEEK ALLTRIM(CHAVE)
        SET SOFTSEEK OFF
      EndIf
      If EOF() .OR. DELE()
        MSGINFO("Não Encontrado !","Alerta")
        GO REG
      Else
        return 
      EndIf
    ENDDO
  return 

  Procedure RELATCOM()
    DEFINE DIALOG oDlgg FROM 1,1 TO 22,70 TITLE "Comercial" BRUSH oBrushTela STYLE WS_CAPTION
    @ 0.4,24 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER OF oDlgg SIZE 80,110
    @   0.44,0.5 SAY " " BOX RAISED SIZE 185,95
    WDE=M->DAT_HOJE
    WATE=M->DAT_HOJE+7
    wclide:=0
    wcliate:=999999
    m->quant=1
    PREVIA=0
    @   0.87,1 say "Período de/Até" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @   2.61,1 say "Cliente de/Até" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
    @ 01,10 GET WDE FONT oFontget SIZE 45,11
    @ 01,16 GET WATE FONT oFontget SIZE 45,11
    @ 03,10 GET owclide VAR wclide PICTURE "999999" FONT oFontget SIZE 40,11 RIGHT
    @ 03,16 GET owcliate VAR wcliate PICTURE "999999" FONT oFontget SIZE 40,11 RIGHT
    @ 4.7,1 COMBOBOX oCbx VAR M->QUANT ITEMS {"Compromissos(A efetuar)","Compromissos(Incluidos)"} OF oDlgg UPDATE font OFONTGET SIZE 100,90
    @ 07,01 SBUTTON PROMPT "&Imprime" RESOURCE "IMPRIME" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) ACTION (PREVIA:=1,oDlgg:end()) FONT oFontbut SIZE 40,35
    @ 07,10 SBUTTON oimp PROMPT "&Visualiza" RESOURCE "PROC" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) ACTION (PREVIA:=2,oDlgg:End()) FONT oFontbut SIZE 40,35
    @ 07,19 SBUTTON PROMPT "&Cancela" RESOURCE "SAIR2" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) ACTION (PREVIA:=0,oDlgg:End()) FONT oFontbut SIZE 40,35
    oDlgg:lhelpicon:=.f.
    ACTIVATE DIALOG oDlgg CENTERED
    If PREVIA=0
      return 
    EndIf
    If PREVIA=1
      PREV=.F.
    Else
      PREV=.T.
    EndIf
    SELE 14
    If M->QUANT=1
      SET ORDER TO 2
      SET SOFTSEEK ON
      SEEK DTOS(WDE)
      SET SOFTSEEK OFF
      ctitle="Follow-up Período de: "+DTOC(WDE)+" a "+DTOC(WATE)
    Else
      ctitle="Compromissos Incluidos no Período de: "+DTOC(WDE)+" a "+DTOC(WATE)
      SET ORDER TO 1
      GO TOP
    EndIf

    DEFINE FONT oFont1 NAME "VERDANA" BOLD SIZE 0,-8
    DEFINE FONT oFont2 NAME "VERDANA" SIZE 0,-8
    DEFINE FONT oFont3 NAME "VERDANA" BOLD SIZE 0,-9

    If PREV
      REPORT oRpt TITLE  " "," "," "," ","  ","Exatus.net",CTITLE FONT oFont1,oFont2,oFont3;
        PREVIEW;
        FOOTER "Data:"+dtoc(dat_hoje)+"   "+time()+" Folha:" + Strzero( oRpt:nPage, 3 )," ",'"Proibida sua reprodução, difusão e conservação (apropriação)."',"Usuário: "+A02NOME;
        CAPTION ctitle;

      Else

      REPORT oRpt TITLE  " "," "," "," ","  ","Exatus.net",CTITLE FONT oFont1,oFont2,oFont3;
        PREVIEW;
        FOOTER "Data:"+dtoc(dat_hoje)+"   "+time()+" Folha:" + Strzero( oRpt:nPage, 3 )," ",'"Proibida sua reprodução, difusão e conservação (apropriação)."',"Usuário: "+A02NOME;
        CAPTION ctitle;

      EndIf

      COLUMN TITLE "Cliente","Usuário" DATA LEFT(ANOTAC->CHAVE,6)+" - "+LEFT(PESQ(VAL(LEFT(ANOTAC->CHAVE,6)),"13",1,"NOME"),30),ANOTAC->USUARIO SIZE 33 FONT 2 GRID
      COLUMN TITLE "Follow-up","Inclusão"  DATA DTOC(ANOTAC->AVISAREM),DTOC(ANOTAC->DATA) FONT 2 SIZE 10 GRID
      COLUMN TITLE "Descrição","Resumo da Ocorrência" DATA ANOTAC->HISTO,LEFT(ANOTAC->ANOTA,100) FONT 2 GRID
      oRpt:CELLVIEW()
      If oRpt:lCreated
        oRpt:oTitle:afont[6]:= { || 3 }
        oRpt:oTitle:afont[7]:= { || 3 }
      EndIf

      ENDREPORT

      If M->QUANT=1
        ACTIVATE REPORT oRpT FOR (AVISAREM>=WDE .And. AVISAREM<=WATE .And. SUBSTR(ANOTAC->CHAVE,16,1)="-") .And. VAL(LEFT(ANOTAC->CHAVE,6))>=WCLIDE .And. VAL(LEFT(ANOTAC->CHAVE,6))<=WCLIATE
      Else
        ACTIVATE REPORT oRpT FOR (DATA>=WDE .And. DATA<=WATE .And. SUBSTR(ANOTAC->CHAVE,16,1)="-") .And. VAL(LEFT(ANOTAC->CHAVE,6))>=WCLIDE .And. VAL(LEFT(ANOTAC->CHAVE,6))<=WCLIATE
      EndIf

      oFont1:End()
      oFont2:End()
      oFont3:End()
    return 

    Procedure BPROCURA()
      Local oD1lg,oFont
      PUBLIC RETORNO
      RETORNO=.T.
      MFONTES()
      DEFINE DIALOG oD1lg FROM 1, 1 TO 13.5 , 55 TITLE "Pesquisa" BRUSH oBrushTela STYLE WS_CAPTION
      @   0.44,7.5 say "" box raised COLOR CLR_BLACK,nRGB(248,248,248) size 135,50 OF oD1lg
      @   0.61,11 say Texto FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,15
      @ 3.0,09.5 get VARIAVEL picture mascara  FONT oFontget size 75,10 OF oD1lg
      @ 0.7,0.5 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER size 46,65 OF oD1lg
      @ 70,80 SBUTTON PIXELS PROMPT "C&onfirma"  TEXT ON_CENTER RESOURCE "OFF","ON","DISABLE" NOBOX  ADJUST OF oD1lg SIZE 40,12 ACTION oD1lg:end() FONT oFontbut DEFAULT
      @ 70,130 SBUTTON PIXELS PROMPT "&Cancela"  TEXT ON_CENTER RESOURCE "OFF","ON","DISABLE" NOBOX  ADJUST OF oD1lg SIZE 40,12 ACTION (VERM(34),oD1lg:End()) FONT oFontbut
      oD1lg:lhelpicon:=.f.
      ACTIVATE DIALOG oD1lg CENTERED
    return  RETORNO

    Procedure OBSERVA(WCHAVE,TITULO)
      Local oDlgW,oFontW,oBrushcer,oFontcer
      DEFINE FONT oFontcer NAME "Arial" SIZE 6, 15 BOLD
      W1CHAVE=TRIM(WCHAVE)
      SET DELE ON
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
        ARQOBS=LOCALNOT1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
        ARQOBS1:=ARQOBS+".DA1"
        ARQOBS2:=ARQOBS+".DA2"
        ARQOBS3:=ARQOBS+".KE3"
        ARQOBS4:=ARQOBS+".FPT"
        If !FILE(ARQOBS1)
          EXIT
        EndIf
      ENDDO
      MAT_STRU={}
      AADD(MAT_STRU,{"CHAVE","C",30,0})
      AADD(MAT_STRU,{"DATA","D",8,0})
      AADD(MAT_STRU,{"USUARIO","C",30,0})
      AADD(MAT_STRU,{"AVISAREM","D",8,0})
      AADD(MAT_STRU,{"HORA","C",10,0})
      AADD(MAT_STRU,{"HISTO","C",60,0})
      AADD(MAT_STRU,{"ANOTA","M",10,0})
      AADD(MAT_STRU,{"REGISTRO","N",10,0})
      DBCREATE(ARQOBS2,MAT_STRU)
      SELE 20
      If .NOT. USEREDE(ARQOBS2,.T.,10,"WOBS")
        MSGSTOP("Não foi possível abrir os arquivos",WSIST)
        SELE 13
        return 
      EndIf
      INDEX ON DTOS(DATA)+HORA TAG HOR TO &ARQOBS1
      SELE 14
      SEEK W1CHAVE
      DO WHILE !EOF() .And. VAL(LEFT(CHAVE,6))=MWCLI
        SysRefresh()
        If CHAVE#W1CHAVE
          SKIP
          LOOP
        EndIf
        DECLARE _var1 [FCOUNT()]
        AFIELDS(_var1)
        FOR t = 1 TO FCOUNT()
          SysRefresh()
          _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
          _var5 = _var1 [t]
          &_var4 = &_var5
        NEXT t
        ZEGISTRO=RECNO()
        SELE 20
        ADIREG(0)
        PUTREC(0,0)
        SELE 14
        SKIP
      ENDDO
      SELE 20
      GO TOP
      FINALIZA=.F.
      MFONTES()
      If SUBSTR(W1CHAVE,16,1)="-"
        TITULO=" "+STRZERO(CLIEUS->CODIGO,6)+"-"+CLIEUS->NOME
      Else
        TITULO=" "+"Cliente:"+STRZERO(MWCLI,6)+"-"+PESQ(MWCLI,"1",1,"NOME")
      EndIf
      DEFINE DIALOG oDlgw FROM 1,2 TO 27,92 TITLE TITULO FONT ofontcer BRUSH oBrushTela STYLE WS_CAPTION
      @   0.5,35 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 143,182
      wscn="N"
      @ 1,27 GET owsanoa VAR WOBS->ANOTA FONT oFontget OF oDlgw  MEMO SIZE 130,170 WHEN WSCN="S"

      @ 1,0.5 LISTBOX oLbxcer1 FIELDS TRANSFORM(WOBS->DATA,"@D"),TRANSFORM(WOBS->AVISAREM,"@D"),WOBS->HISTO HEADERS "Data","Follow-up em","Descrição" SIZE 200, 137  OF oDlgw ON CHANGE (owsanoa:refresh())

      @ 155, 10 SBUTTON PIXELS TOOLTIP "Inclui"   PROMPT "&Inclui"  TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) FONT oFontbut RESOURCE "INCLUI"  OF oDlgw SIZE 35, 35 ACTION (NCADANO("I"),olbxcer1:refresh(),oLbxcer1:Setfocus())
      @ 155, 50 SBUTTON PIXELS TOOLTIP "Altera"   PROMPT "&Altera"  TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) FONT oFontbut RESOURCE "ALTERA"  OF oDlgw SIZE 35, 35 ACTION (NCADANO("A"),olbxCer1:Refresh(),oLbxcer1:Setfocus())
      @ 155, 90 SBUTTON PIXELS TOOLTIP "Exclui"   PROMPT "&Exclui"  TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) FONT oFontbut RESOURCE "EXCLUI"  OF oDlgw SIZE 35, 35 ACTION (DELANO(),oLbxcer1:refresh(),oLbxcer1:Setfocus())
      @ 155,130 SBUTTON PIXELS TOOLTIP "Imprimir" PROMPT "I&mprime" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) FONT oFontbut RESOURCE "IMPRIME" OF oDlgw SIZE 35, 35 ACTION (WREPORT(9,olbxcer1,TITULO),olbxcer1:refresh(),oLbxcer1:Setfocus())
      @ 155,170 SBUTTON PIXELS TOOLTIP "Retorna"  PROMPT "&Sair"    TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) FONT oFontbut RESOURCE "SAIR2"   OF oDlgw SIZE 35, 35 ACTION (GRAVAANO(),IIF(FINALIZA,oDlgw:end(),""))
      oDlgw:lhelpicon:=.f.
      ACTIVATE DIALOG oDlgw CENTERED

      SELE 20
      USE
      Erase &ARQOBS1
      Erase &ARQOBS2
      Erase &ARQOBS3
      Erase &ARQOBS4
      SELE 13
    return 

    Function DELANO()
      If MSGYESNO( "Exclui este Registro ?",WSIST)
        If !PORQUEEXCLUI()
          return 
        EndIf
        If REGLOCK(10)
          DELE
        Else
          MSGINFO("Registro Bloqueado por outro Usuário ",WSIST)
        EndIf
      EndIf
    return  NIL

    Procedure NCADANO(FZ1)
      PUBLIC oWDlg,oWBrush,oFontcombo,oFontw
      DEFINE FONT oFontcombo  NAME "Arial" SIZE 6,12 BOLD
      SELE 20
      If FZ1="I"
        DECLARE _var1 [FCOUNT()],_var2 [FCOUNT()],_var3 [FCOUNT()]
        AFIELDS(_var1,_var2,_var3)
        FOR t = 1 TO FCOUNT()
          SysRefresh()
          _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
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
          EndIf
        NEXT t
        ZSUARIO=A02NOME
        ZATA=M->DAT_HOJE
        ZORA=TIME()
      Else
        If !REGLOCK()
          MSGINFO("Registro sendo alterado por outro Usuário","Tente Mais Tarde")
          return 
        Else
          DECLARE _var1 [FCOUNT()]
          AFIELDS(_var1)
          FOR t = 1 TO FCOUNT()
            SysRefresh()
            _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
            _var5 = _var1 [t]
            &_var4 = &_var5
          NEXT t
        EndIf
      EndIf
      ZHAVE=W1CHAVE
      ZVISAREM=M->DAT_HOJE+30
      MFONTES()
      DEFINE DIALOG oWDlg FROM 0,0 TO 30,81 TITLE "Anotações/Ocorrencias" BRUSH oBrushTela STYLE WS_CAPTION
      @ 0.44,0.5 SAY " " BOX RAISED FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 312,200
      @ 0.87,01 SAY "Data: " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 01,07 GET ZATA FONT oFontget SIZE 45,11 NO MODIFY
      @ 0.87,23 SAY "Hora: " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 01,20 GET ZORA FONT oFontget SIZE 50,11 NO MODIFY
      @ 1.74,01 SAY "Usuário:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 02,07 GET ZSUARIO FONT oFontget SIZE 120,11 NO MODIFY
      @ 2.61,01 SAY "Follow-up em:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 3.48,01 say "Descrição:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 4.35,21 SAY "Anotações/Ocorrencia" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @ 03,07 GET ZVISAREM FONT oFontget SIZE 45,11 OF oWdlg
      @ 04,07 GET ZISTO FONT oFontget SIZE 180,11 OF oWdlg
      @ 06,01 GET ZNOTA FONT oFontget OF oWdlg MEMO SIZE 300,120

      @ 212,110 SBUTTON PIXELS PROMPT "C&onfirma"  TEXT ON_CENTER RESOURCE "OFF","ON","DISABLE" NOBOX ADJUST OF oWDlg SIZE 40,12 ACTION (IIF(FZ1="I",ADIREG(0),REGLOCK(10)),PUTREC(0,0),oWdlg:END()) FONT oFontbut
      @ 212,175 SBUTTON PIXELS PROMPT "&Cancela"   TEXT ON_CENTER RESOURCE "OFF","ON","DISABLE" NOBOX ADJUST  OF OWdLG SIZE 40,12 ACTION (oWDlg:End()) FONT oFontbut
      owDlg:lhelpicon:=.f.
      ACTIVATE DIALOG oWDlg CENTERED
      COMMIT
      SELE 13
      REGLOCK(10)
      REPLACE ULTCONT WITH M->DAT_HOJE
      REPLACE PROXCONT WITH M->DAT_HOJE+30
      SELE 20
    return 

    Function GRAVAANO()
      FINALIZA=.T.
      SELE 20
      SET DELE OFF
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        DECLARE _var1 [FCOUNT()]
        AFIELDS(_var1)
        FOR t = 1 TO FCOUNT()
          SysRefresh()
          _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
          _var5 = _var1 [t]
          &_var4 = &_var5
        NEXT t
        APAGA=.F.
        If DELE()
          APAGA=.T.
        EndIf
        SELE 14
        If ZEGISTRO=0
          ADIREG(0)
        Else
          GO ZEGISTRO
        EndIf
        REGLOCK(10)
        If APAGA
          DELE
          SELE 20
          SKIP
          LOOP
        EndIf
        CHAVE=W1CHAVE
        PUTREC(0,0)
        SELE 20
        SKIP
      ENDDO
      SET DELE ON
      SELE 13
    return  .T.

    Procedure EMPVAL(QTP1)
      If OP7#"S"
        MSGINFO("Aguarde término das rotinas","")
        return 
      EndIf
      qtp=qtp1
      If !USUARIO(15)
        return 
      EndIf
      If QTP=2
        SELE 22
      ElseIf QTP=3
        SELE 30
      EndIf

      If QTP=2
        Browse(14,"Corretoras CambioNet/DOCTOSNET","CambioNet/DOCTOSNET", { || buscacor() },{ ||NCADcorr("I") },{ ||NCADcorr("A") })
      ElseIf QTP=3
        Browse(14,"Corretoras BolsasNet (Novo)","BolsasNet", { || buscacor() },{ ||NCADcorr("I") },{ ||NCADcorr("A") })
      EndIf
    return 

    Procedure BUSCAcor()
      TEXTO="Corretora"
      VARIAVEL=SPACE(8)
      MASCARA="@!"
      REG=RECNO()
      DO WHILE .T.
        SysRefresh()
        If !PROCURA()
          EXIT
        EndIf
        CHAVE=variavel
        If qtp<>3
          LOCATE FOR CORRETORA=VAL(CHAVE)
        ElseIf QTP=3
          LOCATE FOR CORRETORA=CHAVE
        EndIf
        If EOF() .OR. DELE()
          MSGALERT("Não Encontrado !","Alerta")
          GO REG
        Else
          EXIT
        EndIf
      ENDDO
    return 

    Procedure NCADcorr(FS)
      Local oDlg,oBrush
      PUBLIC F
      F=FS
      DECLARE _VAR1[FCOUNT()]
      AFIELDS(_VAR1)
      If F="I"
        GO BOTT
        SKIP
      EndIf
      FOR t = 1 TO FCOUNT()
        SysRefresh()
        _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
        _var5 = _var1 [t]
        &_var4 = &_var5
      NEXT t
      RELEASE _VAR1
      MFONTES()
      If QTP=2
        DEFINE DIALOG oDlg FROM 0,0 TO 18,58 TITLE "Usuários do CambioNet" BRUSH oBrushTela STYLE WS_CAPTION
        @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 222,110
        @   0.87,02 SAY "Corretora" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   0.87,20 SAY "Cod.Bacen" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   1.74,02 SAY "Nome" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,02 SAY "Email" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   3.48,02 SAY "Smtp/Senha(DOCTOSNET)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   4.35,02 SAY "Pasta" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   5.22,02 SAY "Pasta Logs" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @ 01,10 GET ZORRETORA PICTURE "999" RIGHT FONT oFontget SIZE 25,11
        @ 01,21 GET ZNST PICTURE "99999" RIGHT FONT oFontget SIZE 30,11
        @ 02,10 GET ZOMECORR FONT oFontget SIZE 140,11
        @ 03,10 GET ZMAIL FONT oFontget SIZE 140,11
        @ 04,10 GET ZMTPCORR FONT oFontget SIZE 140,11
        @ 05,10 GET ZOCALCOR PICTURE "@!" FONT oFontget SIZE 140,11
        @ 06,10 GET ZASTALOG PICTURE "@!" FONT oFontget SIZE 140,11
        @ 7,02 CHECKBOX ZNVMAIL   PROMPT "Enviar Email"               SIZE  70,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248) WHEN !ZINIMAGE
        @ 7,09 CHECKBOX ZADOS     PROMPT "Guardar Dados(CambioNet)"   SIZE 108,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248) WHEN !ZINIMAGE
        @ 7,22 CHECKBOX ZNSTALA   PROMPT "Implantacao BD"             SIZE  80,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248) WHEN !ZINIMAGE
      ElseIf QTP=3
        DEFINE FONT oFont NAME "courier new" SIZE 0,-12
        DEFINE DIALOG oDlg FROM 0,0 TO 35,85 TITLE "Usuários do BolsasNet" BRUSH oBrushTela STYLE WS_CAPTION
        @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 325,235
        @   0.87,02 SAY "Usuario" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   1.74,02 SAY "Senha" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,02 SAY "Corretora" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,20 SAY "Custo p/1000 impressões" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   3.48,02 SAY "<B>ovespa,BM(F)" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   3.48,20 SAY "Validade" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   3.48,35 SAY "Tipo:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   4.35,02 SAY "Pasta" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   5.22,02 SAY "Email" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   6.09,02 SAY "SMTP" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   6.96,02 say "Login" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   7.83,02 SAY "Senha" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   8.70,02 SAY "Direitos" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)

        @ 01,10 GET ZSUARIO FONT oFontget SIZE 100,11
        @ 02,10 GET ZENHA FONT oFontget SIZE 100,11
        @ 03,10 GET ZORRETORA FONT oFontget SIZE 30,11 PICTURE "@!"
        @ 04,10 GET ZOLSA FONT oFontget SIZE 10,11 PICTURE "!" VALID ZOLSA$"BF"
        @ 04,20 GET ZALIDADE FONT oFontget SIZE 45,11 PICTURE "@D"
        @ 04,30 GET ZIPO FONT oFontget SIZE 10,11 PICTURE "!" VALID ZIPO$"UCB"
        @ 05,10 GET ZASTA FONT oFontget SIZE 120,11 PICTURE "@!" VALID !EMPTY(ZASTA)
        @ 06,10 GET ZMAIL FONT ofontget SIZE 140,11 VALID !EMPTY(ZMAIL)
        @ 07,10 GET ZMTP FONT ofontget SIZE 140,11
        @ 08,10 GET ZOGEMAIL FONT ofontget SIZE 140,11
        If EMPTY(ZATAEXP)
          ZATAEXP=DATE()
        EndIf
        @ 09,10 GET ZENHAEMAIL FONT ofontget SIZE 140,11
        @ 10,10 GET ZIREITOS FONT oFontget SIZE 140,11
        @ 03,25 GET ZUSTOMIL FONT ofontget SIZE 80,11 PICTURE "@E 999,999.99" RIGHT
        @ 10.44,01 SAY "Mensagem Extrato Unificado" FONT oFontsay  COLOR CLR_BLACK,nRGB(248,248,248)

        @ 11.0,02 CHECKBOX ZORREIO    PROMPT "Imprimir/Correio"         SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 11.0,12 CHECKBOX ZNVMAIL    PROMPT "Enviar Email"             SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 11.0,22 CHECKBOX ZADOS      PROMPT "Guardar Dados(DoctosNet)" SIZE 130,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 11.0,35 CHECKBOX ZINIMAGE   PROMPT "Integração Sictur"      SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 12.0,18 CHECKBOX ZCEITA     PROMPT "Aceitar não enviar email e não imprimir" SIZE 200,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)

        //  @ 12.0,35 CHECKBOX ZNVBC      PROMPT "Enviar/Receber Bacen"     SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)

        @ 13,01 GET ZENS1 FONT oFont SIZE 250,11
        @ 14,01 GET ZENS2 FONT ofont SIZE 250,11
        @ 15,01 GET ZENS3 FONT oFont SIZE 250,11
        @ 16,01 GET ZENS4 FONT oFont SIZE 250,11
        @ 17,01 GET ZENS5 FONT oFont SIZE 250,11
        @ 18,01 GET ZENS6 FONT oFont SIZE 250,11

      ElseIf QTP=4

        DEFINE DIALOG oDlg FROM 0,0 TO 18,58 TITLE "Parametros do sistema" BRUSH oBrushTela STYLE WS_CAPTION
        @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 222,110
        @   1.74,02 SAY "Corretora" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   2.61,02 SAY "Pasta arqs.Gerados" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @   3.48,02 SAY "Custo p/1000 impressões" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
        @ 02,12 GET ZORRETORA FONT oFontget SIZE 30,11 PICTURE "@!" NO MODIFY
        @ 03,12 GET ZASTA FONT oFontget SIZE 120,11 PICTURE "@!" VALID !EMPTY(ZASTA) .And. RIGHT(ALLTRIM(ZASTA),1)$"\/"
        @ 04,12 GET ZUSTOMIL FONT ofontget SIZE 80,11 PICTURE "@E 999,999.99" RIGHT

        @ 6,02 CHECKBOX ZCEITA     PROMPT "Aceitar não enviar email e não imprimir" SIZE 200,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 7,02 CHECKBOX ZORREIO    PROMPT "Imprimir/Correio"         SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
        @ 7,12 CHECKBOX ZNVMAIL    PROMPT "Enviar Email"             SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
      EndIf

      MCONF=.F.
      If QTP=2 .OR. QTP=4
        @ 122,065 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (IIF(F="I",ADIREG(0),""),PUTREC(),MCONF:=.T.,oDlg:end())
        @ 122,125 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
      ElseIf QTP=3
        @ 245,065 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (ZORPOEMAIL:=ansitooem(zorpoemail),IIF(F="I",ADIREG(0),""),PUTREC(),MCONF:=.T.,oDlg:end())
        @ 245,125 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
      EndIf
      oDlg:lhelpicon:=.f.
      ACTIVATE DIALOG oDlg CENTERED
      If QTP=2
        If MCONF
          If !VERDIRETORIO(ZOCALCOR)
            MSGINFO("Nao foi criado o diretório : "+ZOCALCOR,"")
          EndIf
        EndIf
      ElseIf QTP=3
        If !VERDIRETORIO("\NOTAS\")
          MSGINFO("Nao foi criado o diretório : \NOTAS\","")
        EndIf
        If !VERDIRETORIO("\NOTAS\"+ALLTRIM(ZORRETORA)+"\")
          MSGINFO("Nao foi criado o diretório : \NOTAS\"+ALLTRIM(ZORRETORA)+"\","")
        EndIf
        If !VERDIRETORIO(ALLTRIM(ZASTA))
          MSGINFO("Nao foi criado o diretório : "+ZASTA,"")
        EndIf
      ElseIf QTP=4
        If !VERDIRETORIO(ALLTRIM(ZASTA))
          MSGINFO("Nao foi criado o diretório : "+ZASTA,"")
        EndIf
      EndIf
      COMMIT
    return 


    Function VERIMPRESSORAS()
      Local aPrn:=GetAllEntrys(),f
      For f:=1 to Len(aPrn)
        SysRefresh()
        If AT("win2pdf",LOWER(aPrn[F]))#0
          TEMPDF=.T.
          TEMPDF1=ALLTRIM(APRN[F])
        EndIf
      Next
    return  Nil

    Function GetAllEntrys()
      Local aDevices:={}, cAllEntries, cEntry, I, cName, cPrn, cPort, J

      cAllEntries := Strtran( GetProfString( "Devices" ), Chr( 0 ), CRLF )

      FOR I:= 1 TO MlCount( cAllEntries )
        SysRefresh()
        cName := MemoLine( cAllEntries,,I)
        cEntry := GetProfString( "Devices",cName,"")
        J := 2
        DO WHILE ! EMPTY(cPort := StrToken(cEntry,J++,","))
          SysRefresh()
          AADD(aDevices,TRIM(cName))
        ENDDO
      NEXT
    return  aDevices

    Procedure ENVIAEMAIL(MQUEMRECEBE,MTITULO,MTEXTO,MANEXO,MANEXO1,MCONTA,msenha,mservidor,mlogin,mforca,MNOMEENV,MWEBCOR,WMPORTA)
      DEFAULT mservidor:=cIPServer,mforca:=.f.,MNOMEENV:=MQUEMRECEBE,MWEBCOR:=MCONTA,WMPORTA:="587"
      cuser=mlogin
      cAssun=ANSITOOEM(mtitulo)
      cAnexo=alltrim(manexo)
      cDe=alltrim(lower(mconta))
      cServ=alltrim(mservidor)
      cBody=mtexto+"<br>"
      cAnexo1=alltrim(manexo1)
      cpara1=ALLTRIM(LOWER(mquemrecebe))
      CPARA1=Strtran(CPARA1," ","")
      cpara1=Strtran(cpara1,";",",")
      while .T.
        SysRefresh()
        M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
        ARQ_EMAIL=LOCNOT1+"EM"+SUBS(STR(M->EX_T+1000,4),2)
        ARQ_EMAIL1=ARQ_EMAIL+".BAT"
        ARQ_EMAIL2=ARQ_EMAIL+".RET"
        ARQ_EMAIL3=ARQ_EMAIL+".HTML"
        ARQ_EMAIL4=ARQ_EMAIL+".TXT"
        If !FILE(ARQ_EMAIL1)
          EXIT
        EndIf
      ENDDO
      cassun=oemtoansi(cassun)
      If AT("amazonaws",LOWER(cserv))=0
        EnviaCOMPEmail(cServ,cde,MWEBCOR,cPara1,Mnomeenv,Cassun,CBody,Canexo,Mlogin,Msenha,WMPORTA,ARQ_EMAIL2,cAnexo1)
      Else
        EnviaCDO(cServ,cde,MWEBCOR,cPara1,Mnomeenv,Cassun,CBody,Canexo,cuser,Msenha,WMPORTA,cAnexo1)
      EndIf
      DECLARE TX1[1]
      TX1[1]=cAssun+" "+MQUEMRECEBE
      If !EMPTY(WDADOS)
        LOGFILE(WDADOS,TX1)
      EndIf
    return  .T.

    Procedure BUSCANOVAS()
      DIF=TIMEDIFF(ULTBUSCA,TIME())
      DIF1=TIMEASSECONDS(DIF)
      If DIF1<(LOCALCOTNOT*60)
        return 
      EndIf
      ULTBUSCA=TIME()
      cmacro="wget.exe -T10 -olog.txt -OCOTAS.dat http://www.exatus.net/busca_cotacoes.asp"
      waitrun(CMACRO,0)
      // ATUANOT()
      CURSORARROW()
    return 

    Procedure CADWINCC()
      If !USUARIO(14)
        return 
      EndIf
      SELE 24
      Browse(13,"Usuários CambioNet","CambioNet", { || buscaval() },{ ||NCADWINC("I") },{ ||NCADWINC("A") })
    return 

    Procedure NCADWINC(FS)
      Local oDlg,oBrush
      PUBLIC F
      F=FS
      DECLARE _VAR1[FCOUNT()]
      AFIELDS(_VAR1)
      If F="I"
        GO BOTT
        SKIP
      EndIf
      FOR t = 1 TO FCOUNT()
        SysRefresh()
        _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
        _var5 = _var1 [t]
        &_var4 = &_var5
      NEXT t
      RELEASE _VAR1
      MFONTES()
      DEFINE DIALOG oDlg FROM 0,0 TO 20,60 TITLE "Dados do Usuário-CambioNet" BRUSH oBrushTela STYLE WS_CAPTION
      @   0.44,0.5 SAY " " BOX RAISED  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 230,125
      @   0.87,02 SAY "Usuário"    FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   1.74,02 SAY "Senha"      FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   2.61,02 SAY "Frase"      FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   3.48,02 SAY "Pasta"      FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   4.35,02 SAY "Email"      FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   6,02    SAY "Ult.Acesso" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   6.90,02 SAY "Acessos"    FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      @   7.80,02 SAY "Corretora"  FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
      A="S"
      @ 01,10 GET ZSUARIO FONT oFontget size 145,11
      @ 02,10 GET ZENHA FONT oFontget SIZE 145,11
      @ 03,10 GET ZRASE FONT oFontget SIZE 145,11
      @ 04,10 GET ZASTA FONT oFontget SIZE 145,11
      @ 05,10 GET ZMAIL FONT oFontget SIZE 145,11
      @ 07,10 GET ZLTACESSO FONT oFontget SIZE 45,11 WHEN A="N"
      @ 08,10 GET ZCESSO FONT oFontget SIZE 45,11 WHEN A="N"
      @ 09,10 GET ZORRETORA FONT oFontget SIZE 145,11

      @ 6,02 CHECKBOX ZLOQUEADO PROMPT "Bloqueado"        SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
      @ 6,08 CHECKBOX ZSADMIN   PROMPT "Administrador"    SIZE 100,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
      @ 6,15.5 CHECKBOX ZNVEXTR   PROMPT "Enviar Email de Extratos Mensalmente"  SIZE 139,11 OF oDlg FONT oFontchk COLOR CLR_BLACK,nRGB(248,248,248)
      GRAVOU=.F.
      @ 136,070 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "&Confirma"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION (IIF(F="I",ADIREG(0),""),PUTREC(),GRAVOU:=.T.,oDlg:end())
      @ 136,130 SBUTTON PIXELS NOBOX RESOURCE "OFF","ON","DISABLE" PROMPT "C&ancela"  ADJUST TEXT ON_CENTER FONT oFontbut OF oDlg SIZE 40,12 ACTION oDlg:End()
      oDlg:lhelpicon:=.f.
      ACTIVATE DIALOG oDlg CENTERED
      EMPRESAVAL=VAL(RIGHT(STRZERO(VAL(ZSUARIO),15),2))
      SELE 22
      LOCATE FOR CORRETORA=EMPRESAVAL
      LOGF=PASTALOG
      SELE 24
      If !EMPTY(LOGF) .And. GRAVOU=.T.
        DECLARE _var1 [FCOUNT()],_VAR2[FCOUNT()],_VAR3[FCOUNT()]
        aFIELDS(_VAR1,_VAR2,_VAR3)
        TX=""
        FOR I=1 TO LEN(_VAR1)
          SysRefresh()
          a=_var1[i]
          A=&A
          If _VAR2[I]="N"
            A=STR(A)
          ElseIf _VAR2[I]="M"
            A=" "
          ElseIf _VAR2[I]="D"
            A=DTOC(A)
          ElseIf _VAR2[I]="L"
            A=" "
          EndIf
          TX=TX+","+A
        NEXT I
        If F="I"
          OBSERVAC=" INCLUIU: "
        Else
          OBSERVAC=" ALTEROU: "
        EndIf
        DECLARE TX1[1]
        TX1[1]="USUARIO: "+A02NOME+OBSERVAC+TX
        LOGFILE(LOGF,TX1)
      EndIf
      COMMIT
    return 

    Procedure ATUAWEBWINCC()
      SELE 22
      WIIPSERVER=CIPSERVER
      GO TOP
      M->LOCAL21=TRIM("/"+CURDIR())
      DO WHILE !EOF()
        SysRefresh()
        If WINIMAGE
          SKIP
          LOOP
        EndIf
        WEBLOCAL=ALLTRIM(LOCALCOR)
        WEBCORR=STRZERO(CORRETORA,2)
        ZNOMECORR=NOMECORR
        SEINSTALA=INSTALA
        MCONTA=LOWER(ALLTRIM(EMAIL))
        MWARQ=WEBLOCAL+"FEITO.FAZ"
        MWARQ2=WEBLOCAL+"FEITO.ESP"
        MWARQ1=WEBLOCAL+"TUDOROF.ZIP"
        If FILE(MWARQ) .And. FILE(MWARQ1)
          LZCOPYFILE(MWARQ,MWARQ2)
          INICIOU1()
        EndIf
        SELE 22
        MWARQ1=WEBLOCAL+"TUDO.ZIP"
        If FILE(MWARQ1) .And. FILE(MWARQ)
          LZCOPYFILE(MWARQ,MWARQ2)
          INICIOU()
          INICIOU2()
        EndIf
        Erase &MWARQ
        NARQ=WEBLOCAL+"QUEMROF.FRO"
        Erase &NARQ
        NARQ=WEBLOCAL+"QUEM.FRO"
        Erase &NARQ
        SELE 22
        SKIP
      ENDDO
      lCHDIR(M->LOCAL21)
      CIPSERVER=WIIPSERVER
    return 

    Procedure INICIOU()
      ZIRE=""
      ZATURA=""
      ZERCADORIA=""
      ZODALIDADE=""
      ZONHECIMENTO=""
      ZTATUS=""
      ZBS=""
      comando=WEBLOCAL+"ATUA.BAT"
      WAITRUN(COMANDO,0)
      ZXIGENCIA=""
      DECLARE TX1[1]
      TX1[1]="Cambio Net"
      LOGFILE(WEBLOCAL+"ATUALIZ.TXT",TX1)
      NARQ=WEBLOCAL+"TUDO.ZIP"
      COMANDO='brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
      WAITRUN(COMANDO,0)
      Erase &NARQ
      CORRENTE=Strtran("/"+TRIM(CURDIR()),"\","/")
      DECLARE VLABEL[ADIR(WEBLOCAL+"EX*.XLS")]
      ADIR(WEBLOCAL+"EX*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"EXPORTAC.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      ADIR(WEBLOCAL+"S1*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"SDOEXPOR.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      ADIR(WEBLOCAL+"S2*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"SDOIMPOR.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      DECLARE VLABEL[ADIR(WEBLOCAL+"IM*.XLS")]
      ADIR(WEBLOCAL+"IM*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"IMPORTAC.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      DECLARE VLABEL[ADIR(WEBLOCAL+"FD*.XLS")]
      ADIR(WEBLOCAL+"FD*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"FINDOEXT.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      DECLARE VLABEL[ADIR(WEBLOCAL+"FP*.XLS")]
      ADIR(WEBLOCAL+"FP*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"FINPARAE.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      DECLARE VLABEL[ADIR(WEBLOCAL+"CT*.XLS")]
      ADIR(WEBLOCAL+"CT*.XLS",Vlabel)
      CURSORWAIT()
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        MZCODCLI=VAL(SUBSTR(VLABEL[I],3,6))
        ARQXLS=WEBLOCAL+ALLTRIM(VLABEL[I])
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+STRZERO(MZCODCLI,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &ARQXLS
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        ARQXLS1=M->LOCAL+"CTAEXTER.XLS"
        //  LZCOPYFILE(ARQXLS,ARQXLS1)
        Erase &ARQXLS
      NEXT I
      LCHDIR(CORRENTE)
      ARQDB=WEBLOCAL+"EMAIL.INC"
      If FILE(ARQDB)
        SENTIMOS()
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL+"*.DB")]
      ADIR(WEBLOCAL+"*.DB",Vlabel)
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        If VLABEL[I]="SCANNER.DB" .OR. VLABEL[I]="DOCTOS.DBF" .OR. VLABEL[I]="ARECEB.DBF" .OR. VLABEL[I]="ACCACE.DB"
          LOOP
        EndIf
        CURSORWAIT()
        lCHDIR(CORRENTE)
        SELE 25
        ARQDB=WEBLOCAL+ALLTRIM(VLABEL[I])
        ARQFT=LEFT(ARQDB,LEN(ARQDB)-3)+".FPT"
        If USEREDE(ARQDB,.T.,10,"TMP")
          GO TOP
          DO WHILE .NOT. EOF()
            SysRefresh()
            SELE 25
            ARQPDF=WEBLOCAL+"C"+STRZERO(INTERNO,6)+".PDF"
            ARQPDF2="C"+STRZERO(INTERNO,6)+".PDF"
            If !FILE(ARQPDF)
              ARQPDF=WEBLOCAL+STRZERO(INTERNO,8)+".PDF"
              ARQPDF2=STRZERO(INTERNO,8)+".PDF"
              If !FILE(ARQPDF)
                SKIP
                LOOP
              EndIf
            EndIf
            M->CLIE2=STRZERO(CODCLI,6)+WEBCORR
            SELE 24
            GO TOP
            LOCATE FOR CLIENTE=M->CLIE2
            If EOF() .And. !SEINSTALA
              Erase &ARQPDF
              SELE 25
              SKIP
              LOOP
            EndIf
            SELE 25
            M->LOCAL1=WEBLOCAL
            M->LOCAL=WEBLOCAL+STRZERO(CODCLI,6)
            If !lCHDIR(M->LOCAL)
              If !lMKDIR(M->LOCAL)
                lCHDIR(CORRENTE)
                SKIP
                LOOP
              EndIf
            EndIf
            lCHDIR(CORRENTE)
            M->LOCAL=M->LOCAL+"/"
            M->LOCAL1=M->LOCAL1+"/"
            M->LOCAL=Strtran(M->LOCAL,"/","\")
            M->LOCAL1=Strtran(M->LOCAL1,"/","\")
            M->ARQPDF1=M->LOCAL+ARQPDF2
            If FILE(M->ARQPDF1)
              Erase &ARQPDF
              SKIP
              LOOP
            EndIf
            ARQ=M->LOCAL+"CONTT.DBF"
            ZDA=0
            ZOVOC=""
            If !FILE(ARQ)
              If !CRIACONT(ARQ)
                Erase &ARQPDF
                SELE 25
                SKIP
                LOOP
              EndIf
            EndIf
            If !CMOV1()
              Erase &ARQPDF
              SELE 25
              SKIP
              LOOP
            EndIf
            SELE 25
            ZONGLO=0
            DECLARE _var1 [FCOUNT()]
            AFIELDS(_var1)
            FOR t = 1 TO FCOUNT()
              SysRefresh()
              _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
              _var5 = _var1 [t]
              &_var4 = &_var5
            NEXT t
            ZOMECLI1=SINAL(ZOMECLI)
            ZOMECLI=ZOMECLI1
            SELE 26
            If !USEREDE(ARQ,.F.,10,"CONTR","S")
              SELE 25
              Erase &ARQPDF
              SKIP
              LOOP
            EndIf

            If ADIREG(0)
              PUTREC()
              LZCOPYFILE(ARQPDF,M->ARQPDF1)
            EndIf
            USE
            SELE 25
            Erase &ARQPDF
            If ZONGLO#0 .And. ZONGLO<>ZODCLI
              M->LOCAL1=WEBLOCAL
              M->LOCAL=WEBLOCAL+STRZERO(ZONGLO,6)
              If !lCHDIR(M->LOCAL)
                If !lMKDIR(M->LOCAL)
                  lCHDIR(CORRENTE)
                  SKIP
                  LOOP
                EndIf
              EndIf
              lCHDIR(CORRENTE)
              M->LOCAL=M->LOCAL+"/"
              M->LOCAL1=M->LOCAL1+"/"
              M->LOCAL=Strtran(M->LOCAL,"/","\")
              M->LOCAL1=Strtran(M->LOCAL1,"/","\")
              ARQ=M->LOCAL+"CONTT.DBF"
              ZDA=0
              If !FILE(ARQ)
                If !CRIACONT(ARQ)
                  SKIP
                  LOOP
                EndIf
              EndIf
              If !CMOV1()
                SKIP
                LOOP
              EndIf
              ZLIENTE=ZONGLO
              ZONGLO=0
              SELE 26
              If !USEREDE(ARQ,.F.,10,"CONTR","S")
                SELE 25
                SKIP
                LOOP
              EndIf
              If ADIREG(0)
                PUTREC()
              EndIf
              USE
              SELE 25
            EndIf
            SKIP
          ENDDO
        EndIf
        SELE 25
        USE
        Erase &ARQDB
        Erase &ARQFT
        lChDir(CORRENTE)
      NEXT I
      lCHDIR(CORRENTE)
      DECLARE VLABEL[ADIR(WEBLOCAL+"*.DAT")]
      ADIR(WEBLOCAL+"*.DAT",Vlabel)
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        SELE 25
        ARQDB=WEBLOCAL+ALLTRIM(VLABEL[I])
        ARQDBCRE=WEBLOCAL+"AVENCER.DBF"
        Erase &ARQDBCRE
        mat_stru={}
        AADD(MAT_STRU,{"DTPAGTO","D",8,0})
        AADD(MAT_STRU,{"DTREF","D",8,0})
        aAdd(mat_stru,{"CONGLO","N",6,0})
        aAdd(mat_stru,{"CGCCPF","C",14,0})
        aAdd(mat_stru,{"DATA","D",8,0})
        aAdd(mat_stru,{"TP","C",1,0})
        aAdd(mat_stru,{"DTCON","D",8,0})
        aAdd(mat_stru,{"CLIENTE","N",6,0})
        aAdd(mat_stru,{"NOMECLI","C",50,0})
        aAdd(mat_stru,{"NOMEPAGREC","C",50,0})
        aAdd(mat_stru,{"NOMEMOE","C",20,0})
        aAdd(mat_stru,{"VENCTO","D",8,0})
        aAdd(mat_stru,{"PRZVINC","D",8,0})
        aAdd(mat_stru,{"DIRE","C",15,0})
        aAdd(mat_stru,{"PAGREC","N",6,0})
        aAdd(mat_stru,{"REFCLI","C",30,0})
        aAdd(mat_stru,{"MOEDA","N",3,0})
        aAdd(mat_stru,{"VALOR","N",19,2})
        aAdd(mat_stru,{"FAT1","C",69,0})
        aAdd(mat_stru,{"MERCAD","C",50,0})
        aAdd(mat_stru,{"MODALID","C",15,0})
        aAdd(mat_stru,{"EXIGENCIA","C",70,0})
        aAdd(mat_stru,{"CONHEC","C",30,0})
        aAdd(mat_stru,{"ADICAO","C",3,0})
        aAdd(mat_stru,{"INTERNO","N",8,0})
        aAdd(mat_stru,{"COMAG","C",30,0})
        aAdd(mat_stru,{"STATUS","C",15,0})
        AADD(MAT_STRU,{"LPINHA1","C",69,0})
        AADD(MAT_STRU,{"LPINHA2","C",69,0})
        AADD(MAT_STRU,{"LPINHA3","C",69,0})
        AADD(MAT_STRU,{"LPINHA4","C",69,0})
        AADD(MAT_STRU,{"LPINHA5","C",69,0})
        AADD(MAT_STRU,{"LPINHA6","C",69,0})
        AADD(MAT_STRU,{"LPINHA7","C",69,0})
        AADD(MAT_STRU,{"LPINHA8","C",69,0})
        DBCREATE(ARQDBCRE,mat_stru,"DBFNTX")
        If USEREDE(ARQDB,.T.,10,"TMP")
          GO TOP
          SELE 26
          If USEREDE(ARQDBCRE,.F.,10,"TMP1")
            SELE 25
            DO WHILE .NOT. EOF()
              SysRefresh()
              SELE 25
              DECLARE _var1 [FCOUNT()]
              AFIELDS(_var1)
              ZODALID=""
              ZDICAO=""
              ZONHEC=""
              ZOMAG=""
              ZPINHA1=""
              ZPINHA2=""
              ZPINHA3=""
              ZPINHA4=""
              ZPINHA5=""
              ZPINHA6=""
              ZPINHA7=""
              ZPINHA8=""
              ZTREF=CTOD("")
              ZTPAGTO=CTOD("")
              ZNTERNO=0
              ZTCON=CTOD("")
              ZGCCPF=""
              zencto=ctod("")
              ZONGLO=0
              FOR t = 1 TO FCOUNT()
                _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
                _var5 = _var1 [t]
                &_var4 = &_var5
              NEXT t
              ZTREF=ZTPAGTO
              If ZTREF=CTOD("")
                ZTREF=ZENCTO
              EndIf
              ZOMECLI1=SINAL(ZOMECLI)
              ZOMECLI=ZOMECLI1
              SELE 26
              If ADIREG(0)
                PUTREC()
              EndIf
              SELE 25
              SKIP
            ENDDO
          EndIf
        EndIf
        SELE 25
        USE
        SELE 26
        USE
        Erase &ARQDB
      NEXT I
      M->FROM=""
      If FILE(WEBLOCAL+"QUEM.FRO")
        SELE 25
        If USEREDE(WEBLOCAL+"QUEM.FRO",.F.,10,"TMP")
          GO TOP
          M->FROM=ALLTRIM(FROM)
        EndIf
        USE
      EndIf
      If !EMPTY(M->FROM)
        SEENVIAEMAIL="S"
        webcor=" "
        ENVIAEMAIL(FROM,"CambioNet","CambioNet atualizado com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
      NARQ=WEBLOCAL+"QUEM.FRO"
      Erase &NARQ
      Erase &WDADOS
    return 

    Procedure INICIOU1()
      DECLARE TX1[1]
      TX1[1]="Cambio Net - ROF"
      LOGFILE(WEBLOCAL+"ATUALIZ.TXT",TX1)
      WDADOS=WEBLOCAL+"ENVIADOS.TXT"
      Erase &WDADOS
      NARQ=WEBLOCAL+"TUDOROF.ZIP"
      COMANDO='brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
      WAITRUN(COMANDO,0)
      Erase &NARQ
      CORRENTE=Strtran("/"+TRIM(CURDIR()),"\","/")
      SELE 25
      ARQDB=WEBLOCAL+"ROFS.DAT"
      If !FILE(ARQDB)
        return 
      EndIf
      If USEREDE(ARQDB,.T.,10,"TMP")
        GO TOP
        DO WHILE .NOT. EOF()
          SysRefresh()
          MZREG=RECNO()
          ARQPDF=WEBLOCAL+STRZERO(TMP->NUM_CONTR,8)+".PDF"
          If !FILE(ARQPDF)
            SKIP
            LOOP
          EndIf
          sysrefresh()
          M->CLIE2=STRZERO(TMP->CLIENTE,6)+WEBCORR
          SELE 24
          LOCATE FOR USUARIO=M->CLIE2
          If EOF()
            Erase &ARQPDF
            SELE 25
            SKIP
            LOOP
          EndIf
          SELE 25
          WEBFAZEMAIL=TMP->EMAIL
          M->LOCAL=WEBLOCAL+STRZERO(TMP->CLIENTE,6)
          If !lCHDIR(M->LOCAL)
            If !lMKDIR(M->LOCAL)
              lCHDIR(CORRENTE)
              Erase &ARQPDF
              SKIP
              LOOP
            EndIf
          EndIf
          M->LOCAL=WEBLOCAL+STRZERO(TMP->CLIENTE,6)+"\ESP"
          If !lCHDIR(M->LOCAL)
            If !lMKDIR(M->LOCAL)
              lCHDIR(CORRENTE)
              Erase &ARQPDF
              SKIP
              LOOP
            EndIf
          EndIf
          lCHDIR(CORRENTE)
          M->LOCAL=M->LOCAL+"/"
          M->LOCAL=Strtran(M->LOCAL,"/","\")
          NW=TMP->CODIGO
          NW=SINAL(NW)
          NW=Strtran(NW,"/","")
          NW=Strtran(NW,"\","")
          NW=Strtran(NW,"-","")
          M->ARQPDF1=M->LOCAL+ALLTRIM(NW)+".PDF"
          COMANDO="COPY "+ALLTRIM(ARQPDF)+" "+ALLTRIM(ARQPDF1)
          MEMOWRIT("COPIA.BAT",COMANDO)
          CURSORWAIT()
          WAITRUN("COPIA.BAT",0)
          SELE 25
          Erase &ARQPDF
          GO MZREG
          SKIP
        ENDDO
        USE
        Erase &ARQDB
        lChDir(CORRENTE)
      EndIf
      DECLARE TX1[1]
      TX1[1]="Cambio Net - ROF Planilhas"
      LOGFILE(WEBLOCAL+"ATUALIZ.TXT",TX1)
      WDADOS=WEBLOCAL+"ENVIADOS.TXT"
      CORRENTE=Strtran("/"+TRIM(CURDIR()),"\","/")
      DECLARE VLABEL[ADIR(WEBLOCAL+"AV*.XLS")]
      ADIR(WEBLOCAL+"AV*.XLS",vlabel)
      I=0
      FOR I=1 TO LEN(VLABEL)
        SysRefresh()
        NARQ=WEBLOCAL+VLABEL[I]
        M->CLIE2=SUBSTR(VLABEL[I],3,6)+WEBCORR
        SELE 24
        LOCATE FOR USUARIO=M->CLIE2
        If EOF()
          Erase &NARQ
          LOOP
        EndIf
        M->LOCAL=WEBLOCAL+SUBSTR(VLABEL[I],3,6)
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &NARQ
            LOOP
          EndIf
        EndIf
        M->LOCAL=WEBLOCAL+SUBSTR(VLABEL[I],3,6)+"\ESP"
        If !lCHDIR(M->LOCAL)
          If !lMKDIR(M->LOCAL)
            lCHDIR(CORRENTE)
            Erase &NARQ
            LOOP
          EndIf
        EndIf
        lCHDIR(CORRENTE)
        M->LOCAL=M->LOCAL+"/"
        M->LOCAL=Strtran(M->LOCAL,"/","\")
        M->ARQPDF1=M->LOCAL+VLABEL[I]
        COMANDO="COPY "+ALLTRIM(NARQ)+" "+ALLTRIM(ARQPDF1)
        MEMOWRIT("COPIA.BAT",COMANDO)
        CURSORWAIT()
        WAITRUN("COPIA.BAT",0)
        Erase &NARQ
      NEXT I
      lChDir(CORRENTE)
      M->FROM=""
      If FILE(WEBLOCAL+"QUEMROF.FRO")
        SELE 25
        If USEREDE(WEBLOCAL+"QUEMROF.FRO",.F.,10,"TMP")
          GO TOP
          M->FROM=ALLTRIM(FROM)
        EndIf
        USE
      EndIf
      If !EMPTY(M->FROM)
        SEENVIAEMAIL="S"
        webcor=" "
        ENVIAEMAIL(FROM,"CambioNet - ROF","CambioNet-ROF atualizado com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        If FILE(WDADOS)
          SEENVIAEMAIL="S"
          webcor=" "
          ENVIAEMAIL(FROM,"Cambio Net-Rof","Relatório dos Envios de Email",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
      EndIf
      NARQ=WEBLOCAL+"QUEMROF.FRO"
      Erase &NARQ
      Erase &WDADOS
    return 

    Procedure INICIOU2()
      ZXIGENCIA=""
      DECLARE TX1[1]
      TX1[1]="Cambio Net - Doctos Scanneados"
      LOGFILE(WEBLOCAL+"ATUALIZ.TXT",TX1)
      CORRENTE=Strtran("/"+TRIM(CURDIR()),"\","/")
      CURSORWAIT()
      lCHDIR(CORRENTE)
      ARQDB=WEBLOCAL+"SCANNER.DB"
      If !FILE(ARQDB)
        return 
      EndIf
      SELE 25
      If USEREDE(ARQDB,.T.,10,"TMP")
        GO TOP
        DO WHILE .NOT. EOF()
          SysRefresh()
          SELE 25
          ARQPDF =WEBLOCAL+STRZERO(INTERNO,8)+".PD1"
          ARQPDF2=WEBLOCAL+STRZERO(CLIENTE,6)+"\SWIFT\"+STRZERO(INTERNO,8)+".PDF"
          If !FILE(ARQPDF)
            SKIP
            LOOP
          EndIf
          M->CLIE2=STRZERO(CLIENTE,6)+WEBCORR
          SELE 24
          LOCATE FOR USUARIO=M->CLIE2
          If EOF() .And. !SEINSTALA
            Erase &ARQPDF
            SELE 25
            SKIP
            LOOP
          EndIf
          SELE 25
          M->LOCAL1=WEBLOCAL
          M->LOCAL=WEBLOCAL+STRZERO(CLIENTE,6)
          If !lCHDIR(M->LOCAL)
            If !lMKDIR(M->LOCAL)
              lCHDIR(CORRENTE)
              SKIP
              LOOP
            EndIf
          EndIf
          lCHDIR(CORRENTE)
          M->LOCAL=WEBLOCAL+STRZERO(CLIENTE,6)+"\SWIFT"
          If !lCHDIR(M->LOCAL)
            If !lMKDIR(M->LOCAL)
              lCHDIR(CORRENTE)
              SKIP
              LOOP
            EndIf
          EndIf
          lCHDIR(CORRENTE)
          M->LOCAL=M->LOCAL+"\"
          lCHDIR(CORRENTE)
          If FILE(ARQPDF2)
            Erase &ARQPDF
            SELE 25
            SKIP
            LOOP
          EndIf
          LZCOPYFILE(ARQPDF,ARQPDF2)
          SELE 25
          Erase &ARQPDF
          SKIP
        ENDDO
      EndIf
      SELE 25
      USE
      Erase &ARQDB
      lChDir(CORRENTE)
    return 

    Procedure CRIACONT(ARQ)
      CURSORWAIT()
      MAT_STRU={}
      AADD(MAT_STRU,{"CONGLO","N",6,0})
      aAdd(mat_stru,{"NOMEPAG","C",69,0})
      aAdd(mat_stru,{"NOMEPRACA","C",20,0})
      aAdd(mat_stru,{"SIMBOLMOE","C",5,0})
      aAdd(mat_stru,{"NOMENAT","C",51,0})
      aAdd(mat_stru,{"GRUPONAT","C",2,0})
      aAdd(mat_stru,{"ANEXO","M",10,0})
      aAdd(mat_stru,{"OUTRAS","M",10,0})
      aAdd(mat_stru,{"LIQUIDACAO","D",8,0})
      aAdd(mat_stru,{"CERTAUT","C",10,0})
      aAdd(mat_stru,{"FIRCE","C",13,0})
      aAdd(mat_stru,{"ENTREGADOC","D",8,0})
      aAdd(mat_stru,{"VLMOEDNAC","C",18,0})
      aAdd(mat_stru,{"PAGREC","N",6,0})
      aAdd(mat_stru,{"NR_CONTRAT","N",  12,  0})
      aAdd(mat_stru,{"INTERNO","N",8,0})
      aAdd(mat_stru,{"TIPO","C",1,0})
      aAdd(mat_stru,{"PRACA","C",4,0})
      aAdd(mat_stru,{"BANCO","N",6,0})
      aAdd(mat_stru,{"NOMEBANCO","C",25,0})
      aAdd(mat_stru,{"CODCLI","N",6,0})
      aAdd(mat_stru,{"NOMECLI","C",59,0})
      aAdd(mat_stru,{"DTCONTR","D",8,0})
      aAdd(mat_stru,{"CODMOEDA","N",3,0})
      aAdd(mat_stru,{"VALMOEDA","C",18,0})
      aAdd(mat_stru,{"CGCCPF","C",14,0})
      aAdd(mat_stru,{"CGCCLI","C", 14,  0})
      aAdd(mat_stru,{"CPFCLI","C", 11,  0})
      aAdd(mat_stru,{"INSTITU","N",5,0})
      aAdd(mat_stru,{"MERCADO","C",1,0})
      aAdd(mat_stru,{"TXCAMBIAL","C",14,0})
      aAdd(mat_stru,{"NATOPERA","N",12,0})
      aAdd(mat_stru,{"INST","N",5,0})
      aAdd(mat_stru,{"DESPBANCO","C",40,0})
      aAdd(mat_stru,{"ASSINA1","C",30,0})
      aAdd(mat_stru,{"ASSINA2","C",30,0})
      aAdd(mat_stru,{"ASSINA3","C",30,0})
      aAdd(mat_stru,{"ASSINA4","C",30,0})
      aAdd(mat_stru,{"ASSINA5","C",30,0})
      aAdd(mat_stru,{"ASSINA6","C",30,0})
      aAdd(mat_stru,{"POSICAO","C",50,0})
      aAdd(mat_stru,{"NDA","N",1,0})
      AADD(MAT_STRU,{"DIRE","C",12,0})
      AADD(MAT_STRU,{"FATURA","C",12,0})
      AADD(MAT_STRU,{"MERCADORIA","C",20,0})
      AADD(MAT_STRU,{"MODALIDADE","C",20,0})
      AADD(MAT_STRU,{"CONHECIMENTO","C",20,0})
      AADD(MAT_STRU,{"OBS","C",50,0})
      AADD(MAT_STRU,{"NOVOC","C",1,0})
      DBCREATE(ARQ,MAT_STRU,"DBFNTX")
    return  .T.

    Procedure GERAXLS()
      PARA WTP1,CTEXTO1
      WTP=WTP1
      CTEXTO=CTEXTO1
      PUBLIC oFileXLS
      PUBLIC nFormat1, nFormat2, nFormat3
      PUBLIC nFont1, nFont2, nFont3, nFont4
      DEFINE XLS FORMAT nFormat1 PICTURE '###,###,###,##0.00'
      DEFINE XLS FORMAT nFormat2 PICTURE '###,##0'
      DEFINE XLS FORMAT nFormat3 PICTURE '###,###,###.000000'
      DEFINE XLS FONT nFont1 NAME "Arial" HEIGHT 16 BOLD ITALIC
      DEFINE XLS FONT nFont2 NAME "Arial" HEIGHT 10 UNDERLINE BOLD
      DEFINE XLS FONT nFont3 NAME "Arial" HEIGHT 8
      DEFINE XLS FONT nFont4 NAME "Arial" HEIGHT 8 BOLD
      DEFINE XLS FONT nFont5 NAME "Arial" HEIGHT 12 BOLD
      XLS oFileXLS FILE &ARQ_XLS
      If WTP=1
        CABECA={"Auditado","Codigo","Nome","Endereço","","Assessor","Folhas","R$ Correio","Enviado Email","Poderia Economizar","Enviado para","LOG"}
        ANCHOS:={1.5,1.5,6.0,7.0,3.5,1.0,1.0,1.2,2.0,2.0,15.0,8.0}
      ElseIf WTP=2
        CABECA={"Cliente","F/J","CNPJ/CPF","Regular","C/V","Cartão","Moeda","Quantidade","Taxa","R$","IOF","Vl.Operação","Natureza","Nr.Sictur","Banco","Agencia","Conta","tipo","Referencia","Atualizou Saldo","Motivo erro"}
        ANCHOS:={5.0,1.0,2.0,1.0,1.0,0.8,1.0,2.0,2.0,2.0,2.0,2.0,2.0,3.0,2.5,0.5,1.0,1.0,1.0,2.0,1.0,8.0}
      ElseIf WTP=3
        CABECA={"Arquivo","Nome","Cep","Logradouro","Numero","Compl","Bairro","Cidade","UF","CPF","DDD","Fone","DDDCEL","Celular","Docto","Dt.Docto","Dt Nasc","Resultado","Situaçao"}
        ANCHOS:={6.0,6.0,2.0,4.0,1.0,1.0,1.5,2.0,0.5,2.0,1.0,2.0,1.0,2.0,2.5,1.0,1.0,8.0,1.0}
      EndIf
      FOR W=1 TO LEN(CABECA)
        SysRefresh()
        XLS COL W WIDTH (ANCHOS[W]/2.54*19) OF oFileXLS
      NEXT W
      GO TOP
      WTOTD=0
      WTOTM=0
      WTOTT=0
      wcabec=0
      WTOTMDM=0
      wlin=0
      WTOTDMA=0
      WTOTDMG=0
      M->ME_SES="Janeiro  FevereiroMarço    Abril    Maio     Junho    Julho    Agosto   Setembro Outubro  Novembro Dezembro "
      If WTP=1
        If WFAZCORREIO
          oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
          ZARQ1=NARQ12
          Erase &ZARQ1
          oReg:Set("PDFFilename",ZARQ1)
          oReg:Close()
          PRINT oPrn NAME "Relatório XLS" TO TEMPDF1
          DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
          DEFINE FONT oFontv NAME "verdana" BOLD SIZE 0,9 OF oPrn
          DEFINE FONT oFont1 NAME "verdana" SIZE 0,16 BOLD OF oPrn
          DEFINE FONT oFont2 NAME "verdana" SIZE 0,6 OF oPrn
          nRowStep = oPrn:nVertRes() / 66
          nColStep = oPrn:nHorzRes() / 80
        EndIf
        ARQMW=WEBLOCAL4+"ASSESSOR.DBF"
        ARQMWK=WEBLOCAL4+"ASSESSOR.KEY"
        If !FILE(ARQMW)
          aStructure:={}
          aAdd( aStructure, { "ASSESSOR","N",5,0})
          aAdd( aStructure, { "CUSTO","N",19,2})
          aAdd( astructure, { "CIMP","N",19,2})
          dbCreate(ARQMW,aStructure,"DBFNTX")
          RELEASE aStructure
        EndIf
        SELE 38
        If !USEREDE(ARQMW,.T.,10,"ASSESSOR","S")
          MSGINFO("Problema na geração do total por Assessor, emails já enviados !",WSIST)
          QUIT
        EndIf
        INDEX ON ASSESSOR TO &ARQMWK
        SELE 37
        @ 1,1 XLS SAY ALLTRIM(CTEXTO)+" "+DTOC(DATE())+TIME() FONT nFont1 OF oFileXLS
        NLINHAS=5
        FOR W=1 TO LEN(CABECA)
          SysRefresh()
          @ 4,W XLS SAY CABECA[W] FONT nFont2 OF oFileXLS
        NEXT W
        GO TOP
        NLINHAS=5
        WTOT=0
        ENVELOPES=0
        IMPRESSOES=0
        WCUSTO=0
        AUDI=0
        ECONOMIA=0
        qecono=0
        QECOFOLHA=0
        QFRENTEV=0
        NAOENCO=0
        DO WHILE .NOT. EOF()
          SysRefresh()
          If !fazer .And. !ENVIADO
            skip
            loop
          EndIf
          WCUSTO=CALCCORREIO()
          If !FAZER
            SKIP
            LOOP
          EndIf
          If wcusto=0 .And. !enviado
            SKIP
            LOOP
          EndIf
          If BAIXA
            SKIP
            LOOP
          EndIf
          If WCABEC=0 .And. WFAZCORREIO
            PAGE
            oprn:cmsay(0.5,0.5,"Relatório "+ALLTRIM(WEBCOR)+" Enviado em: "+DTOC(DATE())+" as "+TIME(),oFont1,,nRGB(0,0,0),,0)
            If AUDIT
              oprn:cmsay(2.0,0.3,"R",oFontv,,nRGB(0,0,0),,0)
              AUDI=AUDI+1
            EndIf
            oprn:cmsay(2.0,0.5,"Código",oFontv,,nRGB(0,0,0),,0)
            oprn:cmsay(2.0,2.0,"Nome",oFontv,,nRGB(0,0,0),,0)
            oprn:cmsay(2.0,8.5,"Endereço",oFontv,,nRGB(0,0,0),,0)
            oprn:cmsay(2.0,18.3,"Ass.",oFontv,,nRGB(0,0,0),,0)
            oprn:cmsay(2.0,19.0,"Folhas",oFontv,,nRGB(0,0,0),,0)
            oprn:cmsay(2.0,20.0,"Correio",oFontv,,nRGB(0,0,0),,0)
            WCABEC=1
            WLIN=2.2
          EndIf
          NLINHAS=NLINHAS+1
          If NOME="NÃO ENCONTRADO NO CADASTRO"
            NAOENCO=NAOENCO+1
          EndIf
          If AUDIT
            @ NLINHAS,1 XLS SAY "SIM" FONT nfont3 OF ofilexls
          EndIf
          @ NLINHAS,2 XLS SAY STRZERO(CODIGO,9) FONT nFont3 OF oFilexls
          @ NLINHAS,3 XLS SAY NOME FONT nFont3 OF oFilexls
          @ NLINHAS,4 XLS SAY ENDE1 FONT nFont3 OF oFilexls
          @ NLINHAS,5 XLS SAY ENDE2 FONT nFont3 OF oFilexls
          @ NLINHAS,6 XLS SAY ASSESSOR FONT nFont3 OF oFilexls
          @ NLINHAS,7 XLS SAY STRZERO(FOLHAS,4) FONT nfont3 OF ofilexls
          @ NLINHAS,8 XLS SAY WCUSTO FONT nFont3 FORMAT NFormat1 OF oFileXLS
          @ NLINHAS,9 XLS SAY IIF(ENVIADO,"Sim","Não") FONT nFont3 OF oFilexls
          @ NLINHAS,10 XLS SAY IIF(ENVIADO,WCUSTO,"") FORMAT nFormat1 OF oFilexls
          @ NLINHAS,11 XLS SAY IIF(ENVIADO,EMAILC,"") FONT nFont3 OF oFilexls
          @ NLINHAS,12 XLS SAY LOGMAIL FONT nFont3 OF oFilexls
          WLIN=WLIN+0.4
          If WFAZCORREIO
            If AUDIT
              oprn:cmsay(wlin,0.3,"*",oFontv,,nRGB(0,0,0),,0)
            EndIf
            oprn:cmsay(wlin,0.5,strzero(codigo,9),oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(wlin,2.0,NOME,oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(wlin,8.5,ENDE1,oFont2,,nRGB(0,0,0),,0)
            oPrn:cmsay(wlin,15.0,left(ENDE2,24),oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(wlin,18.3,strzero(assessor,5),oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(wlin,19.0,strzero(folhas,4),oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(wlin,20.0,transform(wcusto,"@E 999.99"),oFont2,,nRGB(0,0,0),,0)
          EndIf
          If FRENTEV
            QFRENTEV=QFRENTEV+1
          EndIf
          SELE 38
          SEEK ETIQ->ASSESSOR
          If !FOUND()
            ADIREG(0)
            REPLACE ASSESSOR WITH ETIQ->ASSESSOR
          EndIf
          If REGLOCK(10)
            REPLACE CIMP WITH CIMP+(WEBCUSTO/1000*ETIQ->FOLHAS)
            REPLACE CUSTO WITH CUSTO+WCUSTO
          EndIf
          SELE 37
          If fazer
            envelopes=envelopes+1
            impressoes=impressoes+folhas
          EndIf
          wtot=wtot+wcusto
          If ENVIADO
            ECONOMIA=ECONOMIA+WCUSTO
            qecono=qecono+1
            QECOFOLHA=QECOFOLHA+FOLHAS
          EndIf
          SKIP
          If !EOF() .And. WLIN>26 .And. WFAZCORREIO
            ENDPAGE
            WCABEC=0
          EndIf
        ENDDO
        If WFAZCORREIO .And. lastrec()#0
          WLIN=WLIN+0.8
          oprn:cmsay(wlin,0.5,STRZERO(ENVELOPES-NAOENCO,6)+" Envelopes, "+STRZERO(IMPRESSOES,6)+" Impressões, "+STRZERO(QFRENTEV,4)+" Iguais",oFontv,,nRGB(0,0,0),,0)
          oprn:cmsay(wlin,15.0,"Total Correio->",oFontv,,nRGB(0,0,0),,0)
          oprn:cmsay(wlin,19.0,transform(wtot,"@E 99,999.99"),oFontv,,nRGB(0,0,0),,0)
          WLIN=WLIN+0.4
          oPrn:cmsay(wlin,0.5,"Registradas: "+STRZERO(AUDI,4)+"  Não Encontrado no cadastro: "+STRZERO(NAOENCO,4),oFontv,,nRGB(0,0,0),,0)
          oPrn:cmsay(wlin,15.0,"Total Impressões->",oFontv,,nRGB(0,0,0),,0)
          oPrn:cmsay(wlin,19.0,transform(WEBCUSTO/1000*IMPRESSOES,"@E 99,999.99"),oFontv,,nRGB(0,0,0),,0)
          WLIN=WLIN+0.4
          oPrn:cmsay(wlin,15.0,"Valor a Cobrar->",oFontv,,nRGB(0,0,0),,0)
          oPrn:cmsay(wlin,19,transform((WEBCUSTO/1000*IMPRESSOES)+WTOT,"@E 99,999.99"),oFontv,,nRGB(0,0,0),,0)
          ENDPAGE
          ENDPRINT
          tini=time()
          DO WHILE .T.
            CURSORWAIT()
            DIF=TIMEASSECONDS(TIMEDIFF(tini,TIME()))
            If dif>60
              EXIT
            EndIf
            If (nHandle:=FOpen(NARQ12,FO_READ+FO_EXCLUSIVE))!=-1
              fClose(nHandle)
              EXIT
            Else
              fClose(nHandle)
            EndIf
          ENDDO
        EndIf
        NLINHAS=NLINHAS+2
        @ NLINHAS,1 XLS SAY STRZERO(ENVELOPES-NAOENCO,6)+" Envelopes, "+STRZERO(IMPRESSOES,6)+" Impressões" FONT nfont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Correio: " FONT nfont3 OF oFilexls
        @ NLINHAS,7 XLS SAY WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
        @ NLINHAS,9 XLS SAY ECONOMIA FONT nFont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS, 1 XLS SAY "Economizaria "+STRZERO(QECONO,6)+" Envelopes, "+STRZERO(QECOFOLHA,6)+" Impressões" FONT nFont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Impressões: " FONT nFont3 OF oFilexls
        @ NLINHAS,7 XLS SAY WEBCUSTO/1000*IMPRESSOES FONT nfont2 FORMAT nFormat1 OF oFilexls
        @ NLINHAS,9 XLS SAY WEBCUSTO/1000*QECONO FONT nFont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS+1,2 XLS SAY "Custo p/1.000 Impressões:  "+ALLTRIM(TRANSFORM(WEBCUSTO,"@E 999,999.99")) FONT nFont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Valor a Cobrar: " FONT nFont3 OF oFilexls
        @ NLINHAS,7 XLS SAY (WEBCUSTO/1000*IMPRESSOES)+WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
        @ NLINHAS,9 XLS SAY ECONOMIA+(WEBCUSTO/1000*QECONO) FONT nFont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS+1,2 XLS SAY "Iguais:  " +STR(QFRENTEV) FONT nFont3 OF oFilexls
        @ NLINHAS+2,2 XLS SAY "Registradas: "+STRZERO(AUDI,4) FONT nFont3 OF oFilexls
        @ NLINHAS+3,2 XLS SAY "Não Encontrados no Cadastro de Clientes: "+STRZERO(NAOENCO,4) FONT nFont3 OF oFilexls
        SELE 38
        GO TOP
        totasses=0
        totcimp=0
        nLinhas+=3
        @ NLINHAS,6 XLS SAY "Assessor" FONT nFont3 OF ofilexls
        @ NLINHAS,7 XLS SAY "Custo Correio" FONT nFont3 OF oFilexls
        @ NLINHAS,8 XLS SAY "Custo Impressão" FONT nFont3 OF oFilexls
        DO WHILE !EOF()
          nlinhas+=1
          @ NLINHAS,6 XLS SAY STRZERO(ASSESSOR->ASSESSOR,5) FONT nFont3 OF oFilexls
          @ NLINHAS,7 XLS SAY ASSESSOR->CUSTO FONT nFont3 FORMAT nFormat1 OF oFilexls
          @ NLINHAS,8 XLS SAY ASSESSOR->CIMP FONT nFont3 FORMAT nFormat1 OF oFilexls
          totasses=totasses+custo
          totcimp=totcimp+cimp
          SKIP
        ENDDO
        NLINHAS+=1
        @ NLINHAS,6 XLS SAY "Total: " FONT nFont2 OF oFilexls
        @ NLINHAS,7 XLS SAY TOTASSES FONT nfont2 FORMAT nFormat1 OF oFilexls
        @ NLINHAS,8 XLS SAY TOTCIMP FONT nfont2 FORMAT nFormat1 OF oFilexls
        nlinhas+=2
        use
        Erase &arqmw
        Erase &arqmwk
        SELE 37
        go top
        envelopes=0
        impressoes=0
        wtot=0
        NLINHAS=NLINHAS+5
        @ NLINHAS,1 XLS SAY "ENVIADAS POR EMAIL SEM PRECISAR IMPRIMIR" FONT nFont1 OF oFileXLS
        NLINHAS=NLINHAS+2
        DO WHILE .NOT. EOF()
          SysRefresh()
          If fazer
            skip
            loop
          EndIf
          If BAIXA
            SKIP
            LOOP
          EndIf
          NLINHAS=NLINHAS+1
          WCUSTO=CALCCORREIO()
          @ NLINHAS,2 XLS SAY STRZERO(CODIGO,9) FONT nFont3 OF oFilexls
          @ NLINHAS,3 XLS SAY NOME FONT nFont3 OF oFilexls
          @ NLINHAS,4 XLS SAY ENDE1 FONT nFont3 OF oFilexls
          @ NLINHAS,5 XLS SAY ENDE2 FONT nFont3 OF oFilexls
          @ NLINHAS,6 XLS SAY STRZERO(FOLHAS,4) FONT nfont3 OF ofilexls
          @ NLINHAS,7 XLS SAY WCUSTO FONT nFont3 FORMAT NFormat1 OF oFileXLS
          @ NLINHAS,8 XLS SAY EMAILC FONT nFont3 OF oFilexls
          If !EMPTY(EMAILC)
            @ NLINHAS,12 XLS SAY LOGMAIL FONT nFont3 OF oFilexls
          EndIf
          envelopes=envelopes+1
          impressoes=impressoes+folhas
          wtot=wtot+wcusto
          SKIP
        ENDDO
        NLINHAS=NLINHAS+2
        @ NLINHAS,1 XLS SAY STRZERO(ENVELOPES,6)+" Envelopes, "+STRZERO(IMPRESSOES,6)+" Impressões" FONT nfont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Correio: " FONT nfont3 OF oFilexls
        @ NLINHAS,6 XLS SAY WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS,2 XLS SAY "Custo p/1.000 Impressões:  "+ALLTRIM(TRANSFORM(WEBCUSTO,"@E 999,999.99")) FONT nFont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Impressões: " FONT nFont3 OF oFilexls
        @ NLINHAS,6 XLS SAY WEBCUSTO/1000*IMPRESSOES FONT nfont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS,4 XLS SAY "Valor economizado: " FONT nFont3 OF oFilexls
        @ NLINHAS,6 XLS SAY (WEBCUSTO/1000*IMPRESSOES)+WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
        SELE 37
        go top
        envelopes=0
        impressoes=0
        wtot=0
        NLINHAS=NLINHAS+5
        @ NLINHAS,1 XLS SAY "CLIENTES MARCADOS PARA NÃO SEREM IMPRESSOS E NÃO FORAM ENVIADOS OS EMAILS" FONT nFont1 OF oFileXLS
        NLINHAS=NLINHAS+2
        DO WHILE .NOT. EOF()
          SysRefresh()
          If !BAIXA
            SKIP
            LOOP
          EndIf
          NLINHAS=NLINHAS+1
          WCUSTO=CALCCORREIO()
          @ NLINHAS,2 XLS SAY STRZERO(CODIGO,9) FONT nFont3 OF oFilexls
          @ NLINHAS,3 XLS SAY NOME FONT nFont3 OF oFilexls
          @ NLINHAS,4 XLS SAY ENDE1 FONT nFont3 OF oFilexls
          @ NLINHAS,5 XLS SAY ENDE2 FONT nFont3 OF oFilexls
          @ NLINHAS,6 XLS SAY STRZERO(FOLHAS,4) FONT nfont3 OF ofilexls
          @ NLINHAS,7 XLS SAY WCUSTO FONT nFont3 FORMAT NFormat1 OF oFileXLS
          @ NLINHAS,8 XLS SAY EMAILC FONT nFont3 OF oFilexls
          If !EMPTY(EMAILC)
            @ NLINHAS,12 XLS SAY LOGMAIL FONT nFont3 OF oFilexls
          EndIf
          envelopes=envelopes+1
          impressoes=impressoes+folhas
          wtot=wtot+wcusto
          SKIP
        ENDDO
        NLINHAS=NLINHAS+2
        @ NLINHAS,1 XLS SAY STRZERO(ENVELOPES,6)+" Envelopes, "+STRZERO(IMPRESSOES,6)+" Impressões" FONT nfont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Correio: " FONT nfont3 OF oFilexls
        @ NLINHAS,6 XLS SAY WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS,2 XLS SAY "Custo p/1.000 Impressões:  "+ALLTRIM(TRANSFORM(WEBCUSTO,"@E 999,999.99")) FONT nFont3 OF oFilexls
        @ NLINHAS,4 XLS SAY "Total Impressões: " FONT nFont3 OF oFilexls
        @ NLINHAS,6 XLS SAY WEBCUSTO/1000*IMPRESSOES FONT nfont2 FORMAT nFormat1 OF oFilexls
        NLINHAS=NLINHAS+1
        @ NLINHAS,4 XLS SAY "Valor economizado: " FONT nFont3 OF oFilexls
        @ NLINHAS,6 XLS SAY (WEBCUSTO/1000*IMPRESSOES)+WTOT FONT nFont2 FORMAT nFormat1 OF oFilexls
      EndIf
      If WTP=2
        NARQ71=WEBLOCAL+"ARQ_MORE.TXT"
        WCONT=0
        If FILE(NARQ71)
          ARQS=FCREATE(NARQ6,0)
        EndIf
        SELE 33
        GO TOP
        WTOT=0
        @ 1,1 XLS SAY ALLTRIM(CTEXTO)+" "+DTOC(DATE())+" "+TIME() FONT nFont1 OF oFileXLS
        FOR W=1 TO LEN(CABECA)
          SysRefresh()
          @ 3,W XLS SAY CABECA[W] FONT nFont2 OF oFileXLS
        NEXT W
        NLINHAS=5
        GO TOP
        DO WHILE !EOF()
          NLINHAS=NLINHAS+1
          @ NLINHAS,1 XLS SAY NOMECLI FONT nfont3 OF ofilexls
          @ NLINHAS,2 XLS SAY TPPES FONT nfont3 OF ofilexls
          @ NLINHAS,3 XLS SAY CGCCPF FONT nfont3 OF ofilexls
          @ NLINHAS,4 XLS SAY ENVBC FONT nfont3 OF ofilexls
          @ NLINHAS,5 XLS SAY CV FONT nfont3 OF oFilexls
          @ NLINHAS,6 XLS SAY CARTAO FONT nFont3 OF oFilexls
          @ NLINHAS,7 XLS SAY MOEDA FONT nfont3 OF oFilexls
          @ NLINHAS,8 XLS SAY VAL(QUANTM) FONT nfont3 FORMAT NFormat1 OF oFilexls
          @ NLINHAS,9 XLS SAY VAL(TAXA) FONT nFont3 FORMAT nFormat3 OF oFilexls
          @ NLINHAS,10 XLS SAY VAL(REAIS) FONT nfont3 FORMAT nFormat1 of ofilexls
          @ NLINHAS,11 XLS SAY VAL(IOF) FONT nfont3 FORMAT nFormat1 OF oFilexls
          @ NLINHAS,12 XLS SAY IIF(CV="C",VAL(REAIS)-VAL(IOF),VAL(REAIS)+VAL(IOF)) FONT nfont3 FORMAT nFormat1 of ofilexls
          @ NLINHAS,13 XLS SAY NAT1+"."+NAT2+"."+NAT3+"."+NAT4+"."+NAT5 FONT nFont3 OF oFilexls
          @ NLINHAS,14 XLS SAY strzero(NRSICTUR,12) FONT nFont3 OF oFilexls
          @ NLINHAS,15 XLS SAY BCO FONT nFont3 OF oFilexls
          @ NLINHAS,16 XLS SAY AGENCIA FONT nFont3 OF oFilexls
          @ NLINHAS,17 XLS SAY CTA FONT nFont3 OF oFilexls
          @ NLINHAS,18 XLS SAY TPCTA FONT nFont3 OF oFilexls
          @ NLINHAS,19 XLS SAY REFER FONT nFont3 OF oFilexls
          @ NLINHAS,20 XLS SAY ATUALIZOU FONT nFont3 OF oFilexls
          @ NLINHAS,21 XLS SAY MOTIVO FONT nFont3 OF oFilexls
          If FILE(NARQ71)
            WCONT=WCONT+1
            MZLINHA=ALLTRIM(STR(WCONT,10,0))+"|"+ALLTRIM(STR(WCONT,10,0))+"|"+ALLTRIM(REFER)+"|"+IIF(ENVBC="R","PAGA","ERRO")+CHR(13)+CHR(10)
            FWRITE(ARQS,MZLINHA,LEN(MZLINHA))
          EndIf
          WTOT=WTOT+WEBCUSTO
          SKIP
        ENDDO
        NLINHAS=NLINHAS+2
        @ NLINHAS,9 XLS SAY "Custo Integração: " FONT nFont3 OF oFilexls
        @ NLINHAS,10 XLS SAY WTOT font NfONT3 format NfORMAT1 of oFilexls
        If FILE(NARQ71)
          FCLOSE(ARQS)
        EndIf
      EndIf

      If WTP=3
        GO TOP
        @ 1,1 XLS SAY ALLTRIM(CTEXTO)+" "+DTOC(DATE())+" "+TIME() FONT nFont1 OF oFileXLS
        FOR W=1 TO LEN(CABECA)
          SysRefresh()
          @ 4,W XLS SAY CABECA[W] FONT nFont2 OF oFileXLS
        NEXT W
        nlinhas=5
        DO WHILE !EOF()
          NLINHAS=NLINHAS+1
          @ NLINHAS,1 XLS SAY ARQUIVO FONT nFont3 OF oFilexls
          @ NLINHAS,2 XLS SAY NOME FONT nFont3 OF oFilexls
          @ NLINHAS,3 XLS SAY CEP FONT nFont3 OF oFilexls
          @ NLINHAS,4 XLS SAY LOGRAD FONT nFont3 OF oFilexls
          @ NLINHAS,5 XLS SAY NUMERO FONT nFont3 OF oFilexls
          @ NLINHAS,6 XLS SAY COMPL FONT nFont3 OF oFilexls
          @ NLINHAS,7 XLS SAY BAIRRO FONT nFont3 OF oFilexls
          @ NLINHAS,8 XLS SAY CIDADE FONT nFont3 OF oFilexls
          @ NLINHAS,9 XLS SAY UF FONT nFont3 OF oFilexls
          @ NLINHAS,10 XLS SAY CPF FONT nFont3 OF oFilexls
          @ NLINHAS,11 XLS SAY DDD FONT nFont3 OF oFilexls
          @ NLINHAS,12 XLS SAY FONE FONT nFont3 OF oFilexls
          @ NLINHAS,13 XLS SAY DDDC FONT nFont3 OF oFilexls
          @ NLINHAS,14 XLS SAY CELUL FONT nFont3 OF oFilexls
          @ NLINHAS,15 XLS SAY DOCTO FONT nFont3 OF oFilexls
          @ NLINHAS,16 XLS SAY DTDOCTO FONT nFont3 OF oFilexls
          @ NLINHAS,17 XLS SAY DTNASC FONT nFont3 OF oFilexls
          @ NLINHAS,18 XLS SAY RESULT FONT nFont3 OF oFilexls
          @ NLINHAS,19 XLS SAY SITUA FONT nFont3 OF oFilexls
          SKIP
        ENDDO
      EndIf

      XLS PAGE BREAK AT NLINHAS OF oFileXLS

      SET XLS TO DISPLAY OF oFileXLS
      NFOT=EMP2+" - Relatório gerado em "+DTOC(DATE()) + " as "+TIME()
      SET XLS TO PRINTER ;
        HEADER "";
        FOOTER NFOT;
        TOP MARGIN 0.3 ;
        BOTTOM MARGIN 1.1;
        LEFT MARGIN 0.1 ;
        RIGHT MARGIN 0.1 ;
        OF oFileXLS
      ENDXLS oFileXLS
    return  ARQ_XLS

    Procedure GERARELPDF(QUALR)
      If QUALR=1
        cTitle="Relatório Ideal Standard"
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        MLOCAL=ARQ_PDF
        Erase &MLOCAL
        oReg:Set("PDFFilename",MLOCAL)
        oReg:Close()
        PRINT oPrn NAME CTITLE TO TEMPDF1
        oPrn:SetLandscape()
        DEFINE FONT nFont1 NAME "Arial" SIZE 0,16 BOLD ITALIC OF oPrn
        DEFINE FONT nFont2 NAME "courier new negrito" UNDERLINE BOLD SIZE 0,7 OF oPrn
        DEFINE FONT nFont3 NAME "courier new" SIZE 0,7 OF oPrn
        DEFINE FONT nFont4 NAME "Courier new negrito" SIZE 0,7 BOLD OF oPrn
        DEFINE FONT nFont5 NAME "Arial" SIZE 0,12 BOLD OF oPrn
        nRowStep = oPrn:nVertRes() / 66
        nColStep = oPrn:nHorzRes() / 80
        WTOTD:=WTOTM:=WTOTMDM:=WTOTDMA:=WTOTDMG:=WTOTT:=0
        M->ME_SES="Janeiro  FevereiroMarço    Abril    Maio     Junho    Julho    Agosto   Setembro Outubro  Novembro Dezembro "
        CABEC:=CABEC1:=0
        ANTES=LEFT(CHAVE,7)
        MESANTES=LEFT(CHAVE,6)
        ANOANTES=LEFT(CHAVE,4)
        DO WHILE .NOT. EOF()
          SysRefresh()
          If CABEC1=0
            PAGE
            oPrn:Say(1*nRowStep,nColstep*3,ZOMECORR,nFont1)
            oPrn:Say(3*nRowStep,nColstep*15,"Relatório de pagamentos a serem efetuados                Atualizado em "+DTOC(DATE())+" as  "+TIME(),nFont5)
            CABEC1=1
            NLINHAS=5
          EndIf
          If CABEC=0
            If TIPO="D"
              TIPW:="Distribuição - "
            ElseIf TIPO="M"
              TIPW:="Manufatura - "
            ElseIf TIPO="T"
              TIPW:="Terceiros -  "
            EndIf
            oPrn:Say(NLINHAS*nRowStep,ncolstep*3,TIPW+TRIM(SUBS(ME_SES,VAL(SUBSTR(CHAVE,5,2))*9-8,9))+"/"+LEFT(CHAVE,4),nFont1)
            NLINHAS=NLINHAS+2
            oPrn:Say(nLinhas*nRowStep,ncolstep*3,"Ref Cliente                       DI/CI             Fatura            Conhecto.                          Embarque             Valor       Valor da DI  Vencimento    Vcto180 D   Doctos.Faltantes  Exportador",nFont4)
            CABEC=1
            NLINHAS=NLINHAS+2
            ANTES=LEFT(CHAVE,7)
          EndIf
          If !EMPTY(TMPS->DIRE)
            ZDIRE=TMPS->DIRE
          Else
            ZDIRE=TMPS->MODALID
          EndIf
          MTUDO=TMPS->REFCLI+"  "+ZDIRE+"  "+TMPS->FAT1+"  "+TMPS->CONHEC+"  "+DTOC(TMPS->DTCON)+"  "+TRANSFORM(TMPS->VALOR,"@E 999,999,999.99")+"  "+TRANSFORM(TMPS->VALOR,"@ZE 999,999,999.99")+"  "+DTOC(TMPS->VENCTO)+"  "+DTOC(TMPS->PRZVINC)+"  "+LEFT(TMPS->EXIGENCIA,15)+"  "+LEFT(TMPS->NOMEPAGREC,15)
          oPrn:Say(NLINHAS*nRowStep,nColstep*3,MTUDO,nFont3)
          If TMPS->TIPO="D"
            WTOTD+=TMPS->VALOR
          ElseIf TMPS->TIPO="M"
            WTOTM+=TMPS->VALOR
          Else
            WTOTT+=TMPS->VALOR
          EndIf
          WTOTMDM+=TMPS->VALOR
          WTOTDMA+=TMPS->VALOR
          WTOTDMG+=TMPS->VALOR
          NLINHAS=NLINHAS+1
          SKIP
          If LEFT(CHAVE,7)#ANTES .OR. EOF()
            NLINHAS=NLINHAS+1
            CABEC=0
            If SUBSTR(ANTES,7,1)="D"
              oPrn:Say(NLINHAS*nRowStep,nColstep*3,"Valor Total de Distribuição para o Mes  "+(SUBS(ME_SES,VAL(SUBSTR(ANTES,5,2))*9-8,9))+"/"+LEFT(ANTES,4)+SPACE(64)+TRANSFORM(WTOTD,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTD,"@ZE 999,999,999.99"),nFont2)
              WTOTD=0
            ElseIf SUBSTR(ANTES,7,1)="M"
              oPrn:Say(NLINHAS*nRowStep,nColstep*3,"Valor Total de Manufatura para o Mes  "+(SUBS(ME_SES,VAL(SUBSTR(ANTES,5,2))*9-8,9))+"/"+LEFT(ANTES,4)+SPACE(66)+TRANSFORM(WTOTM,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTM,"@ZE 999,999,999.99"),nFont2)
              WTOTM=0
            Else
              oPrn:Say(NLINHAS*nRowStep,nColstep*3,"Valor Total de Terceiros  para o Mes  "+(SUBS(ME_SES,VAL(SUBSTR(ANTES,5,2))*9-8,9))+"/"+LEFT(ANTES,4)+SPACE(66)+TRANSFORM(WTOTT,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTT,"@ZE 999,999,999.99"),nFont2)
              WTOTT=0
            EndIf
            ANTES=LEFT(CHAVE,7)
            NLINHAS=NLINHAS+2
          EndIf
          If CHAVE#MESANTES .OR. EOF()
            CABEC=0
            oPrn:Say(NLINHAS*nRowStep,nColStep*3,"Valor Total do Mes  "+(SUBS(ME_SES,VAL(SUBSTR(MESANTES,5,2))*9-8,9))+"/"+LEFT(MESANTES,4)+SPACE(84)+TRANSFORM(WTOTMDM,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTMDM,"@ZE 999,999,999.99"),nFont2)
            WTOTMDM=0
            MESANTES=LEFT(CHAVE,6)
            NLINHAS=NLINHAS+2
          EndIf
          If CHAVE#ANOANTES .OR. EOF()
            CABEC=0
            oPrn:Say(NLINHAS*nRowStep,nColStep*3,"Valor Total do Ano de "+ANOANTES+SPACE(92)+TRANSFORM(WTOTDMA,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTDMA,"@ZE 999,999,999.99"),nFont2)
            WTOTDMA=0
            ANOANTES=LEFT(CHAVE,4)
            NLINHAS=NLINHAS+2
          EndIf
          If EOF()
            CABEC=0
            oPrn:Say(NLINHAS*nRowStep,nColStep*3,"Total Geral "+space(106)+TRANSFORM(WTOTDMG,"@ZE 999,999,999.99")+"   "+TRANSFORM(WTOTDMG,"@ZE 999,999,999.99"),nFont2)
            NLINHAS=NLINHAS+3
          EndIf
          If NLINHAS>50
            ENDPAGE
            CABEC1=0
          EndIf
        ENDDO
        If CABEC1=0
          ENDPAGE
        EndIf
        ENDPRINT
        DO WHILE .T.
          CURSORWAIT()
          If (nHandle:=FOpen(MLOCAL,FO_READ+FO_EXCLUSIVE))!=-1
            fClose(nHandle)
            EXIT
          Else
            fClose(nHandle)
          EndIf
        ENDDO
      EndIf
    return 

    Procedure CMOV1()
      SELE 26
      WARQ=M->LOCAL+"CONTT.DBF"
      If !USEREDE(WARQ,.F.,10,"CONTR","S")
        return  .F.
      EndIf
      If FCOUNT()<49
        USE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=M->LOCAL+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN4:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        mat_stru={}
        aAdd(MAT_STRU,{"CONGLO","N",6,0})
        aAdd(mat_stru,{"NOMEPAG","C",69,0})
        aAdd(mat_stru,{"NOMEPRACA","C",20,0})
        aAdd(mat_stru,{"SIMBOLMOE","C",5,0})
        aAdd(mat_stru,{"NOMENAT","C",51,0})
        aAdd(mat_stru,{"GRUPONAT","C",2,0})
        aAdd(mat_stru,{"ANEXO","M",10,0})
        aAdd(mat_stru,{"OUTRAS","M",10,0})
        aAdd(mat_stru,{"LIQUIDACAO","D",8,0})
        aAdd(mat_stru,{"CERTAUT","C",10,0})
        aAdd(mat_stru,{"FIRCE","C",13,0})
        aAdd(mat_stru,{"ENTREGADOC","D",8,0})
        aAdd(mat_stru,{"VLMOEDNAC","C",18,0})
        aAdd(mat_stru,{"PAGREC","N",6,0})
        aAdd(mat_stru,{"NR_CONTRAT","N",  12,  0})
        aAdd(mat_stru,{"INTERNO","N",8,0})
        aAdd(mat_stru,{"TIPO","C",1,0})
        aAdd(mat_stru,{"PRACA","C",4,0})
        aAdd(mat_stru,{"BANCO","N",6,0})
        aAdd(mat_stru,{"NOMEBANCO","C",25,0})
        aAdd(mat_stru,{"CODCLI","N",6,0})
        aAdd(mat_stru,{"NOMECLI","C",59,0})
        aAdd(mat_stru,{"DTCONTR","D",8,0})
        aAdd(mat_stru,{"CODMOEDA","N",3,0})
        aAdd(mat_stru,{"VALMOEDA","C",18,0})
        aAdd(mat_stru,{"CGCCPF","C",14,0})
        aAdd(mat_stru,{"CGCCLI","C", 14,  0})
        aAdd(mat_stru,{"CPFCLI","C", 11,  0})
        aAdd(mat_stru,{"INSTITU","N",5,0})
        aAdd(mat_stru,{"MERCADO","C",1,0})
        aAdd(mat_stru,{"TXCAMBIAL","C",14,0})
        aAdd(mat_stru,{"NATOPERA","N",12,0})
        aAdd(mat_stru,{"INST","N",5,0})
        aAdd(mat_stru,{"DESPBANCO","C",40,0})
        aAdd(mat_stru,{"ASSINA1","C",30,0})
        aAdd(mat_stru,{"ASSINA2","C",30,0})
        aAdd(mat_stru,{"ASSINA3","C",30,0})
        aAdd(mat_stru,{"ASSINA4","C",30,0})
        aAdd(mat_stru,{"ASSINA5","C",30,0})
        aAdd(mat_stru,{"ASSINA6","C",30,0})
        aAdd(mat_stru,{"POSICAO","C",50,0})
        aAdd(mat_stru,{"NDA","N",1,0})
        AADD(MAT_STRU,{"DIRE","C",12,0})
        AADD(MAT_STRU,{"FATURA","C",12,0})
        AADD(MAT_STRU,{"MERCADORIA","C",20,0})
        AADD(MAT_STRU,{"MODALIDADE","C",20,0})
        AADD(MAT_STRU,{"CONHECIMENTO","C",20,0})
        AADD(MAT_STRU,{"OBS","C",50,0})
        AADD(MAT_STRU,{"NOVOC","C",1,0})
        DBCREATE(ARQ_PRN1,mat_stru,"DBFNTX")
        ARQ=M->LOCAL+"CONTT.DBF"
        ARQ1=M->LOCAL+"CONTT.DBT"
        SELE 26
        If !USEREDE(ARQ_PRN1,.T.,10,"TTMP","S")
          return  .F.
        EndIf
        APPEND FROM &ARQ VIA "DBFNTX"
        USE
        Erase &ARQ
        Erase &ARQ1
        FRENAME(ARQ_PRN1,ARQ)
        FRENAME(ARQ_PRN4,ARQ1)
      EndIf
      SELE 26
      USE
    return  .T.

    Function OUTROS(TEXTO,VARIA,MASC)
      Local oPar1,oBrush51,oFont
      MFONTES()
      DEFINE DIALOG oPar1 FROM 1, 1 TO 13.5 , 55 TITLE "Outros" BRUSH oBrushTela STYLE WS_CAPTION
      @   0.44,7.5 say "" box raised COLOR CLR_BLACK,nRGB(248,248,248) size 135,50 OF oPar1
      @   0.87,11 say Texto FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248) SIZE 100,13
      @ 3.0,09.5 get VARIA picture masc FONT oFontget size 75,10 OF oPar1
      @ 70,60 SBUTTON oBtn PIXELS PROMPT "&Confirma"  TEXT ON_CENTER RESOURCE "OFF","ON","DISABLE" NOBOX  ADJUST FONT oFontbut OF oPar1 size 40,12 action (oPar1:end()) DEFAULT
      oPar1:lhelpicon:=.f.
      ACTIVATE DIALOG oPar1 CENTERED
    return (VARIA)

    Procedure BUSCAXML()
      CURSORWAIT()
      LINK=""
      CMACRO="wget.exe -T30 -olog.txt -ONOTICA.XML "+ALLTRIM(LINK)
      WAITRUN(CMACRO,0)
      CURSORWAIT()
      MEMOWRIT("NOTIC.XML",Strtran(MEMOREAD("NOTICA.XML"),CHR(10),CHR(13)+CHR(10)))
      Erase "NOTICA.XML"
      oText:=TTxtfile():NEW("NOTIC.XML")
      CONTAD=0
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)#"N"
        CONTAD=1
      EndIf
      TI=0
      If CONTAD>50
        CONTAD=50
      EndIf
      FOR TI=1 TO CONTAD
        SysRefresh()
        ttt=alltrim(oText:Readline())
        If AT("<item>",TTT)=0
          If CONTAD>1
            oText:Skip()
            LOOP
          EndIf
        EndIf
        oText:Skip()
        ZATA=""
        ZORA=""
        ZANCHETE=""
        ZINK1=""
        ZEXTO1=""
        TTT=ALLTRIM(oText:Readline())
        ZANCHETE=Strtran(TTT,"<title>","")
        ZANCHETE=Strtran(ZANCHETE,"</title>","")
        oText:Skip()
        TTT=ALLTRIM(oText:Readline())
        ZINK1=Strtran(TTT,"<link>","")
        ZINK1=Strtran(zink1,"</link>","")
        ZINK1=Strtran(zink1,CHR(10),"")
        ZINK1=Strtran(zink1,CHR(13),"")
        ZINK1=SUBSTR(ZINK1,AT("*",ZINK1)+1,LEN(ZINK1))
        SELE 1
        SET ORDER TO 2
        SEEK ZINK1
        If EOF()
          CMACRO="wget.exe -T10 -olog.txt -ODESC1.HTM "+ALLTRIM(ZINK1)
          WAITRUN(CMACRO,0)
          CURSORWAIT()
          MEMOWRIT("DESC.HTM",Strtran(MEMOREAD("DESC1.HTM"),CHR(10),CHR(13)+CHR(10)))
          Erase "DESC1.HTM"
          TUDO=MEMOREAD("DESC.HTM")
          Erase DESC.HTM
          TUDO=SUBSTR(TUDO,AT("<!--DATA-->",TUDO),LEN(TUDO))
          ZATA=SUBSTR(TUDO,12,10)
          TUDO=SUBSTR(TUDO,AT("<!--HORA-->",TUDO),LEN(TUDO))
          ZORA=SUBSTR(TUDO,12,2)+":"+SUBSTR(TUDO,15,2)
          TUDO=SUBSTR(TUDO,AT("<!--TEXTO-->",TUDO),LEN(TUDO))
          If AT("<!--LINKS-->",TUDO)#0
            ZEXTO1=SUBSTR(TUDO,15,AT("<!--LINKS-->",TUDO)-14)
          Else
            ZEXTO1=SUBSTR(TUDO,15,AT("<!--/TEXTO-->",TUDO)-14)
          EndIf
          ZEXTO1=Strtran(ZEXTO1,"<b>","")
          ZEXTO1=Strtran(ZEXTO1,"</b>","")
          ZEXTO1=Strtran(ZEXTO1,"<B>","")
          ZEXTO1=Strtran(ZEXTO1,"</B>","")
          ZEXTO1=Strtran(ZEXTO1,"<BR>",CHR(13)+CHR(10))
          ZEXTO1=Strtran(ZEXTO1,"<br>",CHR(13)+CHR(10))
          ZEXTO1=Strtran(ZEXTO1,"&quot;","'")
          ZEXTO1=Strtran(ZEXTO1,'<span class="tagline">',"")
          ZEXTO1=Strtran(ZEXTO1,"</span>","")
          ZEXTO1=Strtran(ZEXTO1,"</a>","")
          ZEXTO1=Strtran(ZEXTO1,"<i>","")
          ZEXTO1=Strtran(ZEXTO1,"</i>","")
          CTD=0
          DO WHILE AT("<a href",ZEXTO1)#0
            SysRefresh()
            CTD=CTD+1
            NAS=AT("<a href",ZEXTO1)-1
            ZEXTO=LEFT(ZEXTO1,NAS)
            NAS1=AT(">",ZEXTO1)+1
            ZEXTO2=SUBSTR(ZEXTO1+SPACE(300),NAS1,LEN(ZEXTO1)-LEN(ZEXTO))
            ZEXTO1=ZEXTO+ZEXTO2
            If CTD=50
              EXIT
            EndIf
          ENDDO
          zanchete=Strtran(ZANCHETE,"&quot;","'")
          SELE 3
          ZTNOT=CTOD(ZATA)
          ZRNOT=ZORA+SPACE(3)
          ZEXTO=ANSITOOEM(ZEXTO1)
          ZGENCIA=1
          ZSSUNTO=1
          ZRIORID=1
          ZANCHETE=ALLTRIM(ANSITOOEM(ZANCHETE))
          ZINK=ZINK1
          ZESCAG=DESC
          ZRDIAS=NRDIAS
          ZESCASS="TODAS"
          SELE 1
          SET ORDER TO 1
          SEEK DTOS(ZTNOT)+ZRNOT+STRZERO(ZGENCIA,3)
          If EOF() .And. !EMPTY(ZTNOT) .And. !EMPTY(ZEXTO)
            ADIREG(0)
            ZDNOT=RECNO()
            PUTREC()
          EndIf
        EndIf
      NEXT TI
      oText:CLOSE()
      Erase "NOTIC.XML"
      SELE 1
      SET ORDER TO 1
      CURSORARROW()
    return 

    Procedure SENTIMOS()
      DECLARE TX1[1]
      TX1[1]="Cambio Net - Clientes Inativos"
      LOGFILE(WEBLOCAL+"ATUALIZ.TXT",TX1)
      SELE 22
      WCORRET=STRZERO(CORRETORA,2)
      MTEXTO:=MEMOREAD(ARQDB)
      Erase &ARQDB
      M->FROM=""
      If FILE(WEBLOCAL+"QUEM.FRO")
        SELE 25
        If USEREDE(WEBLOCAL+"QUEM.FRO",.F.,10,"TMP")
          GO TOP
          M->FROM=ALLTRIM(FROM)
        EndIf
        USE
      Else
        SELE 22
        return 
      EndIf
      SELE 22
      If EMPTY(M->FROM)
        return 
      EndIf
      REGS=RECNO()
      SELE 24
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If SUBSTR(USUARIO,7,2)=WCORRET
          If DATE()-ULTACESSO>90 .And. !EMPTY(LOWER(ALLTRIM(EMAIL)))
            SEENVIAEMAIL="S"
            webcor=" "
            ENVIAEMAIL(LOWER(ALLTRIM(EMAIL)),"CambioNet "+ZNOMECORR,"Ref: "+TRIM(USUARIO)+" Ultimo Acesso: "+DTOC(ULTACESSO)+"<br>"+MTEXTO,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
        SKIP
      ENDDO
      SELE 22
      GO REGS
    return 
    //---------------------------------------------------------------------------//
    Function INDI1()
      CLOSE DATABASE
      SELE 6
      If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
        CLOSE DATABASE
        return 
      EndIf
      SET ORDER TO
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        CURSORWAIT()
        wregs=recno()
        _LOCAL=LOCALCART+A01CODE+"/"
        NOMEARQS(1)
        CLOSE DATABASE
        CRIARQC()
        CLOSE DATABASE
        criamovc()
        CLOSE DATABASE
        INDICESC()
        close database
        SELE 6
        If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
          CLOSE DATABASE
          return 
        EndIf
        go wregs
        SKIP
      ENDDO
      CLOSE DATABASE
      CURSORARROW()
      MSGINFO("TERMINADO !")
    return 

    Procedure CRIAMOVC()
      FAZER=.F.
      CLOSE DATABASE
      CURSORWAIT()
      SELE 1
      If !USEREDE(ARQ06,.F.,10,"BDIS")
        MSGStop("O arquivo BDI não está  disponível",WSIST)
        return 
      EndIf
      If FCOUNT()<4
        CLOSE DATABASE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN2:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        MAT_STRU={}
        Aadd(mat_stru,{"DATA","D",8,0})
        Aadd(mat_stru,{"CODPAPEL","C",12,0})
        Aadd(mat_stru,{"PRECOM","N",19,5})
        AADD(MAT_STRU,{"POP","C",1,0})
        DBCREATE(ARQ_PRN1,mat_stru)
        FAZER=.T.
        SELE 1
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
          MSGINFO("Não foi possivel continuar !","Tente mais tarde")
          CLOSE DATABASE
          QUIT
        EndIf
        ARQ=ARQ06
        APPEND FROM &ARQ
        CLOSE DATABASE
        Erase &ARQ
        FRENAME(ARQ_PRN1,ARQ)
        CLOSE DATABASE
      EndIf
      SELE 1
      If !USEREDE(ARQ14,.F.,10,"PAPEIS")
        MSGStop("O arquivo Papeis não está  disponível",WSIST)
        return 
      EndIf
      If FCOUNT()<29
        CLOSE DATABASE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN2:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        MAT_STRU={}
        AADD(MAT_STRU,{"a03codp","C",12,0})  // Codigo do papel
        AADD(MAT_STRU,{"a03desc","C",40,0})  // Descricao
        AADD(MAT_STRU,{"a03tipo","C",01,0})  // Tipo de papel (1-A Vista 2-Opcao 3-A Termo)
        AADD(MAT_STRU,{"a03codt","C",02,0})  // Codigo da tabela
        AADD(MAT_STRU,{"a03lmil","C",01,0})  // Lote de mil (S/N)
        AADD(MAT_STRU,{"a03pm","N",19,2})    // Preco de Mercado
        AADD(MAT_STRU,{"A03DTBD","D",8,0})   // Data Ult.Movimento
        AADD(MAT_STRU,{"A03DTOP","D",8,0})   // Data de Validade da Opção
        AADD(MAT_STRU,{"A03CODPO","C",12,0}) // Código do Papel a Vista
        AADD(MAT_STRU,{"A03POP","C",1,0}) //    S=POP
        DBCREATE(ARQ_PRN1,mat_stru)
        FAZER=.T.
        SELE 1
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
          MSGINFO("Não foi possivel continuar !","Tente mais tarde")
          CLOSE DATABASE
          QUIT
        EndIf
        ARQ=ARQ14
        APPEND FROM &ARQ
        CLOSE DATABASE
        Erase &ARQ
        FRENAME(ARQ_PRN1,ARQ)
        CLOSE DATABASE
      EndIf
      SELE 1
      If !USEREDE(ARQ03,.F.,10,"SALDO")
        MSGStop("O arquivo Saldo Extra não está  disponível",WSIST)
        return 
      EndIf
      If FCOUNT()<29
        CLOSE DATABASE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN2:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        MAT_STRU={}
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
        DBCREATE(ARQ_PRN1,mat_stru)
        FAZER=.T.
        SELE 1
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
          MSGINFO("Não foi possivel continuar !","Tente mais tarde")
          CLOSE DATABASE
          QUIT
        EndIf
        ARQ=ARQ03
        APPEND FROM &ARQ
        CLOSE DATABASE
        Erase &ARQ
        FRENAME(ARQ_PRN1,ARQ)
        SELE 1
        If !USEREDE(ARQ03,.F.,10,"SALDO")
          MSGStop("O arquivo Saldo Extra não está  disponível",WSIST)
          return 
        EndIf
      EndIf
      SET ORDER TO 0
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If A03CODP="INHA3" .And. REGLOCK(10)
          REPLACE A03CODP WITH "GFSA3"
        EndIf
        If A03CODP="ABNOTE" .And. REGLOCK(10)
          REPLACE A03CODP WITH "ABNB3"
        EndIf
        If A03CODP="PLIM4" .And. REGLOCK(10)
          REPLACE A03CODP WITH "NETC4"
        EndIf
        If A03CODP="LIGH3" .And. REGLOCK(10)
          REPLACE A03CODP WITH "LIGT3"
        EndIf
        SKIP
      ENDDO
      CLOSE DATABASE
      If !USEREDE(ARQ10,.F.,10,"MOVEXTRA")
        MSGStop("O arquivo Movimento Extra não está  disponível",WSIST)
        return 
      EndIf
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If A10CODP="INHA3" .And. REGLOCK(10)
          REPLACE A10CODP WITH "GFSA3"
        EndIf
        If A10CODP="ABNOTE" .And. REGLOCK(10)
          REPLACE A10CODP WITH "ABNB3"
        EndIf
        If A10CODP="PLIM4" .And. REGLOCK(10)
          REPLACE A10CODP WITH "NETC4"
        EndIf
        If A10CODP="LIGH3" .And. REGLOCK(10)
          REPLACE A10CODP WITH "LIGT3"
        EndIf
        SKIP
      ENDDO
      If FCOUNT()<16
        CLOSE DATABASE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN2:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        MAT_STRU={}
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
        DBCREATE(ARQ_PRN1,MAT_STRU)
        FAZER=.T.
        SELE 1
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
          MSGINFO("Não foi possivel continuar !","Tente mais tarde")
          CLOSE DATABASE
          QUIT
        EndIf
        ARQ=ARQ10
        APPEND FROM &ARQ
        CLOSE DATABASE
        Erase &ARQ
        FRENAME(ARQ_PRN1,ARQ)
      EndIf
      CLOSE DATABASE
      If !USEREDE(ARQ11,.F.,10,"MOVNOTA")
        MSGStop("O arquivo Movimento de Notas não está  disponível",WSIST)
        return 
      EndIf
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If A11CODP="INHA3" .And. REGLOCK(10)
          REPLACE A11CODP WITH "GFSA3"
        EndIf
        If A11CODP="ABNOTE" .And. REGLOCK(10)
          REPLACE A11CODP WITH "ABNB3"
        EndIf
        If A11CODP="PLIM4" .And. REGLOCK(10)
          REPLACE A11CODP WITH "NETC4"
        EndIf
        If A11CODP="LIGH3" .And. REGLOCK(10)
          REPLACE A11CODP WITH "LIGT3"
        EndIf
        SKIP
      ENDDO
      If FCOUNT()<16
        CLOSE DATABASE
        DO WHILE .T.
          SysRefresh()
          M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
          ARQ_PRN=LOCALCART+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          ARQ_PRN2:=ARQ_PRN+".DBT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        MAT_STRU={}
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
        DBCREATE(ARQ_PRN1,MAT_STRU)
        FAZER=.T.
        SELE 1
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP","S")
          MSGINFO("Não foi possivel continuar !","Tente mais tarde")
          CLOSE DATABASE
          QUIT
        EndIf
        ARQ=ARQ11
        APPEND FROM &ARQ
        CLOSE DATABASE
        Erase &ARQ
        FRENAME(ARQ_PRN1,ARQ)
      EndIf
      If FAZER
        INDICESC()
      EndIf
      CLOSE DATABASE
    return 

    Function INDI21(Q)
      CLOSE DATABASE
      SELE 6
      If !USEREDE(LOCALCART+"EMPRESAS.DBF",.F.,10,"CARTEMP")
        CLOSE DATABASE
        return 
      EndIf
      SET ORDER TO
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If REGLOCK(10)
          REPLACE A01CORRET WITH 1
        EndIf
        SKIP
      ENDDO
      CLOSE DATABASE
      If Q=1
        MSGINFO("TERMINADO !")
      EndIf
    return 

    Func Caja(nArriba,nIzq,nAbajo,nDerecha,oPrn,nTipo,oBrush,oPen)
      Local xCor := {} , yCor := {}

      // Pasamos coordenadas de cms a pixel
      xCor := oPrn:Cmtr2Pix(nArriba,nIzq)
      yCor := oPrn:Cmtr2Pix(nAbajo,nDerecha)

      DO CASE
        CASE nTipo == 0       // Caja Vacia
          oPrn:Box(xCor[1],xCor[2],yCor[1],yCor[2],oPen)
        CASE nTipo == 1       // Caja rellena
          oPrn:FillRect({xCor[1],xCor[2],yCor[1],yCor[2]},oBrush)
      ENDCASE
    return  NIL

    Procedure ATUANOT()
      cmacro="@ECHO OFF"+CHR(13)+CHR(10)+"wget.exe -T10 -olog.txt -ONOTIC.DAT http://agenciabrasil.ebc.com.br/web/ebc-agencia-brasil/enviorss/-/journal/rss/19523/187401?doAsGroupId=19523&refererPlid=71705 -U --user-agent=MOZILLA"+CHR(13)+CHR(10)+"DEL AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
      MEMOWRIT("AGUARDE.TXT","rotina 18"+CMACRO)
      MEMOWRIT("FZNOT.BAT",CMACRO)
      CMACRO="FZNOT.BAT"
      waitrun(CMACRO,0)
      DO while file("AGUARDE.TXT")
      ENDDO
      CURSORWAIT()
      Erase FZNOT.BAT
      WTUDO=MEMOREAD("NOTIC.DAT")
      WTUDO=Strtran(WTUDO,"<entry>",chr(13)+chr(10))
      WTUDO=Strtran(WTUDO,"<published>",chr(13)+chr(10))
      WTUDO=Strtran(WTUDO,"<summary type",CHR(13)+CHR(10)+"<summary type")
      WTUDO=Strtran(WTUDO,"</summary>","</summary>"+CHR(13)+CHR(10))
      WTUDO=Strtran(WTUDO,"&lt;br /&gt;&lt;br /&gt;",CHR(13)+CHR(10))
      WTUDO=Strtran(WTUDO,"&ldquo;","")
      MEMOWRIT("NOTIC.DAT",WTUDO)
      oText:=TTxtfile():NEW("NOTIC.DAT")
      CONTAD=0
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)#"N"
        CONTAD=1
      EndIf
      DO WHILE !oText:eof()
        ttt=oText:Readline()
        If left(ttt,8)#" <title>"
          oText:Skip()
          loop
        EndIf
        inicio=AT("<title>",TTT)+7
        final=at("</title>",TTT)-9
        ZANCHETE=SUBSTR(TTT,inicio,FINAL)
        oText:Skip()
        ttt=oText:Readline()
        ZTNOT=CTOD(SUBSTR(TTT,9,2)+"/"+SUBSTR(TTT,6,2)+"/"+LEFT(TTT,4))
        ZRNOT=VAL(SUBSTR(TTT,12,2))-3
        If ZRNOT=-2
          ZRNOT=1
        ElseIf ZRNOT=-1
          ZRNOT=2
        ElseIf ZRNOT=0
          ZRNOT=3
        EndIf
        ZRNOT=STRZERO(ZRNOT,2)+SUBSTR(TTT,14,3)
        ZGENCIA=1
        ZSSUNTO=1
        ZRIORID=1
        ZESCAG="AGENCIA BRASIL"
        ZGENCIA=1
        ZRDIAS=LOCALdiasnot
        ZESCASS="TODAS"
        SELE 1
        set order to 1
        If ZTNOT<>CTOD("")
          SEEK DTOS(ZTNOT)+ZRNOT+SPACE(3)+STRZERO(ZGENCIA,3)
          If FOUND()
            oText:Skip()
            LOOP
          EndIf
        EndIf
        oText:Skip()
        NOTICIA=""
        DO WHILE .T.
          NOTICIA=NOTICIA+oText:Readline()+CHR(13)+CHR(10)
          If AT("</summary>",NOTICIA)#0 .OR. otext:eof() .OR. AT(" <title>",NOTICIA)#0
            EXIT
          EndIf
          oText:Skip()
        ENDDO
        NOTICIA=Strtran(NOTICIA,'<summary type="html">',"")
        NOTICIA=Strtran(NOTICIA,'</summary>',"")
        NOTICIA=Strtran(NOTICIA,"&lt;","<")
        NOTICIA=Strtran(NOTICIA,"&gt;",">")
        NOTICIA=Strtran(NOTICIA,"&amp;ndash;","-")
        NOTICIA=Strtran(NOTICIA,"&amp;nbsp;"," ")
        MEMOWRIT("NOTA1.TXT",ZANCHETE)
        CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+"iconv.exe -c -s -f utf-8 -t iso-8859-1 nota1.txt > nota2.txt"+CHR(13)+CHR(10)+"EXIT"
        MEMOWRIT("FZ2.BAT",CMACRO)
        CMACRO="FZ2.BAT"
        WAITRUN(CMACRO,0)
        CURSORWAIT()
        Erase FZ2.BAT
        ZEXTO1=MEMOREAD("nota2.txt")
        Erase NOTA2.TXT
        ZEXTO=Strtran(ZEXTO1,'"','""')
        ZEXTO1=Strtran(ZEXTO,"'","''")
        ZEXTO=Strtran(ZEXTO1,CHR(10),"")
        ZEXTO1=Strtran(ZEXTO,CHR(13),"")
        ZEXTO=Strtran(ZEXTO1,"<p>","")
        ZEXTO1=Strtran(ZEXTO,"</p>",CHR(13)+CHR(10))
        ZEXTO=ZEXTO1
        ZANCHETE=ANSITOOEM(ZEXTO)
        MEMOWRIT("NOTA1.TXT",NOTICIA)
        CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+"iconv.exe -c -s -f utf-8 -t iso-8859-1 nota1.txt > nota2.txt"+CHR(13)+CHR(10)+"EXIT"
        MEMOWRIT("FZ2.BAT",CMACRO)
        CMACRO="FZ2.BAT"
        WAITRUN(CMACRO,0)
        CURSORWAIT()
        Erase FZ2.BAT
        ZEXTO1=MEMOREAD("nota2.txt")
        Erase NOTA2.TXT
        ZEXTO=Strtran(ZEXTO1,'"','""')
        ZEXTO1=Strtran(ZEXTO,"'","''")
        ZEXTO=Strtran(ZEXTO1,CHR(10),"")
        ZEXTO1=Strtran(ZEXTO,CHR(13),"")
        ZEXTO=Strtran(ZEXTO1,"<p>","")
        ZEXTO1=Strtran(ZEXTO,"</p>",CHR(13)+CHR(10))
        ZEXTO=ZEXTO1
        ZEXTO=ANSITOOEM(ZEXTO)
        If !EMPTY(ZANCHETE) .And. !EMPTY(ZEXTO) .And. !EMPTY(ZRNOT) .And. ZTNOT<>CTOD("")
          ADIREG(0)
          ZDNOT=RECNO()
          ZINK=""
          PUTREC()
        EndIf
      ENDDO
      oText:close()
      SET(_SET_EOL,cSvSetEOL) //Retorna o caracter de finalização para o default CHR(10)+CHR(13)
    return 

    Procedure BUSCAVAL()
      TEXTO="Usuário"
      VARIAVEL=space(100)
      MASCARA="@!"
      REG=RECNO()
      DO WHILE .T.
        SysRefresh()
        If !PROCURA()
          EXIT
        EndIf
        CHAVE=alltrim(variavel)
        If EMPTY(CHAVE)
          EXIT
        EndIf
        LOCATE FOR USUARIO=CHAVE
        If EOF() .OR. DELE()
          MSGALERT("Não Encontrado !","Alerta")
          GO REG
        Else
          EXIT
        EndIf
      ENDDO
    return 

    Procedure OUVIDORIA()
      DIF=TIMEDIFF(INIOUVI,TIME())
      DIF1=TIMEASSECONDS(DIF)
      If DIF1<600
        return 
      EndIf
      CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+"wget.exe -T120 -olog.txt -ONADA.TXT https://exatus.net/ouvidoria/verificamail.asp --no-check-certificate"+CHR(13)+CHR(10)+"DEL PARAR.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
      memowrit("R1.BAT",CMACRO)
      MEMOWRIT("PARAR.TXT",CMACRO)
      waitrun("R1.BAT",0)
      DO WHILE FILE("PARAR.TXT")
      ENDDO
      INIOUVI=TIME()
    return 

    Procedure CRIARQGUARDA(WEBLOCAL3)
      SELE 36
      If !FILE(WEBLOCAL3+"GRUPODOC.DBF")
        MAT_STRU={}
        AADD(MAT_STRU,{"CODIGO","N",5,0})
        AADD(MAT_STRU,{"DESCRI","C",20,0})
        AADD(MAT_STRU,{"TP","C",1,0})
        AADD(MAT_STRU,{"STATUS","C",1,0})
        DBCREATE(WEBLOCAL3+"GRUPODOC.DBF",MAT_STRU)
        If !USEREDE(WEBLOCAL3+"GRUPODOC.DBF",.F.,10,"TMP")
          USE
          warqs=WEBLOCAL3+"GRUPODOC.DBF"
          Erase &WARQS
        Else
          ADIREG(0)
          REPLACE CODIGO WITH 1
          REPLACE DESCRI WITH "CADASTRO PF"
          REPLACE TP WITH "F"
          ADIREG(0)
          REPLACE CODIGO WITH 2
          REPLACE DESCRI WITH "CADASTRO PJ"
          REPLACE TP WITH "J"
          ADIREG(0)
          REPLACE CODIGO WITH 3
          REPLACE DESCRI WITH "CAMBIO"
          ADIREG(0)
          REPLACE CODIGO WITH 4
          REPLACE DESCRI WITH "BOVESPA"
          ADIREG(0)
          REPLACE CODIGO WITH 5
          REPLACE DESCRI WITH "BM&F"
          ADIREG(0)
          REPLACE CODIGO WITH 6
          REPLACE DESCRI WITH "CONTABEIS"
          ADIREG(0)
          REPLACE CODIGO WITH 7
          REPLACE DESCRI WITH "GERAIS DO DIA"
        EndIf
        USE
        RELEASE mat_stru
      EndIf
      If !FILE(WEBLOCAL3+"TIPODOC.DBF")
        MAT_STRU={}
        AADD(MAT_STRU,{"CODIGO","N",5,0})
        AADD(MAT_STRU,{"GRUPO","N",5,0})
        AADD(MAT_STRU,{"DESCRI","C",40,0})
        AADD(MAT_STRU,{"VENCTO","L",1,0})
        AADD(MAT_STRU,{"OBRIGA","L",1,0})
        AADD(MAT_STRU,{"REFERE","L",1,0})
        AADD(MAT_STRU,{"GUARDAR","L",1,0})
        AADD(MAT_STRU,{"MASCARA","C",20,0})
        AADD(MAT_STRU,{"STATUS","C",1,0})
        DBCREATE(WEBLOCAL3+"TIPODOC.DBF",MAT_STRU)
        If !USEREDE(WEBLOCAL3+"TIPODOC.DBF",.F.,10,"TMP")
          USE
          WARQS=WEBLOCAL3+"TIPODOC.DBF"
          Erase &WARQS
        Else
          MAT_STRU={}
          AADD(MAT_STRU,{1,1,"RG              ",.F.,.T.,.T.,.T.,"999.999.999-A "})
          AADD(MAT_STRU,{2,1,"CPF             ",.F.,.T.,.T.,.T.,"999.999.999-99"})
          AADD(MAT_STRU,{3,1,"COMPROV.ENDEREÇO",.T.,.T.,.T.,.T.,"              "})
          AADD(MAT_STRU,{4,1,"DECLARAÇÃO I.R. ",.T.,.T.,.T.,.T.,"              "})
          AADD(MAT_STRU,{5,1,"FICHA CADASTRAL ",.T.,.T.,.T.,.T.,"              "})
          AADD(MAT_STRU,{6,1,"CARTÃO ASSINATURA",.T.,.T.,.T.,.T.,"             "})
          AADD(MAT_STRU,{7,2,"FICHA CADASTRAL  ",.T.,.T.,.T.,.T.,"             "})
          AADD(MAT_STRU,{8,2,"DECLARAÇÃO I.R.  ",.T.,.T.,.T.,.T.,"             "})
          AADD(MAT_STRU,{9,2,"DOCTOS SOCIOS    ",.T.,.T.,.T.,.T.,"             "})
          AADD(MAT_STRU,{10,2,"PROCURAÇÃO SOCIOS",.T.,.T.,.T.,.T.,"            "})
          AADD(MAT_STRU,{11,2,"COMPROV.CNPJ     ",.T.,.T.,.T.,.T.,"99.999.999/9999-99"})
          AADD(MAT_STRU,{12,2,"BALANÇO          ",.T.,.T.,.T.,.T.,"            "})
          AADD(MAT_STRU,{13,2,"ESTATUTO SOCIAL  ",.T.,.T.,.T.,.T.,"            "})
          AADD(MAT_STRU,{14,2,"CARTÃO ASSINATURAS",.T.,.T.,.T.,.T.,"            "})
          AADD(MAT_STRU,{15,2,"PROCURAÇÃO CORRETORA",.T.,.F.,.T.,.T.,"            "})
          AADD(MAT_STRU,{16,2,"COMPROV.ENDEREÇO",.T.,.T.,.T.,.T.,"            "})
          AADD(MAT_STRU,{17,3,"CONTRATO PREST.SERV",.T.,.T.,.T.,.T.,""})
          AADD(MAT_STRU,{18,4,"NOTA DE CORRETAGEM",.F.,.F.,.T.,.T.,"9999999"})
          AADD(MAT_STRU,{19,5,"NOTA DE CORRETAGEM",.F.,.F.,.T.,.T.,"9999999"})
          AADD(MAT_STRU,{20,6,"EXTRATO CONTA CORRENTE",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{21,6,"AVISO DE CONTA CORRENTE",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{22,6,"AUTORIZACAO DE SUBSCRICAO",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{23,6,"INFORME DE RENDIMENTO",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{24,6,"AUTORIZACAO DE SUBSCRICAO",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{25,6,"POSICOES",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{26,6,"EXIGENCIA DE MARGEM",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{27,6,"DEMONSTRACOES DE ATIVOS",.F.,.F.,.T.,.T.,""})
          AADD(MAT_STRU,{28,7,"GERAL DO DIA",.f.,.f.,.t.,.t.,""})
          FOR I=1 TO 28
            ADIREG(0)
            REPLACE CODIGO WITH MAT_STRU[I,1]
            REPLACE GRUPO WITH MAT_STRU[I,2]
            REPLACE DESCRI WITH MAT_STRU[I,3]
            REPLACE VENCTO WITH MAT_STRU[I,4]
            REPLACE OBRIGA WITH MAT_STRU[I,5]
            REPLACE REFERE WITH MAT_STRU[I,6]
            REPLACE GUARDAR WITH MAT_STRU[I,7]
            REPLACE MASCARA WITH MAT_STRU[I,8]
          NEXT I
        EndIf
        USE
        RELEASE mat_stru
      EndIf
      If !FILE(WEBLOCAL3+"DOCTOS.DBF")
        MAT_STRU={}
        AADD(MAT_STRU,{"CORRETORA","N",4,0})
        AADD(MAT_STRU,{"CLIENTE","N",9,0})
        AADD(MAT_STRU,{"ORDEMGER","N",19,0}) // ORDEM GERAL (SERIA O REGISTRO)
        AADD(MAT_STRU,{"GRUPOD","N",5,0})   // GRUPO DO DOCUMENTO
        AADD(MAT_STRU,{"TIPODOC","N",5,0}) // CODIGO DO DOCUMENTO
        AADD(MAT_STRU,{"VENCTO","D",8,0})
        AADD(MAT_STRU,{"REFERE","D",8,0})
        AADD(MAT_STRU,{"OBS","C",50,0})
        AADD(MAT_STRU,{"DOCNR","C",20,0})
        AADD(MAT_STRU,{"ARQUIVO","N",12,0})
        AADD(MAT_STRU,{"ATIVO","L",1,0}) // SE O DOCUMENTO ESTÁ ATIVO OU DESATIVADO
        AADD(MAT_STRU,{"ESCANEM","D",8,0}) // ESCANEADO EM
        AADD(MAT_STRU,{"STATUS","C",1,0}) // STATUS DO ARQUIVO COM A INTERNET
        AADD(MAT_STRU,{"POR","C",20,0}) // POR QUEM FOI ESCANEDADO
        DBCREATE(WEBLOCAL3+"DOCTOS.DBF",MAT_STRU)
        RELEASE mat_stru
      EndIf
      USE
    return 

    Procedure CRIARQBOLSANET(PASTA)
      ARQ=ALLTRIM(PASTA)+"ASSESSOR.DBF"
      If !FILE(ARQ)
        aStructure:={}
        aAdd( astructure, { "CODIGO"    , "N" , 3 , 0 })
        aAdd( astructure, { "NOME","C", 50,0})
        aAdd( astructure, { "EMAIL","C",250,0})
        aAdd( aStructure, { "ENVIA","L",1,0})
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"MGRUPO.DBF"
      If !FILE(ARQ)
        aStructure := {}
        aAdd( aStructure, { "CODIGO"    , "N",   6, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
        aAdd( aStructure, { "GRUPO"     , "N",   6, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"GRUPO.DBF"
      If !FILE(ARQ)
        aStructure := {}
        aAdd( aStructure, { "CODIGO"    , "N",   6, 0 } )
        aAdd( aStructure, { "NOME"      , "C", 254, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"LOGBOLSA.DBF"
      If !FILE(ARQ)
        aStructure :={}
        aAdd( aStructure, {"ROTINA","C",20,0})
        aAdd( aStructure, {"CAMPO","C",20,0})
        aAdd( astructure, {"CHAVE","C",10,0})
        aAdd( aStructure, {"DESCRICAO","C",50,0})
        aAdd( aStructure, {"ANTES","C",254,0})
        aAdd( astructure, {"ATUAL","C",254,0})
        aAdd( aStructure, {"DATA","C",20,0})
        aAdd( aStructure, {"USUARIO","C",50,0})
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"MCLIENTE.DBF"
      If !FILE(ARQ)
        aStructure := {}
        aAdd( aStructure, { "CODIGO"    , "N",   9, 0 } )
        aAdd( aStructure, { "BMFBOV"    , "C",   1, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "CLIENTE"   , "N",   9, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "FLAG"      , "C",   1, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"LOGS.DBF"
      If !FILE(ARQ)
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    ,"N" ,   9, 0 } )
        aAdd( aStructure, { "DATAENV"   ,"D" ,   8, 0 } )
        aAdd( aStructure, { "HORAENV"   ,"C" ,  20, 0 } )
        aAdd( aStructure, { "OQUE"      ,"C" , 254, 0 } )
        aAdd( aStructure, { "WCEP"      ,"C" ,   8, 0 } )
        aAdd( aStructure, { "NOME"      ,"C" ,  60, 0 } )
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "C",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "FOLHAS"    , "N",   5, 0 } )
        aAdd( aStructure, { "ENDE1","C",100,0})
        aAdd( aStructure, { "ENDE2","C",100,0})
        aAdd( aStructure, { "FAZER"     , "L",   1, 0 } )
        aAdd( aStructure, { "EMAIL"     , "L",   1, 0 } )
        aAdd( aStructure, { "ASSESSOR"  , "N",   5, 0 } )
        aAdd( aStructure, { "Enviado"   , "L",   1, 0 } )
        aAdd( aStructure, { "frentev"   , "L",   1, 0 } )
        aAdd( aStructure, { "EmailC"     , "C",1000, 0 } )
        aAdd( aStructure, { "BAIXA"     , "L",   1, 0 } )
        aAdd( aStructure, { "IMPRESSO"  , "L",   1, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"CLIENTE.DBF"
      If !FILE(ARQ)
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    , "N",   9, 0 } )
        aAdd( aStructure, { "CNPJCPF"   , "C",  14, 0 } )
        aAdd( aStructure, { "INVEST"    , "N",   7, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "APELIDO"   , "C"  ,10, 0 } )
        aAdd( aStructure, { "NASC"      , "D",   8, 0 } )
        aAdd( aStructure, { "TIPO"      , "C",   1, 0 } ) // (J)uridica, (F)isica
        aAdd( aStructure, { "SEXO"      , "C",   1, 0 } ) // Sexo quando Fisica
        aAdd( aStructure, { "ECIVIL"    , "C",   1, 0 } ) // Estado Civil
        aAdd( aStructure, { "CONJUGE"   , "C",  60, 0 } ) // Nome da Esposa
        aAdd( aStructure, { "BANCO"     , "N",   3, 0 } ) // Codigo do Banco Conta Corrente
        aAdd( aStructure, { "ASSESSOR"  , "N",   5, 0 } ) // codigo do Assessor
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "C",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "TIPOINV"   , "N",   5, 0 } ) // Tipo de Investidor
        aAdd( aStructure, { "SENHA"     , "C"  ,20, 0 } )
        aAdd( aStructure, { "FRASE"     , "C"  ,50, 0 } )
        aAdd( aStructure, { "EEMAIL"    , "L",   1, 0 } ) // Envia Email sempre sim,
        aAdd( aStructure, { "EIMPRE"    , "L",   1, 0 } ) // Imprime os relatórios
        aAdd( aStructure, { "EGUARD"    , "L",   1, 0 } ) // Guarda Doctos na Internet
        aAdd( aStructure, { "FRENTEV"   , "L",   1, 0 } ) // Frente/Verso
        aAdd( aStructure, { "MASTER"    , "N",   9, 0 } ) // Codigo Master do Frente/Verso
        aAdd( aStructure, { "ULTNOTABOV", "D",   8, 0 } ) // Data da Ultima nota de Bovespa
        aAdd( aStructure, { "ULTNOTBMF" , "D",   8, 0 } ) // Data da Ultima nota de BMF
        aAdd( aStructure, { "ULTCONTA"  , "D",   8, 0 } ) // Data do Ultimo Extrto de Contas Correntes
        aAdd( aStructure, { "Mailing"   , "L",   1, 0 } ) // Envia Mailing ?
        aAdd( aStructure, { "MAILNEWS"  , "C", 240, 0 } ) // Email do News
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "FLAG"      , "C",   1, 0 } )
        aAdd( aStructure, { "TEMMAIL"   , "L",   1, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        aAdd( aStructure, { "ENDE1","C",100,0})
        aAdd( aStructure, { "ENDE2","C",100,0})
        aAdd( aStructure, { "AUDIT","L",1,0})
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=ALLTRIM(PASTA)+"CONTRS.DBF"
      If !FILE(ARQ)
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    , "N",   5, 0 } ) // O nr. 1 é o Email para retorno ao cliente pessoa Fisica
        aAdd( aStructure, { "DESCRICAO" , "C", 100, 0 } ) // O nr. 2 é o Email para cobrança do cliente pessoa fisica após 10 dias
        aAdd( astructure, { "FIXO"      , "L",   1, 0 } ) // o Nr. 3 é o Email para OK ao cliente após o Dep.Cadastro baixar o Arquivo
        aAdd( aStructure, { "FJ"        , "C",   1, 0 } ) // O Nr. 4 é o Email para retorno ao cliente pessoa Juridica
        aAdd( aStructure, { "MARK"      , "C",   1, 0 } ) // será S ou N se será utilizado no Marketing
        aAdd( aStructure, { "COR"       , "C",  10, 0 } ) // cor para marketing
        aAdd( astructure, { "MARCADO"   , "C",   1, 0 } ) // se deverá ser marcado
        aAdd( astructure, { "CONTEUDO"  , "M",  10, 0 } ) // O Nr. 5 é o email para cobrança do cliente pessoa Juridica após 10 dias
        aAdd( astructure, { "INCLUIDO"  , "C",  20, 0 } ) // o Nr. 6 é o email para OK ao cliente após o Dept.Cadastro Baixar o Arquivo Pessoa Juridica
        aAdd( astructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( astructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( astructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
      ARQ=WEBLOCAL4+"ETIQUETA.DBF"
      If !FILE(ARQ)
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    ,"N" ,   9, 0 } )
        aAdd( aStructure, { "WCEP"      ,"C" ,   8, 0 } )
        aAdd( aStructure, { "NOME"      ,"C" ,  60, 0 } )
        aAdd( aStructure, { "DATAENV"   ,"D" ,   8, 0 } )
        aAdd( aStructure, { "HORAENV"   ,"C" ,  20, 0 } )
        aAdd( aStructure, { "OQUE"      ,"C" , 254, 0 } )
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "C",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "FOLHAS"    , "N",   5, 0 } )
        aAdd( aStructure, { "AUDIT"     , "L",   1, 0 } )
        aAdd( aStructure, { "ENDE1","C",100,0})
        aAdd( aStructure, { "ENDE2","C",100,0})
        aAdd( aStructure, { "FAZER"     , "L",   1, 0 } )
        aAdd( aStructure, { "EMAIL"     , "L",   1, 0 } )
        aAdd( aStructure, { "ASSESSOR"  , "N",   5, 0 } )
        aAdd( aStructure, { "Enviado"   , "L",   1, 0 } )
        aAdd( aStructure, { "frentev"   , "L",   1, 0 } )
        aAdd( aStructure, { "EmailC"     , "C",1000, 0 } )
        aAdd( aStructure, { "BAIXA"     , "L",   1, 0 } )
        aAdd( aStructure, { "IMPRESSO"  , "L",   1, 0 } )
        aAdd( aStructure, { "LOGMAIL"   , "C", 250, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
    return 

    Procedure CRIAMOVBOLSANET(PASTA)
      SELE 34
      USE
      SELE 34
      ARQ4=ALLTRIM(PASTA)+"CLIENTE.KEY"
      Erase &ARQ4
      ARQ4=ALLTRIM(PASTA)+"CLIENTE1.KEY"
      Erase &ARQ4
      If !USEREDE(ALLTRIM(PASTA)+"CLIENTE.DBF",.F.,10,"CLIENTE")
        return 
      EndIf
      If FCOUNT()<42
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    , "N",   9, 0 } )
        aAdd( aStructure, { "CNPJCPF"   , "C",  14, 0 } )
        aAdd( aStructure, { "INVEST"    , "N",   7, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "APELIDO"   , "C"  ,10, 0 } )
        aAdd( aStructure, { "NASC"      , "D",   8, 0 } )
        aAdd( aStructure, { "TIPO"      , "C",   1, 0 } ) // (J)uridica, (F)isica
        aAdd( aStructure, { "SEXO"      , "C",   1, 0 } ) // Sexo quando Fisica
        aAdd( aStructure, { "ECIVIL"    , "C",   1, 0 } ) // Estado Civil
        aAdd( aStructure, { "CONJUGE"   , "C",  60, 0 } ) // Nome da Esposa
        aAdd( aStructure, { "BANCO"     , "N",   3, 0 } ) // Codigo do Banco Conta Corrente
        aAdd( aStructure, { "ASSESSOR"  , "N",   5, 0 } ) // codigo do Assessor
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "C",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "TIPOINV"   , "N",   5, 0 } ) // Tipo de Investidor
        aAdd( aStructure, { "SENHA"     , "C"  ,20, 0 } )
        aAdd( aStructure, { "FRASE"     , "C"  ,50, 0 } )
        aAdd( aStructure, { "EEMAIL"    , "L",   1, 0 } ) // Envia Email
        aAdd( aStructure, { "FRENTEV"   , "L",   1, 0 } ) // Frente/Verso
        aAdd( aStructure, { "MASTER"    , "N",   9, 0 } ) // Codigo Master do Frente/Verso
        aAdd( aStructure, { "EIMPRE"    , "L",   1, 0 } ) // Imprime os relatórios
        aAdd( aStructure, { "EGUARD"    , "L",   1, 0 } ) // Guarda Doctos noa Internet
        aAdd( aStructure, { "Mailing"   , "L",   1, 0 } ) // Envia Mailing ?
        aAdd( aStructure, { "ULTNOTABOV", "D",   8, 0 } ) // Data da Ultima nota de Bovespa
        aAdd( aStructure, { "ULTCONTA"  , "D",   8, 0 } ) // Data do Ultimo Extrto de Contas Correntes
        aAdd( aStructure, { "MAILNEWS"  , "C", 240, 0 } ) // Email do News
        aAdd( aStructure, { "ULTNOTBMF" , "D",   8, 0 } ) // Data da Ultima nota de BMF
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "FLAG"      , "C",   1, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        aAdd( aStructure, { "TEMMAIL"   , "L",   1, 0 } )
        aAdd( aStructure, { "ENDE1","C",100,0})
        aAdd( aStructure, { "ENDE2","C",100,0})
        aAdd( aStructure, { "AUDIT","L",1,0})
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(PASTA)+"CLIENTE.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        GO TOP
        DO WHILE !EOF()
          If EMPTY(SENHA)
            REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
          EndIf
          SKIP
        ENDDO
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
      If !USEREDE(ALLTRIM(PASTA)+"LOGBOLSA.DBF",.F.,10,"LOGBOLSA")
        return 
      EndIf
      If FCOUNT()<8
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure:={}
        aAdd( aStructure, {"ROTINA","C",20,0})
        aAdd( aStructure, {"CAMPO","C",20,0})
        aAdd( astructure, {"CHAVE","C",10,0})
        aAdd( aStructure, {"DESCRICAO","C",50,0})
        aAdd( aStructure, {"ANTES","C",254,0})
        aAdd( astructure, {"ATUAL","C",254,0})
        aAdd( aStructure, {"DATA","C",20,0})
        aAdd( aStructure, {"USUARIO","C",50,0})
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(PASTA)+"LOGBOLSA.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
      SELE 34
      ARQ4=ALLTRIM(PASTA)+"MCLIENTE.KEY"
      Erase &ARQ4
      If !USEREDE(ALLTRIM(PASTA)+"MCLIENTE.DBF",.F.,10,"MCLIENTE")
        return 
      EndIf
      If FCOUNT()<11
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    , "N",   9, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
        aAdd( aStructure, { "BMFBOV"    , "C",   1, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "CLIENTE"   , "N",   9, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "FLAG"      , "C",   1, 0 } ) // Tipo "B"=bovespa "F"=BMF
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(PASTA)+"MCLIENTE.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
      SELE 34
      ARQ4=ALLTRIM(PASTA)+"GRUPO.KEY"
      Erase &ARQ4
      If !USEREDE(ALLTRIM(PASTA)+"GRUPO.DBF",.F.,10,"GRUPO")
        return 
      EndIf
      If FCOUNT()<7
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure:={}
        aAdd( aStructure, { "CODIGO"    , "N",   6, 0 } )
        aAdd( aStructure, { "NOME"      , "C", 254, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(PASTA)+"GRUPO.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
      ARQ4=ALLTRIM(PASTA)+"MGRUPO.KEY"
      Erase &ARQ4
      If !USEREDE(ALLTRIM(PASTA)+"MGRUPO.DBF",.F.,10,"MGRUPO")
        return 
      EndIf
      If FCOUNT()<8
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure := {}
        aAdd( aStructure, { "CODIGO"    , "N",   6, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C", 254, 0 } )
        aAdd( aStructure, { "GRUPO"     , "N",   6, 0 } )
        aAdd( aStructure, { "INCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "ALTERADO"  , "C",  20, 0 } )
        aAdd( aStructure, { "EXCLUIDO"  , "C",  20, 0 } )
        aAdd( aStructure, { "USERINC"   , "C",  50, 0 } )
        aAdd( aStructure, { "USERALT"   , "C",  50, 0 } )
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(PASTA)+"MGRUPO.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
    return 

    Procedure RECPDF(PARAM)
      SEENVIAEMAIL=PARAM
      SELE 30
      GO TOP
      CURSORWAIT()
      DO WHILE .NOT. EOF()
        SysRefresh()
        ARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        If !ENVMAIL .And. !CORREIO .And. !DADOS
          SKIP
          LOOP
        EndIf
        If TIPO="U" .OR. WINIMAGE
          SKIP
          LOOP
        EndIf
        If !FILE(ARQ)
          SKIP
          LOOP
        EndIf
        SELE 30
        ARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        ARQF3=ALLTRIM(PASTA)+"ASSUNTO.TXT"
        If FILE(ARQF3)
          WNOVOASSUNTO=ALLTRIM(MEMOREAD(ARQF3))
        Else
          WNOVOASSUNTO=""
        EndIf
        Erase &ARQF3
        ARQF3=ARQ
        WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\BOLSASNET\"
        VERDIRETORIO(WEBLOCAL1)
        WEBLOCAL2=ALLTRIM(PASTA)+"ARQF\"
        VERDIRETORIO(WEBLOCAL2)
        WEBLOCAL7=alltrim(pasta)+"CONTRS\"
        VERDIRETORIO(WEBLOCAL7)
        DECLARE VLABEL1[ADIR(WEBLOCAL7+"*.*")]
        ADIR(WEBLOCAL7+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL7+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        SELE 30
        WEBLOCAL4=ALLTRIM(PASTA)+"IMPRIME\"
        WEBLOCAL6=ALLTRIM(PASTA)+"FRENTEV\"
        WEBLOCAL9=ALLTRIM(PASTA)+"SEPARADO\"
        WEBLOCAL10=ALLTRIM(PASTA)+"NAOENCO\"
        WWEBLOCA9=ALLTRIM(PASTA)+"ECONOMIAECUSTOS\"
        VERDIRETORIO(WEBLOCAL4)
        VERDIRETORIO(WEBLOCAL6)
        VERDIRETORIO(WEBLOCAL9)
        VERDIRETORIO(WEBLOCAL10)
        VERDIRETORIO(WWEBLOCA9)
        DECLARE VLABEL1[ADIR(WEBLOCAL10+"*.*")]
        ADIR(WEBLOCAL10+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL10+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        DECLARE VLABEL1[ADIR(WEBLOCAL4+"*.*")]
        ADIR(WEBLOCAL4+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL4+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        DECLARE VLABEL1[ADIR(WEBLOCAL6+"*.*")]
        ADIR(WEBLOCAL6+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL6+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        DECLARE VLABEL1[ADIR(WEBLOCAL9+"*.*")]
        ADIR(WEBLOCAL9+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL9+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        SELE 30
        WMENS1=OEMTOANSI(MENS1)
        WMENS2=OEMTOANSI(MENS2)
        WMENS3=OEMTOANSI(MENS3)
        WMENS4=OEMTOANSI(MENS4)
        WMENS5=OEMTOANSI(MENS5)
        WMENS6=OEMTOANSI(MENS6)
        CRIARQBOLSANET(alltrim(BOLSANET->PASTA))
        CRIAMOVBOLSANET(alltrim(BOLSANET->PASTA))
        SELE 30
        WFAZDADOS=DADOS
        WFAZCORREIO=CORREIO
        WEBLOCAL=ALLTRIM(PASTA)
        WEBCUSTO=CUSTOMIL
        WEBCOR=CORRETORA
        WEBASSUNTO=OEMTOANSI(NASSUNTO)
        If !EMPTY(WNOVOASSUNTO)
          WEBASSUNTO=OEMTOANSI(WNOVOASSUNTO)
        EndIf
        WEBNOMECOR=OEMTOANSI(NOMECOR)
        WEBBOLSA=BOLSA
        WEBACEITA=ACEITA
        WEBMAIL=EMAIL
        WEBCORPO=OEMTOANSI(CORPOEMAIL)
        WEBCORPCON=OEMTOANSI(CORPOASS)
        WEBSENHA=SENHAEMAIL
        WEBJUD=MAILJUD
        WEBCAD=MAILCAD
        WEBMAR=MAILMAR
        WEBMKT=MAILMKT
        WEBIMP=MAILIMP
        WEBREL=MAILREL
        WEBAUD=MAILAUDI
        MCONTA=M->FROM
        WEBSERV=ALLTRIM(SMTP)
        WEBFORCA=FORCAEMSTP
        WEBLOGIN=LOGEMAIL
        WEBPDFS=SENHAPDF
        WEBBOV=CODBOV
        If LOCALQTIPO="W" .And. COD_OPER<997
          WEBMAIL=ZEMAIL
          WEBCORPO=OEMTOANSI(ZCORPOEMAIL)
          WEBSENHA=ZSENHAEMAIL
          WEBJUD=ZMAILJUD
          WEBCAD=ZMAILCAD
          WEBMAR=ZMAILMAR
          WEBMKT=ZMAILMKT
          WEBIMP=ZMAILIMP
          WEBAUD=ZMAILAUDI
          WEBREL=ZMAILREL
          MCONTA=ZEMAIL
          WEBAUD=ZMAILAUDI
          WEBSERV=ALLTRIM(ZSMTP)
          WEBFORCA=ZFORCAEMSTP
          WEBLOGIN=ZLOGEMAIL
          WEBPDFS=ZSENHAPDF
          WEBBOV=ZCODBOV
        EndIf
        If WFAZDADOS
          WEBLOCAL3=ALLTRIM(PASTA)+"GUARDA\"
          VERDIRETORIO(WEBLOCAL3)
          CRIARQGUARDA(ALLTRIM(WEBLOCAL3))
        EndIf
        SELE 30
        DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
        ADIR(WEBLOCAL1+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.exe")]
        ADIR(WEBLOCAL+"*.EXE",vlabel)
        I=0
        FOR I=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+VLABEL[I]
          COMANDO=WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
          WAITRUN(COMANDO,0)
          Erase &NARQ
        NEXT I
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.zip")]
        ADIR(WEBLOCAL+"*.zip",vlabel)
        I=0
        FOR I=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+VLABEL[I]
          COMANDO=WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
          WAITRUN(COMANDO,0)
          Erase &NARQ
        NEXT I
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
        ADIR(WEBLOCAL+"*.PDF",vlabel)
        I=0
        WDADOS=WEBLOCAL1+"ENVIADO.TXT"
        DECLARE TX1[1]
        TX1[1]="BolsasNet "+WEBCOR
        LOGFILE(WEBLOCAL+"ROTINAS.TXT",TX1)
        TX1[1]=WNOVOASSUNTO
        LOGFILE(WDADOS,TX1)
        If AUTOMATICO
          DECLARE TX1[1]
          TX1[1]="Executando BolsasNet "+WEBCOR+" "+WEBLOCAL
          LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        EndIf
        FEZALGO=.F.
        RELBMFBOV()
        EMAILRETORNO=WEBREL
        If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
          ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Log dos Relatórios enviados","Em anexo relatorio",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
        If SEENVIAEMAIL="S" .And. FEZALGO .And. AT("EXATUS",WEBCOR)=0 .And. at("PAUL",WEBCOR)=0
          ENVIAEMAIL("bene@exatus.net","BolsasNet "+WEBCOR,"Em anexo relatorio",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
        Erase &ARQF3
        SELE 30
        SKIP
      ENDDO
      cursorarrow()
    return 

    Procedure RELBMFBOV()
      SOPROC=.F.
      CADPLANNER()
      CADCSV()
      CADCSVCLI()

      //ATUACADASTRO()
      EXTRUNIF()
      EXTRANUAL()
      ATUAEMAILGER()
      EXTRGERFUT()
      //APLIRESGERACAO()
      //INFORMEGRENDNOVO()
      //INFORMENG()
      //ATUACCIN()
      INFSIMPAUL=.F.
      DECLARE VLABEL[ADIR(WEBLOCAL+"SIMPAULINF*.PDF")]
      ADIR(WEBLOCAL+"SIMPAULINF*.PDF",VLABEL)
      If len(vlabel)>0
        INFSIMPAUL=.T.
      EndIf
      SELE 30
      //*  Atualiza cadastro para emails sem o codigo do cliente
      NARQ=WEBLOCAL+"CADASTRO.PDF"
      DECLARE VLABEL[ADIR(WEBLOCAL+"CADASTRO*.PDF")]
      ADIR(WEBLOCAL+"CADASTRO*.PDF",VLABEL)
      OZ=0
      If len(vlabel)>0
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+ALLTRIM(VLABEL[OI])
          FAZERCADASTRO(NARQ)
          Erase &NARQ
        NEXT OI
      EndIf
      SELE 30
      //*  Atualiza cadastro de emails pelo relatório do Sinacor
      NARQ=WEBLOCAL+"EMAIL.PDF"
      DECLARE VLABEL[ADIR(WEBLOCAL+"EMAIL*.PDF")]
      ADIR(WEBLOCAL+"EMAIL*.PDF",VLABEL)
      OZ=0
      FAZEREM=.F.
      If len(vlabel)>0
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+ALLTRIM(VLABEL[OI])
          FAZEREMAIL(NARQ)
          Erase &NARQ
        NEXT OI
      EndIf
      SELE 30
      //*  Atualiza cadastro de emails pelo relatório do Sinacor
      NARQ=WEBLOCAL+"EMAIL.PDF"
      DECLARE VLABEL[ADIR(WEBLOCAL+"E-MAIL*.PDF")]
      ADIR(WEBLOCAL+"E-MAIL*.PDF",VLABEL)
      OZ=0
      FAZEREM=.F.
      If len(vlabel)>0
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+ALLTRIM(VLABEL[OI])
          FAZEREMAIL(NARQ)
          Erase &NARQ
        NEXT OI
      EndIf
      SELE 30
      MARKETING()
      //** SEPARA AS PAGINAS
      DECLARE VLABEL[ADIR(WEBLOCAL+"*.PDF")]
      ADIR(WEBLOCAL+"*.PDF",VLABEL)
      OZ=0
      If len(vlabel)>0
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Gerando/Enviando Relatórios "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        FEZALGO=.T.
      EndIf
      WERROARQ=""
      TEMCETIP=.F.
      FOR OI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        If ALLTRIM(UPPER(VLABEL[OI]))="BOLETIM.PDF"
          MBOLET=ALLTRIM(WEBLOCAL)+"BOLETIM.PDF"
          MBOLET1=ALLTRIM(WEBLOCAL2)+"BOLETIM"+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+".PDF"
          LZCOPYFILE(MBOLET,MBOLET1)
          LOOP
        EndIf
        NARQ=UPPER(VLABEL[OI])
        If AT("_CERTIFICACAO",NARQ)#0
          MBOLET=ALLTRIM(WEBLOCAL)+ALLTRIM(VLABEL[OI])
          MBOLET1=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
          NARQ=ALLTRIM(WEBLOCAL1)+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(MBOLET,MBOLET1)
          LZCOPYFILE(MBOLET,NARQ)
          Erase &MBOLET
          TEMCETIP=.T.
          LOOP
        EndIf
        NARQ=WEBLOCAL+ALLTRIM(VLABEL[OI])
        M->LOCAL21=TRIM("/"+CURDIR())
        M->LOCAL3=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
        LZCOPYFILE(NARQ,M->LOCAL3)
        LCHDIR(WEBLOCAL1)
        CURSORWAIT()
        comando=WCORRENTE+"\PDFTK "+Lfn2Sfn(NARQ)+" dump_data output "+WEBLOCAL+"T.TXT"
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+weblocal+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal+"FZ.BAT",mzlinha)
        MEMOWRIT(weblocal+"AGUARDE.TXT","Rotina 01"+MZLINHA)
        WAITRUN(weblocal+"FZ.BAT",0)
        DO WHILE FILE(weblocal+"AGUARDE.TXT")
          SysRefresh()
        ENDDO
        ULTPG=VAL(SUBSTR(MEMOREAD(WEBLOCAL+"T.TXT"),at("NumberOfPages:",MEMOREAD(WEBLOCAL+"T.TXT"))+15,50))
        DARQ=WEBLOCAL+"FZ.BAT"
        Erase &DARQ
        DARQ=WEBLOCAL+"T.TXT"
        Erase &DARQ
        LCHDIR(WEBLOCAL1)
        comando=WCORRENTE+"\PDFTK "+Lfn2Sfn(NARQ)+" burst "
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal1+"FZ.BAT",mzlinha)
        MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 02"+MZLINHA)
        DARQ1=WEBLOCAL1+"AGUARDE.TXT"
        DO WHILE !FILE(DARQ1)
        ENDDO
        WAITRUN(weblocal1+"FZ.BAT",0)
        DO WHILE FILE(weblocal1+"AGUARDE.TXT")
        ENDDO
        CURSORWAIT()
        LCHDIR(M->LOCAL21)
        MTEZ=WEBLOCAL1+"FZ.BAT"
        Erase &MTEZ
        LCHDIR(M->LOCAL21)
        CURSORWAIT()
        Erase &NARQ
        COMANDO="DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 03"+COMANDO)
        DARQ1=WEBLOCAL1+"AGUARDE.TXT"
        DO WHILE !FILE(DARQ1)
        ENDDO
        DARQ=WEBLOCAL+"FZ.BAT"
        ARQ100=FCREATE(DARQ)
        BSAVES=FWRITE(ARQ100,"@ECHO OFF"+CHR(13)+CHR(10),11)
        FOR IO=1 TO ULTPG
          SysRefresh()
          ARQDE=WEBLOCAL1+"PG_"+STRZERO(IO,4)+".PDF"
          If IO>9999
            ARQDE=WEBLOCAL1+"PG_"+STRZERO(IO,5)+".PDF"
          EndIf
          OZ+=1
          ARQATE=WEBLOCAL1+"O"+STRZERO(OZ,6)+".PDF"
          comando="REN "+ARQDE+" O"+STRZERO(OZ,6)+".PDF"+chr(13)+chr(10)
          BSAVE=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
          NARQ2=WEBLOCAL1+"O"+STRZERO(OZ,6)+".TXT"
          COMANDO=WCORRENTE+"\PDFTEXT.EXE -eol dos -layout "+ARQATE+CHR(13)+CHR(10)
          BSAVE=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        NEXT IO
        COMANDO="DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)
        BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        COMANDO="DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)
        BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        COMANDO="DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)
        BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        COMANDO="DEL "+WEBLOCAL+"FZ.BAT"+CHR(13)+CHR(10)
        BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        COMANDO="DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
        FCLOSE(ARQ100)
        CURSORWAIT()
        WAITRUN(DARQ,0)
        CURSORWAIT()
        DO WHILE FILE(DARQ1) .And. FILE(WEBLOCAL+"FZ.BAT")
          SysRefresh()
        ENDDO
        CURSORWAIT()
        Erase &DARQ
      NEXT OI
      ULTIMO=OZ
      CURSORWAIT()
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
        NARQ3=weblocal1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".DAT"
        If !FILE(NARQ3)
          EXIT
        EndIf
      ENDDO
      MAT_STRU={}
      AADD(MAT_STRU,{"LINHA","C",254,0})
      DBCREATE(NARQ3,MAT_STRU)
      SELE 33
      If !USEREDE(NARQ3,.T.,10,"TMP")
        SELE 30
        return 
      EndIf
      SELE 31
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 33
        USE
        SELE 30
        return 
      EndIf
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 33
        USE
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ4
      Erase &NARQ6
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      INDEX ON STRZERO(VAL(CNPJCPF),14)+IIF(TEMMAIL,"1","2") TO &NARQ6
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 30
        Erase &NARQ4
        Erase &NARQ6
        return 
      EndIf
      USE
      NARQ41=WEBLOCAL+"ASSESSOR.KEY"
      Erase &NARQ41
      SELE 39
      If !USEREDE(WEBLOCAL+"ASSESSOR.DBF",.T.,10,"ASSESSOR","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 35
        USE
        Erase &NARQ4
        Erase &NARE6
        SELE 30
        return 
      Else
        INDEX ON CODIGO TO &NARQ41
      EndIf
      USE
      If !USEREDE(WEBLOCAL+"ASSESSOR.DBF",.F.,10,"ASSESSOR","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 35
        USE
        Erase &NARQ4
        Erase &NARE6
        SELE 30
        return 
      Else
        SET INDEX TO &NARQ41
      EndIf
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 39
        USE
        Erase &NARQ41
        SELE 30
        Erase &NARQ4
        Erase &NARQ6
        return 
      EndIf
      SET INDEX TO &NARQ4,&NARQ6
      If WFAZDADOS
        SELE 36
        If !USEREDE(WEBLOCAL3+"DOCTOS.DBF",.F.,10,"DOCTOS","S")
          SELE 32
          USE
          SELE 33
          USE
          SELE 31
          USE
          SELE 39
          USE
          SELE 35
          USE
          Erase &NARQ4
          Erase &NARQ6
          Erase &NARQ41
          SELE 30
          return 
        EndIf
      EndIf
      SELE 31
      INDEX ON CLIENTE TO &NARQ5
      SELE 37
      If !USEREDE(WEBLOCAL4+"ETIQUETA.DBF",.T.,10,"ETIQ","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 35
        USE
        SELE 39
        USE
        SELE 36
        USE
        Erase &NARQ4
        Erase &NARQ41
        Erase &NARQ6
        SELE 30
        return 
      Else
        NARQ7=WEBLOCAL4+"ETIQUETA.KEY"
        INDEX ON CODIGO TO &NARQ7
      EndIf
      NRELAT=""
      WNOMEA=""
      WEND1A=""
      WEND2A=""
      WINFO=0
      CLIANTES=0
      RELINTERNO=.F.
      NARQ20=""
      FOR MMI=1 TO ULTIMO
        FEZALGO=.T.
        SysRefresh()
        CURSORWAIT()
        NARQ1=WEBLOCAL1+"O"+STRZERO(MMI,6)+".PDF"
        NARQ2=WEBLOCAL1+"O"+STRZERO(MMI,6)+".TXT"
        NARQ20=WEBLOCAL10+"O"+STRZERO(MMI,6)+".TXT"
        SELE 33
        ZAP
        If !FILE(NARQ2)
          LOOP
        EndIf
        SELE 33
        If INFSIMPAUL=.F.
          APPEND FROM &NARQ2 SDF
        Else
          oText:=TTxtfile():NEW(NARQ2)
          CONTAD=0
          CONTAD=oText:RecCount()
          If VALTYPE(CONTAD)#"N"
            CONTAD=1
          EndIf
          TI=0
          FOR TI=1 TO CONTAD
            SysRefresh()
            ttt=alltrim(oText:Readline())
            ADIREG(0)
            REPLACE LINHA WITH TTT
            oText:Skip()
          NEXT TI
          oText:CLOSE()
          oText:=nil
        EndIf
        CLIEN=0
        WWASSESSOR=0
        FEZ=.F.
        WWNOME=""
        WWENDE1=""
        WWENDE2=""
        NRELAT=""
        GO TOP
        If LASTREC()<10
          Erase &NARQ1
          Erase &NARQ2
          Erase &NARQ20
          LOOP
        EndIf
        If LASTREC()>10
          GO 5
          If AT("DEMONSTRAÇÃO FINANCEIRA POR ASSESSOR",LINHA)#0
            CLIEN=0
            FEZ=.T.
            NRELAT="022-RELATÓRIOS (0) POR ASSESSOR DA GERAÇÃO"
            GO 8
            If LEFT(LINHA,11)=" Assessor :"
              CLIEN=VAL(SUBSTR(LINHA,12,10))
            EndIf
            If CLIEN=0
              GO 7
              If LEFT(LINHA,11)=" Assessor :"
                CLIEN=VAL(SUBSTR(LINHA,12,10))
              EndIf
            EndIf
            If CLIEN=0
              CLIEN=CLIANTES
            Else
              CLIANTES=CLIEN
            EndIf
            RELINTERNO=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 6
          If AT("RELATÓRIO DE POSIÇÕES DE OPÇÕES POR ASSESSOR",LINHA)#0
            CLIEN=0
            FEZ=.T.
            NRELAT="022-RELATÓRIOS (1) POR ASSESSOR DA GERAÇÃO"
            DO WHILE !EOF()
              If LEFT(LINHA,11)="Assessor:  "
                CLIEN=VAL(SUBSTR(LINHA,12,30))
                EXIT
              Else
                SKIP
              EndIf
            ENDDO
            If CLIEN=0
              CLIEN=CLIANTES
            Else
              CLIANTES=CLIEN
            EndIf
            RELINTERNO=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 4
          If AT("RELATÓRIO DE POSIÇÕES DE TERMO POR ASSESSOR",LINHA)#0
            CLIEN=0
            FEZ=.T.
            NRELAT="022-RELATÓRIOS (2) POR ASSESSOR DA GERAÇÃO"
            DO WHILE !EOF()
              If LEFT(LINHA,11)="Assessor   "
                CLIEN=VAL(SUBSTR(LINHA,12,30))
                EXIT
              Else
                SKIP
              EndIf
            ENDDO
            If CLIEN=0
              CLIEN=CLIANTES
            Else
              CLIANTES=CLIEN
            EndIf
            RELINTERNO=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 2
          NRELAT=""
          If AT("RELATÓRIO DE MARGEM DE GARANTIAS",linha)#0
            NRELAT="044-RELATÓRIO DE MARGEM DE GARANTIAS"
            GO 10
            CLIEN=VAL(SUBSTR(LINHA,10,100))
            ATLI=AT("-",LINHA)+2
            WWNOME=ALLTRIM(SUBSTR(LINHA,ATLI,50))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>14 .And. CLIEN=0
          GO 1
          NRELAT=""
          If AT("Posição em Aberto de Derivativos Fungíveis",linha)#0
            NRELAT="045-Posição em Aberto de Derivativos Fungíveis"
            GO 14
            CLIEN=VAL(SUBSTR(LINHA,7,100))
            GO 16
            WWNOME=ALLTRIM(SUBSTR(LINHA,7,50))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>30 .And. CLIEN=0 .And. WEBCOR="EXATUS"  // Nota
          GO 1
          If AT("Sao Paulo,",LINHA)#0
            GO 2
            CLIEN=alltrim(SUBSTR(LINHA,at("CNPJ:",LINHA),25))
            CLIEN=Strtran(CLIEN,"CNPJ:","")
            CLIEN=Strtran(CLIEN,"CPF:","")
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            WWNOME=alltrim(LEFT(LINHA,43))
            SELE 35
            SET ORDER TO 2
            SEEK STRZERO(CLIEN,14)
            If FOUND() .And. CLIEN=VAL(CNPJCPF)
              CLIEN=CODIGO
              SELE 33
            Else
              DECLARE TX1[1]
              TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
              CLIEN=0
            EndIf
            SELE 35
            SET ORDER TO 1
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>39 .And. CLIEN=0 .And. WEBCOR="EXATUS" // boleto  verificar linha 2 e 34
          GO 6
          If AT("Banco Itaú S.A.",LINHA)#0
            GO 36
            If AT("Sacado",linha)#0
              GO 37
            Else
              GO 38
            EndIf
            CLIEN=alltrim(SUBSTR(LINHA,at("CNPJ -",LINHA),25))
            CLIEN=Strtran(CLIEN,"CNPJ -","")
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            WWNOME=alltrim(LEFT(LINHA,50))
            SELE 35
            SET ORDER TO 2
            SEEK STRZERO(CLIEN,14)
            If FOUND() .And. CLIEN=VAL(CNPJCPF)
              CLIEN=CODIGO
              SELE 33
            Else
              DECLARE TX1[1]
              TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
              CLIEN=0
            EndIf
            SELE 35
            SET ORDER TO 1
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>39 .And. CLIEN=0 .And. WEBCOR="EXATUS" // boleto  verificar linha 2 e 34
          GO 2
          If AT("Banco Itaú S.A.",LINHA)#0
            GO 33
            CLIEN=alltrim(SUBSTR(LINHA,at("CNPJ -",LINHA),25))
            CLIEN=Strtran(CLIEN,"CNPJ -","")
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            WWNOME=alltrim(LEFT(LINHA,50))
            SELE 35
            SET ORDER TO 2
            SEEK STRZERO(CLIEN,14)
            If FOUND() .And. CLIEN=VAL(CNPJCPF)
              CLIEN=CODIGO
              SELE 33
            Else
              DECLARE TX1[1]
              TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
              CLIEN=0
            EndIf
            SELE 35
            SET ORDER TO 1
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0 .And. WEBCOR="PLANNER"
          GO 1
          If AT("COMPROVANTE ANUAL DE RENDIMENTOS PAGOS OU",linha)#0
            GO 18
            If AT("CNPJ",LINHA)#0
              SKIP
              NRELAT="043-Informes - PLANNER"
              CLIEN=TRIM(LINHA)
              CLIEN=RIGHT(CLIEN,20)
              CLIEN=Strtran(CLIEN,"CNPJ:","")
              CLIEN=Strtran(CLIEN,"CPF:","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
              WWNOME=LEFT(LINHA,80)
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                SELE 33
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SELE 35
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        If LASTREC()>20 .And. CLIEN=0 .And. WEBCOR="PLANNER"
          GO 1
          If AT("         INFORME DE ATIVOS ESCRITURAIS PARA",linha)#0
            GO 10
            If AT("CPF / CNPJ:",LINHA)#0
              NRELAT="043-Informes Escriturais - PLANNER"
              CLIEN=alltrim(SUBSTR(LINHA,17,40))
              CLIEN=Strtran(CLIEN,"CNPJ:","")
              CLIEN=Strtran(CLIEN,"CPF:","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
              GO 8
              WWNOME=ALLTRIM(SUBSTR(LINHA,17,200))
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                SELE 33
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SELE 35
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0
          GO 1
          If AT(" INFORME DE RENDIMENTOS FINANCEIROS",LINHA)<>0
            GO 18
            ATLI=AT("Código:",linha)+8
            If ATLI#8
              CLIEN=val(substr(linha,atli,20))
            Else
              ATLI=AT("Código",linha)+7
              CLIEN=val(substr(linha,atli,20))
            EndIf
            If ATLI<9
              CLIEN=0
            EndIf
            If CLIEN#0
              FEZ=.T.
              NRELAT="029-INFORME DE RENDIMENTOS FINANCEIROS PLANNER - "
              SET ORDER TO 1
            EndIf
            SELE 33
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0
          GO 1
          If AT("Informe de Rendimentos Financeiros",LINHA)=0
            SKIP
          EndIf
          If AT("Informe de Rendimentos Financeiros",LINHA)=0
            SKIP
          EndIf
          If AT("Informe de Rendimentos Financeiros",LINHA)<>0
            GO BOTT
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT("Código do cliente:",linha)=0
              SKIP -1
            EndIf
            If AT(" Código do cliente:",linha)<>0
              ATLI=AT("Código do cliente:",linha)+18
              CLIEN=ALLTRIM(SUBSTR(LINHA,atli,20))
              ATLI=AT("-",CLIEN)
              CLIEN=LEFT(CLIEN,ATLI)
              CLIEN=VAL(CLIEN)
            EndIf
            If CLIEN#0
              FEZ=.T.
              NRELAT="036-INFORME DE RENDIMENTOS SIMPAUL -"
              SET ORDER TO 1
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0 .And. WEBCOR="PLANNER"
          GO 1
          If AT("           INFORME DE RENDIMENTOS FINANCEIROS",linha)#0
            GO 17
            If AT("CPF",LINHA)#0 .OR. AT("CNPJ",LINHA)#0
              ATC=AT("NOME COMPLETO",LINHA)
              NRELAT="040-Informe FUNDOS - PLANNER"
              SKIP
              CLIEN=LEFT(LINHA,30)
              CLIEN=Strtran(CLIEN,"CNPJ:","")
              CLIEN=Strtran(CLIEN,"CPF:","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                SELE 33
                WWNOME=TRIM(SUBSTR(LINHA,ATC,100))
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SELE 35
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>12 .And. CLIEN=0 .And. WEBCOR="PLANNER"
          GO 1
          If AT(" Informe de Rendimentos -",linha)#0
            GO 10
            If AT("CPF/CNPJ:",LINHA)#0
              NRELAT="041-Informe Tesouro Direto/Renda fixa - PLANNER"
              CLIEN=SUBSTR(LINHA,10,25)
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                SELE 33
                WWNOME=ALLTRIM(SUBSTR(LINHA,AT("NOME:",LINHA)+5,100))
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SELE 35
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0 .And. WEBCOR="PLANNER"
          GO 1
          If AT("           INFORME DE RENDIMENTOS FINANCEIROS",linha)#0
            GO 21
            If AT("CPF:",LINHA)#0 .OR. AT("CNPJ:",LINHA)#0
              NRELAT="040-Informe FUNDOS - PLANNER"
              CLIEN=SUBSTR(LINHA,6,30)
              CLIEN=Strtran(CLIEN,"CNPJ:","")
              CLIEN=Strtran(CLIEN,"CPF:","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                SELE 33
                GO 17
                WWNOME=ALLTRIM(LINHA)
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SELE 35
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>9 .And. CLIEN=0
          GO 1
          NRELAT=""
          If AT("Extrato Consolidado de Clientes",linha)#0
            GO 10
            NRELAT="034-Extrato Consolidado de clientes"
            go 5
            CLIEN=SUBSTR(LINHA,AT("CPF :",LINHA)+5,30)
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            SELE 35
            SET ORDER TO 2
            SEEK STRZERO(CLIEN,14)
            If FOUND() .And. CLIEN=VAL(CNPJCPF)
              CLIEN=CODIGO
              SELE 33
              GO 5
              WWNOME=SUBSTR(LINHA,4,60)
              SKIP 1
              WWENDE1=ALLTRIM(LINHA)
              SKIP 1
              WWENDE2=ALLTRIM(LINHA)
            Else
              DECLARE TX1[1]
              TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
              CLIEN=0
            EndIf
            SELE 35
            SET ORDER TO 1
          EndIf
          SELE 33
          FEZ=.T.
          If CLIEN=0
            SELE 33
            FEZ=.F.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>18 .And. CLIEN=0
          GO 1
          If AT(" Nota de Negociação de Títulos  ",LINHA)<>0
            GO 17
            If LEFT(LINHA,9)<>"Nome     "
              GO 18
            EndIf
            If LEFT(LINHA,9)="Nome     "
              CLIEN=VAL(substr(LINHA,8,15))
              If CLIEN#0
                NRELAT="031-Renda Fixa - Geração"
                atli=at(" - ",linha)
                WWNOME=alltrim(substr(linha,atli+3,60))
                FEZ=.T.
              EndIf
            EndIf
          EndIf
        EndIf
        SELE 33
        If CLIEN=0 .And. LASTREC()>16
          GO 1
          If AT("   Nota de Balcão",linha)#0
            NRELAT="024-NOTAS DE BALCÃO"
            FEZ=.T.
            GO 20
            CLIEN=VAL(LINHA)
            go 15
            WWNOME=ALLTRIM(LINHA)
            go 17
            WWENDE1=ALLTRIM(LINHA)
            GO 19
            WWENDE2=substr(ALLTRIM(LINHA),6,15)
            If CLIEN=0
              GO 19
              CLIEN=VAL(LINHA)
              go 14
              WWNOME=ALLTRIM(LINHA)
              go 16
              WWENDE1=ALLTRIM(LINHA)
              GO 18
              WWENDE2=substr(ALLTRIM(LINHA),6,15)
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>0 .And. CLIEN=0
          GO 1
          If AT("COMPROVANTE DE RENDIMENTOS PAGOS E D",LINHA)#0 .And. CLIEN=0
            FEZ=.T.
            NRELAT="021-INFORMES(FOLHA)"
            DO WHILE !EOF()
              If AT("Nome Completo",LINHA)#0
                SKIP
                WINFO=WINFO+1
                CLIEN=WINFO
                WWNOME=ALLTRIM(SUBSTR(LINHA,30,50))
                WWENDE1=""
                WWENDE2=""
                EXIT
              EndIf
              SKIP
            ENDDO
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 1
          If AT("         Extrato de Cotistas",LINHA)#0
            GO 5
            CLIEN=VAL(LEFT(STRZERO(VAL(SUBSTR(LINHA,107,147)),10),9))
            If CLIEN#0
              NRELAT="013-Extrato de Cotistas"
              WWNOME=ALLTRIM(LEFT(LINHA,107))
              If EMPTY(WWNOME)
                SKIP
                WWNOME=ALLTRIM(LEFT(LINHA,107))
              EndIf
              SKIP
              WWENDE1=ALLTRIM(LEFT(LINHA,107))
              SKIP
              WWENDE2=ALLTRIM(SUBSTR(LINHA,50,44))
              If !EMPTY(WWENDE2)
                WWENDE2=WWENDE2+"-"+alltrim(substr(linha,100,30))
                SKIP
                WCEP=SUBSTR(LINHA,21,30)
                WCEP1=Strtran(WCEP," ","")
                WWENDE2=WCEP1+" - "+WWENDE2
              Else
                WWD=ALLTRIM(SUBSTR(LINHA,100,30))
                SKIP
                WWENDE2=LEFT(LINHA,82)+" "+WWD
                SKIP
                WCEP=SUBSTR(LINHA,21,30)
                WCEP1=Strtran(WCEP," ","")
                WWENDE2=WCEP1+" - "+WWENDE2
              EndIf
              If wwende1="0    "
                wwende1=""
                wwende2=""
              EndIf
              FEZ=.T.
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>12 .And. CLIEN=0
          GO 2
          If AT("                                       Carteira de Ações",LINHA)#0
            GO 12
            If AT(" Código:",LINHA)#0
              CLIEN=VAL(SUBSTR(LINHA,10,40))
            EndIf
            If CLIEN#0
              NRELAT="016-Carteira de Ações"
              go 7
              WWNOME=ALLTRIM(LINHA)
              SKIP
              WWENDE1=ALLTRIM(LINHA)
              SKIP 2
              WWENDE2=ALLTRIM(LEFT(LINHA,76))+"-"+ALLTRIM(SUBSTR(LINHA,80,60))
              FEZ=.T.
              WNOMEA=WWNOME
              WEND1A=WWENDE1
              WEND2A=WWENDE2
              CLIANTES=CLIEN
            EndIf
            If CLIEN=0
              NRELAT="016-Carteira de Ações-Continua"
              CLIEN=CLIANTES
              WWNOME=WNOMEA
              WWENDE1=WEND1A
              WWENDE2=WEND2A
              FEZ=.T.
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>8 .And. CLIEN=0
          GO 6
          If AT("         Posição Consolidada  ",LINHA)#0
            GO 8
            ATLI=AT("Cliente: ",LINHA)
            If ATLI#0
              CLIEN=VAL(SUBSTR(LINHA,ATLI+11,40))
              If CLIEN#0
                NRELAT="017-Posição Consolidada"
                go 8
                FEZ=.T.
              EndIf
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>15 .And. CLIEN=0
          GO 5
          If AT("  Extrato número : ",LINHA)#0
            GO 4
            ATLI=AT(" Cliente : ",LINHA)
            If ATLI#0
              CLIEN=VAL(SUBSTR(LINHA,ATLI+11,40))
              If CLIEN#0
                NRELAT="018-Extrato Custodia"
                GO 1
                FOL=AT("Folha :",LINHA)
                FOL1=VAL(SUBSTR(LINHA,FOL+8,40))
                If FOL1=1
                  GO 7
                  do while (empty(LINHA) .or. AT("Título",LINHA)<>0)
                    skip
                  enddo
                  WWNOME=ALLTRIM(LINHA)
                  SKIP
                  WWENDE1=ALLTRIM(LINHA)
                  SKIP
                  WWENDE3=ALLTRIM(LINHA)
                  SKIP
                  WWENDE2=ALLTRIM(LINHA)+" "+WWENDE3
                Else
                  WWNOME=" "
                  WWENDE1=" "
                  WWENDE2=" "
                EndIf
                FEZ=.T.
              EndIf
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>12 .And. CLIEN=0
          GO 2
          If AT("        Carteira de Titulos de Renda Fixa",LINHA)#0
            GO 10
            ATLI=AT("   Codigo: ",LINHA)
            If ATLI#0
              CLIEN=VAL(RIGHT(STRZERO(VAL(SUBSTR(LINHA,ATLI+12,40)),6),5))
              If CLIEN#0
                NRELAT="015-Renda Fixa"
                go 7
                WWNOME=ALLTRIM(LINHA)
                SKIP 2
                WWENDE1=ALLTRIM(substr(LINHA,10,100))
                SKIP 2
                WWENDE2=SUBSTR(LINHA,6,11)
                ATLI=AT("Cidade: ",LINHA)
                WWENDE2=WWENDE2+ALLTRIM(SUBSTR(LINHA,ATLI+8,30))
                ATLI=AT("Estado: ",LINHA)
                WWENDE2=WWENDE2+"-"+ALLTRIM(SUBSTR(LINHA,ATLI+8,10))
                FEZ=.T.
              EndIf
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>15 .And. CLIEN=0
          GO 1
          nrelat=""
          If LEFT(LINHA,95)="                                                                             NOTA DE CORRETAGEM" .OR. LEFT(LINHA,91)="                                                                         NOTA DE CORRETAGEM" .OR. LEFT(LINHA,96)="                                                                              NOTA DE CORRETAGEM"
            NRELAT="001-Nota de Corretagem Bovespa"
            LI=12
            GO 12
            CLIEN=VAL(left(LINHA,35))
            SKIP 2
            WWASSESSOR=VAL(RIGHT(ALLTRIM(LINHA),6))
            SKIP -2
            If CLIEN=0
              GO 13
              LI=13
              CLIEN=VAL(LINHA)
              SKIP 2
              WWASSESSOR=VAL(RIGHT(ALLTRIM(LINHA),6))
              SKIP -2
            EndIf
            WWNOME=""
            If CLIEN#0 .And. EMPTY(WWNOME)
              WWNOME=ALLTRIM(SUBSTR(LINHA,20,150))
              SKIP
              WWENDE1=ALLTRIM(SUBSTR(LINHA,20,150))
              SKIP
              WWENDE2=ALLTRIM(SUBSTR(LINHA,20,150))
              WWENDE2=LEFT(WWENDE2,AT(" - ",WWENDE2)+10)
              GO LI
              If !EMPTY(WWNOME)
                SKIP
                WWENDE1=ALLTRIM(SUBSTR(LINHA,20,150))
                SKIP
                WWENDE2=ALLTRIM(SUBSTR(LINHA,20,150))
                WWENDE2=LEFT(WWENDE2,AT(" - ",WWENDE2)+10)
              EndIf
              If EMPTY(WWNOME) .And. LI=12
                SKIP
                WWNOME=ALLTRIM(SUBSTR(LINHA,20,150))
                SKIP
                WWENDE1=ALLTRIM(SUBSTR(LINHA,20,150))
                SKIP
                WWENDE2=ALLTRIM(SUBSTR(LINHA,20,150))
                WWENDE2=LEFT(WWENDE2,AT(" - ",WWENDE2)+10)
              EndIf
              If EMPTY(WWNOME) .And. LI=13
                SKIP -1
                WWNOME=ALLTRIM(SUBSTR(LINHA,20,150))
                SKIP 2
                WWENDE1=ALLTRIM(SUBSTR(LINHA,20,150))
                SKIP
                WWENDE2=ALLTRIM(SUBSTR(LINHA,20,150))
                WWENDE2=LEFT(WWENDE2,AT(" - ",WWENDE2)+10)
              EndIf
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>3 .And. CLIEN=0
          GO 1
          If AT("   CÓDIGO COTISTA:",LINHA)#0
            NRELAT="019-Extrato de Cotistas - S.Barros"
            SKIP
            SKIP
            ATLI=AT("CNPJ/CPF:",LINHA)
            If ATLI=0
              SKIP
              ATLI=AT("CNPJ/CPF:",LINHA)
            EndIf
            CLIEN=SUBSTR(LINHA,ATLI+9,30)
            CLIEN=Strtran(CLIEN,":","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
          EndIf
          If CLIEN#0
            SKIP
            WWNOME=LINHA
            SKIP
            WWENDE1=LINHA
            SKIP
            wlinha=alltrim(linha)
            wlinha=substr(wlinha,2,200)
            wlinha=alltrim(wlinha)
            WWENDE2=LEFT(wLINHA,5)+"-"+SUBSTR(wLINHA,6,200)
            SELE 35
            SET ORDER TO 2
            SEEK STRZERO(CLIEN,14)
            If FOUND() .And. CLIEN=VAL(CNPJCPF)
              CLIEN=CODIGO
            Else
              CLIEN=0
            EndIf
            SET ORDER TO 1
          EndIf
          SELE 33
          FEZ=.T.
        EndIf
        SELE 33
        If CLIEN=0 .And. LASTREC()>9
          GO 1
          If AT("Extrato de Movimento de Conta Corrente",LINHA)#0
            NRELAT="012-Extrato de Contas Correntes"
            GO 9
            CLIEN=VAL(LEFT(STRZERO(VAL(LINHA),10),9))
            If CLIEN=0
              GO 10
              CLIEN=VAL(LEFT(STRZERO(VAL(LINHA),10),9))
            EndIf
            If CLIEN#0
              GO 2
              ATLI=AT("Cliente:",LINHA)
              WWNOME=ALLTRIM(SUBSTR(LINHA,ATLI+10,60))
              SKIP
              WWENDE1=ALLTRIM(SUBSTR(LINHA,ATLI+10,60))
              SKIP 3
              ATLI=AT("Cep: ",LINHA)
              WWCEP=SUBSTR(LINHA,ATLI+5,10)
              SKIP
              WWENDE2=WWCEP+" "+ALLTRIM(LINHA)
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>16 .And. CLIEN=0
          GO 1
          NRELAT=""
          If AT("   NOTA DE CORRETAGEM  ",LINHA)#0 .OR. AT("   NOTA DE BROKERAGE  ",LINHA)#0 .OR. AT("  NOTA DE BROKERAGEM ",LINHA)#0 .or. at("NOTA DE NEGOCIAÇÃO ",LINHA)#0
            NRELAT="004-Nota de Corretagem BM&F"
            GO 15
            If AT("Código do cliente",LINHA)#0
              ATLI=AT("Código do cliente",LINHA)
              SKIP
              CLIEN=VAL(SUBSTR(LINHA,ATLI,40))
              If CLIEN#0
                skip -2
                WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                SKIP
                WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                SKIP
                WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                FEZ=.T.
              EndIf
            EndIf
            GO 14
            If AT("Código do cliente",LINHA)#0 .And. CLIEN=0
              ATLI=AT("Código do cliente",LINHA)
              LI=1
              SKIP
              CLIEN=VAL(SUBSTR(LINHA,ATLI,40))
              If CLIEN=0
                SKIP
                CLIEN=VAL(SUBSTR(LINHA,ATLI,40))
                LI=2
              EndIf
              If CLIEN#0
                If LI=2
                  SKIP -3
                Else
                  SKIP -2
                EndIf
                WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                SKIP
                WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                SKIP
                WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                If EMPTY(WWENDE2)
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
                FEZ=.T.
              EndIf
            EndIf
          EndIf
        EndIf
        SELE 33
        If LASTREC()>15 .And. CLIEN=0
          GO 11
          NRELAT=""
          If LEFT(LINHA,12)="Cliente     "
            NRELAT="001-Nota de Corretagem Bovespa"
            GO 12
            CLIEN=VAL(LINHA)
            SKIP 2
            WWASSESSOR=VAL(RIGHT(ALLTRIM(LINHA),6))
            SKIP -2
            If CLIEN=0
              GO 13
              CLIEN=VAL(LINHA)
              SKIP 2
              WWASSESSOR=VAL(RIGHT(ALLTRIM(LINHA),6))
              SKIP -2
            EndIf
            If CLIEN#0
              WWNOME=ALLTRIM(SUBSTR(LINHA,20,150))
              If EMPTY(WWNOME)
                SKIP
                WWNOME=ALLTRIM(SUBSTR(LINHA,20,150))
              EndIf
              SKIP
              WWENDE1=ALLTRIM(SUBSTR(LINHA,20,150))
              SKIP
              WWENDE2=ALLTRIM(SUBSTR(LINHA,20,150))
              WWENDE2=LEFT(WWENDE2,AT(" - ",WWENDE2)+10)
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>8 .And. CLIEN=0
          GO 1
          NRELAT=""
          If AT("            AVISO DE CTA CORRENTES - SOLICI",LINHA)#0 .or. at("       Aviso de Contas Correntes",LINHA)#0
            NRELAT="002-Aviso de Ctas Correntes"
            GO 7
            If LEFT(LINHA,50)=SPACE(50)
              GO 8
            EndIf
            If AT("(",LINHA)#0
              CLIEN=SUBSTR(LINHA,AT("(",LINHA)+1,50)
              DO WHILE AT("(",CLIEN)#0
                SysRefresh()
                CLIEN=SUBSTR(CLIEN,AT("(",CLIEN)+1,50)
              ENDDO
            EndIf
            If VALTYPE(CLIEN)="C"
              CLIEN=VAL(CLIEN)
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 6
          NRELAT=""
          If AT("               Autorização de Subscrição  ",LINHA)#0
            NRELAT="002-Autorização de Subscrição"
            go bott
            skip -1
            clien=trim(linha)
            clien=right(clien,16)
            clien=Strtran(clien," ","")
            CLIEN=VAL(clien)
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>23 .And. CLIEN=0
          WLIN=6
          GO 6
          If AT("Informe de Rendimentos Financeiros",linha)=0
            WLIN=7
            GO 7
          EndIf
          If AT("Informe de Rendimentos Financeiros",linha)#0
            If WLIN=7
              GO 23
            Else
              GO 22
            EndIf
            CLIEN=SUBSTR(LINHA,80,40)
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=VAL(CLIEN)
            If CLIEN#0
              FEZ=.T.
              NRELAT="003-Apuração de IR Day Trade e Operações"
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SET ORDER TO 1
            EndIf
          EndIf
          SELE 33
        EndIf
        SELE 33
        If LASTREC()>23 .And. CLIEN=0
          WLIN=6
          GO 1
          If AT("Informe de Rendimentos Financeiros",linha)#0
            GO 15
            If EMPTY(LINHA)
              SKIP
            EndIf
            If AT("2 - Pessoa Física Beneficiária dos Rendimentos",linha)#0 .OR. AT("2 - Pessoa Jurídica Beneficiária dos Rendimentos",LINHA)#0
              SKIP 2
              CLIEN=SUBSTR(LINHA,100,40)
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=VAL(CLIEN)
              If CLIEN#0
                FEZ=.T.
                NRELAT="042-Informes Genial Investimentos"
                SELE 35
                SET ORDER TO 2
                SEEK STRZERO(CLIEN,14)
                If FOUND() .And. CLIEN=VAL(CNPJCPF)
                  CLIEN=CODIGO
                Else
                  DECLARE TX1[1]
                  TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                  If !EMPTY(WDADOS)
                    LOGFILE(WDADOS,TX1)
                  EndIf
                  CLIEN=0
                EndIf
                SET ORDER TO 1
              EndIf
            EndIf
          EndIf
          SELE 33
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0
          GO 9
          NRELAT=""
          If AT(" APURAÇÃO DO IRRF",LINHA)=0 .And. AT("Natureza : Total auferido",LINHA)=0
            go 8
          EndIf
          If AT(" APURAÇÃO DO IRRF",LINHA)#0 .OR. AT("Natureza : Total auferido",LINHA)#0
            NRELAT="029-Apuração de IR Day Trade e Operações"
            GO 19
            CLIEN=SUBSTR(LINHA,15,40)
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=Strtran(CLIEN," ","")
            CLIEN=VAL(CLIEN)
            If CLIEN=0
              GO 18
              CLIEN=SUBSTR(LINHA,15,40)
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=Strtran(CLIEN," ","")
              CLIEN=VAL(CLIEN)
            EndIf
            SKIP
            WWNOME=LINHA
            SKIP
            WWENDE1=LINHA
            SKIP
            SKIP
            WWENDE2=ALLTRIM(LINHA)
            SKIP -1
            WWENDE2=WWENDE2+ALLTRIM(LINHA)
            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
            FEZ=.T.
            If CLIEN=0
              SELE 33
              FEZ=.T.
            EndIf
            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                wwnome=nome
                WNOMEA=NOME
                WEND1A=ENDE1
                WEND2A=ENDE2
                WWENDE1=ENDE1
                WWENDE2=ENDE2
                WWNOME=NOME
                CLIANTES=CLIEN
              Else
                DECLARE TX1[1]
                TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
                If !EMPTY(WDADOS)
                  LOGFILE(WDADOS,TX1)
                EndIf
                CLIEN=0
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
            FEZ=.T.
            If CLIEN=0
              SELE 33
              FEZ=.T.
            EndIf
          EndIf
        EndIf
        If CLIEN=0
          GO 1
          If LINHA="Informe Geração:"
            NRELAT="003-Informe Rendimentos Genial Investimentos"
            CLIEN=VAL(SUBSTR(LINHA,18,8))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0
          GO 5
          NRELAT=""
          If AT("Comprovante de Rendimentos Pagos ou Creditados e de Retenção de Imposto de Renda na Fonte",LINHA)#0 .or. at("Comprovante das Operações e de Retenção de Imposto de Renda na Fonte",linha)#0
            NRELAT="029-Comprovante das Operações e de Retenção de Imposto de Renda na Fonte-Fisica"
            GO 23
            CLIEN=SUBSTR(LINHA,15,30)
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            SKIP
            WWNOME=ALLTRIM(LINHA)
            SKIP
            WWENDE1=ALLTRIM(LINHA)
            SKIP
            SKIP
            WWENDE2=ALLTRIM(LINHA)
            SKIP -1
            WWENDE2=WWENDE2+" "+ALLTRIM(LINHA)

            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                FEZ=.T.
              Else
                CLIEN=0
                FEZ=.F.
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
          EndIf
        EndIf
        SELE 33
        If LASTREC()>20 .And. CLIEN=0
          GO 5
          NRELAT=""
          If AT("Comprovante das Operações e de Retenção de Imposto",LINHA)#0 .or. at("Comprovante de Rendimentos Pagos ou Creditados e de Retenção de Imposto de Renda na Fonte",LINHA)#0
            NRELAT="029-Comprovante Tesouro Direto - HP"
            GO 19
            CLIEN=ALLTRIM(Strtran(LINHA,"CPF:",""))
            CLIEN=Strtran(CLIEN,"CNPJ:","")
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=Strtran(CLIEN,"-","")
            CLIEN=Strtran(CLIEN,"/","")
            CLIEN=VAL(CLIEN)
            SKIP
            WWNOME=ALLTRIM(LINHA)
            SKIP
            WWENDE1=ALLTRIM(LINHA)
            SKIP
            SKIP
            WWENDE2=ALLTRIM(LINHA)
            SKIP -1
            WWENDE2=WWENDE2+" "+ALLTRIM(LINHA)
            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
                FEZ=.T.
              Else
                CLIEN=0
                FEZ=.F.
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
          EndIf
        EndIf
        SELE 33
        If LASTREC()>23 .And. CLIEN=0
          GO 5
          NRELAT=""
          If LEFT(LINHA,46)="Comprovante de Rendimentos Pagos ou Creditados"
            NRELAT="003-Comprovante de Rendimentos Pagos ou Creditados  e de Retenção de Imposto de Renda na Fonte-Juridica"
            GO 21
            CLIEN=VAL(SUBSTR(LINHA,6,30))
            If CLIEN=0
              GO 19
              CLIEN=SUBSTR(LINHA,15,30)
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
            EndIf
            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
              Else
                CLIEN=0
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>23 .And. CLIEN=0
          GO 9
          NRELAT=""
          If AT("APURAÇÃO DO IRRF SOBRE OPERAÇÕES",LINHA)#0
            NRELAT="003-Comprovante de Rendimentos Pagos ou Creditados  e de Retenção de Imposto de Renda na Fonte-Juridica"
            GO 21
            CLIEN=VAL(SUBSTR(LINHA,6,30))
            If CLIEN=0
              GO 19
              CLIEN=SUBSTR(LINHA,15,40)
              CLIEN=Strtran(CLIEN,".","")
              CLIEN=Strtran(CLIEN,"-","")
              CLIEN=Strtran(CLIEN,"/","")
              CLIEN=VAL(CLIEN)
            EndIf
            If CLIEN#0
              SELE 35
              SET ORDER TO 2
              SEEK STRZERO(CLIEN,14)
              If FOUND() .And. CLIEN=VAL(CNPJCPF)
                CLIEN=CODIGO
              Else
                CLIEN=0
              EndIf
              SET ORDER TO 1
            EndIf
            SELE 33
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>16 .And. CLIEN=0
          GO 1
          NRELAT=""
          If AT("   NOTA DE CORRETAGEM  ",LINHA)#0 .OR. AT("   NOTA DE BROKERAGE  ",LINHA)#0 .OR. AT("  NOTA DE BROKERAGEM ",LINHA)#0 .or. at("NOTA DE NEGOCIAÇÃO ",LINHA)#0
            NRELAT="004-Nota de Corretagem BM&F"
            If CLIEN=0
              GO 14
              If AT("Código do cliente",LINHA)#0
                ATLI=AT("Código do cliente",LINHA)
                SKIP
                CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              EndIf
            EndIf
            If CLIEN=0
              GO 4
              If LEFT(LINHA,10)="Corretora "
                GO 12
                If AT("Código do cliente",LINHA)#0
                  ATLI=AT("Código do cliente",LINHA)
                  SKIP
                  CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                  If CLIEN#0
                    skip -2
                    WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                  EndIf
                EndIf
              EndIf
            EndIf
            If CLIEN=0
              GO 5
              If LEFT(LINHA,10)="Corretora "
                GO 13
                If AT("Código do cliente",LINHA)#0
                  ATLI=AT("Código do cliente",LINHA)
                  SKIP
                  CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                  If CLIEN#0
                    skip -2
                    WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                  EndIf
                EndIf
              EndIf
            EndIf
            If CLIEN=0
              GO 15
              If AT("Código do cliente",LINHA)#0
                ATLI=AT("Código do cliente",LINHA)
                SKIP
                CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                ElseIf CLIEN=0
                  SKIP
                  CLIEN=VAL(SUBSTR(LINHA,ATLI,40))
                  If CLIEN#0
                    SKIP -3
                    WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                    If EMPTY(WWENDE2)
                      SKIP
                      WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                    EndIf
                  EndIf
                EndIf
              EndIf
            EndIf
            If CLIEN=0
              GO 14
              CLIEN=SUBSTR(LINHA,150,101)
              If AT("-",CLIEN)=0 .And. AT("/",CLIEN)=0
                CLIEN=VAL(SUBSTR(LINHA,150,100))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              Else
                CLIEN=0
              EndIf
            EndIf
            If CLIEN=0
              GO 13
              CLIEN=SUBSTR(LINHA,174,40)
              If AT("-",CLIEN)=0 .And. AT("/",CLIEN)=0
                CLIEN=VAL(SUBSTR(LINHA,174,41))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              Else
                clien=val(clien)
              EndIf
            EndIf
            If CLIEN=0
              CLIEN=SUBSTR(LINHA,202,41)
              If AT("-",CLIEN)=0 .And. AT("/",CLIEN)=0
                clien=val(clien)
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              EndIf
            EndIf
            If CLIEN=0 .And. lastrec()>19
              GO 19
              If AT("-",LINHA)=0 .And. AT("/",LINHA)=0
                CLIEN=VAL(LINHA)
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              EndIf
            EndIf
            If CLIEN=0 .And. lastrec()>19
              GO 20
              If AT("-",LINHA)=0 .And. AT("/",LINHA)=0
                CLIEN=VAL(LINHA)
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              EndIf
            EndIf
            If CLIEN=0 .And. lastrec()>19
              GO 17
              If AT("Código do cliente",LINHA)#0
                ATLI=AT("Código do cliente",LINHA)
                SKIP
                CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  If EMPTY(WWNOME)
                    SKIP
                    WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  EndIf
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                EndIf
              EndIf
            EndIf
            If CLIEN=0 .And. lastrec()>19
              GO 16
              If AT("Código do cliente",LINHA)#0
                ATLI=AT("Código do cliente",LINHA)
                SKIP
                CLIEN=VAL(SUBSTR(LINHA,ATLI,41))
                If CLIEN#0
                  skip -2
                  WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                  SKIP
                  WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                ElseIf CLIEN=0
                  SKIP
                  CLIEN=VAL(SUBSTR(LINHA,ATLI,40))
                  If CLIEN#0
                    SKIP -3
                    WWNOME=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE1=ALLTRIM(SUBSTR(LINHA,20,100))
                    SKIP
                    WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                    If EMPTY(WWENDE2)
                      SKIP
                      WWENDE2=ALLTRIM(SUBSTR(LINHA,20,100))
                    EndIf
                  EndIf
                EndIf
              EndIf
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>10 .And. CLIEN=0
          GO 2
          NRELAT=""
          If AT("RL/KF/DA03/11",LINHA)#0
            NRELAT="005-Posição dos Clientes"
            GO 9
            ATLI=AT("-",linha)
            WCOD=SUBSTR(LINHA,ATLI+1,15)
            ATLI=AT("-",WCOD)
            WCOD=LEFT(WCOD,ATLI-1)
            CLIEN=VAL(WCOD)
            If CLIEN=0
              GO 10
              ATLI=AT("-",linha)
              WCOD=SUBSTR(LINHA,ATLI+1,15)
              ATLI=AT("-",WCOD)
              WCOD=LEFT(WCOD,ATLI-1)
              CLIEN=VAL(WCOD)
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>8 .And. CLIEN=0
          GO 5
          NRELAT=""
          If AT("Demonstrativo da Posição dos Investidores do Participante",LINHA)#0
            NRELAT="006-Demonstrativo da Posição dos Investidores do Participante"
            GO 8
            CLIEN=VAL(SUBSTR(LINHA,12,7))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>7 .And. CLIEN=0
          GO 5
          If AT("DEMONSTRATIVO DA POSICAO DOS CLIENTES DA SOCIEDADE CORRETORA",LINHA)#0
            NRELAT="006-DEMONSTRATIVO DA POSICAO DOS CLIENTES DA SOCIEDADE CORRETORA"
            GO 7
            CLIEN=VAL(Strtran(SUBSTR(LINHA,12,9),".",""))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>11 .And. CLIEN=0
          GO 5
          If AT("EXIGENCIA DE MARGEM DA GARANTIA",LINHA)#0
            NRELAT="007-EXIGENCIA DE MARGEM DA GARANTIA"
            GO 11
            CLIEN=VAL(Strtran(SUBSTR(LINHA,5,7),".",""))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>8 .And. CLIEN=0
          GO 5
          If AT("Liquidações",LINHA)#0
            NRELAT="008-Liquidações Realizadas para o Investidor"
            GO 8
            CLIEN=VAL(SUBSTR(LINHA,13,7))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>7 .And. CLIEN=0
          GO 3
          If AT("DEMONSTRATIVO DE OPERACOES A TERMO",LINHA)#0
            NRELAT="009-DEMONSTRATIVO DE OPERACOES A TERMO"
            GO 6
            CLIEN=SUBSTR(LINHA,16,10)
            If AT("-",CLIEN)#0
              CLIEN=LEFT(CLIEN,AT("-",CLIEN))
            EndIf
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=VAL(Strtran(CLIEN,"-",""))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If CLIEN=0
          GO 2
          If AT("RELAÇÃO DAS GARANTIAS DEPOSITADAS",LINHA)#0
            NRELAT="010-DEMONSTRATIVO DE ATIVOS DEPOSITADOS NO SISTEMA DE GARANTIAS"
            GO 10
            If LEFT(LINHA,9)="Cliente: "
              CLIEN=VAL(SUBSTR(LINHA,11,50))
              FEZ=.T.
            EndIf
          EndIf
        EndIf
        SELE 33
        If CLIEN=0 .And. lastrec()>12
          GO 2
          If AT("Relatório de Garantias Depositadas   ",LINHA)#0
            NRELAT="010-DEMONSTRATIVO DE ATIVOS DEPOSITADOS NO SISTEMA DE GARANTIAS"
            GO 1
            FOR ZNI=1 TO 12
              If AT("Cliente:",LINHA)#0
                CLIEN=VAL(substr(LINHA,9,100))
                WWNOME=alltrim(substr(linha,AT(" - ",LINHA)+3,100))
                FEZ=.T.
                EXIT
              EndIf
              SKIP
            NEXT ZNI
          EndIf
        EndIf
        SELE 33
        If LASTREC()>5 .And. CLIEN=0
          GO 3
          If AT("DEMONSTRATIVO DE ATIVOS",LINHA)#0
            NRELAT="010-DEMONSTRATIVO DE ATIVOS DEPOSITADOS NO SISTEMA DE GARANTIAS"
            GO 5
            NA=AT("CLIENTE :",LINHA)
            CLIEN=SUBSTR(LINHA,NA+11,50)
            If AT("-",CLIEN)#0
              CLIEN=LEFT(CLIEN,AT("-",CLIEN))
            EndIf
            CLIEN=Strtran(CLIEN,".","")
            CLIEN=VAL(Strtran(CLIEN,"-",""))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>6 .And. CLIEN=0
          GO 3
          If AT("ATIVOS",LINHA)#0
            NRELAT="010-Relação dos Ativos por corretora PG:"+NARQ1
            GO 6
            CLIEN=int(VAL(alltrim(SUBSTR(LINHA,64,10))))
            If CLIEN=0
              CLIEN=INT(VAL(ALLTRIM(SUBSTR(LINHA,66,8))))
            EndIf
            If CLIEN=0
              CLIEN=INT(VAL(ALLTRIM(SUBSTR(LINHA,67,9))))
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If LASTREC()>9 .And. CLIEN=0
          GO 3
          If AT("POSICAO",LINHA)#0
            NRELAT="011-Demonstrativo do Movimento / Posição dos Clientes PG:"+NARQ1
            GO 9
            CLIEN=VAL(SUBSTR(LINHA,AT("-",LINHA)+2,6))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If CLIEN=0 .And. LASTREC()>7
          GO 7
          If AT("Extrato de Contas Correntes",LINHA)#0
            NRELAT="023-Extrato de Contas Correntes PG:"+NARQ1
            go 12
            CLIEN=VAL(SUBSTR(LINHA,10,AT("-",LINHA)-1))
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If CLIEN=0 .And. LASTREC()>2
          GO 2
          NFZ=.F.
          If AT("Ctas Correntes - ",LINHA)#0 .or. at("Tesouraria -",LINHA)#0
            NFZ=.T.
          EndIf
          GO 1
          If AT("Ctas Correntes - ",LINHA)#0 .OR. AT("Tesouraria -",LINHA)#0
            NFZ=.T.
          EndIf
          If NFZ
            FEZ=.T.
            na=0
            If lastrec()>=21
              GO 21
              na=21
            EndIf
            If LEFT(LINHA,16)#"CONTA CORRENTE :" .And. LEFT(LINHA,23)#"CONTA DE INVESTIMENTO :" .And. lastrec()>=22
              GO 22
              na=22
            EndIf
            If LEFT(LINHA,16)#"CONTA CORRENTE :" .And. LEFT(LINHA,23)#"CONTA DE INVESTIMENTO :" .And. lastrec()>=17
              GO 17
              NA=17
            EndIf
            If LEFT(LINHA,16)#"CONTA CORRENTE :" .And. LEFT(LINHA,23)#"CONTA DE INVESTIMENTO :" .And. lastrec()>=18
              GO 18
              NA=18
            EndIf
            If LEFT(LINHA,16)="CONTA CORRENTE :" .OR. LEFT(LINHA,23)="CONTA DE INVESTIMENTO :"
              NRELAT="012-Extrato de Contas Correntes PG:"+NARQ1
              CLIEN=VAL(SUBSTR(LINHA,25,AT("-",LINHA)-1))
            Else
              CLIEN=0
            EndIf
            If CLIEN#0
              GO 11
              If NA=22
                GO 12
              EndIf
              If NA=17
                GO 8
              EndIf
              If NA=18
                GO 9
              EndIf
              WWNOME=SUBSTR(LINHA,9,60)
              SKIP 2
              WWENDE1=LEFT(LINHA,70)
              SKIP
              WWENDE2=LEFT(LINHA,70)
              If SUBSTR(WWENDE2,7,1)=" "
                WWENDE2=LEFT(WWENDE2,6)+SUBSTR(WWENDE2,8,60)
              EndIf
              If SUBSTR(WWENDE2,9,1)=" "
                WWENDE2=LEFT(WWENDE2,8)+SUBSTR(WWENDE2,10,60)
              EndIf
              If EMPTY(WWNOME) .OR. EMPTY(WWENDE1) .OR. EMPTY(WWENDE2)
                WWNOME=""
                WWENDE1=""
                WWENDE2=""
              Else
                WNOMEA=WWNOME
                WEND1A=WWENDE1
                WEND2A=WWENDE2
                CLIANTES=CLIEN
              EndIf
            EndIf
            If CLIEN=0
              NRELAT="012-Extrato de Contas Correntes PG:"+NARQ1
              CLIEN=CLIANTES
              WWNOME=WNOMEA
              WWENDE1=WEND1A
              WWENDE2=WEND2A
            EndIf
            FEZ=.T.
          EndIf
        EndIf
        SELE 33
        If left(nrelat,3)#"012" .And. left(nrelat,3)#"016" .And. LEFT(NRELAT,3)#"003" .And. LEFT(NRELAT,3)#"022"
          cliantes=0
        EndIf
        wnre=val(left(nrelat,3))
        If CLIEN#0 .And. FEZ
          If WNRE=4 .And. WEBCOR<>"ATIVA" .And. WEBCOR<>"PLANNER" .And. WEBCOR<>"GERACAO" .And. webcor<>"CONVEC"
            CLIEN=VAL(LEFT(STRZERO(CLIEN,11),10))
          EndIf

          //DECLARE TX1[1]
          //TX1[1]=strzero(clien,9)+"  "+strzero(mmi,6)+" "+LEFT(NRELAT,30)+" "+WWNOME
          //LOGFILE("C:\exatus\teste.txt",TX1)

          WWENDE1=Strtran(WWENDE1,"   "," ")
          WWENDE1=Strtran(WWENDE1,"  "," ")
          WWENDE2=Strtran(WWENDE2,"   "," ")
          WWENDE2=Strtran(WWENDE2,"  "," ")
          If AT("(0",WWENDE1)#0
            WWENDE1=LEFT(WWENDE1,AT("(0",WWENDE1)-1)
          EndIf
          If SUBSTR(WWENDE2,9,1)#" "
            WWENDE2=LEFT(WWENDE2,9)+" "+SUBSTR(WWENDE2,10,80)
          EndIf
          CODCLI=WEBLOCAL1+STRZERO(CLIEN,8)+".PDF"
          CODCLI1=WEBLOCAL1+STRZERO(CLIEN,8)+".PD"
          If wnre=22
            NWBP=WEBLOCAL1+STRZERO(CLIEN,8)+STRZERO(0,4)+".PDF"
            If !FILE(NWBP)
              LZCOPYFILE(NARQ1,NWBP)
            Else
              FOR BWP=1 TO 9999
                NWBP=WEBLOCAL1+STRZERO(CLIEN,8)+STRZERO(BWP,4)+".PDF"
                If !FILE(NWBP)
                  EXIT
                EndIf
              NEXT WBP
              LZCOPYFILE(NARQ1,NWBP)
            EndIf
            Erase &NARQ1
            Erase &NARQ2
            LOOP
          EndIf
          SELE 35
          SET ORDER TO 1
          SEEK CLIEN
          If (wnre=39 .OR. wnre=1 .OR. wnre=2 .or. wnre=19 .or. wnre=4 .OR. wnre=12 .OR. wnre=13 .OR. wnre=14 .or. wnre=15 .or. wnre=16 .or. wnre=18 .OR. WNRE=20 .OR. WNRE=21 .OR. WNRE=24 .OR. WNRE=25 .OR. (WNRE=28 .And. WEBCOR<>"GERACAO") .OR. WNRE=29 .OR. WNRE=30 .OR. WNRE=34 .OR. WNRE=43) .And. !EMPTY(WWNOME)
            If !FOUND()
              ADIREG(0)
              ADICAO=.T.
            Else
              REGLOCK(10)
              ADICAO=.F.
            EndIf
            REPLACE CODIGO WITH CLIEN
            If EMPTY(NOME)
              REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERINC WITH "Arqv.Relatórios"
              REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
            Else
              REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERALT WITH "Arqv.Relatórios"
            EndIf
            If WNRE<>30
              REPLACE NOME WITH WWNOME
            ElseIf WNRE=30 .And. EMPTY(NOME)
              REPLACE NOME WITH WWNOME
            EndIf
            REPLACE ENDE1 WITH WWENDE1
            REPLACE ENDE2 WITH WWENDE2
            If ADICAO
              REPLACE EIMPRE  WITH .T.
              REPLACE EEMAIL  WITH .T.
              REPLACE EGUARD  WITH .T.
              REPLACE MAILING WITH .T.
            EndIf
          EndIf

          SEEK CLIEN
          If !FOUND()
            ADIREG(0)
            REPLACE CODIGO  WITH CLIEN
            REPLACE NOME    WITH WWNOME
            REPLACE EIMPRE  WITH .T.
            REPLACE EEMAIL  WITH .T.
            REPLACE EGUARD  WITH .T.
            REPLACE MAILING WITH .T.
          EndIf
          If (wnre=4 .OR. wnre=1 .OR. WNRE=12 .OR. WNRE=24) .And. FOUND() .And. REGLOCK(10)
            If wnre=4
              REPLACE ULTNOTBMF  WITH DATE()
            ElseIf wnre=12
              REPLACE ULTCONTA   WITH DATE()
            Else
              REPLACE ULTNOTABOV WITH DATE()
            EndIf
          ElseIf FOUND() .And. REGLOCK(10)
            REPLACE ULTCONTA WITH DATE()
          EndIf
          SEEK CLIEN
          If NOME="BB BANCO" .And. reglock(10)
            REPLACE FRENTEV WITH .T.
          EndIf
          If !FOUND() .And. !WFAZDADOS
            Erase &NARQ1
            Erase &NARQ2
            LOOP
          EndIf
          SELE 35
          SEEK CLIEN
          SELE 37
          SEEK CLIEN
          If !FOUND()
            ADIREG(0)
            REPLACE CODIGO WITH CLIEN
            REPLACE NOME WITH CLIENTES->NOME
            REPLACE OQUE WITH NRELAT
            REPLACE DATAENV WITH DATE()
            REPLACE HORAENV WITH TIME()
            REPLACE ENDERECO WITH CLIENTES->ENDERECO
            REPLACE NUMERO WITH CLIENTES->NUMERO
            REPLACE COMPL WITH CLIENTES->COMPL
            REPLACE ASSESSOR WITH CLIENTES->ASSESSOR
            REPLACE BAIRRO WITH CLIENTES->BAIRRO
            REPLACE AUDIT WITH CLIENTES->AUDIT
            REPLACE CIDADE WITH CLIENTES->CIDADE
            REPLACE UF WITH CLIENTES->UF
            REPLACE CEP WITH CLIENTES->CEP
            REPLACE ENDE1 WITH CLIENTES->ENDE1
            REPLACE ENDE2 WITH CLIENTES->ENDE2
            REPLACE FRENTEV WITH CLIENTES->FRENTEV
            If EMPTY(ENDE1)
              REPLACE ENDE1 WITH ALLTRIM(ENDERECO)+", "+ALLTRIM(NUMERO)+"  "+ALLTRIM(COMPL)
              REPLACE ENDE2 WITH LEFT(CEP,5)+"-"+SUBSTR(CEP,6,3)+" "+ALLTRIM(CIDADE)+"-"+UF
            EndIf
            REPLACE WCEP WITH STRZERO(VAL(Strtran(Strtran(LEFT(ENDE2,9),"-","")," ","")),8)
            REPLACE FOLHAS WITH 1
          Else
            REPLACE FOLHAS WITH FOLHAS+1
          EndIf
          NWBP=WEBLOCAL1+STRZERO(CLIEN,8)+STRZERO(0,4)+".PDF"
          If !FILE(NWBP)
            LZCOPYFILE(NARQ1,NWBP)
          Else
            FOR BWP=1 TO 9999
              NWBP=WEBLOCAL1+STRZERO(CLIEN,8)+STRZERO(BWP,4)+".PDF"
              If !FILE(NWBP)
                EXIT
              EndIf
            NEXT WBP
            LZCOPYFILE(NARQ1,NWBP)
          EndIf
        Else   /// CLIEN=0 "CLIENTE NÃO ENCONTRADO"
          NWBP=WEBLOCAL1+STRZERO(0,12)+".PDF"
          If !FILE(NWBP)
            LZCOPYFILE(NARQ1,NWBP)
          Else
            FOR WBP=1 TO 9999
              NWBP=WEBLOCAL1+STRZERO(0,8)+STRZERO(WBP,4)+".PDF"
              If !FILE(NWBP)
                EXIT
              EndIf
            NEXT WBP
            LZCOPYFILE(NARQ1,NWBP)
          EndIf
          SELE 37
          SEEK 0
          If !FOUND()
            ADIREG(0)
            REPLACE NOME   WITH "NÃO ENCONTRADO NO CADASTRO DE CLIENTES"
            REPLACE ENDE1  WITH "NÃO ENCONTRADO NO CADASTRO DE CLIENTES"
            REPLACE ENDE2  WITH "NÃO ENCONTRADO NO CADASTRO DE CLIENTES"
            REPLACE WCEP WITH "00000000"
            REPLACE FOLHAS WITH 1
            REPLACE FAZER  WITH .T.
          Else
            REPLACE FOLHAS WITH FOLHAS+1
            REPLACE FAZER WITH .T.
          EndIf
        EndIf
        Erase &NARQ1
        Erase &NARQ2
      NEXT MMI
      SELE 33
      USE
      Erase &NARQ3
      If TEMCETIP
        DECLARE VLABEL[ADIR(WEBLOCAL1+"??????????????_CERTIFICACAO*.PDF")]
        ADIR(WEBLOCAL1+"??????????????_CERTIFICACAO*.PDF",VLABEL)
        FOR JI=1 TO LEN(VLABEL)
          CLIEN=VAL(SUBSTR(VLABEL[JI],28,14))
          MZLOCAL=ALLTRIM(WEBLOCAL1)+ALLTRIM(VLABEL[JI])
          SELE 35
          SET ORDER TO 2
          SEEK STRZERO(CLIEN,14)
          If FOUND() .And. CLIEN=VAL(CNPJCPF)
            CLIEN=CODIGO
            SET ORDER TO 1
          Else
            DECLARE TX1[1]
            TX1[1]="Cliente não cadastrado no BolsasNet "+STRZERO(CLIEN,14)
            If !EMPTY(WDADOS)
              LOGFILE(WDADOS,TX1)
            EndIf
            CLIEN=0
            Erase &MZLOCAL
            SET ORDER TO 1
            LOOP
          EndIf
          NOMEENV=NOME
          CODCLI=CLIEN
          FZAUDIT=AUDIT
          NRELAT="Posição Cetip "
          SELE 31
          SEEK CODCLI
          ENVIADOPARA=""
          DO WHILE !EOF() .And. CLIENTE=CODCLI
            SysRefresh()
            If CLIENTE#CODCLI
              LOOP
            EndIf
            MREG=RECNO()
            WCORPOENVIO=Strtran(WEBCORPO,";;NOME;:",TRIM(NOME))
            WCORPOENVIO=Strtran(WCORPOENVIO,";;CODIGO;:",STRZERO(CODCLI,8))
            MBOLET=""
            If LEFT(EMAIL,1)#"[" .And. LEFT(EMAIL,1)#"{"
              If AT("@",EMAIL)=0
                SKIP
                LOOP
              EndIf
              QEMAIL=EMAIL
              If FZAUDIT .And. !EMPTY(WEBAUD)
                QEMAIL=ALLTRIM(EMAIL)+";"+ALLTRIM(WEBAUD)
              EndIf
              ENVIADOPARA=QEMAIL
              If SEENVIAEMAIL="S" .And. !EMPTY(EMAIL)
                If EMPTY(WEBASSUNTO)
                  WNASSUNTO="Relatórios Cliente: "+STRZERO(CODCLI,8)
                Else
                  WNASSUNTO=ALLTRIM(WEBASSUNTO)+" Cliente: "+STRZERO(CODCLI,8)
                EndIf
                ENVIAEMAIL(QEMAIL,WNASSUNTO,WCORPOENVIO,MZLOCAL,MBOLET,WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA,NOMEENV,WEBNOMECOR)
              EndIf
            ElseIf LEFT(EMAIL,1)="["
              MODIGO=VAL(Strtran(EMAIL,"[",""))
              SELE 32
              LOCATE FOR MODIGO=GRUPO
              DO WHILE .NOT. EOF()
                If MODIGO<>GRUPO
                  CONT
                  LOOP
                EndIf
                SysRefresh()
                ENVIADOPARA=EMAIL
                If AT("@",EMAIL)=0
                  SELE 32
                  CONT
                  LOOP
                EndIf
                If SEENVIAEMAIL="S" .And. !EMPTY(EMAIL)
                  If EMPTY(WEBASSUNTO)
                    WNASSUNTO="Relatórios Cliente: "+STRZERO(CODCLI,8)
                  Else
                    WNASSUNTO=ALLTRIM(WEBASSUNTO)+" Cliente: "+STRZERO(CODCLI,8)
                  EndIf
                  ENVIAEMAIL(EMAIL,WNASSUNTO,WCORPOENVIO,MZLOCAL,MBOLET,WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA,NOMEENV,WEBNOMECOR)
                EndIf
                SELE 32
                CONT
              ENDDO
            EndIf
            SELE 37
            SEEK CODCLI
            If EOF()
              ADIREG(0)
              REPLACE CODIGO WITH CODCLI
              REPLACE NOME WITH CLIENTES->NOME
              REPLACE OQUE WITH NRELAT
              REPLACE DATAENV WITH DATE()
              REPLACE HORAENV WITH TIME()
              REPLACE ENDERECO WITH CLIENTES->ENDERECO
              REPLACE NUMERO WITH CLIENTES->NUMERO
              REPLACE COMPL WITH CLIENTES->COMPL
              REPLACE ASSESSOR WITH CLIENTES->ASSESSOR
              REPLACE BAIRRO WITH CLIENTES->BAIRRO
              REPLACE AUDIT WITH CLIENTES->AUDIT
              REPLACE CIDADE WITH CLIENTES->CIDADE
              REPLACE UF WITH CLIENTES->UF
              REPLACE CEP WITH CLIENTES->CEP
              REPLACE ENDE1 WITH CLIENTES->ENDE1
              REPLACE ENDE2 WITH CLIENTES->ENDE2
              REPLACE FRENTEV WITH CLIENTES->FRENTEV
              If EMPTY(ENDE1)
                REPLACE ENDE1 WITH ALLTRIM(ENDERECO)+", "+ALLTRIM(NUMERO)+"  "+ALLTRIM(COMPL)
                REPLACE ENDE2 WITH LEFT(CEP,5)+"-"+SUBSTR(CEP,6,3)+" "+ALLTRIM(CIDADE)+"-"+UF
              EndIf
              REPLACE WCEP WITH STRZERO(VAL(Strtran(Strtran(LEFT(ENDE2,9),"-","")," ","")),8)
              REPLACE FOLHAS WITH 1
            Else
              REPLACE FOLHAS WITH FOLHAS+1
            EndIf
            REPLACE ENVIADO WITH .T.
            If !EMPTY(ALLTRIM(EMAILC))
              REPLACE EMAILC WITH alltrim(emailc)+";"+alltrim(ENVIADOPARA)
            Else
              REPLACE EMAILC WITH ENVIADOPARA
            EndIf
            REPLACE LOGMAIL WITH TX2[1]
            REPLACE BAIXA WITH .F.
            SELE 31
            SKIP
          ENDDO
          Erase &MZLOCAL
        NEXT JI
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL1+"????????0000.PDF")]
      ADIR(WEBLOCAL1+"????????0000.PDF",VLABEL)
      FOR JI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        CLIEN=VAL(LEFT(VLABEL[JI],8))
        If !RELINTERNO
          SELE 37
          SEEK CLIEN
          WWCEP=WCEP
        Else
          WCEP=""
        EndIf
        NWBP=WEBLOCAL1+WCEP+LEFT(VLABEL[JI],8)+".PDF"
        CODCLI=WEBLOCAL1+LEFT(VLABEL[JI],8)+"0001.PDF"
        CODCLI1=WEBLOCAL1+LEFT(VLABEL[JI],8)+"0000.PDF"
        If !FILE(CODCLI)
          FRENAME(CODCLI1,NWBP)
        Else
          Erase &nwbp
          comando=WCORRENTE+"\PDFTK "+WEBLOCAL1+LEFT(VLABEL[JI],8)+"*.PDF cat output "+NWBP+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 04"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &CODCLI1
          FOR BWP=1 TO 9999
            CODCLI1=WEBLOCAL1+LEFT(VLABEL[JI],8)+STRZERO(BWP,4)+".PDF"
            If !FILE(CODCLI1)
              EXIT
            EndIf
            Erase &CODCLI1
          NEXT WBP
        EndIf
      NEXT JI
      If WEBCOR="GERACAO" .And. SOPROC
        WWEBLOCAL1=WWEBLOCA9+"DROPBOX\GERACAO\"
        VERDIRETORIO(WWEBLOCAL1)
        comando=WCORRENTE+"\BRAZIP.EXE AH "+WWEBLOCAL1+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+"PORCLIENTE.ZIP "+WEBLOCAL1+"*.PDF"+CHR(13)+CHR(10)
        MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
        MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
        MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina PORCLIENTE "+MZLINHA)
        CURSORWAIT()
        WAITRUN(weblocal1+"FZ.BAT",0)
        CURSORWAIT()
        DO WHILE FILE(weblocal1+"AGUARDE.TXT")
          SysRefresh()
        ENDDO
        MTEZ=WEBLOCAL1+"FZ.BAT"
        Erase &MTEZ
        DECLARE VLABEL[ADIR(WEBLOCAL1+"*.PDF")]
        ADIR(WEBLOCAL1+"*.PDF",vlabel)
        FOR JI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          MLOCAL=WEBLOCAL1+VLABEL[JI]
          Erase &MLOCAL
        NEXT JI
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL1+"*.PDF")]
      ADIR(WEBLOCAL1+"*.PDF",vlabel)
      ZARQ=WEBLOCAL+"FRENTE.fig"
      TEMFRENTEVERSO=.F.
      FOR JI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        MLOCAL=WEBLOCAL1+VLABEL[JI]
        MZLOCAL=WEBLOCAL1+SUBSTR(VLABEL[JI],9,8)+RIGHT(VLABEL[JI],4)
        If !RELINTERNO
          CODCLI=VAL(SUBSTR(VLABEL[JI],9,8))
        Else
          CODCLI=VAL(LEFT(VLABEL[JI],8))
          If CODCLI<>0
            SELE 39
            SEEK CODCLI
            If FOUND()
              ENVIADOPARA=lower(alltrim(EMAIL))
              WCORPOENVIO=Strtran(WEBCORPO,";;NOME;:",TRIM(NOME))
              WCORPOENVIO=Strtran(WCORPOENVIO,";;CODIGO;:",STRZERO(CODCLI,8))
              NOMEENV=NOME
              MBOLET=""
              If SEENVIAEMAIL="S" .And. !EMPTY(EMAIL)
                ENVIAEMAIL(EMAIL,"Relatórios Assessor: "+NOME,WCORPOENVIO,MLOCAL,MBOLET,WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA,NOMEENV,WEBCOR)
              EndIf
            EndIf
          EndIf
          Erase &MLOCAL
          LOOP
        EndIf
        If CODCLI=0
          MLOCAL1=WEBLOCAL10+VLABEL[JI]
          LZCOPYFILE(MLOCAL,MLOCAL1)
          Erase &MLOCAL
          LOOP
        EndIf
        SELE 35
        SET ORDER TO 1
        SEEK CODCLI
        If !FOUND()
          If FILE(ZARQ)
            SELE 37
            SEEK CODCLI
            If !FOUND()
              ADIREG(0)
              REPLACE CODIGO WITH CODCLI
            EndIf
            If !CLIENTES->EIMPRE .And. WEBACEITA
              REPLACE BAIXA WITH .T.
            Else
              If FOLHAS<2
                MLOCAL1=WEBLOCAL6+VLABEL[JI]
                LZCOPYFILE(MLOCAL,MLOCAL1)
                TEMFRENTEVERSO=.T.
              Else
                MLOCAL1=WEBLOCAL4+VLABEL[JI]
                LZCOPYFILE(MLOCAL,MLOCAL1)
              EndIf
            EndIf
          Else
            MLOCAL1=WEBLOCAL4+VLABEL[JI]
            LZCOPYFILE(MLOCAL,MLOCAL1)
          EndIf
          Erase &MLOCAL
          LOOP
        EndIf
        WEIMPRE=EIMPRE
        SELE 35
        SEEK CODCLI
        wfrentev=.F.
        If FOUND()
          WFRENTEV=FRENTEV
        Else
          WFRENTEV=.F.
        EndIf
        SELE 31
        SEEK CODCLI
        If (WEIMPRE .OR. !FOUND())
          If FILE(ZARQ)
            SELE 37
            SEEK CODCLI
            If !FOUND()
              ADIREG(0)
              REPLACE CODIGO WITH CODCLI
            EndIf
            If !CLIENTES->EIMPRE .And. WEBACEITA
              REPLACE BAIXA WITH .T.
            Else
              If FOLHAS<2 .And. !WFRENTEV
                MLOCAL1=WEBLOCAL6+VLABEL[JI]
                LZCOPYFILE(MLOCAL,MLOCAL1)
                TEMFRENTEVERSO=.T.
              Else
                MLOCAL1=WEBLOCAL4+VLABEL[JI]
                LZCOPYFILE(MLOCAL,MLOCAL1)
              EndIf
            EndIf
          Else
            If !CLIENTES->EIMPRE .And. WEBACEITA
              SELE 37
              SEEK CODCLI
              If !FOUND()
                ADIREG(0)
                REPLACE CODIGO WITH CODCLI
              EndIf
              REPLACE BAIXA WITH .T.
            Else
              MLOCAL1=WEBLOCAL4+VLABEL[JI]
              LZCOPYFILE(MLOCAL,MLOCAL1)
            EndIf
          EndIf
        EndIf
        SELE 35
        If EGUARD .And. WFAZDADOS
          SELE 36
          ADIREG(0)
          REPLACE CORRETORA WITH 0
          REPLACE CLIENTE WITH CODCLI
          REPLACE ORDEMGER WITH RECNO()
          REPLACE GRUPOD WITH 7
          REPLACE TIPODOC WITH 28
          REPLACE OBS WITH "DOCUMENTO AUTOMATICO DO BOLSASNET"
          REPLACE ARQUIVO WITH RECNO()
          REPLACE ATIVO WITH .T.
          REPLACE ESCANEM WITH DATE()
          REPLACE STATUS WITH "X"
          REPLACE POR WITH "BOLSASNET"
          WEBLOCAL5=WEBLOCAL3+STRZERO(CODCLI,8)+"\"
          VERDIRETORIO(WEBLOCAL5)
          Warq=WEBLOCAL5+STRZERO(INT(ARQUIVO/1000),8)+"\"
          VERDIRETORIO(WARQ)
          WARQ=WARQ+STRZERO((ARQUIVO-INT(ARQUIVO/1000)),4)+".PDF"
          LZCOPYFILE(MLOCAL,WARQ)
          SELE 35
        EndIf
        SELE 35
        FZAUDIT=AUDIT
        If !EEMAIL
          Erase &MLOCAL
          LOOP
        EndIf
        TEMPSENHA=SENHA
        NOMEENV=NOME
        SELE 31
        SEEK CODCLI
        If FOUND()
          MZARQ=WEBLOCAL1+"SENHA.PDF"
          MZCONT=0
          DO WHILE FILE(MZARQ)
            MZCONT=MZCONT+1
            Erase &MZARQ
            If MZCONT>100
              MSGINFO("Não consigo apagar o arquivo SENHA.PDF",wsist)
            EndIf
          ENDDO
          FRENAME(MLOCAL,MZARQ)
          comando=WCORRENTE+"\gswin32c -sDEVICE=pdfwrite -dCompatibilityLevel=1.4 -dPDFSETTINGS=/prepress -dNOPAUSE -dQUIET -dBATCH -sOutputFile="+MLOCAL+" "+MZARQ+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 05"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &MZARQ
        EndIf
        If WEBPDFS .And. FOUND()
          MZARQ=WEBLOCAL1+"SENHA.PDF"
          MZCONT=0
          DO WHILE FILE(MZARQ)
            MZCONT=MZCONT+1
            Erase &MZARQ
            If MZCONT>100
              MSGINFO("Não consigo apagar o arquivo SENHA.PDF",wsist)
            EndIf
          ENDDO
          FRENAME(MLOCAL,MZARQ)
          comando=WCORRENTE+"\PDFTK "+MZARQ+" OUTPUT "+MLOCAL+" allow AllFeatures user_pw "+alltrim(tempsenha)+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+" DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 05"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &MZARQ
        EndIf
        ENVIADOPARA=""
        LZCOPYFILE(MLOCAL,MZLOCAL)
        DO WHILE !EOF() .And. CLIENTE=CODCLI
          SysRefresh()
          If CLIENTE#CODCLI
            LOOP
          EndIf
          MREG=RECNO()
          WCORPOENVIO=Strtran(WEBCORPO,";;NOME;:",TRIM(NOME))
          WCORPOENVIO=Strtran(WCORPOENVIO,";;CODIGO;:",STRZERO(CODCLI,8))
          MBOLET=ALLTRIM(WEBLOCAL)+"BOLETIM.PDF"
          If LEFT(EMAIL,1)#"[" .And. LEFT(EMAIL,1)#"{"
            If AT("@",EMAIL)=0
              SKIP
              LOOP
            EndIf
            QEMAIL=EMAIL
            If FZAUDIT .And. !EMPTY(WEBAUD)
              QEMAIL=ALLTRIM(EMAIL)+";"+ALLTRIM(WEBAUD)
            EndIf
            ENVIADOPARA=QEMAIL
            If SEENVIAEMAIL="S" .And. !EMPTY(EMAIL)
              If EMPTY(WEBASSUNTO)
                WNASSUNTO="Relatórios Cliente: "+STRZERO(CODCLI,8)
              Else
                WNASSUNTO=ALLTRIM(WEBASSUNTO)+" Cliente: "+STRZERO(CODCLI,8)
              EndIf
              ENVIAEMAIL(QEMAIL,WNASSUNTO,WCORPOENVIO,MZLOCAL,MBOLET,WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA,NOMEENV,WEBNOMECOR)
            EndIf
          ElseIf LEFT(EMAIL,1)="["
            MODIGO=VAL(Strtran(EMAIL,"[",""))
            SELE 32
            LOCATE FOR MODIGO=GRUPO
            DO WHILE .NOT. EOF()
              If MODIGO<>GRUPO
                CONT
                LOOP
              EndIf
              SysRefresh()
              ENVIADOPARA=EMAIL
              If AT("@",EMAIL)=0
                SELE 32
                CONT
                LOOP
              EndIf
              If SEENVIAEMAIL="S" .And. !EMPTY(EMAIL)
                If EMPTY(WEBASSUNTO)
                  WNASSUNTO="Relatórios Cliente: "+STRZERO(CODCLI,8)
                Else
                  WNASSUNTO=ALLTRIM(WEBASSUNTO)+" Cliente: "+STRZERO(CODCLI,8)
                EndIf
                ENVIAEMAIL(EMAIL,WNASSUNTO,WCORPOENVIO,MZLOCAL,MBOLET,WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA,NOMEENV,WEBNOMECOR)
              EndIf
              SELE 32
              CONT
            ENDDO
          EndIf
          SELE 37
          SEEK CODCLI
          REPLACE ENVIADO WITH .T.
          If !EMPTY(ALLTRIM(EMAILC))
            REPLACE EMAILC WITH alltrim(emailc)+";"+alltrim(ENVIADOPARA)
          Else
            REPLACE EMAILC WITH ENVIADOPARA
          EndIf
          REPLACE LOGMAIL WITH TX2[1]
          If TX2[2]=.F. .And. SEENVIAEMAIL="S"
            SELE 37
            SEEK CODCLI
            If !FOUND()
              ADIREG(0)
              REPLACE CODIGO WITH CODCLI
            EndIf
            REPLACE BAIXA WITH .F.
            ZARQ=WEBLOCAL+"FRENTE.fig"
            If FOLHAS<2 .And. !WFRENTEV .And. FILE(ZARQ)
              MLOCAL1=WEBLOCAL6+VLABEL[JI]
              LZCOPYFILE(MLOCAL,MLOCAL1)
              TEMFRENTEVERSO=.T.
            Else
              MLOCAL1=WEBLOCAL4+VLABEL[JI]
              LZCOPYFILE(MLOCAL,MLOCAL1)
            EndIf
          EndIf
          SELE 31
          SKIP
        ENDDO
        Erase &MLOCAL
        Erase &MZLOCAL
      NEXT JI
      SELE 32
      USE
      SELE 36
      USE
      SELE 31
      USE
      Erase &NARQ4
      Erase &NARQ6
      Erase &NARQ5
      NARQ11=WEBLOCAL9+"IMPRIMIR.PDF"
      Erase &NARQ11
      If RELINTERNO
        SELE 37
        USE
        SELE 35
        USE
        SELE 39
        USE
        Erase &NARQ41
        SELE 30
        If SEENVIAEMAIL#"S" .or. cod_oper#999
          If EMPTY(WERROARQ)
            MSGINFO("Terminou !",wsist)
          Else
            MSGSTOP("Erro nos arquivos "+WERROARQ,"Não foram gerados no formato A4")
          EndIf
        EndIf
        return 
      Else
        SELE 39
        USE
        Erase &NARQ41
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL10+"*.PDF")]
      ADIR(WEBLOCAL10+"*.PDF",VLABEL)
      If LEN(VLABEL)#0
        If WFAZCORREIO
          comando=WCORRENTE+"\PDFTK "+WEBLOCAL10+"*.PDF cat output "+WEBLOCAL10+"IMPRIMIR.PDF"+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 06"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
        EndIf
        FOR JI=1 TO LEN(VLABEL)
          If VLABEL[JI]="IMPRIMIR.PDF"
            LOOP
          EndIf
          NARQ5=WEBLOCAL10+ALLTRIM(VLABEL[JI])
          Erase &NARQ5
        NEXT JI
        If WFAZCORREIO
          comando=WCORRENTE+"\BRAZIP.EXE AH "+WEBLOCAL10+"IMPRIME.ZIP "+WEBLOCAL10+"IMPRIMIR.PDF"+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 07"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          MTEZ=WEBLOCAL1+"FZ.BAT"
          Erase &MTEZ
          EMAILRETORNO=WEBIMP
          If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
            ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - impressão (Não Identificado) "+webcor,"Em anexo relatorios para serem impressos",NARQ11,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
          NARQ4=WEBLOCAL10+"IMPRIMIR.PDF"
          Erase &NARQ4
        EndIf
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL4+"*.PDF")]
      ADIR(WEBLOCAL4+"*.PDF",vlabel)
      If LEN(VLABEL)#0
        FOR JI=1 TO LEN(VLABEL)
          CODCLI=VAL(SUBSTR(VLABEL[JI],9,8))
          SELE 35
          SEEK CODCLI
          If FOUND() .And. FRENTEV
            SELE 37
            SEEK CODCLI
            If FOUND()
              REPLACE FAZER WITH .T.
            EndIf
            NARQ9=WEBLOCAL4+VLABEL[JI]
            NARQ10=WEBLOCAL9+VLABEL[JI]
            If WFAZCORREIO
              LZCOPYFILE(NARQ9,NARQ10)
              Erase &NARQ9
            EndIf
          EndIf
        NEXT JI
        DECLARE VLABEL[ADIR(WEBLOCAL9+"*.PDF")]
        ADIR(WEBLOCAL9+"*.PDF",vlabel)
        If LEN(VLABEL)#0
          If WFAZCORREIO
            comando=WCORRENTE+"\PDFTK "+WEBLOCAL9+"*.PDF cat output "+NARQ11+CHR(13)+CHR(10)
            MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
            MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
            MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 08"+MZLINHA)
            CURSORWAIT()
            WAITRUN(weblocal1+"FZ.BAT",0)
            CURSORWAIT()
            DO WHILE FILE(weblocal1+"AGUARDE.TXT")
              SysRefresh()
            ENDDO
          EndIf
          FOR JI=1 TO LEN(VLABEL)
            If VLABEL[JI]="IMPRIMIR.PDF"
              LOOP
            EndIf
            NARQ5=WEBLOCAL9+ALLTRIM(VLABEL[JI])
            Erase &NARQ5
          NEXT JI
          If WFAZCORREIO
            comando=WCORRENTE+"\BRAZIP.EXE AH "+WEBLOCAL9+"IMPRIME.ZIP "+WEBLOCAL9+"IMPRIMIR.PDF"+CHR(13)+CHR(10)
            MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
            MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
            MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 09"+MZLINHA)
            CURSORWAIT()
            WAITRUN(weblocal1+"FZ.BAT",0)
            CURSORWAIT()
            DO WHILE FILE(weblocal1+"AGUARDE.TXT")
              SysRefresh()
            ENDDO
          EndIf
          MTEZ=WEBLOCAL1+"FZ.BAT"
          Erase &MTEZ
          Erase &NARQ11
          NARQ11=WEBLOCAL9+"IMPRIME.ZIP"
          EMAILRETORNO=WEBIMP
          If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO) .And. WFAZCORREIO
            ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - impressão (Separados) "+webcor,"Em anexo relatorios para serem impressos",NARQ11,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
      EndIf
      NARQ4=WEBLOCAL4+"IMPRIMIR.PDF"
      Erase &NARQ4
      DECLARE VLABEL[ADIR(WEBLOCAL4+"*.PDF")]
      ADIR(WEBLOCAL4+"*.PDF",vlabel)
      QLB1=1
      If WFAZCORREIO
        FOR JI=1 TO LEN(VLABEL)
          NARQ9=WEBLOCAL4+VLABEL[JI]
          NARQ10=WEBLOCAL4+"I"+VLABEL[JI]
          QLB1=QLB1+1
          If QLB1=4
            QLB1=1
          EndIf
          If QLB1=1
            LOOP
          ElseIf QLB1=2
            QLB=WCORRENTE+"\QUADRADO.PDF"
          ElseIf QLB1=3
            QLB=WCORRENTE+"\LINHA.PDF"
          EndIf
          comando=WCORRENTE+"\PDFTK "+NARQ9+" background "+qlb+" output "+NARQ10+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 10"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &NARQ9
          FRENAME(NARQ10,NARQ9)
        NEXT JI
      EndIf
      If LEN(VLABEL)#0
        If WFAZCORREIO
          comando=WCORRENTE+"\PDFTK "+WEBLOCAL4+"*.PDF cat output "+NARQ4+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 11"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
        EndIf             **
        FOR JI=1 TO LEN(VLABEL)
          If VLABEL[JI]="IMPRIMIR.PDF"
            LOOP
          EndIf
          NARQ5=WEBLOCAL4+ALLTRIM(VLABEL[JI])
          WWCLI=VAL(SUBSTR(VLABEL[JI],9,8))
          SELE 37
          SEEK WWCLI
          If FOUND()
            REPLACE FAZER WITH .T.
          EndIf
          Erase &NARQ5
        NEXT JI
        If WFAZCORREIO
          comando=WCORRENTE+"\BRAZIP.EXE AH "+WEBLOCAL4+"IMPRIME.ZIP "+WEBLOCAL4+"IMPRIMIR.PDF"+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 12"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          Erase &NARQ4
          MTEZ=WEBLOCAL1+"FZ.BAT"
          Erase &MTEZ
          NARQ4=WEBLOCAL4+"IMPRIME.ZIP"
          EMAILRETORNO=WEBIMP
          If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
            ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - impressão(Sem Frente) "+WEBCOR,"Em anexo relatorios para serem impressos",NARQ4,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
      EndIf
      SELE 37
      GO TOP
      If WFAZCORREIO
        NARQ8=WEBLOCAL4+"ETIQ.DBF"
        Erase &NARQ8
        COPY ALL TO &NARQ8 FOR FOLHAS#0 .And. FOLHAS<=10 .And. CODIGO#0 .And. FAZER=.T. .And. !FRENTEV
        NARQ9=WEBLOCAL4+"ETIQ930.DBF"
        COPY ALL TO &NARQ9 FOR FOLHAS>10 .And. FOLHAS<=30 .And. CODIGO#0 .And. FAZER=.T. .And. !FRENTEV
        NARQ13=WEBLOCAL4+"ETIQM30.DBF"
        COPY ALL TO &NARQ13 FOR FOLHAS>30 .And. CODIGO#0 .And. FAZER=.T. .And. !FRENTEV
        SELE 37
        GO TOP
        FOR NDI=1 TO 3
          SELE 38
          If NDI=1
            If !USEREDE(NARQ8,.T.,10,"ETIQ1")
              LOOP
            Else
              NARQ71=WEBLOCAL4+"ETIQ1.KEY"
              INDEX ON WCEP+STRZERO(CODIGO,8) TO &NARQ71
            EndIf
          EndIf
          If NDI=2
            If !USEREDE(NARQ9,.T.,10,"ETIQ1")
              LOOP
            Else
              NARQ71=WEBLOCAL4+"ETIQ1.KEY"
              INDEX ON WCEP+STRZERO(CODIGO,8) TO &NARQ71
            EndIf
          EndIf
          If NDI=3
            If !USEREDE(NARQ13,.T.,10,"ETIQ1")
              LOOP
            Else
              NARQ71=WEBLOCAL4+"ETIQ1.KEY"
              INDEX ON WCEP+STRZERO(CODIGO,8) TO &NARQ71
            EndIf
          EndIf
          oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
          If NDI=1
            ZARQ1=WEBLOCAL4+"ETIQAT10.PDF"
          ElseIf NDI=2
            ZARQ1=WEBLOCAL4+"ETIQ1130.PDF"
          ElseIf NDI=3
            ZARQ1=WEBLOCAL4+"ETIQM30.PDF"
          EndIf
          Erase &ZARQ1
          oReg:Set("PDFFilename",ZARQ1)
          oReg:Close()
          PRINT oPrn NAME "Envelope" TO TEMPDF1
          DEFINE FONT oFont1 NAME "verdana" SIZE 0,11 OF oPrn
          DEFINE FONT oFont2 NAME "verdana" SIZE 0,12 BOLD OF oPrn
          DEFINE FONT oFont3 NAME "verdana" SIZE 0,12 OF oPrn
          DEFINE FONT oFont4 NAME "verdana" SIZE 0,30 BOLD ITALIC OF oPrn
          DEFINE FONT oFont5 NAME "verdana" SIZE 0,12 BOLD ITALIC OF oPrn
          DEFINE font oFont7 NAME "verdana" SIZE 0,09 ITALIC BOLD OF oPrn
          DEFINE FONT oFont6 NAME "verdana" SIZE 0,08 ITALIC OF oPrn
          DEFINE FONT oFont3 NAME "verdana" SIZE 0,12 ITALIC OF oPrn
          nRowStep = oPrn:nVertRes() / 66
          nColStep = oPrn:nHorzRes() / 80
          oPrn:SetLandscape()
          GO TOP
          DO WHILE !EOF()
            MCODIGO=STRZERO(CODIGO,9)
            MCODIGO1=STRZERO((CODIGO+3)*2,9)
            PAGE
            If webcor="PLANNER"
              saybitmap(4.5,8.0,3.26*6.5,1.3*5.0,WEBLOCAL+"LOGO.fig",oPrn)
              oPrn:cmsay(6.0,8.0,"Av.Brig.Faria Lima, 3900 - 10.andar",oFont6,,nRGB(0,0,0),,0)
              oPrn:cmsay(6.3,8.0,"04538-132 - Itaim Bibi - São Paulo - SP",oFont6,,nRGB(0,0,0),,0)
              oPrn:cmsay(5.2,14.0,"CENTRAL PLANNER",oFont2,,nRGB(0,0,0),,0)
              oPrn:cmsay(5.6,14.0,"DE ATENDIMENTO    OUVIDORIA",oFont2,,nRGB(0,0,0),,0)
              oPrn:cmsay(6.0,14.0,"0800 179444           0800 7722231",oFont2,,nRGB(0,0,0),,0)
            EndIf
            If WEBCOR="CONVEC"
              saybitmap(4.5,8.0,3.26*6.5,1.3*5.0,WEBLOCAL+"LOGO.fig",oPrn)
              oPrn:cmsay(6.0,8.0,"Rua Amauri, 255  8.andar",oFont6,,nRGB(0,0,0),,0)
              oPrn:cmsay(6.3,8.0,"01448-000  São Paulo - SP",oFont6,,nRGB(0,0,0),,0)
            EndIf
            CALCCEP(ENDE2,10.1,8.0)
            If FILE("FAC.BMP")
              saybitmap(14.8,17.00,2.472,2.472,"FAC.BMP",oPrn)
            EndIf
            oprn:cmsay(11.0,8.0,LEFT(NOME,50),oFont2,,nRGB(0,0,0),,0)
            oprn:cmsay(11.5,8.0,ENDE1,oFont1,,nRGB(0,0,0),,0)
            oprn:cmsay(12.0,8.0,ENDE2,ofont1,,nRGB(0,0,0),,0)

            oPrn:cmsay(14.5,8.0,STRZERO(CODIGO,9),oFont3,,nRGB(0,0,0),,0)
            @ nRowStep*30.5,nColStep*30 CODE128 strzero(codigo,9) OF oPrn SIZE 0.8 MODE "A"
            oprn:cmsay(14.5,10.0,STR(FOLHAS,6,0),oFont1,,nRGB(0,0,0),,0)
            ENDPAGE
            SKIP
          ENDDO
          oPrn:SetPortrait()
          ENDPRINT
          If LASTREC()>0
            DO WHILE .T.
              CURSORWAIT()
              If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
                fClose(nHandle)
                EXIT
              Else
                fClose(nHandle)
              EndIf
            ENDDO
          EndIf
          SELE 38
          USE
          If NDI=1
            Erase &NARQ8
          EndIf
          If NDI=2
            Erase &NARQ9
          EndIf
          If NDI=3
            Erase &NARQ13
          EndIf
          NARQ71=WEBLOCAL4+"ETIQ1.KEY"
          Erase &NARQ71
        NEXT NDI
        SELE 37
        EMAILRETORNO=WEBIMP
        If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
          wnarq8=WEBLOCAL4+"ETIQAT10.PDF"
          WNARQ9=WEBLOCAL4+"ETIQ1130.PDF"
          WNARQ13=WEBLOCAL4+"ETIQM30.PDF"
          ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Envelopes "+WEBCOR,"Em anexo relatorios para serem impressos",WNARQ8,WNARQ9,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          If FILE(WNARQ13)
            ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Envelopes mais de 30 folhas "+WEBCOR,"Em anexo relatorios para serem impressos",WNARQ13,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
      EndIf
      If TEMFRENTEVERSO
        squencia=0
        DECLARE VLABEL[ADIR(WEBLOCAL6+"*.PDF")]
        ADIR(WEBLOCAL6+"*.PDF",vlabel)
        If LEN(VLABEL)#0
          FOR JI=1 TO LEN(VLABEL)
            CODCLI=VAL(SUBSTR(VLABEL[JI],9,8))
            SELE 37
            SEEK CODCLI
            If FOUND()
              REPLACE FAZER WITH .T.
            EndIf
            If WFAZCORREIO
              ZARQ1=WEBLOCAL+"FRENTE.PDF"
              oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
              Erase &ZARQ1
              oReg:Set("PDFFilename",ZARQ1)
              oReg:Close()
              PRINT oPrn NAME "frente" TO TEMPDF1
              DEFINE FONT oFont    NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
              DEFINE FONT oFontv   NAME "verdana"             SIZE 0,12 OF oPrn
              DEFINE FONT oFont1   NAME "verdana"             SIZE 0,12 BOLD OF oPrn
              DEFINE FONT oFont2   NAME "verdana"             SIZE 0,9 OF oPrn
              DEFINE FONT oFont3   NAME "verdana"             SIZE 0,9 BOLD OF oPrn
              nRowStep = oPrn:nVertRes() / 66
              nColStep = oPrn:nHorzRes() / 80
              squencia+=1
              FRENTEFV()
              ENDPRINT
              DO WHILE .T.
                CURSORWAIT()
                If (nHandle:=FOpen(ZARQ,FO_READ+FO_EXCLUSIVE))!=-1
                  fClose(nHandle)
                  EXIT
                Else
                  fClose(nHandle)
                EndIf
              ENDDO
              NARQ7=WEBLOCAL6+"TEMP.PDF"
              Erase &NARQ7
              NARQ9=WEBLOCAL6+VLABEL[JI]
              FRENAME(NARQ9,NARQ7)
              M->LOCAL23=TRIM("/"+CURDIR())
              LCHDIR(WEBLOCAL6)
              comando=WCORRENTE+"\PDFTK "+NARQ7+" "+ZARQ1+" cat output "+NARQ9
              MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
              MEMOWRIT(weblocal1+"FZ.BAT",mzlinha)
              MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 13"+MZLINHA)
              WAITRUN(weblocal1+"FZ.BAT",0)
              DO WHILE FILE(weblocal1+"AGUARDE.TXT")
                SysRefresh()
              ENDDO
              CURSORWAIT()
              MTEZ=WEBLOCAL1+"FZ.BAT"
              Erase &MTEZ
              Erase &NARQ7
              Erase &ZARQ1
              LCHDIR(M->LOCAL23)
              CURSORWAIT()
            EndIf
          NEXT JI
        EndIf
      EndIf
      SELE 37
      NARQ9=WEBLOCAL4+"RELAT.XLS"
      NARQ12=WEBLOCAL4+"RELAT.PDF"
      ARQ_XLS=NARQ9
      GERAXLS(1,"Relatório Impressões "+WEBCOR+" "+WEBASSUNTO)
      NARQ8=ALLTRIM(WEBLOCAL)+"LOGS.DBF"
      SELE 38
      If USEREDE(NARQ8,.T.,10,"ETIQ1")
        SELE 37
        GO TOP
        DO WHILE !EOF()
          DECLARE _VAR1[FCOUNT()]
          AFIELDS(_VAR1)
          FOR t = 1 TO FCOUNT()
            SysRefresh()
            _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
            _var5 = _var1 [t]
            &_var4 = &_var5
          NEXT t
          RELEASE _VAR1
          SELE 38
          ADIREG(0)
          PUTREC()
          SELE 37
          SKIP
        ENDDO
      EndIf
      SELE 38
      USE
      SELE 37
      USE
      NARQ7=WEBLOCAL4+"ETIQUETA.KEY"
      Erase &NARQ7
      NARQ7=WEBLOCAL4+"ETIQUETA.DBF"
      EMAILRETORNO=WEBIMP
      If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
        //  ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - (Base para Baixa Codigo de Barras - "+WEBCOR,"Em anexo arquivo para baixa no sistema de código de Barras",NARQ7,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
      NARQ4=WEBLOCAL6+"IMPRIMIR.PDF"
      Erase &NARQ4
      bparq=weblocal1+"EXTRATO.PDF"
      BPARQ1=WEBLOCAL6+"EXTRATO.PDF"
      If FILE(BPARQ)
        LZCOPYFILE(BPARQ,BPARQ1)
        Erase &BPARQ
      EndIf
      DECLARE VLABEL[ADIR(WEBLOCAL6+"*.PDF")]
      ADIR(WEBLOCAL6+"*.PDF",vlabel)
      If LEN(VLABEL)#0
        If WFAZCORREIO
          comando=WCORRENTE+"\PDFTK "+WEBLOCAL6+"*.PDF cat output "+NARQ4+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 14"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(weblocal1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
        EndIf
        FOR JI=1 TO LEN(VLABEL)
          If VLABEL[JI]="IMPRIMIR.PDF"
            LOOP
          EndIf
          NARQ5=WEBLOCAL6+ALLTRIM(VLABEL[JI])
          Erase &NARQ5
        NEXT JI
        If WFAZCORREIO
          comando=WCORRENTE+"\BRAZIP.EXE AH "+WEBLOCAL6+"IMPRIME.ZIP "+WEBLOCAL6+"*.PDF"+CHR(13)+CHR(10)
          MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+weblocal1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
          MEMOWRIT(weblocal1+"FZ.BAT",MZLINHA)
          MEMOWRIT(weblocal1+"AGUARDE.TXT","rotina 15"+MZLINHA)
          CURSORWAIT()
          WAITRUN(weblocal1+"FZ.BAT",0)
          CURSORWAIT()
          DO WHILE FILE(WEBLOCAL1+"AGUARDE.TXT")
            SysRefresh()
          ENDDO
          MTEZ=WEBLOCAL1+"FZ.BAT"
          Erase &MTEZ
          Erase &NARQ4
          NARQ4=WEBLOCAL6+"IMPRIME.ZIP"
          EMAILRETORNO=WEBIMP
          If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
            ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - impressão (Com Frente) "+WEBCOR,"Em anexo relatorios para serem impressos",NARQ4,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
      EndIf
      NARQ9=WEBLOCAL4+"RELAT.XLS"
      If FILE(NARQ9)
        NARQ91=WWEBLOCA9+"ENVIADO "+DTOS(DATE())+" "+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+".XLS"
        LZCOPYFILE(NARQ9,NARQ91)
      EndIf
      EMAILRETORNO=WEBREL
      If EMPTY(EMAILRETORNO)
        EMAILRETORNO=WEBMAIL
      EndIf
      If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
        ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Relatório XLS e PDF das Impressões "+WEBCOR,"Em anexo relatorios das impressões",NARQ9,NARQ12,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        If !EMPTY(WEBMAIL)
          ENVIAEMAIL(ALLTRIM(WEBMAIL),"BolsasNet - Relatório XLS e PDF das Impressões "+WEBCOR,"Em anexo relatorios das impressões",NARQ9,NARQ12,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          ENVIAEMAIL("impressoes@exatus.net","BolsasNet - Relatório XLS e PDF das Impressões "+WEBCOR,"Em anexo relatorios das impressões",NARQ9,NARQ12,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
        ENVIAEMAIL("wanderlei@exatus.net","BolsasNet - Relatório XLS e PDF das Impressões "+WEBCOR,"Em anexo relatorios das impressões",NARQ9,NARQ12,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
      SELE 37
      USE
      SELE 35
      USE
      SELE 39
      USE
      Erase &NARQ41
      SELE 30
      MBOLET=ALLTRIM(WEBLOCAL)+"BOLETIM.PDF"
      Erase &MBOLET
      If SEENVIAEMAIL#"S" .or. cod_oper#999
        If EMPTY(WERROARQ)
          MSGINFO("Terminou !",wsist)
        Else
          MSGSTOP("Erro nos arquivos "+WERROARQ,"Não foram gerados no formato A4")
        EndIf
      EndIf
    return 


    Procedure FRENTEFV()
      PAGE
      oPrn:Saybitmap(0,nColStep*3,ZARQ,oPrn:nHorzRes()*.95,oPrn:nVertRes())
      If WEBCOR="GERACAO"
        CALCCEP(ENDE2,14.2,2.2)
      Else
        CALCCEP(ENDE2,22.3,2.2)
      EndIf
      If WEBCOR="GERACAO"
        oPrn:say(oPrn:nVertres()/66*33.0,nColstep*8,LEFT(NOME,48),oFont1)
        oPrn:say(oPrn:nVertRes()/66*34.0,nColstep*8,ENDE1,oFontv)
        oPrn:say(oPrn:nVertRes()/66*34.8,nColstep*8,ENDE2,oFontv)
      Else
        oPrn:say(oPrn:nVertres()/66*52.0,nColstep*8,LEFT(NOME,48),oFont1)
        oPrn:say(oPrn:nVertRes()/66*53.0,nColstep*8,ENDE1,oFontv)
        oPrn:say(oPrn:nVertRes()/66*53.8,nColstep*8,ENDE2,oFontv)
      EndIf
      If WEBCOR="GERACAO"
        @ nRowStep*35.8,nColStep*8 CODE128 strzero(codigo,9) OF oPrn SIZE 0.6 MODE "A"
      Else
        @ nRowStep*54.8,nColStep*8 CODE128 strzero(codigo,9) OF oPrn SIZE 0.6 MODE "A"
      EndIf
      If WEBCOR="GERACAO"
        oPrn:say(oPrn:nVertRes()/66*37.2,ncolstep*8,strzero(codigo,9)+" "+strzero(squencia,4),oFont)
      Else
        oPrn:say(oPrn:nVertRes()/66*56.2,ncolstep*8,strzero(codigo,9)+" "+strzero(squencia,4),oFont)
      EndIf
      If WEBCOR="PLANNER" .And. FILE(WEBLOCAL+"LOGO.BMP")
        saybitmap(14.5,1.00,3.691*1.5,1.819*1.5,WEBLOCAL+"LOGO.BMP",oPrn)
        saybitmap(14.5,1.01,3.691*1.5,1.819*1.5,WEBLOCAL+"LOGO.BMP",oPrn)
        oPrn:say(oPrn:nVertres()/66*36.0,nColstep*4,"     CENTRAL PLANNER",oFont3)
        oPrn:say(oPrn:nVertres()/66*36.6,nColstep*4,"     DE ATENDIMENTO   OUVIDORIA",oFont3)
        oPrn:say(oPrn:nVertres()/66*37.2,nColstep*4,"     0800 179444         0800 7722231",oFont3)
        oPrn:say(oPrn:nVertres()/66*18.8,nColstep*8,"PLANNER Corretora de Valores S/A",oFont2)
        oPrn:say(oPrn:nVertRes()/66*19.4,nColstep*8,"Av.Brig.Faria Lima, 3900 10.andar",oFont2)
        oPrn:say(oPrn:nVertRes()/66*20.0,nColstep*8,"04538-132 - Itaim Bibi - São Paulo - SP",oFont2)
      ElseIf WEBCOR="CONVEC" .And. FILE(WEBLOCAL+"LOGO.BMP")
        saybitmap(14.5,1.00,3.691*1.5,1.819*1.5,WEBLOCAL+"LOGO.BMP",oPrn)
        saybitmap(14.5,1.01,3.691*1.5,1.819*1.5,WEBLOCAL+"LOGO.BMP",oPrn)
        oPrn:say(oPrn:nVertres()/66*18.8,nColstep*8,"CONVENÇÃO S/A CVC",oFont2)
        oPrn:say(oPrn:nVertRes()/66*19.4,nColstep*8,"Rua Amauri, 255 8.andar",oFont2)
        oPrn:say(oPrn:nVertRes()/66*20.0,nColstep*8,"01448-000 - São Paulo - SP",oFont2)
      EndIf
      ENDPAGE
    return 

    Procedure CADWEB1(PARAM)
      SEENVIAEMAIL=PARAM
      SELE 30
      GO TOP
      CURSORWAIT()
      DO WHILE .NOT. EOF()
        SysRefresh()
        If TIPO="U"
          SKIP
          LOOP
        EndIf
        WARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        If FILE(WARQ)
          SKIP
          LOOP
        EndIf
        WBPARQ1=ALLTRIM(PASTA)+"FICHA.CLI"
        WBPARQ2=ALLTRIM(PASTA)+"FICHA.DAT"
        If !FILE(WBPARQ1) .And. !FILE(WBPARQ2)
          SKIP
          LOOP
        EndIf
        SELE 30
        WEBLOCAL1=ALLTRIM(PASTA)+"TEMP\"
        WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\BOLSASNET\"
        VERDIRETORIO(WEBLOCAL1)
        WEBLOCAL7=alltrim(pasta)+"CONTRS\"
        VERDIRETORIO(WEBLOCAL7)
        DECLARE VLABEL1[ADIR(WEBLOCAL7+"*.*")]
        ADIR(WEBLOCAL7+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL7+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        SELE 30
        CRIAMOVFICHA(alltrim(BOLSANET->PASTA))
        CRIARQFICHA(alltrim(BOLSANET->PASTA))
        SELE 30
        WFAZDADOS=DADOS
        WFAZCORREIO=CORREIO
        WEBLOCAL=ALLTRIM(PASTA)
        WEBCUSTO=CUSTOMIL
        WEBCOR=CORRETORA
        WEBNOMECOR=NOMECOR
        WEBMAIL=EMAIL
        WEBCORPO=OEMTOANSI(CORPOEMAIL)
        WEBSENHA=SENHAEMAIL
        WEBJUD=MAILJUD
        WEBCAD=MAILCAD
        WEBMAR=MAILMAR
        WEBMKT=MAILMKT
        WEBIMP=MAILIMP
        WEBREL=MAILREL
        MCONTA=M->FROM
        WEBSERV=ALLTRIM(SMTP)
        WEBFORCA=FORCAEMSTP
        WEBLOGIN=LOGEMAIL
        WEBPDFS=SENHAPDF
        WEBBOV=CODBOV
        SELE 30
        DECLARE TX1[1]
        TX1[1]="BolsasNet Cadastro Web"+WEBCOR
        LOGFILE(WEBLOCAL+"ROTINAS.TXT",TX1)
        FEZALGO=.F.
        CADWEB()
        SELE 30
        SKIP
      ENDDO
      Erase &WBPARQ2
      cursorarrow()
    return 

    Procedure CADWEB()
      //* SE FOR PARA GERAR O ARQUIVO PCIN
      NARQ=WEBLOCAL+"FICHA.DAT"
      NARQ1=WEBLOCAL+"FICHA.CLI"
      If FILE(NARQ) .OR. FILE(NARQ1)
        SELE 32
        NARQ5=WEBLOCAL+"CONTRS.DBF"
        If !USEREDE(NARQ5,.F.,10,"CONTRS")
          SELE 30
          return 
        EndIf
        LOCATE FOR CODIGO=2
        If EOF()
          WCORPOF_CAD=""
        Else
          WCORPOF_CAD=CONTEUDO
        EndIf
        LOCATE FOR CODIGO=5
        If EOF()
          WCORPOJ_CAD=""
        Else
          WCORPOJ_CAD=CONTEUDO
        EndIf
        USE
        SELE 31
        NARQ4=WEBLOCAL+"FICHA.DBF"
        If !USEREDE(NARQ4,.F.,10,"FICHA")
          SELE 30
          return 
        EndIf
        GO TOP
        DO WHILE !EOF()
          If WEBCOR#"SLW"
            If APROVADO .And. CODCLI#0 .And. M->DAT_HOJE-DTAPROV>60 .And. !empty(dtaprov) .And. REGLOCK(10)
              DELE
            EndIf
          EndIf
          If !DELE() .And. !APROVADO .And. CODCLI=0 .And. M->DAT_HOJE-COBRADO10>10 .And. M->DAT_HOJE-CTOD(LEFT(GERADO,10))>10 .And. !gerar .And. REGLOCK(10)
            If IDTIT#0
              SKIP
              LOOP
            EndIf
            If CODCLI#0
              SKIP
              LOOP
            EndIf
            REPLACE COBRADO10 WITH M->DAT_HOJE

            If SEENVIAEMAIL="S"
              //        ENVIAEMAIL(EMAIL,webcor,IIF(PESSOA="F",WCORPOF_CAD,WCORPOJ_CAD),"","",WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA)
            EndIf

            EMAILRETORNO=WEBCAD
            If EMPTY(EMAILRETORNO)
              EMAILRETORNO=WEBMAIL
            EndIf
            If SEENVIAEMAIL="S"
              //       ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Documentos Cadastro:"+NOME,"Enviado email de cobrança de documento ","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            EndIf
          EndIf
          SKIP
        ENDDO
      EndIf
      NARQ=WEBLOCAL+"FICHA.DAT"
      If FILE(NARQ)
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Gerando PCIN do arquivo Ficha "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        SELE 32
        NARQ5=WEBLOCAL+"CONTRS.DBF"
        If !USEREDE(NARQ5,.F.,10,"CONTRS")
          SELE 30
          return 
        EndIf
        LOCATE FOR CODIGO=3
        If EOF()
          WCORPOF_CAD=""
        Else
          WCORPOF_CAD=CONTEUDO
        EndIf
        LOCATE FOR CODIGO=6
        If EOF()
          WCORPOJ_CAD=""
        Else
          WCORPOJ_CAD=CONTEUDO
        EndIf
        USE
        SELE 31
        NARQ4=WEBLOCAL+"FICHA.DBF"
        If !USEREDE(NARQ4,.F.,10,"FICHA")
          SELE 30
          return 
        EndIf
        GO TOP
        WDADOSP=WEBLOCAL1+"PCIN"+STRZERO(WEBBOV,4)+".TXT"
        ARQ=FCREATE(WDADOSP)
        PL=CHR(13)+CHR(10)
        M->DAT_HOJE=DATE()
        LINHA="00PCIN"+STRZERO(WEBBOV,4)+"CBLC    "+strzero(WEBBOV,4)+STRZERO(1,9)+STRZERO(YEAR(M->DAT_HOJE),4)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(DAY(M->DAT_HOJE),2)+STRZERO(YEAR(M->DAT_HOJE),4)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(DAY(M->DAT_HOJE),2)+SPACE(653)+PL
        FWRITE(ARQ,LINHA,LEN(LINHA))
        WCONT=1
        DO WHILE !EOF()
          If CODCLI#0 .And. IDTIT=0 .And. REGLOCK(10) .And. DTAPROV=CTOD("")
            REPLACE APROVADO WITH .T.
            UNLOCK
          EndIf
          If !DELE() .And. APROVADO .And. CODCLI#0 .And. REGLOCK(10) .And. IDTIT=0
            REPLACE APROVADO WITH .F.
            REPLACE DTAPROV WITH M->DAT_HOJE
            WCPF=Strtran(CNPJCPF,".","")
            WCPF=Strtran(wcpf,"-","")
            WCPF=Strtran(wcpf,"/","")
            WDOC=Strtran(DOC,"-","")
            WDOC=Strtran(WDOC," ","")
            WDOC=Strtran(WDOC,"/","")
            WDOC=UPPER(WDOC)
            MD=(16-LEN(WDOC))
            If MD>0
              WDOC=WDOC+SPACE(MD)
            EndIf
            WCONTA=ALLTRIM(LEFT(CONTA,12))+LEFT(DIGCONTA,1)
            WCONTA=Strtran(WCONTA,"-","")
            WCONTA=Strtran(WCONTA,"/","")
            WCONTA=Strtran(WCONTA," ","")
            MD=13-LEN(WCONTA)
            If MD>0
              WCONTA=WCONTA+SPACE(MD)
            EndIf
            WCONT=WCONT+1
            LINHA="01I"+STRZERO(VAL(WCPF),15)+STRZERO(DAY(M->DAT_HOJE),2)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(YEAR(M->DAT_HOJE),4)+STRZERO(DAY(M->DAT_HOJE),2)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(YEAR(M->DAT_HOJE),4)
            FWRITE(ARQ,LINHA,LEN(LINHA))
            cddepend=strzero(coddepend,2)
            TPINVEST=strzero(investidor,5)
            CDASSESSOR=strzero(assessor,5)
            TPDOCTOC=LEFT(TIPODOC,2)
            LINHA=LEFT(DTNASC,2)+SUBSTR(DTNASC,4,2)+RIGHT(DTNASC,4)+CDDEPEND+PESSOA+UPPER(SINAL(OEMTOANSI(NOME)))+WDOC+"  "+UPPER(SINAL(OEMTOANSI(ORGDOC)))+SEXO+LEFT(ESTCIVIL,1)+VINCULO+UPPER(SINAL(OEMTOANSI(CONJUGE)))+"0"+SPACE(60)+"000"+UPPER(SINAL(LEFT(PROFISSAO,3)))+"  0"+UPPER(SINAL(LEFT(NACIONAL,1)))+UPPER(SINAL(LEFT(PAIS,3)))+STRZERO(CODCLI,7)+SPACE(15)+STRZERO(0,15)+LEFT(BANCO,3)+STRZERO(AGENCIA,4)+strzero(val(LEFT(AGDIG,1)),1)+wconta+STRZERO(DIGITO,1)
            FWRITE(ARQ,LINHA,LEN(LINHA))
            WNUMER=ALLTRIM(TRANSFORM(NUMERO,"99999"))
            WNUMER=WNUMER+SPACE(5-LEN(WNUMER))
            LINHA=CDASSESSOR+SPACE(24)+" "+UPPER(SINAL(OEMTOANSI(ENDERECO)))+WNUMER+UPPER(SINAL(OEMTOANSI(COMPL)))+UPPER(SINAL(OEMTOANSI(BAIRRO)))+UPPER(SINAL(OEMTOANSI(CIDADE)))+UF+STRZERO(VAL(Strtran(CEP,"-","")),9)+STRZERO(DDD,7)+STRZERO(VAL(Strtran(TEL,"-","")),9)+STRZERO(0,5)+STRZERO(DDDFAX,7)+STRZERO(VAL(Strtran(FAX,"-","")),9)+"0      "+"SS"+STRZERO(0,5)
            FWRITE(ARQ,LINHA,LEN(LINHA))
            LINHA=STRZERO(0,9)+SPACE(8)+STRZERO(WEBBOV,5)+STRZERO(0,7)+STRZERO(0,7)+"A "+STRZERO(0,15)+STRZERO(0,19)+"N"+STRZERO(0,7)+TPINVEST+"F"+UFDOC+TPDOCTOC+WDOC+ORGDOC+UFDOC+STRZERO(0,5)+LOWER(EMAIL)+PEXPOSTA+SPACE(26)+PL
            FWRITE(ARQ,LINHA,LEN(LINHA))
            If SEENVIAEMAIL="S"
              If PESSOA="F"
                WCORPO1=WCORPOF_CAD
                WCORPO1=Strtran(WCORPO1,";;NOME;:",TRIM(NOME))
                WCORPO1=Strtran(WCORPO1,";;CODIGO;:",STRZERO(CODCLI,8))
              Else
                WCORPO1=WCORPOJ_CAD
                WCORPO1=Strtran(WCORPO1,";;NOME;:",TRIM(NOME))
                WCORPO1=Strtran(WCORPO1,";;CODIGO;:",STRZERO(CODCLI,8))
              EndIf
              ENVIAEMAIL(EMAIL,"Cadastro Realizado com Sucesso: "+STRZERO(codcli,8),"Codigo de Acesso na corretora: "+STRZERO(CODCLI,8)+"-"+STRZERO(DIGITO,1)+"<br><br>"+WCORPO1,"","",WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA)
            EndIf
          EndIf
          SELE 31
          SKIP
        ENDDO
        SELE 31
        USE
        LINHA="99PCIN"+STRZERO(WEBBOV,4)+"CBLC    "+STRZERO(WEBBOV,4)+STRZERO(1,9)+STRZERO(YEAR(M->DAT_HOJE),4)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(DAY(M->DAT_HOJE),2)+STRZERO(WCONT+1,9)+STRZERO(YEAR(M->DAT_HOJE),4)+STRZERO(MONTH(M->DAT_HOJE),2)+STRZERO(DAY(M->DAT_HOJE),2)+SPACE(644)+PL
        FWRITE(ARQ,LINHA,LEN(LINHA))
        FCLOSE(ARQ)
        EMAILRETORNO=WEBCAD
        If EMPTY(EMAILRETORNO)
          EMAILRETORNO=WEBMAIL
        EndIf
        If SEENVIAEMAIL="S"
          ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - PCIN ","Arquivo para inclusão de dados no Sinacor",WDADOSP,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
        WBPARQ2=ALLTRIM(bolsanet->PASTA)+"FICHA.DAT"
        Erase &WBPARQ2
      EndIf
      SELE 30
      //* SE FOR PARA GERAR OS CADASTROS E ENVIAR POR EMAIL AO CLIENTE
      NARQ=WEBLOCAL+"FICHA.CLI"
      FEZCERTO=.F.
      If FILE(NARQ)
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Gerando Contratos para enviar ao Cliente "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        SELE 31
        NARQ4=WEBLOCAL+"FICHA.DBF"
        If !USEREDE(NARQ4,.F.,10,"FICHA","S")
          SELE 30
          return 
        EndIf
        GO TOP
        DO WHILE !EOF()
          If IDTIT#0
            SKIP
            LOOP
          EndIf
          CONTAD=0
          If GERAR .And. REGLOCK(10)
            REPLACE GERAR WITH .F.
            replace IMOVEIS WITH IMOVVLR5+IMOVVLR4+IMOVVLR3+IMOVVLR2+IMOVVLR1
            REPLACE MOVEIS WITH BENSVLR5+BENSVLR4+BENSVLR3+BENSVLR2+BENSVLR1
            REPLACE GERADO WITH DTOC(M->DAT_HOJE)+" "+TIME()
            REPLACE ENVPCLI WITH M->DAT_HOJE
            NARQ5=WEBLOCAL+"CONTRS.DBF"
            If WEBCOR="SLW"
              If TIPOCOR="F"
                NARQ5="C:\NOTAS\SLW\BMF\CONTRS.DBF"
              Else
                NARQ5="C:\NOTAS\SLW\BOVESPA\CONTRS.DBF"
              EndIf
            Else
              NARQ5=WEBLOCAL+"CONTRS.DBF"
            EndIf
            SELE 32
            If !USEREDE(NARQ5,.F.,10,"CONTRS","S")
              SELE 30
              return 
            EndIf
            SELE 32
            GO TOP
            DO WHILE !EOF()
              If FICHA->PESSOA=FJ .And. FIXO .And. CODIGO>6 .And. !DELE()
                CONTAD=CONTAD+1
                ORIGEM=WEBLOCAL7+"BASE.TXT"
                Erase &ORIGEM
                MZCONTEUDO=CONTEUDO
                GERAARQ(CONTAD)
              EndIf
              SELE 32
              SKIP
            ENDDO
            SELE 31
            WCONTRATOS=ALLTRIM(CONTRATOS)
            If LEN(WCONTRATOS)#0
              If RIGHT(WCONTRATOS,1)#","
                WCONTRATOS=WCONTRATOS+","
              EndIf
            EndIf
            SELE 32
            GO TOP
            DO WHILE !EOF()
              If DELE()
                SKIP
                LOOP
              EndIf
              WTS=ALLTRIM(STR(CODIGO))+","
              If FICHA->PESSOA=FJ .And. !FIXO .And. CODIGO>6 .And. AT(WTS,WCONTRATOS)#0
                CONTAD=CONTAD+1
                ORIGEM=WEBLOCAL7+"BASE.TXT"
                Erase &ORIGEM
                MZCONTEUDO=CONTEUDO
                GERAARQ(CONTAD)
              EndIf
              SELE 32
              SKIP
            ENDDO
            SELE 31
            If IDCTIT#0
              MREGS=RECNO()
              MIDCTIT=IDCTIT
              LOCATE FOR CODIGO=MIDCTIT
              SELE 32
              GO TOP
              DO WHILE !EOF()
                If FJ="C" .And. FIXO .And. CODIGO>6
                  CONTAD=CONTAD+1
                  ORIGEM=WEBLOCAL7+"BASE.TXT"
                  Erase &ORIGEM
                  MZCONTEUDO=CONTEUDO
                  GERAARQ(CONTAD)
                EndIf
                SELE 32
                SKIP
              ENDDO
              SELE 31
              GO MREGS
            EndIf
            NARQ4=WEBLOCAL7+"Documentos_de_Cadastros_"+ALLTRIM(webcor)+".PDF"
            comando=WCORRENTE+"\PDFTK "+WEBLOCAL7+"*.PDF cat output "+NARQ4+CHR(13)+CHR(10)
            MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+WEBLOCAL1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
            MEMOWRIT(WEBLOCAL1+"FZ.BAT",MZLINHA)
            MEMOWRIT(WEBLOCAL1+"AGUARDE.TXT","rotina 16"+MZLINHA)
            CURSORWAIT()
            WAITRUN(WEBLOCAL1+"FZ.BAT",0)
            CURSORWAIT()
            DO WHILE FILE(WEBLOCAL1+"AGUARDE.TXT")
              SysRefresh()
            ENDDO
            EMAILRETORNO=WEBCAD
            If EMPTY(EMAILRETORNO)
              EMAILRETORNO=WEBMAIL
            EndIf
            SELE 32
            LOCATE FOR CODIGO=1
            If EOF()
              WCORPOF_CAD=""
            Else
              WCORPOF_CAD=CONTEUDO
            EndIf
            LOCATE FOR CODIGO=4
            If EOF()
              WCORPOJ_CAD=""
            Else
              WCORPOJ_CAD=CONTEUDO
            EndIf
            NOMAARQUIVO:=DIRECTORY(narq4)
            If NOMAARQUIVO[1][2]<10000
              SELE 31
              If REGLOCK(10)
                REPLACE GERAR WITH .T.
                FEZCERTO=.F.
              EndIf
            Else
              FEZCERTO=.T.
            EndIf
            If SEENVIAEMAIL="S" .And. FEZCERTO
              ENVIAEMAIL(LOWER(ALLTRIM(FICHA->EMAIL)),"Documentos "+TRIM(WebCor)+" Cliente: "+FICHA->NOME,IIF(FICHA->PESSOA="F",WCORPOF_CAD,WCORPOJ_CAD),NARQ4,"",WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA)
              ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Cadastro Cliente-WEB","Segue em anexo os Documentos preenchidos",NARQ4,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            EndIf
            DECLARE VLABEL[ADIR(WEBLOCAL7+"*.*")]
            ADIR(WEBLOCAL7+"*.*",vlabel)
            I=0
            FOR I=1 TO LEN(VLABEL)
              SysRefresh()
              NARQ=WEBLOCAL7+VLABEL[I]
              Erase &NARQ
            NEXT I
          EndIf
          SELE 32
          USE
          SELE 31
          SKIP
        ENDDO
        SELE 31
        USE
      EndIf
      If FEZCERTO
        WBPARQ1=ALLTRIM(bolsanet->PASTA)+"FICHA.CLI"
        Erase &WBPARQ1
      EndIf
      SELE 30
    return 

    Procedure GERAARQ(CONTAD)
      SELE 31
      FNAME100=WEBLOCAL7+STRZERO(CONTAD,5)+".HTM"
      ARQPDF=WEBLOCAL7+STRZERO(CONTAD,5)+".PDF"
      MZCONTEUDO=OEMTOANSI(MZCONTEUDO)
      MZCONTEUDO=GERA1ARQ(MZCONTEUDO)
      MEMOWRIT(FNAME100,MZCONTEUDO)
      CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+"\HTML2PDF.EXE "+FNAME100+" "+ARQPDF+CHR(13)+CHR(10)+"DEL "+WCORRENTE+"\PRONTO.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
      MEMOWRIT(wcorrente+"FZ.BAT",CMACRO)
      MEMOWRIT(WCORRENTE+"\PRONTO.TXT","NADA")
      CMACRO=wcorrente+"FZ.BAT"
      WAITRUN(CMACRO,0)
      CURSORWAIT()
      DO WHILE FILE(WCORRENTE+"\PRONTO.TXT") .And. !file(arqpdf)
        SysRefresh()
      ENDDO
      Erase &FNAME100
    return 

    Function GERA1ARQ(LINH)
      CURSORWAIT()
      while AT(";;",LINH)#0
        SysRefresh()
        IN=AT(";;",LINH)
        FI=AT(";:",LINH)
        TAM=LEN(LINH)
        TPC=ALLTRIM(SUBSTR(LINH,IN+2,FI-2-IN))
        TPC=Strtran(TPC,";","")
        TPC=Strtran(TPC,":","")
        If EMPTY(TPC)
          return  LINH
        EndIf
        If !VERFIELD(TPC)
          LINH=TRIM(LEFT(LINH,IN-1)+SUBSTR(LINH,FI+2,TAM))
          return  LINH
        EndIf
        If VALTYPE(&TPC)="U"
          return  LINH
        EndIf
        VARI1=""
        If VALTYPE(&TPC)="C"
          VARI1=&TPC
          VARI1=UPPER(oemtoansi(vari1))
          VARI1=SINAL(UPPER(VARI1)) // AGORA
          VARI1=UPPER(VARI1)
          If LEN(VARI1)=1
            If VARI1="S"
              VARI1=ansitooem("(X) Sim ( ) Não")
            EndIf
            If VARI1="N"
              VARI1=ansitooem("( ) Sim (X) Não")
            EndIf
          EndIf
          If TPC="ORDEM" .And. EMPTY(VARI1)
            VARI1=ANSITOOEM("( ) Concordo  ( ) Não Concordo ( ) Concordo Sob Consulta")
          EndIf
          If TPC="ORDCOR" .And. EMPTY(VARI1)
            VARI1=ANSITOOEM("( ) Concordo  ( ) Não Concordo ( ) Concordo Sob Consulta")
          EndIf
          If TPC="GERADO"
            VARI1=LEFT(&TPC,10)
          EndIf
          If TPC="PROFISSAO"
            VARI1=SUBSTR(VARI1,5,50)
          EndIf
          If TPC="NACIONAL"
            VARI1=SUBSTR(VARI1,5,50)
          EndIf
          If TPC="ESTCIVIL"
            VARI1=SUBSTR(VARI1,5,50)
          EndIf
        ElseIf VALTYPE(&TPC)="D"
          VARI1=DTOC(&TPC)
        ElseIf VALTYPE(&TPC)="N"
          VARI1=ALLTRIM(STR(&TPC))
          If VAL(VARI1)=0
            VARI1=""
          EndIf
        EndIf
        LINH=TRIM(LEFT(LINH,IN-1)+OEMTOANSI(VARI1)+SUBSTR(LINH,FI+2,TAM))
      ENDDO
      If AT(CHR(13),LINH)=0
        LINH=LINH+CHR(13)+CHR(10)
      EndIf
    return  LINH

    Function VERFIELD(CAMPO)
      EXISTE=.F.
      FOR nField := 1 TO FCOUNT()
        BDFIELD=PADR(FIELD(nField),10)
        If BDFIELD = CAMPO
          EXISTE=.T.
        EndIf
      NEXT
    return  EXISTE

    Procedure FAZEREMAIL(NARQ)
      M->WLOCAL4=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
      LZCOPYFILE(NARQ,M->WLOCAL4)
      NWARQ=LOWER(NARQ)
      NWARQ1=LOWER(WEBLOCAL+"EMAIL.PDF")
      If NWARQ#NWARQ1
        FRENAME(NARQ,NWARQ1)
      EndIf
      NARQ=NWARQ1
      FEZALGO=.T.
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Atualizando emails EMAIL.PDF "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      COMANDO=WCORRENTE+"\PDFTEXT.EXE -eol dos -layout "+NARQ+" "+WEBLOCAL+"EMAIL.TXT"
      WAITRUN(COMANDO,0)
      CURSORWAIT()
      oText:=TTxtfile():NEW(WEBLOCAL+"EMAIL.TXT")
      cursorwait()
      CONTAD=0
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)#"N"
        CONTAD=1
      EndIf
      SEMNUMERO=.F.
      TP="1"
      FOR ZXC=1 TO 10
        ttt=oText:ReadLine()
        If AT("SCEMAIL.QRP",TTT)#0
          SEMNUMERO=.T.
        EndIf
        If AT("ATIVIDADE BOLSA",TTT)#0
          TP="BOVESPA"
        ElseIf AT("ATIVIDADE BM&F",TTT)#0
          TP="BMF"
        EndIf
        oText:skip()    /// MUDAR PARA BAIXO (ACERTAR CORVAL PRIMEIRO).
      NEXT ZXC
      TTT=oText:Readline()
      SELE 31
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      If SEMNUMERO
        FZMAILSEM()
        return 
      EndIf
      SELE 31
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If LEFT(EMAIL,1)#"["
          If TP="BOVESPA" .And. (BMFBOV="B" .OR. BMFBOV=" ")
            DELE
          EndIf
          If TP="BMF" .And. BMFBOV="F"
            DELE
          EndIf
        EndIf
        SKIP
      ENDDO
      PACK
      SELE 31
      USE
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      INDEX ON strzero(cliente,9)+LEFT(EMAIL,100) TO &NARQB
      SELE 32
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 32
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTE","S")
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      If TP="BMF"
        EBMF=.T.
      Else
        EBMF=.F.
      EndIf
      GO TOP
      If !FAZEREM
        DO WHILE !EOF()
          If REGLOCK(10)
            REPLACE TEMMAIL WITH .F.
          EndIf
          SKIP
        ENDDO
        FAZEREM=.T.
      EndIf
      INDEX ON CODIGO TO &NARQ4
      CODANT=0
      CURSORWAIT()
      novo=.f.
      M->NOME=""
      cursorwait()
      NCOL=35
      ncol1=12
      NCONTAD=0
      DO WHILE CONTAD>=NCONTAD
        cursorwait()
        TTT=oText:Readline()
        NCONTAD=NCONTAD+1
        If AT("CPF/CNPJ",TTT)#0
          NCOL=AT("CPF/CNPJ",TTT)-2
          NCOL1=AT("Nom",TTT)-1
          oText:Skip()
          loop
        EndIf
        CODATU=VAL(Strtran((LEFT(TTT,NCOL1))," ",""))
        If TP="BMF" .And. (WEBCOR<>"GERACAO" .And. WEBCOR<>"PLANNER" .And. WEBCOR<>"CONVEC")
          CODATU=VAL(LEFT(STRZERO(CODATU,11),10))
        EndIf
        If CODATU#0
          SELE 32
          SEEK CODATU
          If !FOUND()
            ADIREG(0)
            REPLACE CODIGO WITH CODATU
            REPLACE EEMAIL WITH .T.
            REPLACE MAILING WITH .T.
            REPLACE EIMPRE WITH .F.
            REPLACE EGUARD WITH .T.
          EndIf
          If REGLOCK(10)
            REPLACE TEMMAIL WITH .T.
          EndIf
          TT1=RIGHT(TTT,LEN(TTT)-NCOL)
          TT2=""
          FOR WAN=1 TO LEN(TT1)
            If SUBSTR(TT1,WAN,1)#" "
              If SUBSTR(TT1,WAN,1)$"0123456789/"
                TT2=TT2+SUBSTR(TT1,WAN,1)
              EndIf
            EndIf
          NEXT WAN
          do while .t.
            If left(tt2,1)="/"
              TT2=SUBSTR(TT2,2,LEN(TT2))
            Else
              EXIT
            EndIf
          ENDDO
          WNASC=CTOD(LEFT(RIGHT(TT2,11),10))
          WCNPJ=VAL(LEFT(TT2,AT("/",TT2)-3))
          If REGLOCK(10)
            REPLACE CODIGO WITH CODATU
            REPLACE NOME WITH ALLTRIM(SUBSTR(TTT,11,50))
            REPLACE CNPJCPF WITH STRZERO(WCNPJ,14)
            If webcor="GERACAO"
              WCNPJ=ALLTRIM(TRANSFORM(VAL(CNPJCPF),"99999999999999"))
              If val(cnpjcpf)>99999999999
                REPLACE SENHA WITH LEFT(STRZERO(VAL(CNPJCPF),14),4) // aqui o sistema coloca zero a esquerda(14) digitos (para a Geração).
                REPLACE SENHA WITH LEFT(WCNPJ,4)
              Else
                If CIC(STRZERO(VAL(CNPJCPF),11))
                  REPLACE SENHA WITH LEFT(STRZERO(VAL(CNPJCPF),11),4)
                Else
                  REPLACE SENHA WITH LEFT(STRZERO(VAL(CNPJCPF),14),4)
                EndIf
              EndIf
            EndIf
            If WNASC#CTOD("")
              REPLACE NASC WITH WNASC
            EndIf
            If EMPTY(INCLUIDO)
              REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERINC WITH "Arqv.Email Sinacor"
            EndIf
            If EMPTY(SENHA)
              REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
            EndIf
          EndIf
          M->NOME=NOME
          CODANT=CODATU
          oText:skip()
          NCONTAD=NCONTAD+1
          TTT=oText:Readline()
        EndIf
        If CODATU=0
          CODATU=CODANT
        EndIf
        CODANT=CODATU
        m->EMAIL=SUBSTR(TTT,11,107)
        M->EMAIL=Strtran(M->EMAIL," ","")
        M->EMAIL=Strtran(M->EMAIL,";/;",";")
        M->EMAIL=LOWER(M->EMAIL)
        If EMPTY(M->EMAIL) .or. at("@",M->EMAIL)=0
          oText:skip()
          LOOP
        EndIf
        SELE 31
        SEEK STRZERO(CODATU,9)
        FAZE=.T.
        DO WHILE !EOF() .And. CODATU=CLIENTE
          If ALLTRIM(M->NOME)<>ALLTRIM(NOME) .And. REGLOCK(10)
            REPLACE NOME WITH M->NOME
          EndIf
          If TRIM(LOWER(M->EMAIL))=TRIM(LOWER(EMAIL))
            FAZE=.F.
            DECLARE TX1[1]
            TX1[1]="Email JÁ CADASTRADO: "+STRZERO(CLIENTE,8)+"-"+NOME+CHR(13)+CHR(10)+m->EMAIL+CHR(13)+CHR(10)
            If !EMPTY(WDADOS)
              LOGFILE(WDADOS,TX1)
            EndIf
          EndIf
          SKIP
        ENDDO
        If FAZE
          SEEK STRZERO(CODATU,9)+LEFT(M->EMAIL,100)
          If !FOUND() .OR. DELE()
            ADIREG(0)
            REPLACE CLIENTE WITH CODATU
            REPLACE EMAIL WITH M->EMAIL
            REPLACE NOME WITH M->NOME
            If TP="BOVESPA"
              REPLACE BMFBOV WITH "B"
            EndIf
            If TP="BMF"
              REPLACE BMFBOV WITH "F"
            EndIf
            LINHA=oText:Readline()
            REPLACE CODIGO WITH RECNO()
            REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()+" EMAIL"
            REPLACE USERALT WITH "Relatório Email"
            DECLARE TX1[1]
            TX1[1]="Email Cadastrado : "+STRZERO(CLIENTE,8)+"-"+NOME+CHR(13)+CHR(10)+EMAIL+CHR(13)+CHR(10)
            If !EMPTY(WDADOS)
              LOGFILE(WDADOS,TX1)
            EndIf
          EndIf
        EndIf
        oText:skip()
        sysrefresh()
      ENDDO
      SELE 31
      USE
      SELE 32
      USE
      Erase &NARQ4
      Erase &NARQB
      SELE 30
      oText:CLOSE()
      M->LOCAL2=TRIM("/"+CURDIR())
      M->LOCAL3=ALLTRIM(WEBLOCAL2)+"EMAIL.PDF"
      M->WLOCAL4=ALLTRIM(WEBLOCAL2)+ALLTRIM(VLABEL[OI])
      Erase &NARQ
      LCHDIR(M->LOCAL2)
      NARQ=WEBLOCAL+"EMAIL.TXT"
      Erase &NARQ
      EMAILRETORNO=WEBCAD
      If EMPTY(EMAILRETORNO)
        EMAILRETORNO=WEBMAIL
      EndIf
      If SEENVIAEMAIL="S"
        ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - EMAIL(Cadastro de Email) ","Recebido arquivo EMAIL",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
    return 

    Procedure FAZERCADASTRO(NARQ)
      M->WLOCAL4=ALLTRIM(WEBLOCAL2)+ALLTRIM(VLABEL[OI])
      LZCOPYFILE(NARQ,M->WLOCAL4)
      NWARQ=LOWER(NARQ)
      NWARQ1=LOWER(WEBLOCAL+"CADASTRO.PDF")
      If NWARQ#NWARQ1
        FRENAME(NARQ,NWARQ1)
      EndIf
      NARQ=NWARQ1
      FEZALGO=.T.
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Atualizando CADASTRO PARA emails SEM CODIGO "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      COMANDO=WCORRENTE+"\PDFTEXT.EXE -eol dos -layout "+NARQ+" "+WEBLOCAL+"EMAIL.TXT"
      WAITRUN(COMANDO,0)
      CURSORWAIT()
      oText:=TTxtfile():NEW(WEBLOCAL+"EMAIL.TXT")
      cursorwait()
      CONTAD=0
      CONTAD=oText:RecCount()
      If VALTYPE(CONTAD)#"N"
        CONTAD=1
      EndIf
      SELE 32
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 32
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTE","S")
        SELE 30
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      GO TOP
      cursorwait()
      NCOL=77
      ncol1=81
      DO WHILE !oText:EOF()
        cursorwait()
        TTT=oText:Readline()
        CODATU=VAL(LEFT(TTT,8))
        If AT("Nasc/Fund",TTT)#0
          NCOL=AT("Nasc/Fund",TTT)+1
          ncol1=at("CPF/CNPJ",TTT)-6
        EndIf
        If CODATU#0
          SELE 32
          SEEK CODATU
          If !FOUND()
            ADIREG(0)
            REPLACE CODIGO WITH CODATU
            REPLACE MAILING WITH .T.
            REPLACE EEMAIL WITH .T.
            REPLACE EIMPRE WITH .F.
            REPLACE EGUARD WITH .T.
          EndIf
          WNOME=SUBSTR(TTT,12,47)
          WNASC=CTOD(SUBSTR(TTT,NCOL,10))
          WCNPJ=SUBSTR(TTT,NCOL1,19)
          WCNPJ=Strtran(WCNPJ,".","")
          WCNPJ=Strtran(WCNPJ,"-","")
          WCNPJ=Strtran(WCNPJ,"/","")
          WCNPJ=VAL(Strtran(WCNPJ," ",""))
          If REGLOCK(10)
            REPLACE CODIGO WITH CODATU
            REPLACE NOME WITH WNOME
            REPLACE CNPJCPF WITH STRZERO(WCNPJ,14)
            If WNASC#CTOD("")
              REPLACE NASC WITH WNASC
            EndIf
            If EMPTY(INCLUIDO)
              REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERINC WITH "Arqv.Cadastro - Sinacor"
            EndIf
            If EMPTY(SENHA)
              REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
            EndIf
          EndIf
        EndIf
        oText:skip()
        sysrefresh()
      ENDDO
      SELE 32
      USE
      Erase &NARQ4
      SELE 30
      oText:CLOSE()
      M->LOCAL2=TRIM("/"+CURDIR())
      M->LOCAL3=ALLTRIM(WEBLOCAL2)+"CADASTRO.PDF "
      M->WLOCAL4=ALLTRIM(WEBLOCAL2)+ALLTRIM(VLABEL[OI])
      Erase &NARQ
      LCHDIR(M->LOCAL2)
      NARQ=WEBLOCAL+"CADASTRO.TXT"
      Erase &NARQ
      EMAILRETORNO=WEBCAD
      If EMPTY(EMAILRETORNO)
        EMAILRETORNO=WEBMAIL
      EndIf
      If SEENVIAEMAIL="S"
        ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - EMAIL(Cadastrol) ","Recebido arquivo CADASTRO",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
    return 

    Procedure FZMAILSEM()
      SELE 31
      GO TOP
      DO WHILE !EOF()
        SysRefresh()
        If LEFT(EMAIL,1)#"["
          DELE
        EndIf
        SKIP
      ENDDO
      PACK
      SELE 31
      USE
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      INDEX ON strzero(cliente,9)+LEFT(EMAIL,100) TO &NARQB
      SELE 32
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 32
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTE","S")
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      GO TOP
      DO WHILE !EOF()
        If REGLOCK(10)
          REPLACE TEMMAIL WITH .F.
        EndIf
        SKIP
      ENDDO
      INDEX ON strzero(val(CNPJCPF),14) TO &NARQ4
      GO TOP
      CODANT=0
      CODATU=0
      CURSORWAIT()
      novo=.f.
      M->NOME=""
      cursorwait()
      NCOL=35
      NCONTAD=0
      CONTAD=oText:RecCount()
      oText:skip((otext:recno()-1)*-1)
      mcga=0
      DO WHILE !oText:EOF()
        cursorwait()
        TTT=oText:Readline()
        NCONTAD=NCONTAD+1
        If AT("CPF/CNPJ",TTT)#0 .or. at(" CP F / CNP J",TTT)#0
          If AT("CPF/CNPJ",TTT)#0
            NCOL=AT("CPF/CNPJ",TTT)
          Else
            NCOL=AT(" CP F / CNP J",TTT)
          EndIf
          oText:Skip()
          loop
        EndIf
        MCNPJCPF=VAL(SUBSTR(TTT,NCOL,15))
        If MCNPJCPF=0
          oText:Skip()
          loop
        EndIf
        If MCNPJCPF#0
          mcga=mcnpjcpf
          SELE 32
          SEEK strzero(MCNPJCPF,14)
          If !FOUND() .or. mcnpjcpf<>val(cnpjcpf)
            oText:skip()
            loop
          EndIf
          CODATU=CODIGO
          If REGLOCK(10)
            REPLACE TEMMAIL WITH .T.
          EndIf
          If REGLOCK(10) .And. (EMPTY(SENHA))
            REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
          EndIf
          M->NOME=NOME
          CODANT=CODATU
          oText:skip()
          NCONTAD=NCONTAD+1
          TTT=oText:Readline()
        EndIf
        If CODATU=0
          CODATU=CODANT
        EndIf
        CODANT=CODATU
        DO WHILE !oText:eof()
          cursorwait()
          TTT=oText:Readline()
          NCONTAD=NCONTAD+1
          If AT("CPF/CNPJ",TTT)#0 .OR. AT(" CP F / CNP J",TTT)#0
            If AT(" CP F / CNP J",TTT)#0
              NCOL=AT(" CP F / CNP J",TTT)
            Else
              NCOL=AT("CPF/CNPJ",TTT)
            EndIf
            oText:Skip()
            EXIT
          EndIf
          MCNPJCPF=VAL(SUBSTR(TTT,NCOL,15))
          If MCNPJCPF#0
            EXIT
          EndIf
          m->EMAIL=ALLTRIM(TTT)
          M->EMAIL=Strtran(M->EMAIL," ","")
          M->EMAIL=Strtran(M->EMAIL,";/;",";")
          M->EMAIL=LOWER(M->EMAIL)
          If EMPTY(M->EMAIL) .or. at("@",M->EMAIL)=0
            oText:skip()
            LOOP
          EndIf
          SELE 31
          If CODATU>0
            SEEK STRZERO(CODATU,9)+LEFT(M->EMAIL,100)
            If !FOUND()
              ADIREG(0)
              REPLACE CLIENTE WITH CODATU
              REPLACE EMAIL WITH M->EMAIL
              REPLACE NOME WITH M->NOME
              LINHA=oText:Readline()
              REPLACE CODIGO WITH RECNO()
              REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERINC WITH "ARQ.CSV"
              REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERALT WITH "ARQ.CSV"
              DECLARE TX1[1]
              TX1[1]="Email Cadastrado : "+STRZERO(CLIENTE,8)+"-"+NOME+CHR(13)+CHR(10)+EMAIL+CHR(13)+CHR(10)
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
            EndIf
          EndIf
          oText:skip()
          sysrefresh()
        ENDDO
      ENDDO
      SELE 31
      USE
      SELE 32
      USE
      Erase &NARQ4
      Erase &NARQB
      SELE 30
      oText:CLOSE()
      M->LOCAL2=TRIM("/"+CURDIR())
      M->LOCAL3=ALLTRIM(WEBLOCAL2)+"EMAIL.PDF"
      M->WLOCAL4=ALLTRIM(WEBLOCAL2)+ALLTRIM(VLABEL[OI])
      Erase &NARQ
      LCHDIR(M->LOCAL2)
      NARQ=WEBLOCAL+"EMAIL.TXT"
      Erase &NARQ
      EMAILRETORNO=WEBCAD
      If EMPTY(EMAILRETORNO)
        EMAILRETORNO=WEBMAIL
      EndIf
      If SEENVIAEMAIL="S"
        ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - EMAIL(Cadastro de Email SEM CODIGO) ","Recebido arquivo EMAIL SEM CODIGO",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
    return 

    Procedure CALCCORREIO()
      If FOLHAS<4
        return  2.05
      ElseIf FOLHAS<9
        return  2.85
      ElseIf FOLHAS<18
        return  3.96
      ElseIf FOLHAS<28
        return  4.80
      ElseIf FOLHAS<38
        return  5.65
      ElseIf FOLHAS<48
        return  6.55
      ElseIf FOLHAS<58
        return  7.50
      ElseIf FOLHAS<68
        return  8.35
      ElseIf FOLHAS<78
        return  9.25
      ElseIf FOLHAS<88
        return  10.10
      ElseIf FOLHAS<92
        return  11.00
      Else
        return  95.93
      EndIf
    return  0

    Procedure ATUACADASTRO()
      DECLARE VLABEL[ADIR(WEBLOCAL+"*.CSV")]
      ADIR(WEBLOCAL+"*.CSV",VLABEL)
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      If LEN(VLABEL)#0
        SELE 35
        If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTE","S")
          SELE 30
          return 
        EndIf
        INDEX ON CODIGO TO &NARQ4
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Arq. CSV Clientes "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
          NARQ1=ALLTRIM(WEBLOCAL2)+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(NARQ,NARQ1)
          NARQ2=ALLTRIM(WEBLOCAL1)+"ATUACAD.DBF"
          Erase &NARQ2
          SELE 37
          mat_stru={}
          Aadd(mat_stru,{"LINHA","C",255,0})
          DBCREATE(NARQ2,mat_stru)
          USE &NARQ2
          APPEND FROM &NARQ1 SDF
          GO TOP
          DO WHILE !EOF()
            SysRefresh()
            WLI=ALLTRIM(LINHA)+"; ; ; ;"+SPACE(150)
            COLU=AT(";",WLI)
            ZODIGO=LEFT(WLI,COLU-1)
            WLI=SUBSTR(WLI,COLU+1,150)
            COLU=AT(";",WLI)
            WNOME=LEFT(WLI,COLU-1)
            WLI=SUBSTR(WLI,COLU+1,150)
            COLU=AT(";",WLI)
            WEMAIL=LEFT(WLI,COLU-1)
            WLI=SUBSTR(WLI,COLU+1,150)
            COLU=AT(";",WLI)
            WIMPRIME=LEFT(WLI,COLU-1)
            WLI=SUBSTR(WLI,COLU+1,150)
            COLU=AT(";",WLI)
            WMAILING=LEFT(WLI,COLU-1)
            If val(zodigo)=0
              skip
              loop
            EndIf
            If at("@",wemail)=0
              skip
              loop
            EndIf
            SELE 35
            SEEK VAL(ZODIGO)
            If val(zodigo)>0 .And. !found() .And. val(zodigo)<999999
              adireg(0)
              REPLACE CODIGO WITH VAL(ZODIGO)
              REPLACE NOME WITH WNOME
              SEEK VAL(ZODIGO)
            EndIf
            If FOUND() .And. VAL(ZODIGO)#0 .And. REGLOCK(10) .And. val(zodigo)<999999
              If UPPER(WEMAIL)="S"
                REPLACE EEMAIL WITH .T.
              Else
                REPLACE EEMAIL WITH .F.
              EndIf
              If UPPER(WIMPRIME)="S"
                REPLACE EIMPRE WITH .T.
              Else
                REPLACE EIMPRE WITH .F.
              EndIf
              If UPPER(WMAILING)="N"
                REPLACE MAILING WITH .F.
              Else
                REPLACE MAILING WITH .T.
              EndIf
              REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERALT WITH "Arqv.CSV Cadastro"
              DECLARE TX1[1]
              TX1[1]="CSV de Cadastro: "+STRZERO(VAL(ZODIGO),8)+" - Email: "+IIF(EEMAIL,"S","N")+" Impressão: "+IIF(EIMPRE,"S","N")+" Mailing: "+IIF(MAILING,"S","N")+" - "+WNOME
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
            EndIf
            SELE 37
            SKIP
          ENDDO
          SELE 37
          USE
          Erase &NARQ2
          Erase &NARQ
        NEXT OI
        SELE 35
        USE
        Erase &NARQ4
      EndIf
      SELE 30
    return 

    Procedure EXTRUNIF()
      DECLARE VLABEL[ADIR(WEBLOCAL+"*.TXT")]
      ADIR(WEBLOCAL+"*.TXT",VLABEL)
      FOR OI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
        oText:=TTxtfile():NEW(NARQ)
        CONTAD=oText:RecCount()
        If VALTYPE(CONTAD)<>"N"
          CONTAD=1
        EndIf
        linha=oText:Readline()
        oText:Close()
        If LEFT(LINHA,23)="01_____________________"
          DECLARE TX1[1]
          TX1[1]="BolsasNet - Extrato Unificado "+WEBCOR
          LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
          If SEENVIAEMAIL="S"
            EXTRUNIF1(1) //enviar emails
          Else
            EXTRUNIF1(2) // gerar arquivo
          EndIf
        EndIf
      NEXT OI
      SELE 30
    return 

    Procedure EXTRUNIF1(QTP)
      FEZALGO=.T.
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 30
        Erase &NARQ4
        return 
      EndIf
      SET INDEX TO &NARQ4
      SELE 31
      USE
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 35
        USE
        Erase &NARQ4
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQB
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 30
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",155,0})
      ARQTXTO=WEBLOCAL1+"EXTRATO.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      narq7=WEBLOCAL4+"ETIQUETA.KEY"
      Erase &NARQ7
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"EXTRATO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 37
        USE
        Erase &NARQ7
        SELE 32
        USE
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      APPEND FROM &NARQ SDF
      NARQ1=WEBLOCAL2+alltrim(VLABEL[OI])
      LZCOPYFILE(NARQ,NARQ1)
      Erase &NARQ
      GO TOP
      If QTP=1
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        ZARQ1=WEBLOCAL+"EXTRATO"+DTOS(DATE())+".PDF"
        Erase &ZARQ1
        oReg:Set("PDFFilename",ZARQ1)
        oReg:Close()
        PRINT oPrn NAME "Extrato Consolidado" TO TEMPDF1
        DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-10 OF oPrn //8
        DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-10 BOLD OF oPrn //8
        DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-7.5 OF oPrn  // 6
        DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-7.5 BOLD OF oprn // 6
        DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-10 BOLD OF oPrn //8
        DEFINE FONT oFont NAME "courier new negrito" SIZE 0,8.3 OF oPrn //8.3
        DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
        DEFINE BRUSH oCinza       COLOR nRGB(205,205,205)
        DEFINE BRUSH oCinzaEscuro COLOR nRGB(099,099,099)
        DEFINE BRUSH oAzul COLOR nRGB(35,35,142)
        DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
        DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      EndIf
      If QTP=2
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        ZARQ1=weblocal+"EXTRATO"+DTOS(DATE())+".PDF"
        Erase &ZARQ1
        oReg:Set("PDFFilename",ZARQ1)
        oReg:Close()
        PRINT oPrn NAME "Extrato Consolidado" TO TEMPDF1
        DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-10 BOLD OF oPrn //8 NBOLD
        DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-10 BOLD OF oPrn //8
        DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-7.5 BOLD OF oPrn //6 NBOLD
        DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-7.5 BOLD OF oprn //6
        DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-10 BOLD OF oPrn //8
        DEFINE FONT oFont NAME "courier new negrito" SIZE 0,8.3 BOLD OF oPrn //8.3 NBOLD
        DEFINE FONT oFontv NAME "verdana" SIZE 0,12 OF oPrn // NBOLD
        DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
        DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
        DEFINE BRUSH oCinza       COLOR nRGB(205,205,205)
        DEFINE BRUSH oCinzaEscuro COLOR nRGB(0,0,0) // 168,168,168)
        DEFINE BRUSH oAzul COLOR nRGB(0,0,0) // 35,35,142)
        DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
        DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
        nRowStep = oPrn:nVertRes() / 66
        nColStep = oPrn:nHorzRes() / 80
      EndIf
      DO WHILE !EOF()
        SKIP 2
        CLIE=VAL(SUBSTR(LINHA,33,30))
        CLIE1=ALLTRIM(STR(CLIE,10,0))+"  "
        WLINHA=SUBSTR(LINHA,33,150)
        WWNOME=SUBSTR(WLINHA,AT(CLIE1,WLINHA)+LEN(CLIE1),60)
        SKIP
        wende1=substr(linha,33,100)
        SKIP
        wwende2=alltrim(substr(linha,33,100))
        wende2=right(wwende2,9)+"  "+left(wwende2,len(wwende2)-9)
        SKIP
        wemail=substr(linha,33,100)
        SKIP
        If LEFT(LINHA,4)="01__"
          LOOP
        EndIf
        SELE 35
        SEEK CLIE
        If !FOUND()
          ADIREG(0)
          REPLACE CODIGO WITH CLIE
          REPLACE NOME WITH WWNOME
          REPLACE ENDE1 WITH WENDE1
          REPLACE ENDE2 WITH WENDE2
          REPLACE ULTCONTA WITH DATE()
          If !EMPTY(WEMAIL)
            REPLACE EEMAIL WITH .T.
            REPLACE EIMPRE WITH .F.
            REPLACE TEMMAIL WITH .T.
            REPLACE MAILING WITH .T.
          Else
            REPLACE EIMPRE WITH .T.
          EndIf
          REPLACE EGUARD WITH .T.
          REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
          REPLACE USERINC WITH "Relat.Extrato Unificado"
        Else
          If REGLOCK(10)
            REPLACE CODIGO WITH CLIE
            REPLACE NOME WITH WWNOME
            REPLACE ULTCONTA WITH DATE()
            REPLACE ENDE1 WITH WENDE1
            REPLACE ENDE2 WITH ALLTRIM(WENDE2)+" "+UF
            REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
            REPLACE USERALT WITH "Relat.Extrato Unificado"
          EndIf
        EndIf
        WENDE2=ALLTRIM(WENDE2)+" "+UF
        WASSESSOR=STRZERO(ASSESSOR,5)
        If !EMPTY(WEMAIL)
          SELE 31
          SEEK CLIE
          MGRAVEM=.T.
          DO WHILE CLIENTE=CLIE .And. !EOF()
            If TRIM(LOWER(EMAIL))=TRIM(LOWER(WEMAIL))
              MGRAVEM=.F.
            EndIf
            SKIP
          ENDDO
          If MGRAVEM
            ADIREG(0)
            REPLACE CODIGO WITH RECNO()
            REPLACE BMFBOV WITH "B"
            REPLACE EMAIL WITH LOWER(WEMAIL)
            REPLACE NOME WITH WWNOME
            REPLACE CLIENTE WITH CLIE
          EndIf
        EndIf
        SELE 36
        CX=0
        CABEC=0
        WANT="02"
        WTOT=0
        contfol=1
        SELE 35
        SEEK CLIE
        If NOME="BB BANCO" .And. reglock(10)
          REPLACE FRENTEV WITH .T.
        EndIf
        SELE 36
        contlin=0
        DO WHILE !EOF() .And. LEFT(LINHA,4)#"01__"
          If CONTLIN>25
            CABEC=0
            contlin=0
            contfol=contfol+1
            oprn:cmsay(26.5,18.0,"Continua...",oFontV8B,,nRGB(0,0,0),,0) //Data de Emissão
            SELE 36
            ENDPAGE
          EndIf
          If CABEC=0
            PAGE
            If QTP=1
              saybitmap(0.0,0.0,3.26*6.5,1.3*5.0,WEBLOCAL+"EXTRATO.fig",oPrn)
            ElseIf QTP=2
              saybitmap(0.0,0.0,3.26*6.5,1.3*5.0,WEBLOCAL+"EXTR1.fig",oPrn)
            EndIf

            caja(6.5,0.97,6.51,20.1,oPrn,0,oPreto,oPen1) //linha cabecalho
            caja(6.5,0.97,29,0.96,oPrn,0,oPreto,oPen1) //linha esquerda
            caja(6.5,20.10,29,20.11,oPrn,0,oPreto,oPen1) //linha direita
            caja(28.99,0.97,29,20.1,oPrn,0,oPreto,oPen1) //linha rodapé

            caja(6.7,1.1,7.2,19.9,oPrn,1,oCinzaEscuro,oPen1) //caixa cinza escuro

            oprn:cmsay(5.8,1.1,"Cliente: "+ALLTRIM(TRANSFORM(CLIE,"999999999")),oFontV8,,nRGB(255,255,255),,0) //Código do Cliente
            oprn:cmsay(5.8,7.0,"Assessor - "+WASSESSOR,oFontV8,,nRGB(255,255,255),,0) //Assessor


            oprn:cmsay(6.8,1.1,"Cliente: "+ALLTRIM(TRANSFORM(CLIE,"999999999"))+"  "+WWNOME,oFontV8,,nRGB(255,255,255),,0) //Código do Cliente
            oprn:cmsay(6.8,16.3,"Assessor - "+WASSESSOR,oFontV8,,nRGB(255,255,255),,0) //Assessor
            oprn:cmsay(7.5,12.5,"   Emissão: "+DTOC(M->DAT_HOJE)+"     Folha: "+strzero(Contfol,3),oFontV8B,,nRGB(0,0,0),,0) //Data de Emissão
            CONTLIN=8.0

            caja(29.2,1.8,29.8,7.5,oPrn,1,oAzul,oPen1) //caixa Cinza
            oprn:cmsay(29.3,1.9,"Ouvidoria: 0800 724 30 20",oFontV8b,,nRGB(255,255,255),,0) //Titulo
            caja(27.00,1.4,27.01,19.50,oPrn,0,oPreto,oPen1) //linha Cabeçalho Mensagem
            caja(27.00,1.4,28.5,1.41,oPrn,0,oPreto,oPen1) //linha esquerda Mensagem
            caja(27.00,19.50,28.5,19.51,oPrn,0,oPreto,oPen1) //linha Direita Mensagem
            caja(28.49,1.4,28.5,19.5,oPrn,0,oPreto,oPen1) //linha Mensagem rodapé
            If QTP=1
              oprn:cmsay(27.1,1.99,WMENS1,oFont,,nRGB(35,35,142),,0)
              oprn:cmsay(27.3,1.99,WMENS2,oFont,,nRGB(35,35,142),,0)
              oprn:cmsay(27.5,1.99,WMENS3,oFont,,nRGB(35,35,142),,0)
              oprn:cmsay(27.7,1.99,WMENS4,oFont,,nRGB(35,35,142),,0)
              oprn:cmsay(27.9,1.99,WMENS5,oFont,,nRGB(35,35,142),,0)
              oprn:cmsay(28.1,1.99,WMENS6,oFont,,nRgb(35,35,142),,0)
            ElseIf QTP=2
              oprn:cmsay(27.1,1.99,WMENS1,oFont,,nRGB(0,0,0),,0)
              oprn:cmsay(27.3,1.99,WMENS2,oFont,,nRGB(0,0,0),,0)
              oprn:cmsay(27.5,1.99,WMENS3,oFont,,nRGB(0,0,0),,0)
              oprn:cmsay(27.7,1.99,WMENS4,oFont,,nRGB(0,0,0),,0)
              oprn:cmsay(27.9,1.99,WMENS5,oFont,,nRGB(0,0,0),,0)
              oprn:cmsay(28.1,1.99,WMENS6,oFont,,nRgb(0,0,0),,0)
            EndIf
            CABEC=1
          EndIf
          SELE 36
          If LINHA="02______"
            SKIP
            If AT("Resumo do Extrato de Conta Corrente",LINHA)=0
              mesref=RIGHT(trim(linha),7)
              caja(CONTLIN,1.1,CONTLIN+0.6,8.5,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"Extrato C/C - Mês "+mesref,oFontV8B,,nRGB(255,255,255),,0) //Titulo
              SKIP 2
              CONTLIN=CONTLIN+0.57
              caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
              If AT("Laçamento",LINHA)#0
                REPLACE LINHA WITH "02 Data de Ref.   Descrição do Lançamento                                                       Débito                 Crédito                   Saldo"
              EndIf
              oprn:cmsay(CONTLIN+0.23,1.0,SUBSTR(LINHA,3,69)+SUBSTR(LINHA,90,18)+SUBSTR(LINHA,114,18)+SUBSTR(LINHA,138,18),oFontC8B,,nRGB(0,0,0),,0) //Titulo do Movimento
              CONTLIN=CONTLIN+0.68
              caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
              CONTLIN=CONTLIN+0.2
              SKIP
              LOOP
            Else
              CONTLIN=CONTLIN+0.2
              caja(CONTLIN,1.1,CONTLIN+0.6,8.5,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"Extrato C/C - Resumo ",oFontV8B,,nRGB(255,255,255),,0) //Titulo
              CONTLIN=CONTLIN+0.57
              caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
              SKIP
              oPrn:cmsay(CONTLIN+0.23,1.9,"      Saldo Anterior         Saldo Total    Saldo Disponível     Total de Débito    Total de Crédito",OfONTC8B,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.50
              caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
              CONTLIN=CONTLIN+0.2
              WLINHA=LINHA
              SKIP
              oPrn:cmsay(CONTLIN,1.9,SUBSTR(WLINHA,27,20)+SUBSTR(WLINHA,78,20)+SUBSTR(WLINHA,134,20)+SUBSTR(LINHA,38,20)+SUBSTR(LINHA,94,20),oFontC8,,nRGB(0,0,0),,0) //Titulo do Movimento
              CONTLIN=CONTLIN+0.57
              skip
              want=left(linha,2)
              wtot=0
              LOOP
            EndIf
          EndIf

          If LINHA="03______"
            SKIP
            mesref=RIGHT(trim(linha),10)
            caja(CONTLIN,1.1,CONTLIN+0.6,8.5,oPrn,1,oAzul,oPen1) //caixa azul
            oprn:cmsay(CONTLIN+0.1,1.3,"Posição de Ações em "+mesref,oFontV8B,,nRGB(255,255,255),,0) //Titulo
            SKIP 2
            CONTLIN=CONTLIN+0.57
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            oprn:cmsay(CONTLIN+0.23,1.0,SUBSTR(LINHA,3,16)+"      "+SUBSTR(LINHA,32,16)+"      "+SUBSTR(LINHA,62,9)+"      "+SUBSTR(LINHA,89,17)+"      "+SUBSTR(LINHA,116,11)+"          "+SUBSTR(LINHA,136,16),oFontC8B,,nRGB(0,0,0),,0) //Titulo do Movimento
            CONTLIN=CONTLIN+0.68
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            CONTLIN=CONTLIN+0.2
            SKIP
            want=left(linha,2)
            WTOT=0
            LOOP
          EndIf
          If LINHA="04______"
            SKIP
            mesref=RIGHT(trim(linha),10)
            caja(CONTLIN,1.1,CONTLIN+0.6,8.5,oPrn,1,oAzul,oPen1) //caixa azul
            oprn:cmsay(CONTLIN+0.1,1.3,"Posição na BM&F em "+mesref,oFontV8B,,nRGB(255,255,255),,0) //Titulo
            SKIP 2
            CONTLIN=CONTLIN+0.57
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            oprn:cmsay(CONTLIN+0.23,1.0,SUBSTR(LINHA,3,54)+"  "+SUBSTR(LINHA,71,16)+"  "+SUBSTR(LINHA,101,19)+"       "+SUBSTR(LINHA,133,18),oFontC8B,,nRGB(0,0,0),,0) //Titulo do Movimento
            CONTLIN=CONTLIN+0.68
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            CONTLIN=CONTLIN+0.2
            SKIP
            want=left(linha,2)
            WTOT=0
            LOOP
          EndIf
          If LINHA="05_____"
            SKIP
            mesref=RIGHT(trim(linha),10)
            caja(CONTLIN,1.1,CONTLIN+0.6,8.5,oPrn,1,oAzul,oPen1) //caixa azul
            oprn:cmsay(CONTLIN+0.1,1.3,"Posição de Fundos em "+mesref,oFontV8B,,nRGB(255,255,255),,0) //Titulo
            SKIP 2
            CONTLIN=CONTLIN+0.57
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            oprn:cmsay(CONTLIN+0.23,1.0,SUBSTR(LINHA,3,40)+SUBSTR(LINHA,47,20)+SUBSTR(LINHA,72,15)+SUBSTR(LINHA,95,14)+SUBSTR(LINHA,118,11)+"  "+SUBSTR(LINHA,135,16),oFontC8B,,nRGB(0,0,0),,0) //Titulo do Movimento
            CONTLIN=CONTLIN+0.68
            caja(CONTLIN,1.0,CONTLIN+0.004,20.0,oPrn,0,oPreto,oPen1) //linha cabecalho  movimento
            CONTLIN=CONTLIN+0.2
            want=left(linha,2)
            WTOT=0
            SKIP
            LOOP
          EndIf
          If CX=0
            caja(CONTLIN,1.0,CONTLIN+0.4,20.0,oPrn,1,oBranco,oPen1) //caixa branca
            CX=1
          Else
            caja(CONTLIN,1.0,CONTLIN+0.4,20.0,oPrn,1,oCinza,oPen1) //caixa cinza
            CX=0
          EndIf

          If LEFT(linha,2)="02"
            oprn:cmsay(CONTLIN+0.1,1.0,SUBSTR(LINHA,3,72)+SUBSTR(LINHA,88,15)+SUBSTR(LINHA,112,15)+SUBSTR(LINHA,135,16),oFontC8B,,nRGB(0,0,0),,0)
          ElseIf LEFT(LINHA,2)="03"
            oprn:cmsay(CONTLIN+0.1,1.0,SUBSTR(LINHA,3,16)+"      "+SUBSTR(LINHA,32,16)+"      "+SUBSTR(LINHA,62,9)+"      "+SUBSTR(LINHA,89,17)+"      "+SUBSTR(LINHA,116,11)+"          "+SUBSTR(LINHA,136,16),oFontC8B,,nRGB(0,0,0),,0)
          ElseIf LEFT(LINHA,2)="04"
            oprn:cmsay(CONTLIN+0.1,1.0,SUBSTR(LINHA,3,54)+"  "+SUBSTR(LINHA,71,16)+"  "+SUBSTR(LINHA,101,19)+"       "+SUBSTR(LINHA,133,18),oFontC8B,,nRGB(0,0,0),,0)
          ElseIf LEFT(LINHA,2)="05"
            oprn:cmsay(CONTLIN+0.1,1.0,SUBSTR(LINHA,3,40)+SUBSTR(LINHA,47,20)+SUBSTR(LINHA,72,15)+SUBSTR(LINHA,95,14)+SUBSTR(LINHA,118,11)+"  "+SUBSTR(LINHA,135,16),oFontC8B,,nRGB(0,0,0),,0) //Titulo do Movimento
          Else
            oPrn:cmsay(CONTLIN+0.1,1.0,SUBSTR(LINHA,3,148),oFontC8,,nRGB(0,0,0),,0)
          EndIf

          CONTLIN=CONTLIN+0.5
          WVALOR=RIGHT(LINHA,20)
          WVALOR=Strtran(WVALOR,".","")
          WVALOR=VAL(Strtran(WVALOR,",","."))
          WTOT=WTOT+WVALOR
          SKIP
          If EOF() .OR. WANT#LEFT(LINHA,2)
            If WANT#"02" .And. WTOT#0
              caja(CONTLIN,13.5,CONTLIN+0.5,19.9,oPrn,1,oCinzaEscuro,oPen1) //caixa Cinza
              oPrn:cmsay(CONTLIN+0.1,13.7,"   Total:  "+TRANSFORM(WTOT,"@E 999,999,999,999.99"),oFontt8,,nRGB(255,255,255),,0) //Titulo
              CONTLIN=CONTLIN+0.5
            EndIf
            WANT=LEFT(LINHA,2)
            WTOT=0
          EndIf
          If LEFT(LINHA,15)#"01_____________" .And. !EOF()
            If CONTLIN>25
              CABEC=0
              contfol=contfol+1
              oprn:cmsay(26.5,17.0,"Continua...",oFontV8B,,nRGB(0,0,0),,0) //Data de Emissão
              contlin=0
              SELE 36
              ENDPAGE
            EndIf
          EndIf
        ENDDO
        ENDPAGE
        SELE 36
      ENDDO
      ENDPRINT
      DO WHILE .T.
        CURSORWAIT()
        If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
          fClose(nHandle)
          EXIT
        Else
          fClose(nHandle)
        EndIf
      ENDDO
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      Erase &NARQ4
      SELE 31
      USE
      Erase &NARQB
    return 

    Function Code128(nRow,nCol,cCode,oPrint,cMode,Color,lHorz,nWidth,nHeigth)
      aCode :={"212222",;
        "222122",;
        "222221",;
        "121223",;
        "121322",;
        "131222",;
        "122213",;
        "122312",;
        "132212",;
        "221213",;
        "221312",;
        "231212",;
        "112232",;
        "122132",;
        "122231",;
        "113222",;
        "123122",;
        "123221",;
        "223211",;
        "221132",;
        "221231",;
        "213212",;
        "223112",;
        "312131",;
        "311222",;
        "321122",;
        "321221",;
        "312212",;
        "322112",;
        "322211",;
        "212123",;
        "212321",;
        "232121",;
        "111323",;
        "131123",;
        "131321",;
        "112313",;
        "132113",;
        "132311",;
        "211313",;
        "231113",;
        "231311",;
        "112133",;
        "112331",;
        "132131",;
        "113123",;
        "113321",;
        "133121",;
        "313121",;
        "211331",;
        "231131",;
        "213113",;
        "213311",;
        "213131",;
        "311123",;
        "311321",;
        "331121",;
        "312113",;
        "312311",;
        "332111",;
        "314111",;
        "221411",;
        "431111",;
        "111224",;
        "111422",;
        "121124",;
        "121421",;
        "141122",;
        "141221",;
        "112214",;
        "112412",;
        "122114",;
        "122411",;
        "142112",;
        "142211",;
        "241211",;
        "221114",;
        "213111",;
        "241112",;
        "134111",;
        "111242",;
        "121142",;
        "121241",;
        "114212",;
        "124112",;
        "124211",;
        "411212",;
        "421112",;
        "421211",;
        "212141",;
        "214121",;
        "412121",;
        "111143",;
        "111341",;
        "131141",;
        "114113",;
        "114311",;
        "411113",;
        "411311",;
        "113141",;
        "114131",;
        "311141",;
        "411131",;
        "211412",;
        "211214",;
        "211232",;
        "2331112"}

      go_code(_code128(cCode,cMode),nRow,nCol,oPrint,lHorz,Color,nWidth,nHeigth)
    return  nil

    Function _code128(cCode,cMode)

      Local nSum:=0, cBarra, cCar
      Local cTemp, n, nCAr, nCount:=0
      Local lCodeC := .f. ,lCodeA:= .f.

      // control de errores
      If valtype(cCode) !='C'
        alert('Barcode c128 required a Character value. ')
        return  nil
      end
      If !empty(cMode)
        If valtype(cMode)='C' .And. Upper(cMode)$'ABC'
          cMode := Upper(cMode)
        Else
          alert('Code 128 Modes are A,B o C. Character values.')
        end
      end
      If empty(cMode) // modo variable
        // an lisis de tipo  de c¢digo...
        If str(val(cCode),len(cCode))=cCode // s¢lo n£meros
          lCodeC := .t.
          cTemp:=aCode[106]
          nSum := 105
        Else
          for n:=1 to len(cCode)
          nCount+=if(substr(cCode,n,1)>31,1,0) // no cars. de control
        end
          If nCount < len(cCode) /2
            lCodeA := .t.
            cTemp := aCode[104]
            nSum := 103
          Else
            cTemp := aCode[105]
            nSum := 104
          end
        end
      Else
        If cMode =='C'
          lCodeC := .t.
          cTemp:=aCode[106]
          nSum := 105
        ElseIf cMode =='A'
          lCodeA := .t.
          cTemp := aCode[104]
          nSum := 103
        Else
          cTemp := aCode[105]
          nSum := 104
        end
      end

      nCount := 0 // caracter registrado
      for n:= 1 to len(cCode)
        nCount ++
        cCar := substr(cCode,n,1)
        If lCodeC
          If len(cCode)=n  // ultimo caracter
            CTemp += aCode[101] // SHIFT Code B
            nCar := asc(cCar)-31
          Else
            nCar := Val(substr(cCode,n,2))+1
            n++
          end
        ElseIf lCodeA
          If cCar> '_' // Shift Code B
            cTemp += aCode[101]
            nCar := asc(cCar)-31
          ElseIf cCar <= ' '
            nCar := asc(cCar)+64
          Else
            nCar := asc(cCar)-31
          EndIf
        Else // code B standard
          If cCar <= ' ' // shift code A
            cTemp += aCode[102]
            nCar := asc(cCar)+64
          Else
            nCar := asc(cCar)-31
          end
        EndIf
        nSum += (nCar-1)*nCount
        cTemp := cTemp +aCode[nCar]
      next
      nSum := nSum%103 +1
      cTemp := cTemp + aCode[ nSum ] +aCode[107]
      cBarra := ''
      for n:=1 to len(cTemp) step 2
        cBarra+=replicate('1',val(substr(cTemp,n,1)))
        cBarra+=replicate('0',val(substr(cTemp,n+1,1)))
      next
    return  cBarra


    Function go_code( cBarra, nx,ny,oPrint,lHoRz, nColor, nWidth, nLen)
      Local n, oBr
      If empty(nColor)
        nColor := CLR_BLACK
      end
      default lHorz := .t.

      default nWidth := 0.025 // 1/3 M/mm

      default nLen := 1.5 // Cmm.

      define brush oBr color nColor
      //    Width of Bar
      If !lHorz
        nWidth :=round ( nWidth * 10 * oPrint:nVertRes() / oPrint:nVertSize() ,0 )
      Else
        nWidth :=round  ( nWidth * 10 * oPrint:nHorzRes() / oPrint:nHorzSize(), 0 )
      end
      // Len of bar
      If lHorz
        nLen :=round ( nLen * 10 * oPrint:nVertRes() / oPrint:nVertSize() ,0 )
      Else
        nLen :=round  ( nLen * 10 * oPrint:nHorzRes() / oPrint:nHorzSize(), 0 )
      end

      for n:=1 to len(cBarra)
        If substr(cBarra,n,1) ='1'
          If lHorz
            oPrint:fillRect({nx,ny,nx+nLen,(ny+=nWidth)},oBr)
          Else
            oPrint:fillRect({nx,ny,(nx+=nWidth),ny+nLen},oBr)
          end
        Else
          If lHorz
            ny+=nWidth
          Else
            nx += nWidth
          end
        end
      next

      oBr:end()
    return  nil

    Function ORDLIST()
      If ORDLIST=0
        ORDLIST=1
        return  1
      Else
        ORDLIST=0
        return  2
      EndIf
    return  1

    Procedure CRIAMOVFICHA()
      SELE 34
      ARQ4=ALLTRIM(BOLSANET->PASTA)+"FICHA.KEY"
      Erase &ARQ4
      If !USEREDE(ALLTRIM(BOLSANET->PASTA)+"FICHA.DBF",.F.,10,"FICHA")
        return 
      EndIf
      If FCOUNT()<188
        USE
        DO WHILE .T.
          M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
          ARQ_PRN=alltrim(BOLSANET->PASTA)+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)
          ARQ_PRN1:=ARQ_PRN+".DAT"
          If !FILE(ARQ_PRN1)
            EXIT
          EndIf
        ENDDO
        aStructure:={}
        AaDD( aStructure, { "ORDELET","C",1,0})
        aAdd( aStructure, { "LOGIN","C",50,0})
        aAdd( aStructure, { "SENHA","C",20,0})
        aAdd( astructure, { "CODCLI","N",8,0})
        aAdd( aStructure, { "APROVADO","L",1,0}) // APROVADO
        aAdd( aStructure, { "DTAPROV","D",8,0}) // DATA DA GERAÇÃO DO ARQUIVO
        aAdd( aStructure, { "COBRADO10","D",8,0}) // DATA DA COBRANÇA DOS 10 DIAS EM 10 DIAS
        aAdd( aStructure, { "ENVPCLI","D",8,0}) // Enviado para o Cliente o arquivo de contratos se zerar o Worker enviará novamente
        aAdd( aStructure, { "DIGITO"    , "N",   1, 0 } )
        aAdd( aStructure, { "CONTRATOS","C",254,0}) // nr. dos campos de contratos
        aAdd( aStructure, { "GERAR","L",1,0}) // True = Gerar Ficha
        aAdd( aStructure, { "GERADO","C",20,0}) // data e hora da geração dos contratos
        aAdd( aStructure, { "CODIGO"    , "N",  10, 0 } )
        aAdd( aStructure, { "PESSOA"    , "C",   1, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "SEXO"      , "C",   1, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C",  50, 0 } )
        aAdd( aStructure, { "CNPJCPF"   , "C",  18, 0 } )
        aAdd( aStructure, { "NACIONAL"  , "C",  30, 0 } )
        aAdd( aStructure, { "NATURALA"  , "C",  30, 0 } )
        aAdd( aStructure, { "DOC"       , "C",  16, 0 } )
        aAdd( aStructure, { "TIPODOC"   , "C",   4, 0 } )
        aAdd( aStructure, { "UFDOC"     , "C",   2, 0 } )
        aAdd( aStructure, { "ORGDOC"    , "C",   4, 0 } )
        aAdd( aStructure, { "DTDOC"     , "C",  10, 0 } )
        aAdd( aStructure, { "DTNASC"    , "C",  10, 0 } )
        aAdd( aStructure, { "MAE"       , "C",  60, 0 } )
        aAdd( aStructure, { "PAI"       , "C",  60, 0 } )
        aAdd( aStructure, { "ESTCIVIL"  , "C",  50, 0 } )
        aAdd( aStructure, { "CONJUGE"   , "C",  60, 0 } )
        aAdd( aStructure, { "REGCASA"   , "N",   1, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "N",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "PAIS"      , "C",  60, 0 } )
        aAdd( aStructure, { "DDD"       , "N",   2, 0 } )
        aAdd( aStructure, { "TEL"       , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFAX"    , "N",   2, 0 } )
        aAdd( aStructure, { "FAX"       , "C",   9, 0 } )
        aAdd( aStructure, { "DDDCEL"    , "N",   2, 0 } )
        aAdd( aStructure, { "CELULAR"   , "C",   9, 0 } )
        aAdd( aStructure, { "PROFISSAO" , "C",  80, 0 } )
        aAdd( aStructure, { "EMPRESA"   , "C",  60, 0 } )
        aAdd( aStructure, { "CEPEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "ENDEMP"    , "C",  30, 0 } )
        aAdd( aStructure, { "NREMP"     , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLEMP"  , "C",  10, 0 } )
        aAdd( aStructure, { "BAIREMP"   , "C",  18, 0 } )
        aAdd( aStructure, { "CIDEMP"    , "C",  28, 0 } )
        aAdd( aStructure, { "UFEMP"     , "C",   2, 0 } )
        aAdd( aStructure, { "PAISEMP"   , "C",  60, 0 } )
        aAdd( aStructure, { "DDDEMP"    , "N",   2, 0 } )
        aAdd( aStructure, { "TELEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFAXE"   , "N",   2, 0 } )
        aAdd( aStructure, { "FAXEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "RAMALEMP"  , "N",   5, 0 } )
        aAdd( aStructure, { "CORRESP"   , "C",  11, 0 } )
        aAdd( aStructure, { "CEPCOR"    , "C",   9, 0 } )
        aAdd( aStructure, { "ENDCOR"    , "C",  30, 0 } )
        aAdd( aStructure, { "NRCOR"     , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLCOR"  , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRCOR"   , "C",  18, 0 } )
        aAdd( aStructure, { "CIDCOR"    , "C",  28, 0 } )
        aAdd( aStructure, { "UFCOR"     , "C",   2, 0 } )
        aAdd( aStructure, { "PAISCOR"   , "C",  60, 0 } )
        aAdd( aStructure, { "BANCO"     , "C",  60, 0 } )
        aAdd( aStructure, { "AGENCIA"   , "N",   5, 0 } )
        aAdd( aStructure, { "AGDIG"     , "C",   2, 0 } )
        aAdd( aStructure, { "CONTA"     , "C",  13, 0 } )
        aAdd( aStructure, { "DIGCONTA"  , "C",   2, 0 } )
        aAdd( aStructure, { "CONTAINV"  , "C",  13, 0 } )
        aAdd( aStructure, { "DIGCCIN"   , "C",   2, 0 } )
        aAdd( aStructure, { "CONJUNTA"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCONJT"  , "C",  16, 0 } )
        aAdd( aStructure, { "IMOVEIS"   , "N",  19, 2 } )
        aAdd( aStructure, { "MOVEIS"    , "N",  19, 2 } )
        aAdd( aStructure, { "SALDO"     , "N",  19, 2 } )
        aAdd( aStructure, { "SALARIO"   , "N",  19, 2 } )
        aAdd( aStructure, { "OUTRREND"  , "N",  19, 2 } )
        aAdd( aStructure, { "NIRE"      , "C",  20, 0 } )
        aAdd( aStructure, { "FORCONST"  , "C", 100, 0 } )
        aAdd( aStructure, { "CETIP"     , "C",  20, 0 } )
        aAdd( aStructure, { "NOMEEO1"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO1"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMEEO2"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO2"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMEEO3"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO3"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMECON1"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON1"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON1"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON1"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON2"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON2"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON2"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON2"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON3"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON3"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON3"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON3"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON4"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON4"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON4"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON4"   , "C",  50, 0 } )
        aAdd( aStructure, { "VINCULO"   , "C",   1, 0 } )
        aAdd( aStructure, { "CNTPRO"    , "C",   1, 0 } )
        aAdd( aStructure, { "PEXPOSTA"  , "C",   1, 0 } )
        aAdd( aStructure, { "PEMOTIVO"  , "C", 200, 0 } )
        aAdd( astructure, { "REXPOSTA"  , "C",   1, 0 } )
        aAdd( aStructure, { "REMOTIVO"  , "C", 200, 0 } )
        aAdd( aStructure, { "procurador", "C",   1, 0 } )
        aAdd( aStructure, { "ORDEM"     , "C",  21, 0 } )
        aAdd( aStructure, { "QUEM"      , "C",  90, 0 } )
        aAdd( aStructure, { "CTITULAR"  , "C",   1, 0 } )
        aAdd( astructure, { "IDTIT"     , "N",  10, 0 } )
        aAdd( astructure, { "IDCTIT"    , "N",  10, 0 } )
        aAdd( aStructure, { "TIPOCOR"   , "C",   1, 0 } )
        aAdd( aStructure, { "CARTORIO"  , "C",  50, 0 } )
        aAdd( aStructure, { "TPORDEM"   , "C",   6, 0 } )
        aAdd( aStructure, { "NOMEPROC"  , "C",  60, 0 } )
        aAdd( aStructure, { "CNPJCPFPRO", "N",  14, 0 } )
        aAdd( aStructure, { "DOCPROC"   , "C",  20, 0 } )
        aAdd( aStructure, { "DTNASCPROC", "C",  10, 0 } )
        aAdd( aStructure, { "EMAILPROC" , "C",  50, 0 } )
        aAdd( aStructure, { "CEPPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "ENDPROC"   , "C",  30, 0 } )
        aAdd( aStructure, { "NRPROC"    , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLPROC" , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRPROC"  , "C",  18, 0 } )
        aAdd( astructure, { "CIDPROC"   , "C",  28, 0 } )
        aAdd( aStructure, { "UFPROC"    , "C",   2, 0 } )
        aAdd( aStructure, { "PAISPROC"  , "C",  60, 0 } )
        aAdd( astructure, { "DDDPROC"   , "N",   2, 0 } )
        aAdd( aStructure, { "TELPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFPROC"  , "N",   2, 0 } )
        aAdd( aStructure, { "FAXPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "IMOVESP1"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND1"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF1"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR1"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP2"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND2"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF2"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR2"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP3"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND3"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF3"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR3"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP4"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND4"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF4"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR4"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP5"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND5"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF5"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR5"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP1"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC1"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR1"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP2"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC2"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR2"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP3"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC3"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR3"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP4"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC4"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR4"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP5"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC5"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR5"  , "N",  19, 2 } )
        aAdd( aStructure, { "CODDEPEND" , "N",   1, 0 } )
        aAdd( astructure, { "INVESTIDOR", "N",   4, 0 } )
        aAdd( astructure, { "ASSESSOR"  , "N",   5, 0 } )
        aAdd( aStructure, { "CCONH"     , "C",  50, 0 } )
        aAdd( aStructure, { "QUALOUTRO" , "C",  20, 0 } )
        aAdd( aStructure, { "NEWS"      , "C",   1, 0 } )
        aAdd( aStructure, { "LIVRE"     , "C",  50, 0 } )
        aAdd( aStructure, { "IP"        , "C",  16, 0 } )
        aAdd( aStructure, { "ORI1"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI2"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI3"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI4"      , "C", 120, 0 } )
        aAdd( aStructure, { "IPALT"     , "C",  16, 0 } )
        dbCreate(ARQ_PRN1,aStructure,"DBFNTX")
        RELEASE aStructure
        NARQ=ALLTRIM(BOLSANET->PASTA)+"FICHA.DBF"
        FAZER=.T.
        SELE 34
        If !USEREDE(ARQ_PRN1,.T.,10,"TMP")
          MSGINFO("Não foi alterar a estrutura do arquivo "+NARQ+"!","")
          QUIT
        EndIf
        APPEND FROM &NARQ
        USE
        Erase &NARQ
        FRENAME(ARQ_PRN1,NARQ)
      EndIf
      SELE 34
      USE
    return 

    Procedure CRIARQFICHA()
      ARQ=ALLTRIM(BOLSANET->PASTA)+"FICHA.DBF"
      If !FILE(ARQ)
        aStructure:={}
        AaDD( aStructure, { "ORDELET","C",1,0})
        aAdd( aStructure, { "LOGIN","C",50,0})
        aAdd( aStructure, { "SENHA","C",20,0})
        aAdd( astructure, { "CODCLI","N",8,0})
        aAdd( aStructure, { "APROVADO","L",1,0}) // APROVADO
        aAdd( aStructure, { "DTAPROV","D",8,0}) // DATA DA GERAÇÃO DO ARQUIVO
        aAdd( aStructure, { "COBRADO10","D",8,0}) // DATA DA COBRANÇA DOS 10 DIAS EM 10 DIAS
        aAdd( aStructure, { "ENVPCLI","D",8,0}) // Enviado para o Cliente o arquivo de contratos se zerar o Worker enviará novamente
        aAdd( aStructure, { "GERAR","L",1,0}) // True = Gerar Ficha quando o cliente se cadastra
        aAdd( aStructure, { "GERADO","C",20,0}) // data e hora da geração dos contratos
        aAdd( aStructure, { "DIGITO"    , "N",   1, 0 } )
        aAdd( aStructure, { "CONTRATOS","C",254,0}) // nr. dos campos de contratos
        aAdd( aStructure, { "CODIGO"    , "N",  10, 0 } )
        aAdd( aStructure, { "PESSOA"    , "C",   1, 0 } )
        aAdd( aStructure, { "NOME"      , "C",  60, 0 } )
        aAdd( aStructure, { "SEXO"      , "C",   1, 0 } )
        aAdd( aStructure, { "EMAIL"     , "C",  50, 0 } )
        aAdd( aStructure, { "CNPJCPF"   , "C",  18, 0 } )
        aAdd( aStructure, { "NACIONAL"  , "C",  30, 0 } )
        aAdd( aStructure, { "NATURALA"  , "C",  30, 0 } )
        aAdd( aStructure, { "DOC"       , "C",  16, 0 } )
        aAdd( aStructure, { "TIPODOC"   , "C",   4, 0 } )
        aAdd( aStructure, { "UFDOC"     , "C",   2, 0 } )
        aAdd( aStructure, { "ORGDOC"    , "C",   4, 0 } )
        aAdd( aStructure, { "DTDOC"     , "C",  10, 0 } )
        aAdd( aStructure, { "DTNASC"    , "C",  10, 0 } )
        aAdd( aStructure, { "MAE"       , "C",  60, 0 } )
        aAdd( aStructure, { "PAI"       , "C",  60, 0 } )
        aAdd( aStructure, { "ESTCIVIL"  , "C",  50, 0 } )
        aAdd( aStructure, { "CONJUGE"   , "C",  60, 0 } )
        aAdd( aStructure, { "REGCASA"   , "N",   1, 0 } )
        aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
        aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
        aAdd( aStructure, { "NUMERO"    , "N",   5, 0 } )
        aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
        aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
        aAdd( aStructure, { "UF"        , "C",   2, 0 } )
        aAdd( aStructure, { "PAIS"      , "C",   60, 0 } )
        aAdd( aStructure, { "DDD"       , "N",   2, 0 } )
        aAdd( aStructure, { "TEL"       , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFAX"    , "N",   2, 0 } )
        aAdd( aStructure, { "FAX"       , "C",   9, 0 } )
        aAdd( aStructure, { "DDDCEL"    , "N",   2, 0 } )
        aAdd( aStructure, { "CELULAR"   , "C",   9, 0 } )
        aAdd( aStructure, { "PROFISSAO" , "C",  80, 0 } )
        aAdd( aStructure, { "EMPRESA"   , "C",  60, 0 } )
        aAdd( aStructure, { "CEPEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "ENDEMP"    , "C",  30, 0 } )
        aAdd( aStructure, { "NREMP"     , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLEMP"  , "C",  10, 0 } )
        aAdd( aStructure, { "BAIREMP"   , "C",  18, 0 } )
        aAdd( aStructure, { "CIDEMP"    , "C",  28, 0 } )
        aAdd( aStructure, { "UFEMP"     , "C",   2, 0 } )
        aAdd( aStructure, { "PAISEMP"   , "C",   60, 0 } )
        aAdd( aStructure, { "DDDEMP"    , "N",   2, 0 } )
        aAdd( aStructure, { "TELEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFAXE"   , "N",   2, 0 } )
        aAdd( aStructure, { "FAXEMP"    , "C",   9, 0 } )
        aAdd( aStructure, { "RAMALEMP"  , "N",   5, 0 } )
        aAdd( aStructure, { "CORRESP"   , "C",  11, 0 } )
        aAdd( aStructure, { "CEPCOR"    , "C",   9, 0 } )
        aAdd( aStructure, { "ENDCOR"    , "C",  30, 0 } )
        aAdd( aStructure, { "NRCOR"     , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLCOR"  , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRCOR"   , "C",  18, 0 } )
        aAdd( aStructure, { "CIDCOR"    , "C",  28, 0 } )
        aAdd( aStructure, { "UFCOR"     , "C",   2, 0 } )
        aAdd( aStructure, { "PAISCOR"   , "C",   60, 0 } )
        aAdd( aStructure, { "BANCO"     , "C",   60, 0 } )
        aAdd( aStructure, { "AGENCIA"   , "N",   5, 0 } )
        aAdd( aStructure, { "AGDIG"     , "C",   2, 0 } )
        aAdd( aStructure, { "CONTA"     , "C",  13, 0 } )
        aAdd( aStructure, { "DIGCONTA"  , "C",   2, 0 } )
        aAdd( aStructure, { "CONTAINV"  , "C",  13, 0 } )
        aAdd( aStructure, { "DIGCCIN"   , "C",   2, 0 } )
        aAdd( aStructure, { "CONJUNTA"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCONJT"  , "C",  16, 0 } )
        aAdd( aStructure, { "IMOVEIS"   , "N",  19, 2 } )
        aAdd( aStructure, { "MOVEIS"    , "N",  19, 2 } )
        aAdd( aStructure, { "SALDO"     , "N",  19, 2 } )
        aAdd( aStructure, { "SALARIO"   , "N",  19, 2 } )
        aAdd( aStructure, { "OUTRREND"  , "N",  19, 2 } )
        aAdd( aStructure, { "NIRE"      , "C",  20, 0 } )
        aAdd( aStructure, { "FORCONST"  , "C", 100, 0 } )
        aAdd( aStructure, { "CETIP"     , "C",  20, 0 } )
        aAdd( aStructure, { "NOMEEO1"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO1"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMEEO2"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO2"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMEEO3"   , "C",  60, 0 } )
        aAdd( aStructure, { "CPFEO3"    , "C",  14, 0 } )
        aAdd( aStructure, { "NOMECON1"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON1"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON1"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON1"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON2"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON2"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON2"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON2"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON3"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON3"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON3"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON3"   , "C",  50, 0 } )
        aAdd( aStructure, { "NOMECON4"  , "C",  60, 0 } )
        aAdd( aStructure, { "CPFCON4"   , "C",  14, 0 } )
        aAdd( aStructure, { "IDENCON4"  , "C",  14, 0 } )
        aAdd( aStructure, { "ORGCON4"   , "C",  50, 0 } )
        aAdd( aStructure, { "VINCULO"   , "C",   1, 0 } )
        aAdd( aStructure, { "CNTPRO"    , "C",   1, 0 } )
        aAdd( aStructure, { "PEXPOSTA"  , "C",   1, 0 } )
        aAdd( aStructure, { "PEMOTIVO"  , "C", 200, 0 } )
        aAdd( astructure, { "REXPOSTA"  , "C",   1, 0 } )
        aAdd( aStructure, { "REMOTIVO"  , "C", 200, 0 } )
        aAdd( aStructure, { "procurador", "C",   1, 0 } )
        aAdd( aStructure, { "ORDEM"     , "C",  21, 0 } )
        aAdd( aStructure, { "QUEM"      , "C",  90, 0 } )
        aAdd( aStructure, { "CTITULAR"  , "C",   1, 0 } )
        aAdd( astructure, { "IDTIT"     , "N",  10, 0 } )
        aAdd( astructure, { "IDCTIT"    , "N",  10, 0 } )
        aAdd( aStructure, { "TIPOCOR"   , "C",   1, 0 } )
        aAdd( aStructure, { "CARTORIO"  , "C",  50, 0 } )
        aAdd( aStructure, { "TPORDEM"   , "C",   6, 0 } )
        aAdd( aStructure, { "NOMEPROC"  , "C",  60, 0 } )
        aAdd( aStructure, { "CNPJCPFPRO", "N",  14, 0 } )
        aAdd( aStructure, { "DOCPROC"   , "C",  20, 0 } )
        aAdd( aStructure, { "DTNASCPROC", "C",  10, 0 } )
        aAdd( aStructure, { "EMAILPROC" , "C",  50, 0 } )
        aAdd( aStructure, { "CEPPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "ENDPROC"   , "C",  30, 0 } )
        aAdd( aStructure, { "NRPROC"    , "N",   5, 0 } )
        aAdd( aStructure, { "COMPLPROC" , "C",  10, 0 } )
        aAdd( aStructure, { "BAIRPROC"  , "C",  18, 0 } )
        aAdd( astructure, { "CIDPROC"   , "C",  28, 0 } )
        aAdd( aStructure, { "UFPROC"    , "C",   2, 0 } )
        aAdd( aStructure, { "PAISPROC"  , "C",  60, 0 } )
        aAdd( astructure, { "DDDPROC"   , "N",   2, 0 } )
        aAdd( aStructure, { "TELPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "DDDFPROC"  , "N",   2, 0 } )
        aAdd( aStructure, { "FAXPROC"   , "C",   9, 0 } )
        aAdd( aStructure, { "IMOVESP1"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND1"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF1"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR1"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP2"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND2"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF2"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR2"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP3"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND3"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF3"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR3"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP4"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND4"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF4"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR4"  , "N",  19, 2 } )
        aAdd( aStructure, { "IMOVESP5"  , "C",  30, 0 } )
        aAdd( aStructure, { "IMOVEND5"  , "C",  60, 0 } )
        aAdd( aStructure, { "IMOVUF5"   , "C",   2, 0 } )
        aAdd( aStructure, { "IMOVVLR5"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP1"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC1"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR1"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP2"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC2"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR2"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP3"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC3"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR3"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP4"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC4"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR4"  , "N",  19, 2 } )
        aAdd( aStructure, { "BENSTP5"   , "C",  30, 0 } )
        aAdd( aStructure, { "BENSDSC5"  , "C",  60, 0 } )
        aAdd( aStructure, { "BENSVLR5"  , "N",  19, 2 } )
        aAdd( aStructure, { "CODDEPEND" , "N",   1, 0 } )
        aAdd( astructure, { "INVESTIDOR", "N",   4, 0 } )
        aAdd( astructure, { "ASSESSOR"  , "N",   5, 0 } )
        aAdd( aStructure, { "CCONH"     , "C",  50, 0 } )
        aAdd( aStructure, { "QUALOUTRO" , "C",  20, 0 } )
        aAdd( aStructure, { "NEWS"      , "C",   1, 0 } )
        aAdd( aStructure, { "LIVRE"     , "C",  50, 0 } )
        aAdd( aStructure, { "IP"        , "C",  16, 0 } )
        aAdd( aStructure, { "IPALT"     , "C",  16, 0 } )
        aAdd( aStructure, { "ORI1"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI2"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI3"      , "C", 120, 0 } )
        aAdd( aStructure, { "ORI4"      , "C", 120, 0 } )
        dbCreate(ARQ,aStructure,"DBFNTX")
        RELEASE aStructure
      EndIf
    return 

    Procedure ROT()
      If !SOUADM
        If FILE("PARAR.TXT")
          WINEXEC("AUTOMA.EXE 1")
          COMMIT
          CLOSE DATABASE
          QUIT
        EndIf
      EndIf
    return 

    Procedure TODASTAXAS(PACA)
      DECLARE TX1[1]
      NOMEARQ=""
      If dow(date())=7 .or. dow(date())=1
        return 
      EndIf
      If PACA="S"
        CDT=DTOS(DATE())+".CSV"
        If PACA="S"
          HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
          If HORAS<1400 .OR. DATE()=LOCALDTAXAS
            return 
          EndIf
        EndIf
        Erase &cdt
        oPg := CreateObject("MSXML2.ServerXMLHTTP.6.0")
        xComando := "http://www4.bcb.gov.br/Download/fechamento/"+CDT
        oPg:Open("GET",xComando,.f.)
        oPg:Send()
        v_erro1=opg:responsebody
        oPg:=nil
        memowrit(cdt,v_erro1)
        deucerto=dtoc(date())
        If AT(deucerto,V_ERRO1)=0 .OR. LEN(V_ERRO1)<7000
          Erase log.txt
          Erase cdt
          return 
        EndIf
        localdtaxas=date()
        SEENVIAEMAIL="S"
        NOMEARQ=CDT
      Else
        NOMEARQ="C:\EXATUS\TAXAS.CSV"
        SEENVIAEMAIL="N"
      EndIf
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",200,0})
      DBCREATE("C:\EXATUS\nottemp\taxas.dbf",mat_stru)
      SELE 35
      If !USEREDE("C:\EXATUS\NOTTEMP\TAXAS.DBF",.T.,10,"TXA","S")
        return 
      EndIf
      TODASTX=MEMOREAD(NOMEARQ)
      TODASTX=Strtran(TODASTX,CHR(13),"")
      TODASTX=Strtran(TODASTX,CHR(10),CHR(13)+CHR(10))
      TODASTX=TODASTX+CHR(10)+CHR(13)+CHR(10)+CHR(13)
      MEMOWRIT(NOMEARQ,TODASTX)
      APPEND FROM &NOMEARQ SDF
      GO TOP
      CONTAD=lastrec()
      If CONTAD<30
        use
        Erase C:\EXATUS\NOTTEMP\TAXAS.DBF
        localdtaxas=date()-1
        Erase LOG.TXT
        Erase CDT
        return 
      EndIf
      FNAME="\SITES\EXATUS\GRAFICO\TAXASC.CSV"
      ARQ=FCREATE(FNAME)
      wlin1="9999"+chr(13)+chr(10)+TIME()+CHR(13)+CHR(10)
      FWRITE(ARQ,WLIN1,LEN(WLIN1))
      gravou=0
      go top
      ULTMOE="   "
      do while !eof()
        MD1=LINHA+";"
        ZMDT=ALLTRIM(LELINHA(MD1))
        ZMCD=ALLTRIM(LELINHA(MD1))
        ZTPL=ALLTRIM(LELINHA(MD1))
        ZCMO=ALLTRIM(LELINHA(MD1))
        ZVAL1=ALLTRIM(LELINHA(MD1))
        ZVAL2=ALLTRIM(LELINHA(MD1))
        ZVAL3=ALLTRIM(LELINHA(MD1))
        ZVAL4=ALLTRIM(LELINHA(MD1))
        WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";"+REPLICATE("0",3-LEN(ZMCD))+ZMCD+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
        If val(substr(wlin,12,3))=0
          skip
          loop
        EndIf
        WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
        FWRITE(ARQ,WLIN1,LEN(WLIN1))
        gravou=gravou+1
        ULTMOE=SUBSTR(WLIN,12,3)
        If VAL(ULTMOE)=220
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";222"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";221"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=150
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";151"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=165
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";166"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=245
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";246"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=978
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";979"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=425
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";426"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=470
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";471"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=540
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";541"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=706
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";707"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=715
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";716"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=741
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";742"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        If VAL(ULTMOE)=745
          WLIN=REPLICATE("0",10-LEN(ZMDT))+ZMDT+";746"+";"+ZTPL+";"+ZCMO+";"+REPLICATE("0",14-LEN(ZVAL1))+ZVAL1+";"+REPLICATE("0",14-LEN(ZVAL2))+ZVAL2+";"+REPLICATE("0",14-LEN(ZVAL3))+ZVAL3+";"+REPLICATE("0",14-LEN(ZVAL4))+ZVAL4
          WLIN1=SUBSTR(WLIN,12,3)+"-"+SUBSTR(WLIN,18,3)+";"+SUBSTR(WLIN,52,14)+";"+SUBSTR(WLIN,67,14)+";"+SUBSTR(WLIN,22,14)+";"+SUBSTR(WLIN,37,14)+CHR(13)+CHR(10)
          FWRITE(ARQ,WLIN1,LEN(WLIN1))
          gravou=gravou+1
        EndIf
        skip
      enddo
      wlin1=chr(13)+chr(10)+chr(13)+chr(10)+dtoc(date())+chr(13)+chr(10)
      FWRITE(ARQ,WLIN1,LEN(WLIN1))
      FCLOSE(ARQ)
      USE
      Erase C:\EXATUS\NOTTEMP\TAXAS.DBF
      Erase log.txt
      Erase cdt
      If gravou<30
        localdtaxas=date()-1
      Else
        If SEENVIAEMAIL="S"
          ENVIAEMAIL("bene@exatus.net;wanderlei@exatus.net","BolsasNet "+ULTMOE,"todas taxas atualizadas com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
      EndIf
    return 

    Procedure APLIRESGERACAO()
      DECLARE VLABEL[ADIR(WEBLOCAL+"REC_*.*")]
      ADIR(WEBLOCAL+"REC_*.*",vlabel)
      VLABEL:=ASORT(VLABEL)
      If LEN(VLABEL)=0
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Aplicações/Resgates Geração"+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      FEZALGO=.T.
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 30
        Erase &NARQ4
        return 
      EndIf
      SET INDEX TO &NARQ4
      SELE 31
      USE
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 35
        USE
        Erase &NARQ4
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQB
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 30
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",500,0})
      ARQTXTO=WEBLOCAL1+"APLICRESG.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      narq7=WEBLOCAL4+"ETIQUETA.KEY"
      Erase &NARQ7
      SELE 37
      If !USEREDE(WEBLOCAL4+"ETIQUETA.DBF",.T.,10,"ETIQ","S")
        SELE 32
        USE
        SELE 31
        USE
        SELE 35
        USE
        SELE 30
        Erase &NARQ4
        Erase &NARQB
        return 
      Else
        NARQ7=WEBLOCAL4+"ETIQUETA.KEY"
        INDEX ON CODIGO TO &NARQ7
      EndIf
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"APLIC","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 37
        USE
        Erase &NARQ7
        SELE 32
        USE
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      SELE 36
      FOR I=1 TO LEN(VLABEL)
        NARQ=WEBLOCAL+VLABEL[I]
        APPEND FROM &NARQ SDF
        NARQ1=WEBLOCAL2+VLABEL[I]
        LZCOPYFILE(NARQ,NARQ1)
        Erase &NARQ
      NEXT I
      SELE 36
      GO TOP
      oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
      ZARQ1=WEBLOCAL+"APLRESGGERACAO.PDF"
      Erase &ZARQ1
      oReg:Set("PDFFilename",ZARQ1)
      oReg:Close()
      PRINT oPrn NAME "Aplicação / Resgates Genial Investimentos" TO TEMPDF1
      DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-9 OF oPrn
      DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-9 BOLD OF oPrn
      DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-12 OF oPrn
      DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-6 BOLD OF oprn
      DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-8 BOLD OF oPrn
      DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
      DEFINE FONT oFontv NAME "verdana" SIZE 0,12 OF oPrn
      DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
      DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
      DEFINE BRUSH oCinza       COLOR nRGB(220,220,220)
      DEFINE BRUSH oCinzaEscuro COLOR nRGB(168,168,168)
      DEFINE BRUSH oAzul COLOR nRGB(04,43,78) // 35,35,142)
      DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
      DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      DEFINE BRUSH oareia COLOR nRGB(220,207,190)
      nRowStep = oPrn:nVertRes() / 66
      nColStep = oPrn:nHorzRes() / 80
      DO WHILE !EOF()
        CLIE=VAL(SUBSTR(LINHA,42,6))
        If CLIE=0
          SKIP
          LOOP
        EndIf
        WWNOME=SUBSTR(LINHA,2,40)
        WENDE1=SUBSTR(LINHA,48,60)
        WENDE2=SUBSTR(LINHA,160,5)+"-"+SUBSTR(LINHA,165,3)+" "+ALLTRIM(SUBSTR(LINHA,128,32))+"-"+SUBSTR(LINHA,168,2)
        SOLIC=SPACE(6)+SUBSTR(LINHA,272,10)
        CONVE=SPACE(6)+SUBSTR(LINHA,250,10)
        DTPAG=SPACE(6)+SUBSTR(LINHA,342,10)
        QUANTCOTAS=SUBSTR(LINHA,234,16)
        APLICRES=LEFT(LINHA,1)
        UMACOTA=SPACE(4)+SUBSTR(LINHA,260,12)
        BRUTO=SUBSTR(LINHA,202,16)
        IR=SPACE(6)+SUBSTR(LINHA,170,10)
        IOF=SPACE(6)+SUBSTR(LINHA,180,10)
        LIQ=SUBSTR(LINHA,218,16)
        EMISSAO=SUBSTR(LINHA,352,10)
        TITULO=SUBSTR(LINHA,282,60)
        WEMAIL=LOWER(ALLTRIM(SUBSTR(LINHA,362,254)))
        SELE 35
        SEEK CLIE
        If !FOUND()
          ADIREG(0)
          REPLACE CODIGO WITH CLIE
          REPLACE NOME WITH WWNOME
          REPLACE ENDE1 WITH WENDE1
          REPLACE ENDE2 WITH WENDE2
          If !EMPTY(WEMAIL)
            REPLACE EEMAIL WITH .T.
            REPLACE EIMPRE WITH .F.
            REPLACE TEMMAIL WITH .T.
            REPLACE MAILING WITH .T.
          Else
            REPLACE EIMPRE WITH .T.
          EndIf
          REPLACE EGUARD WITH .T.
          REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
          REPLACE USERINC WITH "Arqv.Aplicação/Resgate"
        Else
          If REGLOCK(10)
            REPLACE ENDE1 WITH WENDE1
            REPLACE ENDE2 WITH ALLTRIM(WENDE2)
            REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
            REPLACE USERALT WITH "Arqv.Aplicação/Resgate"
          EndIf
        EndIf
        If !EMPTY(WEMAIL)
          SELE 31
          SEEK CLIE
          MGRAVEM=.T.
          DO WHILE CLIENTE=CLIE .And. !EOF()
            If TRIM(LOWER(EMAIL))=TRIM(LOWER(WEMAIL))
              MGRAVEM=.F.
            EndIf
            SKIP
          ENDDO
          If MGRAVEM
            ADIREG(0)
            REPLACE CODIGO WITH RECNO()
            REPLACE BMFBOV WITH "B"
            REPLACE EMAIL WITH LOWER(WEMAIL)
            REPLACE NOME WITH WWNOME
            REPLACE CLIENTE WITH CLIE
          EndIf
        EndIf
        SELE 36
        SKIP
        PAGE
        If APLICRES="1"
          saybitmap(0.0,0.0,21,1.72,WEBLOCAL+"SUB.fig",oPrn)
        Else
          saybitmap(0.0,0.0,21,1.72,WEBLOCAL+"RESG.fig",oPrn)
        EndIf
        caja(2.3,1.0,4.6,19.9,oPrn,1,oCINZA,oPen1) //caixa Nome
        oprn:cmsay(2.8,1.1,"Cliente: "+ALLTRIM(TRANSFORM(CLIE,"999999999"))+" - "+WWNOME,oFontV8b,,nRGB(0,0,0),,0) //Código do Cliente
        oprn:cmsay(3.3,1.1,wende1,oFontv8,,nRgb(0,0,0),,0)
        oPrn:cmsay(3.6,1.1,wende2,oFontv8,,nRgb(0,0,0),,0)
        oprn:cmsay(3.3,13.0,"   Emissão: "+EMISSAO,oFontV8,,nRGB(0,0,0),,0) //Data de Emissão
        saybitmap(28.5,0,21,1.72,WEBLOCAL+"RODAPE.fig",oPrn)
        CONTLIN=5.5
        SELE 36
        If APLICRES="1"
          oPrn:cmsay(CONTLIN+0.1,1.3,"Prezado(a) Cliente,",oFontv8,,nrgb(0,0,0),,0)
          oPrn:cmsay(CONTLIN+1.0,1.3,"Informamos que, nesta data, foi subscrito em cotas no:",oFontv8,,nrgb(0,0,0),,0)
        Else
          oPrn:cmsay(CONTLIN+0.1,1.3,"Prezado(a) Cliente,",oFontv8,,nrgb(0,0,0),,0)
          oPrn:cmsay(CONTLIN+1.0,1.3,"Conforme sua solicitação, informamos que foram resgatadas cotas no:",oFontv8,,nrgb(0,0,0),,0)
        EndIf
        CONTLIN=CONTLIN+2.0
        caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
        oprn:cmsay(CONTLIN+0.1,1.3,TITULO,oFontV8B,,nRGB(255,255,255),,0) //Titulo
        CONTLIN=CONTLIN+1.5
        If APLICRES="1"
          oPrn:cmsay(CONTLIN,1.3,"DADOS DA SUBSCRIÇÃO:",ofontv8b,,nRgb(0,0,0),,0)
          CONTLIN=CONTLIN+1.5
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Data da Subscrição:                                       "+SOLIC,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Data da Conversão em cotas:                          "+CONVE,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Valor da subscrição:                                    R$"+BRUTO,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Valor unitário da cota:                                 R$"+UMACOTA,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Total de Cotas:                                           "+QUANTCOTAS,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
        Else
          oPrn:cmsay(CONTLIN,1.3,"DADOS DO RESGATE:",ofontv8b,,nRgb(0,0,0),,0)
          CONTLIN=CONTLIN+1.5
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,    "Data da Solicitação:                                                "+SOLIC,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Data da Conversão em cotas:                                  "+CONVE,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,    "Quantidade de Cotas:                                          "+QUANTCOTAS,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Data do Pagamento:                                               "+DTPAG,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,    "Valor Unitário da Cota                                          R$"+UMACOTA,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"Valor Bruto do Resgate:                                       R$  "+BRUTO,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,    "Imposto de Renda:                                              R$   "+IR,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,"IOF:                                                                   R$    "+IOF,oFontv8,,nRGB(0,0,0),,0)
          contlin=contlin+1
          caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
          oPrn:cmsay(CONTLIN+0.004,1.3,    "Valor Líquido do Resgate:                                     R$ "+LIQ,oFontv8,,nRGB(0,0,0),,0)
        EndIf
        ENDPAGE
        SELE 36
      ENDDO
      ENDPRINT
      DO WHILE .T.
        CURSORWAIT()
        If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
          fClose(nHandle)
          EXIT
        Else
          fClose(nHandle)
        EndIf
      ENDDO
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      Erase &NARQ4
      SELE 31
      USE
      Erase &NARQB
      SELE 30
    return 

    Procedure EnviaCOMPEmail
      PARA host_,from_,from_name_,para_,to_name_,titulo_,corpo_,anexo_,usersmtp_,senhasmtp_,portasmtp_,arqlog_,anexo1_
      Mail_ := TOleAuto():New("IceWarpCOM.Mailer")
      _from_=ALLTRIM(substr(from_,at("@",from_)+1,len(from_)))
      Mail_:RemoteHost := host_+":"+alltrim(portasmtp_)
      Mail_:Helo=_from_
      Mail_:FromName := from_name_
      Mail_:FromAddress := from_
      Mail_:MailFrom := from_
      USERSMTP_:=TRIM(USERSMTP_)
      SENHASMTP_:=TRIM(SENHASMTP_)
      Mail_:UserName := usersmtp_
      Mail_:Password := senhasmtp_
      para_=Strtran(para_,",",";")
      Mail_:Subject := titulo_
      Mail_:IsHTML := "True"
      Mail_:BodyText := corpo_
      cPara1=ALLTRIM(LOWER(PARA_))+";"
      cPara=""
      DO WHILE .T.
        ACHE=AT(";",cPara1)
        ACHE1=LEN(cPara1)
        If ACHE=0
          EXIT
        EndIf
        wmmmail=alltrim(substr(cPara1,1,ACHE-1))
        Mail_:AddRecipient(wmmmail, to_name_)
        cPara1=substr(cPara1,ACHE+1,ACHE1)
      ENDDO
      If FILE(anexo_)
        Mail_:AddAttachment := anexo_
      EndIf
      If FILE(anexo1_)
        Mail_:addAttachment := anexo1_
      EndIf
      TX1=[1]="enviado"
      If Mail_:SendMail
        tx2[1]=" "+Mail_:Response+" PARA: "+para_+"  DE: "+from_
        tx2[2]=.T.
      Else
        DECLARE TX1[1]
        TX1[1]=" ERRO: "+Mail_:Response+" PARA: "+para_+"  DE: "+from_
        tx2[1]=" ERRO: "+Mail_:Response+" PARA: "+para_+"  DE: "+from_
        TX2[2]=.F.
        LOGFILE("c:\exatus\nao_enviado.ret",TX1)
        tx1[1]=""
      EndIf
      Mail_:=nil
    return  Nil

    Procedure FRENTE()
      DEFINE BRUSH oCinzaEscuro COLOR nRGB(50,50,50)
      DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
      DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-10 OF oPrn
      DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-8 BOLD OF oPrn
      DEFINE FONT oFontV6  NAME "Verdana" SIZE 0,-6 OF oPrn
      DEFINE PEN oPen2 COLOR nRGB(0,0,0) WIDTH 4
      DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      DEFINE FONT oFontV6B  NAME "Verdana" SIZE 0,-6 BOLD OF oPrn

      DEFINE FONT oFontV10B NAME "Verdana" SIZE 0,-10 BOLD OF oPrn
      DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
      PAGE

      saybitmap(14.3,1.0,3.26*5.33,1.3*4.1,ALLTRIM(BOLSANET->PASTA)+"LOGO.fig",oPrn)
      //caja(linhaincial,colunainicial,linhafinal,colunafinal,oPrn,vazio(0)/cheia(0),oBrush(cor),oPen)
      caja(1,0.8,10.8,20.2,oPrn,1,oCinzaEscuro,oPen1)
      caja(16.5,0.8,24.7,20.2,oPrn,1,oCinzaEscuro,oPen1)
      caja(27,1.3,28.55,20,oPrn,0,oCinzaEscuro,oPen2)

      caja(8,2,10,9.6,oPrn,1,oBranco,oPen1)
      caja(20.7,1.8,24,14.5,oPrn,1,oBranco,oPen1)

      caja(27.15,1.5,27.40,2,oPrn,0,oBranco,oPen2)
      caja(27.50,1.5,27.75,2,oPrn,0,oBranco,oPen2)
      caja(27.85,1.5,28.05,2,oPrn,0,oBranco,oPen2)
      caja(28.15,1.5,28.40,2,oPrn,0,oBranco,oPen2)

      oprn:cmsay(27.15,2.1,"MUDOU-SE",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.50,2.1,"ENDEREÇO INSUFICIENTE",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.85,2.1,"NÚMERO INEXISTENTE",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(28.15,2.1,"DESCONHECIDO",oFontV6,,nRGB(0,0,0),,0)

      caja(27.15,5.3,27.40,5.8,oPrn,0,oBranco,oPen2)
      caja(27.50,5.3,27.75,5.8,oPrn,0,oBranco,oPen2)
      caja(27.85,5.3,28.05,5.8,oPrn,0,oBranco,oPen2)
      caja(28.15,5.3,28.40,5.8,oPrn,0,oBranco,oPen2)

      oprn:cmsay(27.15,5.9,"NÃO PROCURADO",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.50,5.9,"AUSENTE",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.85,5.9,"FALECIDO",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(28.15,5.9,"RECUSADO",oFontV6,,nRGB(0,0,0),,0)

      caja(28.75,1,28.80,20.2,oPrn,1,oPreto,oPen1)
      oprn:cmsay(28.85,7.6,"PARA USO DO CORREIO",oFontV10B,,nRGB(0,0,0),,0)
      caja(29.3,1,29.35,20.2,oPrn,1,oPreto,oPen1)


      caja(27.05,8.3,28.55,8.31,oPrn,1,oPreto,oPen1)
      caja(27.05,14.3,28.55,14.31,oPrn,1,oPreto,oPen1)
      caja(27.05,17.3,28.55,17.31,oPrn,1,oPreto,oPen1)

      oprn:cmsay(27.15,8.4,"ASSINATURA DO ENTREGADOR NR:",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.15,14.4,"DATA:",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.15,17.4,"REINICIADO SERVIÇO",oFontV6,,nRGB(0,0,0),,0)
      oprn:cmsay(27.50,17.4,"POSTAL EM:",oFontV6,,nRGB(0,0,0),,0)

      //REMETENTE

      If WEBCOR="GERACAO"
        oprn:cmsay(8.2,2.2,"REMETENTE:",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.2,4.4,"GENIAL INVESTIMENTOS",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.5,4.4,"AV PAULISTA, 287",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.8,4.4,"10 E 11 ANDARES",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(9.1,4.4,"SÃO PAULO - SP",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(9.4,4.4,"CEP: 01311-000",oFontV6B,,nRGB(0,0,0),,0)

        //DESTINATÁRIO
        oprn:cmsay(20.9,2.1,left(wbene,48),oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(21.3,2.1,wende1,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(21.7,2.1,wcep1+" "+trim(wcid)+"-"+wuf,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(22.5,2.1,"controle:"+strzero(squencia,4),oFontv8,,nRgb(0,0,0),,0)

      ElseIf WEBCOR="CONVEC"
        // REMETENTE
        oprn:cmsay(8.2,2.2,"REMETENTE:",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.2,4.4,"Rua Amauri, 255 8.andar",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.5,4.4,"São Paulo - SP",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.8,4.4,"Cep: 01448-000",oFontV6B,,nRGB(0,0,0),,0)
        //DESTINATÁRIO
        oprn:cmsay(20.9,2.1,left(NOME,48),oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(21.3,2.1,ENDE1,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(21.7,2.1,ENDE2,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(22.5,2.1,"controle:"+strzero(squencia,4),oFontv8,,nRgb(0,0,0),,0)
      ElseIf WEBCOR="PLANNER"
        // REMETENTE
        oprn:cmsay(8.2,2.2,"REMETENTE:",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.2,4.4,"Av.Brig.Faria Lima, 3900 10.andar",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.5,4.4,"São Paulo - SP",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.8,4.4,"Cep: 04538-132",oFontV6B,,nRGB(0,0,0),,0)
        //DESTINATÁRIO
        oprn:cmsay(20.9,2.1,left(NOME,48),oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(21.3,2.1,ENDE1,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(21.7,2.1,ENDE2,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(22.5,2.1,"controle:"+strzero(squencia,4),oFontv8,,nRgb(0,0,0),,0)
      EndIf
      ENDPAGE
    return 

    Procedure ATUAEMAILGER()
      DECLARE VLABEL[ADIR(WEBLOCAL+"EMAILGERAC*.TXT")]
      ADIR(WEBLOCAL+"EMAILGERAC*.TXT",vlabel)
      If LEN(VLABEL)=0
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Atualização emails Geração "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      FEZALGO=.T.
      mat_stru={}
      Aadd(mat_stru,{"CODIGO","N",9,0})
      Aadd(mat_stru,{"NOME","C",50,0})
      Aadd(mat_stru,{"CNPJCPF","C",14,0})
      AaDD(mat_stru,{"SENHA","C",4,0})
      AaDD(mat_stru,{"DTNASC","C",10,0})
      AaDD(mat_stru,{"EMAIL","C",254,0})
      AaDD(mat_stru,{"ENVMAIL","C",1,0})
      AaDD(mat_stru,{"IMPMAIL","C",1,0})
      AaDD(mat_stru,{"ENDER1","C",60,0})
      AaDD(mat_stru,{"BAIRR","C",20,0})
      AaDD(mat_stru,{"CIDADE","C",32,0})
      AaDD(mat_stru,{"CEP","C",8,0})
      AaDD(mat_stru,{"NADA","C",2,0})
      AaDD(mat_stru,{"UF","C",2,0})
      ARQTXTO=WEBLOCAL1+"EMAIL.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"EMAIL","S")
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      SELE 36
      FOR I=1 TO LEN(VLABEL)
        NARQ=WEBLOCAL+VLABEL[I]
        APPEND FROM &NARQ SDF
        NARQ1=WEBLOCAL2+VLABEL[I]
        LZCOPYFILE(NARQ,NARQ1)
        Erase &NARQ
      NEXT I
      SELE 31
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
        SELE 36
        USE
        SELE 30
        return 
      EndIf
      go top
      do while !eof()
        dele
        skip
      ENDDO
      PACK
      USE
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 36
        USE
        SELE 30
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQ5
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTES","S")
        SELE 36
        USE
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 36
        USE
        SELE 31
        USE
        SELE 30
        Erase &NARQ4
        return 
      EndIf
      SET INDEX TO &NARQ4
      GO TOP
      DO WHILE !EOF()
        If REGLOCK(10)
          REPLACE EIMPRE WITH .T.
          REPLACE EEMAIL WITH .F.
        EndIf
        SKIP
      ENDDO
      SELE 36
      GO TOP
      DO WHILE !EOF()
        If SENHA="####" .OR. CNPJCPF="####"
          SKIP
          LOOP
        EndIf
        SELE 35
        SEEK email->CODIGO
        If !FOUND()
          ADIREG(0)
          REPLACE INCLUIDO WITH dtoc(DATE())+" "+TIME()
          REPLACE USERINC WITH "Arqv.Email - Geração"
        Else
          REGLOCK(10)
          REPLACE ALTERADO WITH dtoc(DATE())+" "+TIME()
          REPLACE USERALT WITH "Arqv.Email - Geração"
        EndIf
        REPLACE CODIGO WITH EMAIL->CODIGO
        REPLACE NOME WITH EMAIL->NOME
        REPLACE CNPJCPF WITH EMAIL->CNPJCPF
        REPLACE SENHA WITH EMAIL->SENHA
        REPLACE NASC WITH CTOD(EMAIL->DTNASC)
        If EMAIL->ENVMAIL="S"
          REPLACE EEMAIL WITH .T.
        Else
          REPLACE EEMAIL WITH .F.
        EndIf
        If EMAIL->IMPMAIL="S"
          REPLACE EIMPRE WITH .T.
        Else
          REPLACE EIMPRE WITH .F.
        EndIf
        REPLACE EGUARD WITH .F.
        REPLACE MAILING WITH .T.
        If EMPTY(EMAIL->EMAIL)
          REPLACE TEMMAIL WITH .F.
        Else
          REPLACE TEMMAIL WITH .T.
        EndIf
        REPLACE ENDE1 WITH EMAIL->ENDER1
        REPLACE ENDE2 WITH LEFT(EMAIL->CEP,5)+"-"+RIGHT(EMAIL->CEP,3)+" "+ALLTRIM(EMAIL->CIDADE)+"-"+EMAIL->UF
        If EMPTY(EMAIL->EMAIL)
          SELE 36
          SKIP
          LOOP
        EndIf
        SELE 31
        SEEK EMAIL->CODIGO
        If !FOUND()
          ADIREG(0)
          REPLACE INCLUIDO WITH dtoc(DATE())+" "+TIME()
          REPLACE USERINC WITH "Arqv.Email - Geração"
        Else
          REGLOCK(10)
          REPLACE ALTERADO WITH dtoc(DATE())+" "+TIME()
          REPLACE USERALT WITH "Arqv.Email - Geração"
        EndIf
        REPLACE CODIGO WITH RECNO()
        REPLACE BMFBOV WITH "B"
        REPLACE EMAIL WITH LOWER(EMAIL->EMAIL)
        REPLACE NOME WITH EMAIL->NOME
        REPLACE CLIENTE WITH EMAIL->CODIGO
        REPLACE INCLUIDO WITH dTOc(DATE())+TIME()
        REPLACE USERINC WITH "Arqv.Email - Geração"
        SELE 36
        SKIP
      ENDDO
      SELE 31
      USE
      SELE 36
      USE
      SELE 35
      USE
      SELE 30
      Erase &NARQ4
      Erase &NARQ5
      Erase &ARQTXTO
    return 

    Procedure EXTRGERFUT()
      NARQ=WEBLOCAL+"EXTRATOGERACAO.TXT"
      wnovo=.F.
      If !FILE(NARQ)
        NARQ=WEBLOCAL+"EXTRATOGERACAO_IMPRIMIR.TXT"
        wnovo=.F.
        If !FILE(NARQ)
          NARQ=WEBLOCAL+"EXTRATOV3GERACAO.TXT"
          wnovo=.T.
          If !FILE(NARQ)
            NARQ=WEBLOCAL+"EXTRATOV4GERACAO.TXT"
            wnovo=.T.
            If !FILE(NARQ)
              return 
            EndIf
          EndIf
        EndIf
      EndIf
      SysRefresh()
      CURSORWAIT()
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Extrato Unificado "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      FEZALGO=.T.
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 30
        Erase &NARQ4
        return 
      EndIf
      SET INDEX TO &NARQ4
      SELE 31
      USE
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 35
        USE
        Erase &NARQ4
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQB
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 30
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",500,0})
      ARQTXTO=WEBLOCAL1+"EXTRATO.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      narq7=WEBLOCAL4+"ETIQUETA.KEY"
      Erase &NARQ7
      SELE 37
      If !USEREDE(WEBLOCAL4+"ETIQUETA.DBF",.T.,10,"ETIQ","S")
        SELE 32
        USE
        SELE 31
        USE
        SELE 35
        USE
        SELE 30
        Erase &NARQ4
        Erase &NARQB
        return 
      Else
        NARQ7=WEBLOCAL4+"ETIQUETA.KEY"
        INDEX ON CODIGO TO &NARQ7
      EndIf
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"EXTRATO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 37
        USE
        Erase &NARQ7
        SELE 32
        USE
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      SELE 36
      APPEND FROM &NARQ SDF
      If NARQ=WEBLOCAL+"EXTRATOV4GERACAO.TXT"
        SOPROC=.T.
      EndIf
      NARQ1=WEBLOCAL2+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+"EXTRATO.TXT"
      LZCOPYFILE(NARQ,NARQ1)
      Erase &NARQ
      GO TOP
      DO WHILE .T.
        If LEFT(LINHA,1)<>"1"
          SKIP
          LOOP
        EndIf
        EXIT
      ENDDO
      contarqs=0
      feitoarq=0
      CONTP=0
      DO WHILE !EOF()
        CONTP=CONTP+1
        If CONTP>5000
          FEITOARQ=0
          CONTP=0
          ENDPRINT
          DO WHILE .T.
            CURSORWAIT()
            If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
              fClose(nHandle)
              EXIT
            Else
              fClose(nHandle)
            EndIf
          ENDDO
        EndIf
        If feitoarq=0
          contarqs=contarqs+1
          oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
          ZARQ1=WEBLOCAL+"EX_GERACAO"+strzero(contarqs,4)+".PDF"
          Erase &ZARQ1
          oReg:Set("PDFFilename",ZARQ1)
          oReg:Close()
          PRINT oPrn NAME "Extrato Genial Investimentos" TO TEMPDF1
          DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-8 OF oPrn
          DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-8 BOLD OF oPrn
          DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-6.5 BOLD OF oPrn
          DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-6.5 BOLD OF oprn
          DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-8 BOLD OF oPrn
          DEFINE FONT oFontt81 NAME "Courier New" SIZE 0,-7 BOLD OF oPrn
          DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
          DEFINE FONT oFontv NAME "verdana" SIZE 0,12 OF oPrn
          DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
          DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
          DEFINE BRUSH oCinza       COLOR nRGB(220,220,220)
          DEFINE BRUSH oCinzaEscuro COLOR nRGB(168,168,168)
          DEFINE BRUSH oAzul COLOR nRGB(69,107,180) // 04,43,78
          DEFINE BRUSH onAzul       COLOR nRGB(92,100,129)
          DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
          DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
          DEFINE BRUSH oareia COLOR nRGB(220,207,190)
          DEFINE BRUSH onCinzaclaro COLOR nRGB(241,241,241)
          DEFINE BRUSH onBranco     COLOR nRGB(255,255,255)
          DEFINE BRUSH onBege       COLOR nRGB(176,165,135)
          DEFINE BRUSH onBegeClaro  COLOR nRGB(244,242,238)
          DEFINE BRUSH onCinzaEscuro COLOR nRGB(128,130,134)
          DEFINE FONT onFontCourier NAME "courier new" BOLD SIZE 0, -10.0 OF oPrn
          DEFINE FONT onFontVerdana NAME "Verdana"  SIZE 0, -11.0 BOLD OF oPrn
          DEFINE FONT onFont1Courier NAME "courier new negrito" SIZE 0, -10.0 BOLD OF oPrn
          nRowStep = oPrn:nVertRes() / 66
          nColStep = oPrn:nHorzRes() / 80
          feitoarq=1
        EndIf
        CLIE=VAL(SUBSTR(LINHA,60,6))
        WWNOME=SUBSTR(LINHA,10,50)
        PERDE=SUBSTR(LINHA,180,10)
        PERATE=SUBSTR(LINHA,190,10)
        EMISSAO=SUBSTR(LINHA,200,10)
        SELE 35
        SEEK CLIE
        If !FOUND()
          ADIREG(0)
          REPLACE CODIGO WITH CLIE
          REPLACE NOME WITH WWNOME
          REPLACE EGUARD WITH .T.
          REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
          REPLACE USERINC WITH "Arqv.Extrato Geração"
        EndIf
        WENDE1=ENDE1
        WENDE2=ENDE2
        WASSESSOR=STRZERO(ASSESSOR,5)
        SELE 36
        CX=0
        CABEC=0
        contfol=1
        contlin=0
        FEZ2=0
        FEZ21=0
        FEZ22=0
        FEZ3=0
        FEZ4=0
        FEZ5=0
        FEZ6=0
        FEZ7=0
        CX=0
        ME_SES="Janeiro  FevereiroMarço    Abril    Maio     Junho    Julho    Agosto   Setembro Outubro  Novembro Dezembro "
        SKIP
        PAPANTES="   "
        DO WHILE !EOF() .And. LEFT(LINHA,1)#"1"
          If val(left(linha,1))=0
            skip
            loop
          EndIf
          If CONTLIN>24.5
            CABEC=0
            FEZ2=0
            FEZ21=0
            FEZ22=0
            FEZ3=0
            FEZ4=0
            FEZ5=0
            FEZ6=0
            contlin=0
            contfol=contfol+1
            SELE 37
            SEEK CLIE
            REPLACE FOLHAS WITH FOLHAS+1
            oprn:cmsay(26.0,18.0,"Continua...",oFontV8B,,nRGB(0,0,0),,0) //Data de Emissão
            SELE 36
            ENDPAGE
          EndIf
          If CABEC=0
            PAGE
            If SEENVIAEMAIL="S" .or. NARQ=WEBLOCAL+"EXTRATOV4GERACAO.TXT"
              saybitmap(0.0,0.0,21,1.72,WEBLOCAL+"EXTRATO.fig",oPrn)
            Else
              saybitmap(0.0,0.0,21,1.72,WEBLOCAL+"EXTRATO.BMP",oPrn)
            EndIf
            caja(2.3,1.0,4.6,19.9,oPrn,1,oCINZA,oPen1) //caixa Nome
            oprn:cmsay(2.8,1.1,"Cliente: "+ALLTRIM(TRANSFORM(CLIE,"999999999"))+" - "+WWNOME,oFontV8b,,nRGB(0,0,0),,0) //Código do Cliente
            oprn:cmsay(3.3,1.1,wende1,oFontv8,,nRgb(0,0,0),,0)
            oPrn:cmsay(3.6,1.1,wende2,oFontv8,,nRgb(0,0,0),,0)
            oprn:cmsay(2.8,13.0,"   Página : "+STRZERO(CONTFOL,3),oFontV8,,nRGB(0,0,0),,0) //Página
            oprn:cmsay(3.3,13.0,"   Emissão: "+EMISSAO,oFontV8,,nRGB(0,0,0),,0) //Data de Emissão
            oprn:cmsay(3.6,13.0,"   Período: "+PERDE+" a "+PERATE,oFontv8,,nRgb(0,0,0),,0) // Período

            CONTLIN=5.5

            oPrn:cmsay(27.2,1.3,"RPA: Informamos que houve alteração no documento RPA - Regra e Parâmetros de Atuação, sendo que o mesmo encontra-se",oFontv8,,nRgb(0,0,0),,0)
            oPrn:cmsay(27.5,1.3,"disponível em www.genialinvestimentos/legislacao",oFontV8,,nRgb(0,0,0),,0)
            //      oPrn:cmsay(27.8,1.3,'Perfil de Investimento: Comunicamos que houve alteração em nossa metodologia para definição do Perfil de Risco "Suitability"',oFontv8,,nRgb(0,0,0),,0)
            //      oPrn:cmsay(28.1,1.3,'consulte seu perfil na área de clientes no site www.genialinvestimentos.com.br no menu "Meus Dados".',oFontV8,,nRgb(0,0,0),,0)

            cx=1
            If SEENVIAEMAIL="S" .or. NARQ=WEBLOCAL+"EXTRATOV4GERACAO.TXT"
              saybitmap(28.5,0,21,1.72,WEBLOCAL+"RODAPE.fig",oPrn)
            Else
              oprn:cmsay(29.0,15.5,"ouvidoria@genialinvestimentos.com.br",oFontV8,,nRGB(0,0,0),,0)
              saybitmap(28.5,0,21,1.72,weblocal+"rodape.bmp",oprn)
            EndIf
            CABEC=1
          EndIf
          SELE 36

          If left(LINHA,1)="2"
            If FEZ2=0
              CX=1
              caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"POSIÇÃO DE FUNDOS E CLUBES DE INVESTIMENTOS",oFontV8B,,nRGB(255,255,255),,0) //Titulo
              CONTLIN=CONTLIN+0.8

              oprn:cmsay(CONTLIN+0.0,1.3,"         Quantidades             Cota              Saldo        IR Previsto                IOF             Saldo *",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.2
              oprn:cmsay(CONTLIN+0.0,1.3,"            de Cotas            Atual              Bruto                                                 Líquido",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.5
              caja(CONTLIN-0.1,0,CONTLIN+0.1,21.0,oPrn,1,onBege,oPen1) //caixa bege
              CONTLIN=CONTLIN+0.2
              FEZ2=1
              FEZ21=0
              FEZ22=0
            EndIf
            If SUBSTR(LINHA,2,6)<>"102905" .And. FEZ21=0
              caja(CONTLIN-0.1,0,CONTLIN+1.0,21.0,oPrn,1,onBegeClaro,oPen1) //caixa bege
              oprn:cmsay(CONTLIN+0.0,1.3,"Administrador: GENIAL INVESTIMENTOS S.A.",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.3
              oprn:cmsay(CONTLIN+0.0,1.3,"               CNPJ: 27.652.684/0001-62",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.3
              oprn:cmsay(CONTLIN+0.0,1.3,"               Praça XV de Novembro,nº20 - 12ºandar,Grupo 1201-B,Bairro:Centro,Rio de Janeiro/RJ,CEP 20010-010",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.7
              FEZ21=1
            ElseIf SUBSTR(LINHA,2,6)="102905" .And. FEZ22=0
              If FEZ21=1
                caja(CONTLIN-0.1,0,CONTLIN+0.2,21.0,oPrn,1,onBege,oPen1) //caixa bege
                CONTLIN=CONTLIN+0.15
              EndIf
              caja(CONTLIN-0.1,0,CONTLIN+1.0,21.0,oPrn,1,onBegeClaro,oPen1) //caixa bege
              oprn:cmsay(CONTLIN+0.0,1.3,"Administrador: BRB Distribuidora de Títulos e Valores Mobiliários S.A.",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.3
              oprn:cmsay(CONTLIN+0.0,1.3,"               CNPJ: 33.850.686/0001-69",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.3
              oprn:cmsay(CONTLIN+0.0,1.3,"               SBS Quadra 01, Bloco E Edifício Brasília, 7º Andar | DF - Brasília",OfONTt8,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+0.7
              FEZ22=1
            EndIf
            If FEZ2=0
              //       oprn:cmsay(CONTLIN+0.0,1.3,"         Quantidades             Cota              Saldo        IR Previsto                IOF             Saldo *",OfONTt8,,nrGB(0,0,0),,0)
              //       CONTLIN=CONTLIN+0.2
              //       oprn:cmsay(CONTLIN+0.0,1.3,"            de Cotas            Atual              Bruto                                                 Líquido",OfONTt8,,nrGB(0,0,0),,0)
              //      CONTLIN=CONTLIN+0.5
              //    FEZ2=1
            EndIf
            caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
            oprn:cmsay(CONTLIN+0.004,1.3,alltrim(SUBSTR(LINHA,10,52)),oFontV8b,,nRGB(0,0,0),,0)
            oprn:cmsay(CONTLIN+0.004,15,"CNPJ: "+TRANSFORM(SUBSTR(LINHA,169,14),"@R 99.999.999/9999-99"),oFontV8b,,nRGB(0,0,0),,0)
            CONTLIN=CONTLIN+0.5
            oprn:cmsay(CONTLIN+0.004,1.3," "+SUBSTR(LINHA,78,19)+" "+SUBSTR(LINHA,62,16)+" "+SUBSTR(LINHA,97,18)+" "+SUBSTR(LINHA,115,18)+" "+SUBSTR(LINHA,133,18)+" "+SUBSTR(LINHA,151,18),OfONTt8,,nRGB(0,0,0),,0)
            CONTLIN=CONTLIN+0.5
          EndIf

          If left(LINHA,1)="6"
            caja(CONTLIN-0.1,0,CONTLIN+0.2,21.0,oPrn,1,onBege,oPen1) //caixa bege
            CONTLIN=CONTLIN+0.15
            If CX=0
              caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oBranco,oPen1) //caixa branca
              CX=1
            Else
              caja(CONTLIN-0.1,0.0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
              CX=0
            EndIf
            oprn:cmsay(CONTLIN+0.004,1.3,SUBSTR(LINHA,10,37)+" "+SUBSTR(LINHA,92,18)+" "+SUBSTR(LINHA,110,18)+" "+SUBSTR(LINHA,128,18)+" "+SUBSTR(LINHA,146,18),oFontt8,,nRGB(0,0,0),,0)
            CONTLIN=CONTLIN+0.4
            caja(CONTLIN-0.1,0,CONTLIN+0.05,21.0,oPrn,1,onBege,oPen1) //caixa bege
            CONTLIN=CONTLIN+0.1
          EndIf
          If left(LINHA,1)="4"
            If FEZ4=0
              FEZ4=1
              CX=1
              CONTLIN=CONTLIN+0.5
              caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"MOVIMENTAÇÃO NO PERÍODO",oFontV8B,,nRGB(255,255,255),,0) //Titulo
              CONTLIN=CONTLIN+1
            EndIf

            If SUBSTR(LINHA,8,2)="AA" .OR. PAPANTES<>SUBSTR(LINHA,2,6)
              PAPANTES=SUBSTR(LINHA,2,6)
              contlin=contlin-0.1

              If CX=0
                caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oBranco,oPen1) //caixa branca
                CONTLIN=CONTLIN+0.5
              EndIf

              caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza

              oPrn:cmsay(CONTLIN+0.0,1.3,SUBSTR(LINHA,10,52),oFontV8b,,nRGB(0,0,0),,0)
              oprn:cmsay(CONTLIN+0.5,1.3,"Data da        Lançamento               Valor             Valor          Quantidade                  IR                IOF              Saldo",OfONTC8B,,nrGB(0,0,0),,0)
              oprn:cmsay(CONTLIN+0.7,1.3,"Operação                                                da cota            de cotas                                                   Líquido",OfONTC8B,,nrGB(0,0,0),,0)
              contlin=contlin+1.1
              CX=1
              OBS="Saldo Anterior  "

            ElseIf SUBSTR(LINHA,8,2)="AP"
              OBS="     Aplicação  "
            ElseIf SUBSTR(LINHA,8,2)="AT"
              OBS="   Saldo Atual  "
            ElseIf SUBSTR(LINHA,8,2)="RG"
              OBS="       Resgate  "
            ElseIf SUBSTR(LINHA,8,2)="AI"
              OBS="Ap.Incorporada  "
            ElseIf SUBSTR(LINHA,8,2)="RI"
              OBS="Rg.Incorporado  "
            ElseIf SUBSTR(LINHA,8,2)="CC"
              OBS="IR LEI. 9.532   "
            ElseIf substr(linha,8,2)="TA"
              OBS="Transf.Cotas Apl"
            ElseIf SUBSTR(LINHA,8,2)="TR"
              OBS="Transf.Cotas Rgt"
            Else
              OBS="                "
            EndIf
            If CX=0

              caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oBranco,oPen1) //caixa branca
              CX=1
            Else
              caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
              CX=0
            EndIf
            oprn:cmsay(CONTLIN+0.004,1.3,SUBSTR(LINHA,62,10)+" "+OBS+SUBSTR(LINHA,91,18)+"  "+SUBSTR(LINHA,109,16)+" "+SUBSTR(LINHA,72,19)+"  "+SUBSTR(LINHA,125,18)+" "+SUBSTR(LINHA,143,18)+" "+SUBSTR(LINHA,161,18),oFontC8,,nRGB(0,0,0),,0)
            CONTLIN=CONTLIN+0.5
          EndIf

          If left(LINHA,1)="3" .And. !WNOVO
            If FEZ3=0
              FEZ3=1
              REGISJ=RECNO()
              DO WHILE !EOF()
                If LINHA="1"
                  EXIT
                EndIf
                If LINHA="7"
                  CREFIS=SUBSTR(LINHA,2,20)
                  CREFIS=Strtran(CREFIS,".","")
                  CREFIS=VAL(Strtran(CREFIS,",","."))
                  If CREFIS>0 .And. FEZ7=0
                    CREFIS=TRANSFORM(CREFIS,"@E 999,999,999.99")
                    FEZ7=1
                    CX=1
                    caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
                    oprn:cmsay(CONTLIN+0.1,1.3,"* CRÉDITO FISCAL E/OU PROVISÃO FISCAL",oFontV8B,,nRGB(255,255,255),,0) //Titulo
                    oprn:cmsay(CONTLIN+0.7,1.3,"Valor Acumulado: "+CREFIS,oFontt8,,nrGB(0,0,0),,0)
                    oprn:cmsay(CONTLIN+1.4,1.3,"O Valor relativo ao Crédito Fiscal será utilizado na compensação de impostos a pagar de Fundos sob a",OfONTt8,,NrGB(0,0,0),,0)
                    oPrn:cmsay(CONTLIN+1.7,1.3,"administração da Genial Investimentos. A provisão será utilizada no próprio produto.",OfONTt8,,nRgb(0,0,0),,0)
                    CONTLIN=CONTLIN+1.9
                  EndIf
                  EXIT
                EndIf
                SKIP
              ENDDO
              GO REGISJ
              CX=1
              CONTLIN=CONTLIN+0.5
              caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"INDICADORES DE PERFORMANCE",oFontV8B,,nRGB(255,255,255),,0) //Titulo
              wlin1="Produto                              Mes("+TRIM(SUBS(ME_SES,VAL(SUBSTR(perate,4,2))*9-8,9))+")                  Acumulado("+RIGHT(PERATE,4)+")                             Desde o Início"
              wlin2="                             Produto       CDI  Ibovespa   Produto       CDI  Ibovespa Data Início     Produto       CDI  Ibovespa"
              oPrn:cmsay(CONTLIN+0.7,1.3,wlin1,OfONTt81,,nrGB(0,0,0),,0)
              oPrn:cmsay(CONTLIN+0.9,1.3,wlin2,OfONTt81,,nrGB(0,0,0),,0)
              CONTLIN=CONTLIN+1.5
            EndIf
            If CX=0
              caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oBranco,oPen1) //caixa branca
              CX=1
            Else
              caja(CONTLIN-0.1,0,CONTLIN+0.4,21.0,oPrn,1,oCinza,oPen1) //caixa cinza
              CX=0
            EndIf
            oprn:cmsay(CONTLIN+0.004,1.3,SUBSTR(LINHA,10,25)+" "+SUBSTR(LINHA,35,60)+"  "+SUBSTR(LINHA,125,10)+"  "+SUBSTR(LINHA,95,30),oFontt81,,nRGB(0,0,0),,0)
            CONTLIN=CONTLIN+0.5
          EndIf
          If left(LINHA,1)="3" .And. WNOVO
            If SUBSTR(LINHA,3,1)="0"
              onCor=onCinzaclaro
            ElseIf SUBSTR(LINHA,3,1)="1"
              onCor=onBranco
            ElseIf SUBSTR(LINHA,3,1)="2"
              onCor=onBegeClaro
            ElseIf SUBSTR(LINHA,3,1)="3"
              onCor=onAzul
            ElseIf SUBSTR(LINHA,3,1)="4"
              onCor=onBege
            ElseIf SUBSTR(LINHA,3,1)="5"
              onCor=onCinzaEscuro
            ElseIf SUBSTR(LINHA,3,1)="6"
              onCor=onCinzaEscuro
            Else
              onCor=onCinzaEscuro
            EndIf
            If SUBSTR(LINHA,2,1)="S"
              If FEZ3=0
                FEZ3=1
                CONTLIN=CONTLIN+0.1
                caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
                oprn:cmsay(CONTLIN+0.1,1.3,"INDICADORES DE PERFORMANCE",oFontV8B,,nRGB(255,255,255),,0) //Titulo
                CONTLIN=CONTLIN+0.5
                oprn:cmsay(CONTLIN+0.1,8.0,"                                                                                    Desde o",oFontV8B,,RGB(04,43,78),,0)
                CONTLIN=CONTLIN+0.3
                CLUB=ALLTRIM(SUBSTR(ME_SES,VAL(SUBSTR(perate,4,2))*9-8,9))+"(%)"
                CLUB=CLUB+SPACE(9-LEN(CLUB))
                oprn:cmsay(CONTLIN+0.1,8.0,"       "+CLUB+"           Ano(%)           12M(%)        Início(%)        Data de Início",oFontV8B,,RGB(04,43,78),,0)
                CONTLIN=CONTLIN+0.5
              EndIf
              caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,onCor,oPen1) //caixa Cor onCor (Cabecalho)
              NCLUB=ALLTRIM(SUBSTR(LINHA,11,50))
              oprn:cmsay(CONTLIN+0.1,9.5,NCLUB,oFontV8B,,RGB(255,255,255),,0)
              CONTLIN=CONTLIN+0.6
            Else
              NCLUB0=ALLTRIM(SUBSTR(LINHA,11,50))
              NCLUB=SUBSTR(LINHA,11,25)
              NCLUB1=SUBSTR(LINHA,36,25)
              If RIGHT(NCLUB,1)<>" " .And. LEFT(NCLUB1,1)<>" "
                NCLUB1=ALLTRIM(SUBSTR(NCLUB,RAT(" ",NCLUB)+1,25))+NCLUB1
                NCLUB=LEFT(NCLUB,RAT(" ",NCLUB))
              EndIf
              NCLUB=ALLTRIM(NCLUB)
              NCLUB1=ALLTRIM(NCLUB1)
              //        NCLUB=SPACE(25-LEN(NCLUB))+NCLUB
              //        NCLUB1=SPACE(25-LEN(NCLUB1))+NCLUB1
              If LEN(NCLUB0)>25
                caja(CONTLIN-0.1,0,CONTLIN+1.0,21.0,oPrn,1,onCor,oPen1) //caixa onCor (informação)
                oprn:cmsay(CONTLIN+0.004,1.3,NCLUB,oFontV8B,,nRGB(0,0,0),,0)
                oPRN:cmsay(CONTLIN+0.3,1.3,NCLUB1,oFontV8B,,nRGB(0,0,0),,0)
                CONTLIN=CONTLIN+0.2
              Else
                caja(CONTLIN-0.1,0,CONTLIN+0.8,21.0,oPrn,1,onCor,oPen1) //caixa onCor (informação)
                oprn:cmsay(CONTLIN+0.004,1.3,NCLUB,oFontV8B,,nRGB(0,0,0),,0)
              EndIf
              If len(nclub0)>25
                oprn:cmsay(CONTLIN-0.1+0.004,8.0," "+SUBSTR(LINHA,71,10)+" "+SUBSTR(LINHA,81,10)+" "+SUBSTR(LINHA,91,10)+" "+SUBSTR(LINHA,101,10)+"    "+SUBSTR(LINHA,61,10),onFontCourier,,nRGB(0,0,0),,0)
              Else
                oprn:cmsay(CONTLIN+0.004,8.0," "+SUBSTR(LINHA,71,10)+" "+SUBSTR(LINHA,81,10)+" "+SUBSTR(LINHA,91,10)+" "+SUBSTR(LINHA,101,10)+"    "+SUBSTR(LINHA,61,10),onFontCourier,,nRGB(0,0,0),,0)
              EndIf
              If len(NCLUB0)>25
                caja(CONTLIN-0.3,0,CONTLIN-0.3+0.01,21.0,oPrn,1,onCinzaEscuro,oPen1) //caixa onCor (informação)
              Else
                caja(CONTLIN-0.1,0,CONTLIN-0.1+0.01,21.0,oPrn,1,onCinzaEscuro,oPen1) //caixa onCor (informação)
              EndIf
              CONTLIN=CONTLIN+0.6
            EndIf
          EndIf

          If LINHA="7" .And. WNOVO
            CREFIS=SUBSTR(LINHA,2,20)
            CREFIS=Strtran(CREFIS,".","")
            CREFIS=VAL(Strtran(CREFIS,",","."))
            If CREFIS>0 .And. FEZ7=0
              CREFIS=TRANSFORM(CREFIS,"@E 999,999,999.99")
              FEZ7=1
              CX=1
              caja(CONTLIN,0,CONTLIN+0.6,21.0,oPrn,1,oAzul,oPen1) //caixa azul
              oprn:cmsay(CONTLIN+0.1,1.3,"* CRÉDITO FISCAL E/OU PROVISÃO FISCAL",oFontV8B,,nRGB(255,255,255),,0) //Titulo
              oprn:cmsay(CONTLIN+0.7,1.3,"Valor Acumulado: "+CREFIS,oFontt8,,nrGB(0,0,0),,0)
              oprn:cmsay(CONTLIN+1.4,1.3,"O Valor relativo ao Crédito Fiscal será utilizado na compensação de impostos a pagar de Fundos sob a",OfONTt8,,NrGB(0,0,0),,0)
              oPrn:cmsay(CONTLIN+1.7,1.3,"administração da Genial Investimentos. A provisão será utilizada no próprio produto.",OfONTt8,,nRgb(0,0,0),,0)
              CONTLIN=CONTLIN+1.9
            EndIf
          EndIf

          SKIP

          TOTALINHAS=26.5
          If LEFT(LINHA,1)#"1" .And. !EOF()
            If CONTLIN>TOTALINHAS
              CABEC=0
              contfol=contfol+1
              FEZ2=0
              FEZ21=0
              FEZ22=0
              FEZ3=0
              FEZ4=0
              FEZ5=0
              FEZ6=0
              SELE 37
              SEEK CLIE
              REPLACE FOLHAS WITH FOLHAS+1
              oprn:cmsay(TOTALINHAS+2.0,18.0,"Continua...",oFontV8B,,nRGB(0,0,0),,0) // Continua
              contlin=0
              SELE 36
              ENDPAGE
            EndIf
          EndIf
        ENDDO
        ENDPAGE
        SELE 36
      ENDDO
      ENDPRINT
      DO WHILE .T.
        CURSORWAIT()
        If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
          fClose(nHandle)
          EXIT
        Else
          fClose(nHandle)
        EndIf
      ENDDO
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      Erase &NARQ4
      SELE 31
      USE
      Erase &NARQB
      SELE 30
    return 

    Procedure CALCCEP(VARI,WLINHA,WCOLUNA)
      DEFINE BITMAP onr1 RESOURCE "1"
      DEFINE BITMAP onr2 RESOURCE "2"
      DEFINE BITMAP onr3 RESOURCE "3"
      DEFINE BITMAP onr4 RESOURCE "4"
      DEFINE BITMAP onr5 RESOURCE "5"
      DEFINE BITMAP onr6 RESOURCE "6"
      DEFINE BITMAP onr7 RESOURCE "7"
      DEFINE BITMAP onr8 RESOURCE "8"
      DEFINE BITMAP onr9 RESOURCE "9"
      DEFINE BITMAP onr0 RESOURCE "10"
      WCEP1=Strtran(Strtran(LEFT(VARI,9),"-","")," ","")+"00000000"
      WSOMA:=WDIG:=0
      FOR I=1 TO 8
        WSOMA=WSOMA+VAL(SUBSTR(WCEP1,I,1))
      NEXT I
      WSOMA=STRZERO(WSOMA,2)
      FOR I=1 TO 10
        If RIGHT(WSOMA,1)="0"
          EXIT
        EndIf
        WDIG:=WDIG+1
        WSOMA=STRZERO(VAL(WSOMA)+1,2)
      NEXT I
      WCEP1=LEFT(WCEP1,8)+STRZERO(WDIG,1)
      WCOLUNA1=WCOLUNA
      saybitmap(WLINHA,WCOLUNA1,3.26/4,1.3/2,"10.BMP",oPrn)
      WCOLUNA1=WCOLUNA1+(0.92/5)
      FOR I=1 TO 9
        saybitmap(WLINHA,WCOLUNA1,3.26/4,1.3/2,SUBSTR(WCEP1,I,1)+".BMP",oPrn)
        WCOLUNA1=WCOLUNA1+0.92
      NEXT I
      saybitmap(WLINHA,WCOLUNA1,3.26/4,1.3/2,"10.BMP",oPrn)
      If FILE("FAC.BMP")
        oPrn:cmsay(WLINHA+0.8,WCOLUNA,"CTC MOOCA SPM PL1",oFont,,nRGB(0,0,0),,0)
      EndIf
    return 

    Procedure MARKETING()
      //** Verifica se tem texto a ser enviado aos clientes
      NARQ=WEBLOCAL+"MAILING.DAT"
      If FILE(NARQ)
        cursorwait()
        FEZALGO=.T.
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Enviando Mailing "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        oText:=TTxtfile():NEW(NARQ)
        cursorwait()
        CONTAD=0
        CONTAD=oText:RecCount()
        If VALTYPE(CONTAD)#"N"
          CONTAD=1
        EndIf
        WOPEDE:=WOPATE:=WNASCDE:=WNASCATE:=CTOD("")
        WTIPO:=WSEXO:=MZCORPO:=""
        WDIANIVER:=WMESNIVER:=0
        If CONTAD>=7
          TTT=oText:Readline()+SPACE(30)
          WOPEDE:=CTOD(SUBSTR(TTT,15,10))
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WOPATE:=CTOD(SUBSTR(TTT,16,10))
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WNASCDE:=CTOD(SUBSTR(TTT,17,10))
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WNASCATE:=CTOD(SUBSTR(TTT,18,10))
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WTIPO=SUBSTR(TTT,7,1)
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WSEXO=SUBSTR(TTT,7,1)
          oText:skip()
          TTT=oText:Readline()+SPACE(30)
          WNIVER=SUBSTR(TTT,18,5)
          WDIANIVER=VAL(LEFT(WNIVER,2))
          WMESNIVER=VAL(RIGHT(WNIVER,2))
          oText:skip()
          TTT=oText:ReadLine()+SPACE(500)
          WSSUNTO=SUBSTR(TTT,10,500)
          oText:skip()
          TTT=oText:Readline()+SPACE(50)
          WWANEXO=SUBSTR(TTT,8,50)
          oText:skip()
          oText:Skip()
          MZCORPO=""
          MZCORPO=MEMOREAD(NARQ)
          NL=AT("Corpo: ",MZCORPO)
          NL1=LEN(MZCORPO)
          MZCORPO=SUBSTR(MZCORPO,NL+7,NL1)
          oText:CLOSE()
          M->LOCAL2=TRIM("/"+CURDIR())
          M->LOCAL3=ALLTRIM(WEBLOCAL2)+"MAILING"+alltrim(dtos(date()))+left(time(),2)+substr(time(),4,2)+".DAT"
          LZCOPYFILE(NARQ,M->LOCAL3)
          CURSORWAIT()
          SELE 31
          If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
            SELE 30
            return 
          EndIf
          NARQ5=WEBLOCAL+"MCLIENTE.KEY"
          Erase &NARQ5
          SELE 31
          INDEX ON EMAIL TO &NARQ5
          SELE 32
          If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
            SELE 31
            USE
            Erase &NARQ5
            SELE 30
            return 
          EndIf
          NARQ4=WEBLOCAL+"CLIENTE.KEY"
          Erase &NARQ4
          SELE 35
          If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTE","S")
            SELE 31
            USE
            SELE 32
            USE
            SELE 30
            return 
          EndIf
          INDEX ON CODIGO TO &NARQ4
          SELE 31
          GO TOP
          MLOCAL=ALLTRIM(WEBLOCAL)+ALLTRIM(WWANEXO)
          ANTES="      "
          DO WHILE !EOF()
            SysRefresh()
            MREG=RECNO()
            CODCLI=CLIENTE
            If ANTES=EMAIL
              SELE 31
              SKIP
              LOOP
            EndIf
            ANTES=EMAIL
            SELE 35
            SEEK CODCLI
            If !FOUND()
              SELE 31
              SKIP
              LOOP
            EndIf
            If !MAILING
              SELE 31
              SKIP
              LOOP
            EndIf
            If dele()
              SELE 31
              skip
              loop
            EndIf
            If WOPEDE#CTOD("") .And. ULTNOTABOV<WOPEDE .And. ULTNOTBMF<WOPEDE
              SELE 31
              SKIP
              LOOP
            EndIf
            If WOPATE#CTOD("") .And. ULTNOTABOV>WOPATE .And. ULTNOTBMF>WOPATE
              SELE 31
              SKIP
              LOOP
            EndIf
            If WNASCDE#CTOD("") .And. NASC<WNASCDE
              SELE 31
              SKIP
              LOOP
            EndIf
            If WNASCATE#CTOD("") .And. NASC>WNASCATE
              SELE 31
              SKIP
              LOOP
            EndIf
            If WTIPO#"T" .And. TIPO#WTIPO
              SELE 31
              SKIP
              LOOP
            EndIf
            If WSEXO#"T" .And. SEXO#WSEXO
              SELE 31
              SKIP
              LOOP
            EndIf
            If WDIANIVER#0 .And. DAY(NASC)#WDIANIVER
              SELE 31
              SKIP
              LOOP
            EndIf
            If WMESNIVER#0 .And. MONTH(NASC)#WMESNIVER
              SELE 31
              SKIP
              LOOP
            EndIf
            SELE 31
            CORPOPENVIO=WEBCORPO
            If !EMPTY(MZCORPO)
              CORPOPENVIO=MZCORPO
            EndIf
            EMAILRETORNO=CLIENTE->MAILNEWS
            If EMPTY(EMAILRETORNO)
              EMAILRETORNO=EMAIL
            EndIf
            WCORPOENVIO=Strtran(mzcorpo,";;NOME;:",TRIM(NOME))
            WCORPOENVIO=Strtran(WCORPOENVIO,";;CODIGO;:",STRZERO(CODCLI,8))
            If CODCLI#0
              WCORPOENVIO=WCORPOENVIO+'<div align=CENTER>Para deixar de receber a newsletter <a href="https://exatus.net/administrativo/bolsasnet/news_Erase.asp?corretora='+WEBCOR+'&bolsa='+WEBBOLSA+'&codigo='+STRZERO(CODCLI,8)+'">Clique aqui</a></div>'
            EndIf
            DECLARE TX1[1]
            TX1[1]="Enviado Mailing : "+STRZERO(CODCLI,8)+" - "+CLIENTE->NOME+CHR(13)+CHR(10)+EMAILRETORNO+chr(13)+chr(10)
            mwzn=tx1[1]
            If LEFT(EMAIL,1)#"[" .And. LEFT(EMAIL,1)#"{"
              If AT("@",EMAIL)=0
                SELE 31
                SKIP
                LOOP
              EndIf
              WWASSUNTO=Strtran(WSSUNTO,";;NOME;:",TRIM(NOME))
              WWASSUNTO=Strtran(WWASSUNTO,";;CODIGO;:",STRZERO(CODCLI,8))
              If EMPTY(WWASSUNTO) .OR. EMPTY(WCORPOENVIO)
                SELE 31
                SKIP
                LOOP
              EndIf
              WCORPOENVIO=TRIM(WCORPOENVIO)
              If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
                ENVIAEMAIL(EMAILRETORNO,WWASSUNTO,WCORPOENVIO,MLOCAL,"",WEBMAIL,WEBSENHA,WEBSERV,WEBLOGIN,WEBFORCA)
                If !EMPTY(WDADOS)
                  DECLARE TX1[1]
                  TX1[1]=mwzn
                  LOGFILE(WDADOS,TX1)
                EndIf
              EndIf
            EndIf
            SELE 31
            SKIP
          ENDDO
          M->LOCAL21=TRIM("/"+CURDIR())
          M->LOCAL3=ALLTRIM(WEBLOCAL2)+ALLTRIM(WWANEXO)+alltrim(dtos(date()))+left(time(),2)+substr(time(),4,2)
          LZCOPYFILE(MLOCAL,M->LOCAL3)
          LCHDIR(M->LOCAL21)
          Erase &MLOCAL
          SELE 31
          USE
          Erase &NARQ5
          SELE 32
          USE
          SELE 35
          USE
          Erase &NARQ4
          SELE 30
        EndIf
        Erase &NARQ
        LCHDIR(M->LOCAL2)
        EMAILRETORNO=WEBMKT
        If EMPTY(EMAILRETORNO)
          EMAILRETORNO=WEBREL
        EndIf
        If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
          ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - Newsletter ","Newsletter",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
        cursorarrow()
      EndIf
      SELE 30
    return 

    Procedure ATUACCIN()
      SELE 30
      //** Arquivo CCIN (cadastro completo)
      DECLARE VLABEL[ADIR(WEBLOCAL+"CCIN*.*")]
      ADIR(WEBLOCAL+"CCIN*.*",VLABEL)
      FOR OI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
        Erase &narq
      NEXT OI
      return 


      If LEN(VLABEL)#0
        FEZALGO=.T.
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Recendo arquivos CCIN "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        NARQ=WEBLOCAL+"BACENJUD.DAT"
        If !FILE(NARQ)
          MEMOWRIT(NARQ,"0"+CHR(13)+CHR(10))
        EndIf
        SELE 32
        NARQ4=WEBLOCAL+"CLIENTE.KEY"
        Erase &NARQ4
        If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTE","S")
          SELE 34
          USE
          Erase &NARQ4
          SELE 30
          return 
        EndIf
        INDEX ON CODIGO TO &NARQ4
      EndIf
      FOR OI=1 TO LEN(VLABEL)
        SysRefresh()
        CURSORWAIT()
        NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
        If RIGHT(NARQ,4)<>".TXT"
          ARQZ=NARQ+".TXT"
          FRENAME(NARQ,ARQZ)
          NARQ=ARQZ
        EndIf
        NARQ1=ALLTRIM(WEBLOCAL)+"TEMP.DBF"
        Erase &NARQ1
        mat_stru={}
        aAdd(mat_stru,{"C01","N",02,0})
        aAdd(mat_stru,{"C02","C",01,0})
        aAdd(mat_stru,{"C03","N",15,0})
        aAdd(mat_stru,{"C04","C",08,0})
        aAdd(mat_stru,{"C05","N",02,0})
        aAdd(mat_stru,{"C06","C",01,0})
        aAdd(mat_stru,{"C07","C",60,0})
        aAdd(mat_stru,{"C08","C",16,0})
        aAdd(mat_stru,{"C09","C",02,0})
        aAdd(mat_stru,{"C10","C",04,0})
        aAdd(mat_stru,{"C11","C",01,0})
        aAdd(mat_stru,{"C12","N",01,0})
        aAdd(mat_stru,{"C13","C",01,0})
        aAdd(mat_stru,{"C14","C",60,0})
        aAdd(mat_stru,{"C15","N",01,0})
        aAdd(mat_stru,{"C16","C",60,0})
        aAdd(mat_stru,{"C17","N",01,0})
        aAdd(mat_stru,{"C18","C",02,0})
        aAdd(mat_stru,{"C19","N",03,0})
        aAdd(mat_stru,{"C20","C",01,0})
        aAdd(mat_stru,{"C21","C",01,0})
        aAdd(mat_stru,{"C22","N",01,0})
        aAdd(mat_stru,{"C23","N",01,0})
        aAdd(mat_stru,{"C24","C",03,0})
        aAdd(mat_stru,{"C25","N",05,0})
        aAdd(mat_stru,{"C26","N",07,0})
        aAdd(mat_stru,{"C27","C",13,0})
        aAdd(mat_stru,{"C28","C",02,0})
        aAdd(mat_stru,{"C29","N",05,0})
        aAdd(mat_stru,{"C30","N",10,0})
        aAdd(mat_stru,{"C31","N",03,0})
        aAdd(mat_stru,{"C32","N",05,0})
        aAdd(mat_stru,{"C33","C",13,0})
        aAdd(mat_stru,{"C34","N",01,0})
        aAdd(mat_stru,{"C35","N",01,0})
        aAdd(mat_stru,{"C36","N",05,0})
        aAdd(mat_stru,{"C37","C",24,0})
        aAdd(mat_stru,{"C38","C",01,0})
        aAdd(mat_stru,{"C39","C",30,0})
        aAdd(mat_stru,{"C40","C",05,0})
        aAdd(mat_stru,{"C41","C",10,0})
        aAdd(mat_stru,{"C42","C",18,0})
        aAdd(mat_stru,{"C43","C",28,0})
        aAdd(mat_stru,{"C44","C",02,0})
        aAdd(mat_stru,{"C45","N",09,0})
        aAdd(mat_stru,{"C46","N",07,0})
        aAdd(mat_stru,{"C47","C",09,0})
        aAdd(mat_stru,{"C48","N",05,0})
        aAdd(mat_stru,{"C49","N",07,0})
        aAdd(mat_stru,{"C50","N",09,0})
        aAdd(mat_stru,{"C51","C",07,0})
        aAdd(mat_stru,{"C52","C",01,0})
        aAdd(mat_stru,{"C53","C",01,0})
        aAdd(mat_stru,{"C54","N",05,0})
        aAdd(mat_stru,{"C55","N",07,0})
        aAdd(mat_stru,{"C56","N",01,0})
        aAdd(mat_stru,{"C57","N",01,0})
        aAdd(mat_stru,{"C58","N",15,0})
        aAdd(mat_stru,{"C59","C",01,0})
        aAdd(mat_stru,{"C60","N",07,0})
        aAdd(mat_stru,{"C61","N",05,0})
        aAdd(mat_stru,{"C62","C",01,0})
        aAdd(mat_stru,{"C63","C",02,0})
        aAdd(mat_stru,{"C64","C",02,0})
        aAdd(mat_stru,{"C65","C",16,0})
        aAdd(mat_stru,{"C66","C",04,0})
        aAdd(mat_stru,{"C67","C",02,0})
        aAdd(mat_stru,{"C68","N",19,0})
        aAdd(mat_stru,{"C69","C",50,0})
        aAdd(mat_stru,{"C70","C",01,0})
        aAdd(mat_stru,{"C71","C",70,0})
        DBCREATE(NARQ1,mat_stru,"DBFNTX")
        RELEASE MAT_STRU
        SELE 34
        If !USEREDE(NARQ1,.T.,10,"TEMPC","S")
          SELE 34
          USE
          Erase &NARQ1
          SELE 30
          return 
        EndIf
        APPEND FROM &NARQ SDF
        GO TOP
        SELE 34
        DO WHILE !EOF()
          CURSORWAIT()
          If !(C06$"FJ") .OR. C26=0
            SKIP
            LOOP
          EndIf
          If C26#0
            SELE 32
            SEEK TEMPC->C26
            If !FOUND() .And. TEMPC->C02="D"
              SELE 34
              SKIP
              LOOP
            EndIf
            If !FOUND()
              ADIREG(0)
              REPLACE EEMAIL WITH .T.
              REPLACE MAILING WITH .T.
              REPLACE EIMPRE WITH .F.
              REPLACE EGUARD WITH .T.
              If TEMPC->C07="BB BANCO"
                REPLACE FRENTEV WITH .T.
              EndIf
            EndIf
            If REGLOCK(10)
              If TEMPC->C02="D"
                REPLACE EXCLUIDO WITH DTOC(DATE())+" "+TIME()
                DELE
              EndIf
              If TEMPC->C02$"IACVXR"
                If TEMPC->C02="R"
                  RECALL
                EndIf
                REPLACE CODIGO WITH TEMPC->C26
                REPLACE CNPJCPF WITH strzero(TEMPC->C03,14)
                REPLACE NOME WITH TEMPC->C07
                REPLACE APELIDO WITH ""
                REPLACE NASC WITH CTOD(LEFT(TEMPC->C04,2)+"/"+SUBSTR(TEMPC->C04,3,2)+"/"+RIGHT(TEMPC->C04,4))
                REPLACE TIPO WITH TEMPC->C06
                REPLACE SEXO WITH TEMPC->C11
                REPLACE ECIVIL WITH strzero(TEMPC->C12,1)
                REPLACE CONJUGE WITH TEMPC->C14
                REPLACE BANCO WITH TEMPC->C31
                REPLACE ASSESSOR WITH TEMPC->C36
                REPLACE ENDERECO WITH TEMPC->C39
                REPLACE NUMERO WITH TEMPC->C40
                REPLACE COMPL WITH TEMPC->C41
                REPLACE BAIRRO WITH TEMPC->C42
                REPLACE CIDADE WITH TEMPC->C43
                REPLACE UF WITH TEMPC->C44
                REPLACE CEP WITH strzero(TEMPC->C45,8)
                REPLACE TIPOINV WITH TEMPC->C61
                If EMPTY(SENHA)
                  REPLACE SENHA WITH STRZERO(NRANDOM()*10,8)
                EndIf
                If TEMPC->C02="I"
                  REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
                  REPLACE USERINC WITH "CCIN"
                Else
                  REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
                  REPLACE USERALT WITH "CCIN"
                EndIf
                UNLOCK
              EndIf
              DECLARE TX1[1]
              TX1[1]=TEMPC->C02+STRZERO(CODIGO,8)+"-"+NOME
              If !EMPTY(WDADOS)
                LOGFILE(WDADOS,TX1)
              EndIf
            EndIf
          EndIf
          SELE 34
          SKIP
        ENDDO
        SELE 34
        USE
        M->LOCAL2=TRIM("/"+CURDIR())
        M->LOCAL3=ALLTRIM(WEBLOCAL2)+VLABEL[OI]
        LZCOPYFILE(NARQ,M->LOCAL3)
        LCHDIR(M->LOCAL2)
        Erase &NARQ
        Erase &NARQ1
      NEXT OI
      If LEN(VLABEL)#0
        SELE 32
        USE
        Erase &NARQ4
        EMAILRETORNO=WEBCAD
        If SEENVIAEMAIL="S" .And. !EMPTY(EMAILRETORNO)
          ENVIAEMAIL(ALLTRIM(EMAILRETORNO),"BolsasNet - CCIN (Log) ","Recebido arquivos CCIN",WDADOS,"","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        EndIf
      EndIf
    return 

    Procedure EXTRANUAL()
      DECLARE VLABEL[ADIR(WEBLOCAL+"extratotarifa*.TXT")]
      ADIR(WEBLOCAL+"*.TXT",VLABEL)
      FOR OI=1 TO LEN(VLABEL)
        narq=alltrim(weblocal)+alltrim(vlabel[oi])
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Extrato Anual de Taxas "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        If SEENVIAEMAIL="S"
          EXTRANUAL1(1) //enviar emails
        Else
          EXTRANUAL1(1) // gerar arquivo /// COLOCAR 2...
        EndIf
      NEXT OI
      SELE 30
    return 

    Procedure EXTRANUAL1(QTP)
      FEZALGO=.T.
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 30
        Erase &NARQ4
        return 
      EndIf
      SET INDEX TO &NARQ4
      SELE 31
      USE
      NARQB=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQB
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 35
        USE
        Erase &NARQ4
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQB
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 30
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",321,0})
      ARQTXTO=WEBLOCAL1+"TAXAS.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      narq7=WEBLOCAL4+"ETIQUETA.KEY"
      Erase &NARQ7
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"EXTRATO","S")
        SELE 35
        USE
        Erase &NARQ4
        SELE 31
        USE
        Erase &NARQB
        SELE 37
        USE
        Erase &NARQ7
        SELE 32
        USE
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      APPEND FROM &NARQ SDF
      NARQ1=WEBLOCAL2+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
      LZCOPYFILE(NARQ,NARQ1)
      Erase &NARQ
      GO TOP
      If QTP=1
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        ZARQ1=WEBLOCAL+"TAXAS"+DTOS(DATE())+".PDF"
        Erase &ZARQ1
        oReg:Set("PDFFilename",ZARQ1)
        oReg:Close()
        PRINT oPrn NAME "Extrato anual de Tarifas" TO TEMPDF1
        DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-9 OF oPrn //8
        DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-8.0 BOLD OF oPrn //8
        DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-7.5 OF oPrn  // 6
        DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-7.5 BOLD OF oprn // 6
        DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-10 BOLD OF oPrn //8
        DEFINE FONT oFont NAME "courier new negrito" SIZE 0,8.3 OF oPrn //8.3
        DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
        DEFINE BRUSH oCinza       COLOR nRGB(205,205,205)
        DEFINE BRUSH oCinzaEscuro COLOR nRGB(168,168,168) // 99,99,99)
        DEFINE BRUSH oAzul COLOR nRGB(0,75,133)
        DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
        DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      EndIf
      If QTP=2
        oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
        ZARQ1=weblocal+"TAXAS"+DTOS(DATE())+".PDF"
        Erase &ZARQ1
        oReg:Set("PDFFilename",ZARQ1)
        oReg:Close()
        PRINT oPrn NAME "Extrato anual de Tarifas" TO TEMPDF1
        DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-9 BOLD OF oPrn //8 NBOLD
        DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-8.0 BOLD OF oPrn //8
        DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-7.5 BOLD OF oPrn //6 NBOLD
        DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-7.5 BOLD OF oprn //6
        DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-10  OF oPrn //8
        DEFINE FONT oFont NAME "courier new negrito" SIZE 0,8.3 BOLD OF oPrn //8.3 NBOLD
        DEFINE FONT oFontv NAME "verdana" SIZE 0,12 OF oPrn // NBOLD
        DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
        DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
        DEFINE BRUSH oCinza       COLOR nRGB(205,205,205)
        DEFINE BRUSH oCinzaEscuro COLOR nRGB(0,0,0) // 168,168,168)
        DEFINE BRUSH oAzul COLOR nRGB(0,0,0) // 35,35,142)
        DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
        DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
        nRowStep = oPrn:nVertRes() / 66
        nColStep = oPrn:nHorzRes() / 80
      EndIf
      SELE 36
      go top
      DO WHILE !EOF()
        If LEFT(LINHA,11)#"0001Cliente"
          SKIP
          LOOP
        EndIf
        CLIE=VAL(SUBSTR(LINHA,12,7))
        CLIE1=ALLTRIM(STR(CLIE,10,0))+"  "
        WWNOME=SUBSTR(LINHA,19,60)
        SKIP
        wende1=substr(linha,12,100)
        SKIP
        wende2=SUBSTR(LINHA,12,100)
        SKIP
        anobase=SUBSTR(LINHA,12,4)
        SKIP
        SELE 35
        SEEK CLIE
        If !FOUND()
          ADIREG(0)
          REPLACE CODIGO WITH CLIE
          REPLACE NOME WITH WWNOME
          REPLACE ENDE1 WITH WENDE1
          REPLACE ENDE2 WITH WENDE2
          REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
          REPLACE USERINC WITH "Relat.Anual de Tarifas"
        Else
          If REGLOCK(10)
            REPLACE CODIGO WITH CLIE
            REPLACE NOME WITH WWNOME
            REPLACE ENDE1 WITH WENDE1
            REPLACE ENDE2 WITH WENDE2
            REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
            REPLACE USERALT WITH "Relat.Anual de Tarifas"
          EndIf
        EndIf
        SELE 35
        SEEK CLIE
        If NOME="BB BANCO" .And. reglock(10)
          REPLACE FRENTEV WITH .T.
        EndIf
        SELE 36
        CX=0
        CABEC=0
        WANT="02"
        contfol=1
        contlin=0
        DECLARE TIPOS[19]
        TIPOS[1] ="Corretagem - Bovespa"
        TIPOS[2] ="Corretagem s/Operações Bovespa"
        TIPOS[3] ="Corretagem s/Liq.de Operação"
        TIPOS[4] ="Taxa Operacional - BMF"
        TIPOS[5] ="Corretagem - BM&F"
        TIPOS[6] ="Custódia - Bovespa"
        TIPOS[7] ="Custódia - Ouro"
        TIPOS[8] ="Custódia - Tesouro Direto"
        TIPOS[9] ="Comissão BTC"
        TIPOS[10]="Emissão de TED"
        TIPOS[11]="Transferência de Títulos"
        TIPOS[12]="Juros"
        TIPOS[13]="Multa s/ Saldo Devedor"
        TIPOS[14]="Calculadora de IR"
        TIPOS[15]="Custódia & Tesouro Direto"
        TIPOS[16]="Custódia & Tesouro Direto"
        TIPOS[17]="Custódia & Tesouro Direto"
        TIPOS[18]="Transferência de Posição"
        TIPOS[19]="Total Geral"
        PAGE
        If QTP=1
          saybitmap(0.0,0.0,3.26*6.5,1.3*5.0,WEBLOCAL+"TAXAS.BMP",oPrn)
        ElseIf QTP=2
          saybitmap(0.0,0.0,3.26*6.5,1.3*5.0,WEBLOCAL+"TAXAS.BMP",oPrn)
        EndIf

        oprn:cmsay(2.0,1.1,"Cliente: "+ALLTRIM(TRANSFORM(CLIE,"999999999")),oFontV8,,nRGB(255,255,255),,0) //Código do Cliente
        oprn:cmsay(2.0,7.0,"Assessor - 000",oFontV8,,nRGB(255,255,255),,0) //Assessor

        caja(4.0,1.0,4.5,1.2,oPrn,1,oCinza,oPen1) //caixa cinza escuro
        caja(4.0,5.5,4.5,20.5,oPrn,1,oCinza,oPen1) //caixa cinza escuro

        oprn:cmsay(4.1,1.3,"Cód. Cliente - "+strzero(clie,7),oFontV8,,nRGB(0,0,0),,0) //Código do Cliente
        oprn:cmsay(4.1,5.6,wwnome,oFontv8,,nRGB(0,0,0),,0)
        oprn:cmsay(4.1,17.5,"Ano Base: "+ANOBASE,oFontv8,,nRGB(0,0,0),,0) //Código do Cliente
        CONTLIN=5.0
        caja(29.2,1.1,30,17,oPrn,1,oCinza,oPen1) //caixa cinza
        oprn:cmsay(29.3,1.9,"Ouvidoria - 0800 724 30 20",oFontV8b,,nRGB(0,0,0),,0) //Titulo
        oPrn:cmsay(29.3,17,"  Folha: 001 de 001",oFontV8B,,nRGB(0,0,0),,0)
        CABEC=1
        FOR OZ=1 TO 19
          If TIPOS[OZ]="NAOFAZ"
            LOOP
          EndIf
          WTOTS=0
          If VAL(LEFT(LINHA,4))=OZ .or. (val(left(linha,4))=9999 .And. oz=19)
            WTOTS=VAL(Strtran(Strtran(SUBSTR(LINHA,54,14),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,80,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,103,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,126,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,149,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,172,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,192,14),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,218,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,241,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,264,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,287,11),".",""),",","."))+VAL(Strtran(Strtran(SUBSTR(LINHA,310,11),".",""),",","."))
          EndIf
          If WTOTS<>0
            caja(CONTLIN,1.0,CONTLIN+0.6,6.5,oPrn,1,oAzul,oPen1) //caixa azul
            oprn:cmsay(CONTLIN+0.1,1.3,tipos[oz],oFontV8B,,nRGB(255,255,255),,0) //Titulo
            oprn:cmsay(CONTLIN+0.1,7.5,"Jan            Fev            Mar           Abr            Mai           Jun",oFontc8B,,nRGB(0,0,0),,0) //Titulo
            CONTLIN=CONTLIN+0.20
            If VAL(LEFT(LINHA,4))=OZ .or. (val(left(linha,4))=9999 .And. oz=19)
              MZ1=SUBSTR(LINHA,54,14)+SUBSTR(LINHA,80,11)+SUBSTR(LINHA,103,11)+SUBSTR(LINHA,126,11)+SUBSTR(LINHA,149,11)+SUBSTR(LINHA,172,11)
              oprn:cmsay(CONTLIN+0.23,6.0,MZ1,oFontt8,,nRGB(0,0,0),,0) // Conteudo Jan-Jun
            Else
              oPrn:cmsay(CONTLIN+0.23,6.0,"          0,00       0,00       0,00       0,00       0,00       0,00",oFontt8,,nRGB(0,0,0),,0)
            EndIf
            CONTLIN=CONTLIN+0.43
            oprn:cmsay(CONTLIN+0.1,7.5,"Jul            Ago            Set           Out            Nov           Dez",oFontc8B,,nRGB(0,0,0),,0) //Titulo
            CONTLIN=CONTLIN+0.20
            caja(contlin,1.0,contlin+0.6,3.0,oPrn,1,oCinza,oPen1) //caixa do total
            caja(contlin,3.0,contlin+0.6,6.5,oPrn,1,oCinzaescuro,oPen1) // caixa
            oPrn:cmsay(CONTLIN+0.23,1.3,"Total",oFontV8b,,nRGB(0,0,0),,0)
            oPrn:cmsay(CONTLIN+0.23,3.0,TRANSFORM(WTOTS,"@E 9999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
            If VAL(LEFT(LINHA,4))=OZ  .or. (val(left(linha,4))=9999 .And. oz=19)
              MZ1=SUBSTR(LINHA,192,14)+SUBSTR(LINHA,218,11)+SUBSTR(LINHA,241,11)+SUBSTR(LINHA,264,11)+SUBSTR(LINHA,287,11)+SUBSTR(LINHA,310,11)
              oprn:cmsay(CONTLIN+0.23,6.0,MZ1,oFontt8,,nRGB(0,0,0),,0) // Conteudo Jul-Dez
            Else
              oPrn:cmsay(CONTLIN+0.23,6.0,"          0,00       0,00       0,00       0,00       0,00       0,00",oFontt8,,nRGB(0,0,0),,0)
            EndIf
          EndIf
          If VAL(LEFT(LINHA,4))=OZ .or. (val(left(linha,4))=9999 .And. oz=19)
            SKIP
          EndIf
          If WTOTS<>0
            CONTLIN=CONTLIN+0.58
            caja(CONTLIN,1.0,CONTLIN+0.02,20.6,oPrn,0,oPreto,oPen1) //linha preta
            CONTLIN=CONTLIN+0.20
          EndIf
        NEXT OZ
        ENDPAGE
      ENDDO
      ENDPRINT

      DO WHILE .T.
        CURSORWAIT()
        If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
          fClose(nHandle)
          EXIT
        Else
          fClose(nHandle)
        EndIf
      ENDDO
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      Erase &NARQ4
      SELE 31
      USE
      Erase &NARQB
    return 

    Procedure CADCSV()
      CURSORWAIT()
      DECLARE VLABEL[ADIR(WEBLOCAL+"NEWS*.CSV")]
      ADIR(WEBLOCAL+"NEWS*.CSV",VLABEL)
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ6
      If LEN(VLABEL)=0
        DECLARE VLABEL[ADIR(WEBLOCAL+"5_NEWS*.CSV")]
        ADIR(WEBLOCAL+"5_NEWS*.CSV",VLABEL)
      EndIf
      If LEN(VLABEL)#0
        SELE 35
        If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTE","S")
          SELE 30
          return 
        EndIf
        ZAP
        SELE 32
        If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
          SELE 35
          USE
          SELE 30
          return 
        EndIf
        ZAP
        SysRefresh()
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Arq. NEWS CSV Clientes "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
          NARQ1=ALLTRIM(WEBLOCAL2)+UPPER(ALLTRIM(VLABEL[OI]))
          LZCOPYFILE(NARQ,NARQ1)
          NARQ2=ALLTRIM(WEBLOCAL1)+"ATUACAD.DBF"
          Erase &NARQ2
          SELE 37
          mat_stru={}
          Aadd(mat_stru,{"LINHA","C",500,0})
          DBCREATE(NARQ2,mat_stru)
          USE &NARQ2
          APPEND FROM &NARQ1 SDF
          GO TOP
          DO WHILE !EOF()
            SysRefresh()
            WLI=ALLTRIM(LINHA)+SPACE(450)
            WCOD=0
            If AT("5_NEWS",UPPER(ALLTRIM(VLABEL[OI])))=0
              COLU=AT(";",WLI)
              WNOME=ALLTRIM(SINAL(UPPER(LEFT(WLI,COLU-1))))
              WLI=SUBSTR(WLI,COLU+1,450)
              COLU=AT(";",WLI)
              WEMAIL=ALLTRIM(sinal(LOWER(WLI)))
              WLI=LEN(WEMAIL)
            Else
              COLU=AT(";",WLI)
              WCOD=VAL(ALLTRIM(SINAL(UPPER(LEFT(WLI,COLU-1)))))
              WLI=SUBSTR(WLI,COLU+1,450)
              COLU=AT(";",WLI)
              WNOME=ALLTRIM(SINAL(UPPER(LEFT(WLI,COLU-1))))
              WLI=SUBSTR(WLI,COLU+1,450)
              WEMAIL=ALLTRIM(sinal(LOWER(WLI)))
              WEMAIL=Strtran(WEMAIL,'"',"")
              WEMAIL=Strtran(WEMAIL,' ',"")
              WLI=LEN(WEMAIL)
            EndIf
            If WLI>0 .And. AT(" ",WEMAIL)=0 .And. LEN(WNOME)>0
              If RIGHT(WEMAIL,1)=";"
                WEMAIL=LEFT(WEMAIL,WLI-1)
              EndIf
              SELE 35
              ADIREG(0)
              If WCOD=0
                WCOD=RECNO()+9999
              EndIf
              REPLACE CODIGO WITH WCOD
              REPLACE NOME WITH WNOME
              REPLACE MAILING WITH .T.
              REPLACE EIMPRE  WITH .T.
              REPLACE EEMAIL  WITH .T.
              REPLACE EGUARD  WITH .T.
              REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
              REPLACE USERALT WITH "Arqv.CSV news"
              DECLARE TX1[1]
              SELE 32
              ADIREG(0)
              REPLACE CODIGO WITH RECNO()
              REPLACE CLIENTE WITH WCOD
              REPLACE NOME WITH WNOME
              REPLACE EMAIL WITH WEMAIL
            EndIf
            SELE 37
            SKIP
          ENDDO
          SELE 37
          USE
          Erase &NARQ2
          Erase &NARQ
        NEXT OI
        SELE 35
        USE
        SELE 32
        USE
      EndIf
      SELE 30
    return 

    Procedure CADPLANNER()
      CURSORWAIT()
      DECLARE VLABEL[ADIR(WEBLOCAL+"CSVEMAIL*.CSV")]
      ADIR(WEBLOCAL+"CSVEMAIL*.CSV",VLABEL)
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ6
      If LEN(VLABEL)#0
        SELE 35
        If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTE","S")
          SELE 30
          return 
        EndIf
        INDEX ON CODIGO TO &NARQ4
        SELE 32
        If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
          SELE 35
          USE
          SELE 30
          return 
        EndIf
        zap
        SELE 35
        SysRefresh()
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Arq. PLANNER EMAIL CSV Clientes "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
          NARQ1=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(NARQ,NARQ1)
          NARQ2=ALLTRIM(WEBLOCAL1)+"ATUACAD.DBF"
          Erase &NARQ2
          SELE 37
          mat_stru={}
          Aadd(mat_stru,{"LINHA","C",450,0})
          DBCREATE(NARQ2,mat_stru)
          USE &NARQ2
          APPEND FROM &NARQ1 SDF
          GO TOP
          SKIP
          DO WHILE !EOF()
            SysRefresh()
            WLI=ALLTRIM(LINHA)+',"'+SPACE(450)
            COLU=AT(',"',WLI)
            WCOD=VAL(LEFT(WLI,COLU-1))
            WLI=SUBSTR(WLI,COLU+2,450)
            COLU=AT('","',WLI)
            WNOME=ALLTRIM(SINAL(UPPER(LEFT(WLI,COLU-1))))
            WLI=SUBSTR(WLI,COLU+3,450)
            COLU=AT('","',WLI)
            WCPFCNPJ=ALLTRIM(LEFT(WLI,COLU-1))
            WLI=SUBSTR(WLI,COLU+3,450)
            COLU=AT('",',WLI)
            WDTNASC=ALLTRIM(LEFT(WLI,COLU-1))
            WLI=SUBSTR(WLI,COLU+2,450)
            COLU=AT(',"',WLI)
            WDTNASC=ALLTRIM(LEFT(WLI,COLU-1))
            WLI=SUBSTR(WLI,COLU+2,450)
            COLU=AT('","',WLI)
            WEMAIL=LOWER(SINAL(ALLTRIM(LEFT(WLI,COLU-1))))
            WLI=SUBSTR(WLI,COLU+2,450)
            COLU=AT('","',WLI)
            WCNPJ=ALLTRIM(TRANSFORM(VAL(WCPFCNPJ),"99999999999999"))
            WSENHA=""
            If val(WCNPJ)>99999999999
              WSENHA=STRZERO(VAL(WCNPJ),14) // aqui o sistema coloca zero a esquerda(14) digitos
            Else
              If CIC(STRZERO(VAL(WCNPJ),11))
                WSENHA=STRZERO(VAL(WCNPJ),11)
              Else
                WSENHA=STRZERO(VAL(WCNPJ),14)
              EndIf
            EndIf
            WLI=LEN(WEMAIL)
            If WLI>0 .And. AT(" ",WEMAIL)=0 .And. LEN(WNOME)>0 .And. WCOD>0 .And. at("@",wemail)>0
              If RIGHT(WEMAIL,1)=";"
                WEMAIL=LEFT(WEMAIL,WLI-1)
              EndIf
              SELE 35
              seek wcod
              If !found()
                ADIREG(0)
                REPLACE CODIGO WITH WCOD
                REPLACE NOME WITH WNOME
                REPLACE CNPJCPF WITH STRZERO(VAL(WCPFCNPJ),14)
                REPLACE EEMAIL  WITH .T.
                REPLACE EGUARD  WITH .T.
                REPLACE MAILING WITH .T.
                REPLACE SENHA WITH WSENHA
                REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
                REPLACE USERALT WITH "Arqv.CSV emails"
              Else
                If REGLOCK(10)
                  REPLACE NOME WITH WNOME
                  REPLACE SENHA WITH WSENHA
                  REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
                EndIf
              EndIf
              DECLARE TX1[1]
              SELE 32
              ADIREG(0)
              REPLACE CODIGO WITH RECNO()
              REPLACE CLIENTE WITH WCOD
              REPLACE NOME WITH WNOME
              REPLACE EMAIL WITH WEMAIL
            EndIf
            SELE 37
            SKIP
          ENDDO
          SELE 37
          USE
          Erase &NARQ2
          Erase &NARQ
        NEXT OI
        SELE 35
        USE
        SELE 32
        USE
      EndIf
      SELE 30
    return 

    Procedure CADCSVCLI()
      CURSORWAIT()
      DECLARE VLABEL[ADIR(WEBLOCAL+"EMAIL*.CSV")]
      ADIR(WEBLOCAL+"EMAIL*.CSV",VLABEL)
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      Erase &NARQ4
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ6
      If LEN(VLABEL)#0
        SELE 35
        If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTE","S")
          SELE 30
          return 
        EndIf
        INDEX ON CODIGO TO &NARQ4
        SELE 32
        If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
          SELE 35
          USE
          SELE 30
          return 
        EndIf
        zap
        SELE 35
        SysRefresh()
        DECLARE TX1[1]
        TX1[1]="BolsasNet - Arq. EMAIL CSV Clientes "+WEBCOR
        LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
        FOR OI=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=alltrim(WEBLOCAL)+alltrim(VLABEL[OI])
          NARQ1=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(NARQ,NARQ1)
          NARQ2=ALLTRIM(WEBLOCAL1)+"ATUACAD.DBF"
          Erase &NARQ2
          SELE 37
          mat_stru={}
          Aadd(mat_stru,{"LINHA","C",450,0})
          DBCREATE(NARQ2,mat_stru)
          USE &NARQ2
          APPEND FROM &NARQ1 SDF
          GO TOP
          SKIP
          DO WHILE !EOF()
            SysRefresh()
            WLI=ALLTRIM(LINHA)+";"+SPACE(450)
            COLU=AT(";",WLI)
            WCOD=VAL(LEFT(WLI,COLU-1))
            WLI=SUBSTR(WLI,COLU+1,450)
            COLU=AT(";",WLI)
            WNOME=ALLTRIM(SINAL(UPPER(LEFT(WLI,COLU-1))))
            WLI=SUBSTR(WLI,COLU+1,450)
            COLU=AT(";",WLI)
            WEMAIL=LOWER(SINAL(ALLTRIM(LEFT(WLI,COLU-1))))
            WLI=LEN(WEMAIL)
            If WLI>0 .And. AT(" ",WEMAIL)=0 .And. LEN(WNOME)>0
              If RIGHT(WEMAIL,1)=";"
                WEMAIL=LEFT(WEMAIL,WLI-1)
              EndIf
              SELE 35
              seek wcod
              If !found()
                ADIREG(0)
                REPLACE CODIGO WITH WCOD
                REPLACE NOME WITH WNOME
                REPLACE EEMAIL  WITH .T.
                REPLACE EGUARD  WITH .T.
                REPLACE MAILING WITH .T.
                REPLACE ALTERADO WITH DTOC(DATE())+" "+TIME()
                REPLACE USERALT WITH "Arqv.CSV emails"
              EndIf
              DECLARE TX1[1]
              SELE 32
              ADIREG(0)
              REPLACE CODIGO WITH RECNO()
              REPLACE CLIENTE WITH WCOD
              REPLACE NOME WITH WNOME
              REPLACE EMAIL WITH WEMAIL
            EndIf
            SELE 37
            SKIP
          ENDDO
          SELE 37
          USE
          Erase &NARQ2
          Erase &NARQ
        NEXT OI
        SELE 35
        USE
        SELE 32
        USE
      EndIf
      SELE 30
    return 

    Procedure INFORMENG()
      DECLARE VLABEL[ADIR(WEBLOCAL+"INFORME_REND_*.*")]
      ADIR(WEBLOCAL+"INFORME_REND_*.*",vlabel)
      VLABEL:=ASORT(VLABEL)
      If LEN(VLABEL)=0
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Informes de Rendimento Geração "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      FEZALGO=.T.
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",900,0})
      Aadd(mat_stru,{"CLUBE","C",50,0})
      Aadd(mat_stru,{"CNPJ","C",14,0})
      Aadd(mat_stru,{"ENDCLU","C",60,0})
      Aadd(mat_stru,{"BAIRCLU","C",40,0})
      Aadd(mat_stru,{"CIDCLU","C",40,0})
      Aadd(mat_stru,{"UFCLU","C",2,0})
      Aadd(mat_stru,{"CEPCLU","C",8,0})
      Aadd(mat_stru,{"CODBEN","C",9,0})
      Aadd(mat_stru,{"NOMEBEN","C",100,0})
      Aadd(mat_stru,{"ENDEBEN","C",80,0})
      Aadd(mat_stru,{"BAIRBEN","C",30,0})
      Aadd(mat_stru,{"CIDBEN","C",30,0})
      Aadd(mat_stru,{"UFBEN","C",2,0})
      Aadd(mat_stru,{"CEPBEN","C",8,0})
      Aadd(mat_stru,{"SALDOANT","C",17,0})
      Aadd(mat_stru,{"SALDOATU","C",17,0})
      Aadd(mat_stru,{"RENDIMENT","C",17,0})
      Aadd(mat_stru,{"DIVIDENDO","C",17,0})
      Aadd(mat_stru,{"CREDITO","C",17,0})
      Aadd(mat_stru,{"CPF","C",14,0})
      ARQTXTO=WEBLOCAL1+"INFREND.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      SELE 31
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQ5
      USE
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      SET INDEX TO &NARQ5
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ4
      Erase &NARQ6
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      INDEX ON STRZERO(VAL(CNPJCPF),14)+IIF(TEMMAIL,"1","2") TO &NARQ6
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 31
        USE
        SELE 30
        Erase &NARQ4
        Erase &NARQ6
        return 
      EndIf
      SET INDEX TO &NARQ4,&NARQ6
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"APLIC","S")
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      SELE 36
      FOR I=1 TO LEN(VLABEL)
        NARQ=WEBLOCAL+VLABEL[I]
        APPEND FROM &NARQ SDF
        NARQ1=WEBLOCAL2+VLABEL[I]
        LZCOPYFILE(NARQ,NARQ1)
        Erase &NARQ
      NEXT I
      SELE 36
      GO TOP
      DO WHILE !EOF()
        If LINHA="CLUBE/FUNDO;CNPJ;ENDERECO;"
          DELE
          SKIP
          LOOP
        EndIf
        MD1=LINHA+";"
        REPLACE CLUBE WITH LELINHA(MD1)
        REPLACE CNPJ WITH LELINHA(MD1)
        REPLACE ENDCLU WITH LELINHA(MD1)
        REPLACE BAIRCLU WITH LELINHA(MD1)
        REPLACE CIDCLU WITH LELINHA(MD1)
        REPLACE UFCLU WITH LELINHA(MD1)
        REPLACE CEPCLU WITH LELINHA(MD1)
        REPLACE CODBEN WITH LELINHA(MD1)
        REPLACE NOMEBEN WITH LELINHA(MD1)
        REPLACE ENDEBEN WITH LELINHA(MD1)
        REPLACE BAIRBEN WITH LELINHA(MD1)
        REPLACE CIDBEN WITH LELINHA(MD1)
        REPLACE UFBEN WITH LELINHA(MD1)
        REPLACE CEPBEN WITH LELINHA(MD1)
        REPLACE SALDOANT WITH LELINHA(MD1)
        REPLACE SALDOATU WITH LELINHA(MD1)
        REPLACE RENDIMENT WITH LELINHA(MD1)
        REPLACE DIVIDENDO WITH LELINHA(MD1)
        REPLACE CREDITO WITH LELINHA(MD1)
        REPLACE CPF WITH LELINHA(MD1)
        SKIP
      ENDDO
      PACK
      GO TOP

      oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
      ZARQ1=WEBLOCAL+"INFORMES"+DTOS(DATE())+".PDF"
      Erase &ZARQ1
      oReg:Set("PDFFilename",ZARQ1)
      oReg:Close()

      PRINT oPrn NAME "Informe de Rendimentos Geração" TO TEMPDF1
      DEFINE FONT oFontV8  NAME "Verdana"     SIZE 0,-10 OF oPrn
      DEFINE FONT oFontV8S  NAME "Verdana"     SIZE 0,-8 OF oPrn
      DEFINE FONT oFontV6  NAME "Verdana" SIZE 0,-6 OF oPrn
      DEFINE FONT oFontV6B  NAME "Verdana" SIZE 0,-6 BOLD OF oPrn
      DEFINE FONT oFontV8B NAME "Verdana"     SIZE 0,-8 BOLD OF oPrn
      DEFINE PEN oPen2 COLOR nRGB(0,0,0) WIDTH 4
      DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-12 OF oPrn
      DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-6 BOLD OF oprn
      DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-8 BOLD OF oPrn
      DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
      DEFINE FONT oFontv NAME "verdana" SIZE 0,12 OF oPrn
      DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
      DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
      DEFINE BRUSH oCinza       COLOR nRGB(220,220,220)
      DEFINE FONT oFontV10B NAME "Verdana" SIZE 0,-10 BOLD OF oPrn
      DEFINE BRUSH oCinzaEscuro COLOR nRGB(180,180,180)
      DEFINE BRUSH oAzul COLOR nRGB(04,43,78) // 35,35,142)
      DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
      DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      DEFINE BRUSH oareia COLOR nRGB(220,207,190)
      nRowStep = oPrn:nVertRes() / 66
      nColStep = oPrn:nHorzRes() / 80

      DO WHILE !EOF()
        WCLUBE=CLUBE
        WCNPJclube=TRANSFORM(STRZERO(VAL(CNPJ),14),"@R 99.999.999/9999-99")
        WENDEC=ENDCLU
        WBAIRR=BAIRCLU
        WCIDAD=CIDCLU
        WUFCLUB=UFCLU
        WCEP=LEFT(STRZERO(val(CEPCLU),8),5)+"-"+RIGHT(STRZERO(val(CEPCLU),8),3)
        WBENE=NOMEBEN
        WENDE1=ENDEBEN
        WBAIR=BAIRBEN
        WCID=CIDBEN
        WUF=UFBEN
        WCEP1=LEFT(STRZERO(val(CEPBEN),8),5)+"-"+RIGHT(STRZERO(val(CEPBEN),8),3)
        WSALDOANT=TRANSFORM(VAL(SALDOANT),"@E 999,999,999.99")
        WSALDOATU=TRANSFORM(VAL(SALDOATU),"@E 999,999,999.99")
        WREND=TRANSFORM(VAL(RENDIMENT),"@E 999,999,999.99")
        WDIVID=TRANSFORM(VAL(DIVIDENDO),"@E 999,999,999.99")
        WCRED=TRANSFORM(VAL(CREDITO),"@E 999,999,999.99")
        WCPF=TRANSFORM(STRZERO(val(CPF),11),"@R 999.999.999-99")
        WCLIE=VAL(CODBEN)

        PAGE

        oprn:cmsay(0.5,0.5,"Cliente: "+ALLTRIM(TRANSFORM(WCLIE,"999999999")),oFontV8,,nRGB(255,255,255),,0) //Código do Cliente
        oprn:cmsay(0.5,7.0,"Assessor - ",oFontV8,,nRGB(255,255,255),,0) //Assessor

        caja(1.5,1,3.8,20,oPrn,0,nrgb(0,0,0)) // primeiro quadro
        caja(1.5,10,3.8,10.01,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        saybitmap(1.8,1.35,2.6,1.35,WEBLOCAL+"LOGOREC.fig",oPrn)
        oprn:cmsay(2.3,4.5,"MINISTÉRIO DA FAZENDA",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(2.8,4,"SECRETARIA DA RECEITA FEDERAL",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(2,10.5,"INFORME DE RENDIMENTOS FINANCEIROS",oFontV10B,,nRGB(0,0,0),,0)
        oprn:cmsay(2.5,10.5,"ANO CALENDÁRIO DE 2011",oFontV10B,,nRGB(0,0,0),,0)
        oprn:cmsay(3,10.5,"IMPOSTO DE RENDA",oFontV10B,,nRGB(0,0,0),,0)

        oprn:cmsay(4.3,1,"IDENTIFICAÇÃO DA FONTE PAGADORA",oFontV8B,,nRGB(0,0,0),,0)
        caja(4.9,1,7.4,20,oPrn,0,nRGB(0,0,0),oPen1) //SEGUNDO quadro
        caja(5.5,1,5.51,20,oPrn,0,nRGB(0,0,0),oPen1) //LINHA DIVISORIA
        caja(4.9,6.7,7.4,6.71,oPrn,0,nrgb(0,0,0),open1) //COLUNA DIVISORIA
        oprn:cmsay(5.05,3.5,"CNPJ",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(5.05,7,"CLUBE/FUNDO",oFontV8B,,nRGB(0,0,0),,0)
        //DADOS SEGUNDO QUADRO
        oprn:cmsay(6,2.5,WCNPJCLUBE,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(5.7,7,WCLUBE,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(6.2,7,WENDEC,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(6.7,7,WCEP+" "+trim(WCIDAD)+"-"+WUFCLUB,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(7.9,1,"PESSOA FÍSICA BENEFICIARIA DOS RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
        caja(8.5,1,11.6,20,oPrn,0,nrgb(0,0,0),oPen1) //TERCEIRO quadro
        caja(9.1,1,9.11,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(10,1,10.01,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(8.5,8.7,10,8.71,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        oprn:cmsay(8.65,3.9,"C.P.F",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(8.65,12,"NOME COMPLETO",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(10.2,1.1,"ENDEREÇO:",oFontV8B,,nRGB(0,0,0),,0)
        //DADOS TERCEIRO QUADRO
        oprn:cmsay(9.36,2.5,WCPF,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(9.36,9.2,WBENE,oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(10.2,4,WENDE1,oFontV8,,nRGB(0,0,0),,0)
        If EMPTY(WBAIR)
          oprn:cmsay(10.6,4,wcep1+" "+trim(wcid)+"-"+wuf,oFontV8,,nRGB(0,0,0),,0)
        Else
          oprn:cmsay(10.6,4,wbair,oFontV8,,nRGB(0,0,0),,0)
          oprn:cmsay(11.0,4,wcep1+" "+trim(wcid)+"-"+wuf,oFontV8,,nRGB(0,0,0),,0)
        EndIf
        oprn:cmsay(12,1,"RENDIMENTOS ISENTOS - (VALORES EM REAIS)",oFontV8B,,nRGB(0,0,0),,0)
        caja(12.6,1,16.5,20,oPrn,0,nrgb(0,0,0),oPen1) //QUARTO quadro
        caja(13.2,1,13.21,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(14.5,1,14.501,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(12.6,6.2,16.5,6.21,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        caja(12.6,10.9,14.5,10.91,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        caja(12.6,15.7,16.5,15.71,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        oprn:cmsay(12.75,2.5,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(12.75,6.7,"SALDO EM 31/12/2010",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(12.75,11.4,"SALDO EM 31/12/2011",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(12.75,16.7,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(13.5,1.1,"CONTAS DE POUPANÇA E",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(13.9,1.1,"LETRAS  HIPOTECÁRIAS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(14.7,1.1,"LUCROS E DIVIDENDOS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(15.1,1.1,"APURADOS A PARTIR DE",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(15.5,1.1,"1996 E DISTRIBUIDOS NO",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(15.9,1.1,"ANO CALENDARIO",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(15.5,17.1,WDIVID,oFontV8S,,nRGB(0,0,0),,0)

        oprn:cmsay(17,1,"RENDIMENTOS SUJEITOS À TRIBUTAÇÃO EXCLUSIVA - (VALORES EM REAIS)",oFontV8B,,nRGB(0,0,0),,0)
        caja(17.6,1,21.5,20,oPrn,0,nrgb(0,0,0),oPen1) //QUINTO quadro
        caja(18.2,1,18.21,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(19.5,1,19.51,20,oPrn,0,nrgb(0,0,0),oPen1) //LINHA DIVISORIA
        caja(17.6,6.2,19.5,6.21,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        caja(17.6,10.9,19.5,10.91,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        caja(17.6,15.7,21.5,15.71,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
        oprn:cmsay(17.75,2.5,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(17.75,6.7,"SALDO EM 31/12/2010",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(17.75,11.4,"SALDO EM 31/12/2011",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(17.75,16.7,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(20.3,2.9,"TOTAL DOS RENDIMENTOS SUJEITOS A TRIBUTAÇÃO EXCLUSIVA",oFontV8B,,nRGB(0,0,0),,0)
        //DADOS QUINTO QUADRO
        oprn:cmsay(18.5,1.1,LEFT(WCLUBE,25),oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(18.9,1.1,SUBSTR(WCLUBE,26,25),oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(18.5,8.5,WSALDOANT,oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(18.5,13.0,WSALDOATU,oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(18.5,17.1,WREND,oFontV8S,,nRGB(0,0,0),,0)
        oprn:cmsay(20.3,17.1,WREND,oFontV8S,,nRGB(0,0,0),,0)

        oprn:cmsay(22.3,1,"INFORMAÇÕES COMPLEMENTARES",oFontV8B,,nRGB(0,0,0),,0)
        caja(22.9,1,25.9,20,oPrn,0,nrgb(0,0,0),oPen1) //QUINTO quadro

        //    oPrn:cmsay(24.2,1.1,"CRÉDITO FISCAL E/OU PROVISÃO FISCAL - Valor Acumulado: "+WCRED,oFontV8,,nRGB(0,0,0),,0)
        oprn:cmsay(24.6,1.1,"EMISSÃO: 22/02/2012",oFontV8,,nRGB(0,0,0),,0)

        //   oPrn:cmsay(24.6,1.1,"Este Informe de Rendimentos substitui qualquer outro enviado anteriormente.",oFontv8,,nRGB(0,0,0),,0)

        ENDPAGE
        SELE 36
        SKIP
      ENDDO
      ENDPRINT
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      SELE 31
      USE
      Erase &NARQ4
      Erase &NARQ6
      Erase &NARQ5
      SELE 30
    return 

    Procedure LELINHA(QL)
      COLU=AT(";",QL)
      WCOL=LEFT(QL,COLU-1)
      MD1=SUBSTR(QL,COLU+1,900)
    return  WCOL

    Procedure ATUACONFIDENCE()
      DIRTERENV="c:\notas\confide\bmf\mens\env\"
      D_IRTERENV="c:\notas\confide\bmf\mens\env\arqf\"
      If !VERDIRETORIO(D_IRTERENV)
        MSGINFO("Não foi possivel criar o diretorio "+D_IRTERENV,WSIST)
        QUIT
      EndIf
      NOMAARQUIVO:=DIRECTORY(D_IRTERENV+"*.*")
      FOR WI=1 TO LEN(NOMAARQUIVO)
        If NOMAARQUIVO[wI][3]<(DATE()-4)
          NARQ=D_IRTERENV+NOMAARQUIVO[wI][1]
          Erase &NARQ
        EndIf
      NEXT
      DECLARE VLABEL1[ADIR(dirterenv+"*.ZIP")]
      ADIR(DIRTERENV+"*.ZIP",vlabel1)
      EI=0
      If LEN(VLABEL1)#0
        ATBASECONF("confidence","14203_")
      EndIf
      /*
DIRTERENV="c:\notas\BCONFIDE\bmf\mens\env\"
D_IRTERENV="c:\notas\BCONFIDE\bmf\mens\env\arqf\"
If !VERDIRETORIO(D_IRTERENV)
  MSGINFO("Não foi possivel criar o diretorio "+D_IRTERENV,WSIST)
  QUIT
EndIf
NOMAARQUIVO:=DIRECTORY(D_IRTERENV+"*.*")
FOR WI=1 TO LEN(NOMAARQUIVO)
 If NOMAARQUIVO[wi][3]<(DATE()-4)
   NARQ=D_IRTERENV+NOMAARQUIVO[wI][1]
   Erase &NARQ
 EndIf
NEXT
DECLARE VLABEL1[ADIR(dirterenv+"*.ZIP")]
ADIR(DIRTERENV+"*.ZIP",vlabel1)
EI=0
If LEN(VLABEL1)#0
  ATBASECONF("bconfide","27857_")
EndIf
*/
    return 

    Procedure ATBASECONF(qbas,qarq)
      CONEX1ON="mysqlcorretoras.exatus.net"
      CONEX1DTBAS=qbas
      //CONEX1USER="root"
      //CONEX1SENHA=""
      CONEX1USER="atconfidence#."
      CONEX1SENHA="-3OXURG1M2QCPA"
      // apagar 14203.txt
      CONEX1PORT=3306
      SQL CONNECT ON CONEX1ON;
        PORT CONEX1PORT;
        DATABASE CONEX1DTBAS;
        USER CONEX1USER;
        PASSWORD CONEX1SENHA;
        OPTIONS SQL_NO_WARNING;
        LIB 'MySQL';
        INTO CONEXEX

      If SQL_ErrorNO() > 0
        msginfo("erro na conexão base confidence",wsist)
        return 
      EndIf

      SELE 41
      If !USEREDE("filiais",.F.,10,"FILIAIS","M")
        CLOSE DATABASE
        return 
      EndIf
      SELE 40
      If !USEREDE(qarq+"movimento",.F.,10,"CONTR","M")
        CLOSE DATABASE
        return 
      EndIf

      FOR EI=1 TO LEN(VLABEL1)
        NARQ=DIRTERENV+VLABEL1[EI]
        NARQ_1=D_IRTERENV+VLABEL1[EI]
        NOMAARQUIVO:=DIRECTORY(NARQ)
        If (nHandle:=FOpen(NARQ,FO_READ+FO_EXCLUSIVE))!=-1
          fClose(nHandle)
        Else
          fClose(nHandle)
          LOOP
        EndIf
        If NOMAARQUIVO[1][2]<10
          LOOP
        EndIf
        SELE 48
        If !USEREDE(NARQ,.F.,10,"CONFID")
          QUIT
        EndIf
        GO TOP
        DO WHILE !EOF()
          If LEFT(NUMCTRLIF,5)="10031" .And. reglock(10)
            REPLACE NUMCTRLIf WITH SUBSTR(NUMCTRLIF,2,12)
          Else
            If val(NumCtrlIF)>9999999999
              skip
              loop
            EndIf
          EndIf
          WCODF=1
          WCNPJF=""
          WPAISEMI=""
          If FCOUNT()>57
            WCNPJF=Strtran(CNPJFILCON,".","")
            WCNPJF=Strtran(WCNPJF,"/","")
            WCNPJF=ALLTRIM(Strtran(WCNPJF,"-",""))
            SELE 41
            LOCATE FOR VAL(WCNPJF)=VAL(CNPJ)
            If VAL(WCNPJF)=0
              WCODF=1
            Else
              If !EOF()
                WCODF=CODIGO
              Else
                ADIREG(0)
                REPLACE APELIDO WITH "CONFIDENCE"
                REPLACE NOME WITH "CONFIDENCE CC"
                REPLACE INSTITUICA WITH "14203"
                REPLACE DEPENDENCI WITH "0001"
                REPLACE EMPRESA WITH 18
                REPLACE CNPJ WITH WCNPJF
                REPLACE CORRETORA WITH 12
                REPLACE TIPO WITH "F"
                REPLACE CCONSOLIDADO WITH "N"
                REPLACE IMPRESSAO WITH "B"
                LOCATE FOR VAL(WCNPJF)=VAL(CNPJ)
                WCODF=CODIGO
              EndIf
            EndIf
          EndIf
          SELE 48
          If FCOUNT()>59
            WCORRESP=VAL(CONFID->CNPJCORRE)
          Else
            WCORRESP=0
          EndIf
          If FCOUNT()>61
            WNomeIExt=CONFID->NomeIExt
            WISPBDest=CONFID->ISPBDest
          Else
            WNomeIExt=""
            WISPBDest=""
          EndIf
          SELE 40
          WNRC="SELECT CODIGO,EVENTOCANCEL FROM "+qarq+"movimento WHERE sql_deleted='T' AND codigo="+STRZERO(VAL(CONFID->NumCtrliF),10)
          mz=SQLArrayAssoc(WNRC)
          If LEN(mz)#0
            wcodigo=MZ[1,1]
            WFILT="DELETE FROM "+qarq+"movimento where codigo="+wcodigo
            If VAL(MZ[1,2])=0
              SQL EXECUTE wfilt
              SQL EXECUTE "COMMIT"
            EndIf
          EndIf
          WNRC="SELECT CODIGO FROM "+qarq+"movimento WHERE sql_deleted='F' AND codigo="+STRZERO(VAL(CONFID->NumCtrliF),10)
          mz=SQLArrayAssoc(WNRC)
          If LEN(mz)#0
            WFILT="CODIGO="+STRZERO(VAL(CONFID->NumCtrliF),10)
            SQL FILTER TO "sql_deleted='F' AND CODIGO="+STRZERO(VAL(CONFID->NumCtrliF),10)
            GO TOP
            If CODIGO=VAL(CONFID->NumCtrlIF) .And. REGLOCK(10) .And. obsliquid<>STRZERO(WCORRESP,14)
              REPLACE obsliquid WITH STRZERO(WCORRESP,14)
            EndIf
            SELE 48
            SKIP
            LOOP
          EndIf
          SELE 40
          WFILT="sql_deleted='F' AND CODIGO="+STRZERO(VAL(CONFID->NumCtrliF),10)
          SQL FILTER TO "sql_deleted='F' AND CODIGO="+STRZERO(VAL(CONFID->NumCtrliF),10)
          GO TOP
          If CODIGO=VAL(CONFID->NumCtrlIF)  .OR. CONFID->CDMOE1=0
            SELE 48
            SKIP
            LOOP
          EndIf
          mfo=VAL(CONFID->FormaMn)
          WOUTRAS="Forma de Pagamento: "
          If mfo=5
            MFO1=1
            WOUTRAS=WOUTRAS+" Cheque"+CHR(13)
          ElseIf mfo=1
            MFO1=2
            WOUTRAS=WOUTRAS+" Deposito"+CHR(13)
          ElseIf mfo=3
            MFO1=3
            WOUTRAS=WOUTRAS+" Doc"+CHR(13)
          ElseIf mfo=4
            MFO1=4
            WOUTRAS=WOUTRAS+" Em especie"+CHR(13)
          ElseIf mfo=9
            MFO1=5
            WOUTRAS=WOUTRAS+" Outras"+CHR(13)
          ElseIf mfo=2
            MFO1=6
            WOUTRAS=WOUTRAS+" TED"+CHR(13)
          Else
            MFO1=5
            WOUTRAS=" "
          EndIf
          If !EMPTY(CONFID->ISPBIF)
            WOUTRAS=WOUTRAS+"Banco ISPB: "+CONFID->ISPBIF+CHR(13)
          EndIf
          If !EMPTY(CONFID->AgIFPai)
            WOUTRAS=WOUTRAS+"Agencia: "+CONFID->AgIFPai
          EndIf
          If !EMPTY(CONFID->CtIFPai)
            WOUTRAS=WOUTRAS+" Conta: "+CONFID->CtIFPai
          EndIf
          If CONFID->TpPessoa="J"
            WCGCCLI=CONFID->CGCCPFCL
            WCPFCLI=""
          Else
            WCGCCLI=""
            WCPFCLI=CONFID->CGCCPFCL
          EndIf
          If CONFID->NAT3="N"
            WNAT3=0
          Else
            WNAT3=1
          EndIf
          If CONFID->ACIC=" "
            WACIC="N"
          Else
            WACIC=CONFID->ACIC
          EndIf
          ADIREG(0)
          BEGIN TRANSACTION
          If CONFID->TpPessoa="F"
            REPLACE CLIENTE WITH 2
          ElseIf CONFID->TpPessoa="J"
            REPLACE CLIENTE WITH 4
          EndIf
          If CONFID->TpPessoa="E"
            REPLACE CLIENTE WITH 3
            REPLACE TpPessoa WITH "E"
            WCPFCLI="ISENTO"
          EndIf
          If CONFID->TpOpCam="C"
            REPLACE TIPO WITH "3"
          ElseIf CONFID->TpOpCam="V"
            REPLACE TIPO WITH "4"
          ElseIf CONFID->TpOpCam="5"
            REPLACE CLIENTE WITH 1
            REPLACE TIPO WITH "5"
          ElseIf CONFID->TpOpCam="6"
            REPLACE TIPO WITH "6"
            REPLACE CLIENTE WITH 1
          EndIf
          REPLACE ENDCLI1 WITH CONFID->ENDCLI1
          REPLACE ENDCLI2 WITH CONFID->ENDCLI2
          If TIPO$"56"
            REPLACE ENDCLI1 WITH CONFID->ENDEBANCO
            REPLACE PRACA WITH VAL(CONFID->PRACA)
          EndIf
          If CONFID->SIMPL="1"
            REPLACE VTM WITH "0"
          ElseIf CONFID->SIMPL="2"
            REPLACE VTM WITH "0"
          ElseIf CONFID->SIMPL="3"
            REPLACE VTM WITH "1"
          ElseIf CONFID->SIMPL="4"
            REPLACE VTM WITH "0"
            REPLACE ATM WITH "S"
          EndIf
          REPLACE DOCUMENTO WITH CONFID->DOCEST,NR_BACEN WITH CONFID->NR_BACEN,EVENTOCONTR WITH CONFID->EVENTOCON
          If WACIC="S"
            REPLACE envbc with .F.
          Else
            REPLACE ENVBC WITH .T.
          EndIf
          REPLACE PAISEMISSAO WITH CONFID->BPAISEMISI
          REPLACE TAXAORG WITH VAL(CONFID->TaxCam)
          REPLACE INSTITUICA WITH val(CONFID->INSTIT),TPPESSOA WITH CONFID->TPPESSOA
          REPLACE TPLIQUID WITH MFO1,CGCCLI WITH WCGCCLI,CPFCLI WITH WCPFCLI,NAT3 WITH WNAT3
          REPLACE DTENTREAIS WITH STOD(Strtran(CONFID->DtEntrMN,"-","")),FILIAL WITH WCODF
          REPLACE DTENTMOEDA WITH STOD(Strtran(CONFID->LIQUIDA,"-","")),NOMECLI WITH SINAL(CONFID->NOMECLI),DATA WITH STOD(Strtran(CONFID->DtEvtCAM,"-",""))
          REPLACE MOEDA WITH CONFID->CDMOE1,ESTRANGEIR WITH VAL(CONFID->VlrME),REAIS WITH VAL(CONFID->VlrMN),TAXA WITH VAL(CONFID->TaxCam),CONTA WITH 1,CAIXA WITH 0,AUTORIZADO WITH "S",FECHADO WITH "S"
          If DATA<CTOD("03/05/2015")
            REPLACE IOF WITH (VAL(CONFID->VlrMN)*0.38)/100
          Else
            REPLACE IOF WITH (VAL(CONFID->VlrMN)*1.1)/100
          EndIf
          REPLACE vlroper with VAL(CONFID->VlrMN),NAT1 WITH VAL(CONFID->NAT1),NAT2 WITH VAL(CONFID->NAT2),NAT4 WITH VAL(CONFID->NAT4),NAT5 WITH VAL(CONFID->NAT5)
          REPLACE OUTRAS WITH WOUTRAS,LIQUIDACAO WITH STOD(Strtran(CONFID->LIQUIDA,"-","")),FORMAENTRE WITH VAL(CONFID->FENTR),OPERADOR WITH 2,NOMEPAG WITH CONFID->NOMEPAG,VINCULO WITH CONFID->VINC,CONTAVTM WITH 1
          REPLACE OUTRAS WITH CONFID->TxtOtr
          REPLACE obsliquid WITH STRZERO(WCORRESP,14)
          REPLACE USERINC WITH "MSGBC - AUTOMATICO",INCLUIDO WITH DTOC(DATE())+" "+TIME(),INTEGRADO WITH "N",PAISPAG WITH strzero(CONFID->CDPAIS1,4),ADIG WITH "N",GEROUACIC WITH "N",ACIC WITH WACIC
          REPLACE ISPBDEST WITH WISPBDEST
          REPLACE NomeIExt WITH WNomeIExt
          If (NAT1<>32009 .And. NAT1<>32016 .And. NAT1<>32023) .And. VAL(PAISPAG)=0
            REPLACE PAISPAG WITH "2496"
          EndIf
          If VAL(PAISPAG)=7803
            REPLACE PAISPAG WITH "7806"
          EndIf
          If VAL(PAISPAG)=1511
            REPLACE PAISPAG WITH "2453"
          EndIf
          If IOF=0
            REPLACE IOF WITH 0.01
          EndIf
          If CONFID->TPPESSOA="E"
            REPLACE CLIENTE WITH 3
            REPLACE TpPessoa with "E"
            WCPFCLI="ISENTO"
          EndIf
          If WACIC="S"
            REPLACE LIQUIDADO WITH "C"
          EndIf
          If len(alltrim(WNomeIExt))>0 .And. WACIC="S"
            REPLACE ACIC WITH "N"
            REPLACE INGDIR WITH "S"
            REPLACE NAT1 WITH 37114
          EndIf
          If TIPO="5" .OR. TIPO="6"
            REPLACE IOF WITH 0
          ElseIf TIPO="4"
            REPLACE VLROPER WITH VAL(CONFID->VlrMN)+IOF
          ElseIf TIPO="3"
            REPLACE VLROPER WITH VAL(CONFID->VlrMN)-IOF
          EndIf
          REPLACE CODIGO WITH VAL(CONFID->NumCtrlIF)
        END TRANSACTION
        SELE 48
        SKIP
        ENDDO
        SELE 48
        USE
        LZCOPYFILE(NARQ,NARQ_1)
        Erase &NARQ
      NEXT EI
      SELE 41
      USE
      SELE 40
      USE
      SQLDISCONNECT(CONEXEX)
    return 

    Procedure INFORMEGRENDNOVO()
      DECLARE VLABEL[ADIR(WEBLOCAL+"REND_INFORME_*.*")]
      ADIR(WEBLOCAL+"REND_INFORME_*.*",vlabel)
      VLABEL:=ASORT(VLABEL)
      If LEN(VLABEL)=0
        return 
      EndIf
      SysRefresh()
      CURSORWAIT()
      DECLARE TX1[1]
      TX1[1]="BolsasNet - Informes de Rendimento Geração "+WEBCOR
      LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
      FEZALGO=.T.
      SysRefresh()
      CURSORWAIT()
      mat_stru={}
      Aadd(mat_stru,{"LINHA","C",900,0})
      Aadd(mat_stru,{"CHAVE","C",34,0})
      Aadd(mat_stru,{"QUANTOS","N",4,0})
      AaDD(mat_stru,{"PRIM","L",1,0})
      ARQTXTO=WEBLOCAL1+"INFREND.DAT"
      DBCREATE(ARQTXTO,MAT_STRU,"DBFNTX")
      SELE 31
      NARQ5=WEBLOCAL+"MCLIENTE.KEY"
      Erase &NARQ5
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.T.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      INDEX ON CLIENTE TO &NARQ5
      USE
      If !USEREDE(WEBLOCAL+"MCLIENTE.DBF",.F.,10,"MCLIENTE","S")
        SELE 30
        return 
      EndIf
      SET INDEX TO &NARQ5
      SELE 32
      If !USEREDE(WEBLOCAL+"MGRUPO.DBF",.F.,10,"MGRUPO","S")
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      NARQ4=WEBLOCAL+"CLIENTE.KEY"
      NARQ6=WEBLOCAL+"CLIENTE1.KEY"
      Erase &NARQ4
      Erase &NARQ6
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.T.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 33
        USE
        SELE 31
        USE
        SELE 30
        return 
      EndIf
      INDEX ON CODIGO TO &NARQ4
      INDEX ON STRZERO(VAL(CNPJCPF),14)+IIF(TEMMAIL,"1","2") TO &NARQ6
      USE
      SELE 35
      If !USEREDE(WEBLOCAL+"CLIENTE.DBF",.F.,10,"CLIENTES","S")
        SELE 32
        USE
        SELE 31
        USE
        SELE 30
        Erase &NARQ4
        Erase &NARQ6
        return 
      EndIf
      SET INDEX TO &NARQ4,&NARQ6
      SELE 36
      If !USEREDE(ARQTXTO,.T.,10,"APLIC","S")
        SELE 36
        Erase &ARQTXTO
        return 
      EndIf
      SELE 36
      FOR I=1 TO LEN(VLABEL)
        NARQ=WEBLOCAL+VLABEL[I]
        APPEND FROM &NARQ SDF
        NARQ1=WEBLOCAL2+VLABEL[I]
        LZCOPYFILE(NARQ,NARQ1)
        Erase &NARQ
      NEXT I
      SELE 36
      GO TOP

      oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
      ZARQ1=WEBLOCAL+"INFORMES"+DTOS(DATE())+".PDF"
      Erase &ZARQ1
      oReg:Set("PDFFilename",ZARQ1)
      oReg:Close()

      PRINT oPrn NAME "Informe de Rendimentos Geração" TO TEMPDF1
      DEFINE FONT oFontV8   NAME "Verdana"     SIZE 0,-10 OF oPrn
      DEFINE FONT oFontV8S  NAME "Verdana"     SIZE 0,-8 OF oPrn
      DEFINE FONT oFontV6   NAME "Verdana" SIZE 0,-6 OF oPrn
      DEFINE FONT oFontV6B  NAME "Verdana" SIZE 0,-6 BOLD OF oPrn
      DEFINE FONT oFontV8B  NAME "Verdana"     SIZE 0,-8 BOLD OF oPrn
      DEFINE PEN oPen2 COLOR nRGB(0,0,0) WIDTH 4
      DEFINE FONT oFontC8  NAME "Courier New" SIZE 0,-12 OF oPrn
      DEFINE FONT oFontc8B NAME "Courier New" SIZE 0,-6 BOLD OF oprn
      DEFINE FONT oFontt8  NAME "Courier New" SIZE 0,-10 OF oPrn
      DEFINE FONT oFont NAME "courier new negrito" SIZE 0, 8.3 OF oPrn
      DEFINE FONT oFontv NAME "verdana" SIZE 0,-7 OF oPrn
      DEFINE FONT oFont1 NAME "verdana" SIZE 0,12 BOLD OF oPrn
      DEFINE PEN oPen1 COLOR nRGB(0,0,0) WIDTH 1
      DEFINE BRUSH oCinza       COLOR nRGB(220,220,220)
      DEFINE FONT oFontV10B NAME "Verdana" SIZE 0,-10 BOLD OF oPrn
      DEFINE BRUSH oCinzaEscuro COLOR nRGB(180,180,180)
      DEFINE BRUSH oAzul COLOR nRGB(04,43,78) // 35,35,142)
      DEFINE BRUSH oBranco COLOR nRGB(255,255,255)
      DEFINE BRUSH oPreto COLOR nRGB(0,0,0)
      DEFINE BRUSH oareia COLOR nRGB(220,207,190)
      nRowStep = oPrn:nVertRes() / 66
      nColStep = oPrn:nHorzRes() / 80

      WANOCALEND=""
      WANOCALANT=""
      WANOCALATU=""
      WFONTNOME=""
      WFONTCNPJ=""
      WTPPESS=""
      WCPFCNPJ=""
      WNOME=""
      WCOTISTA=""
      WCAB=0
      contl=0
      DO WHILE !EOF()
        If LEFT(LINHA,2)="00"
          WANOCALEND=SUBSTR(LINHA,3,4)
          WANOCALANT=SUBSTR(LINHA,7,8)
          WANOCALATU=SUBSTR(LINHA,15,8)
          WCAB=0
          contl=0
          SKIP
          LOOP
        EndIf
        If LEFT(LINHA,2)="01"
          WFONTNOME=SUBSTR(LINHA,3,50)
          WFONTCNPJ=SUBSTR(LINHA,53,14)
          WFONTCNPJ=STRZERO(VAL(WFONTCNPJ),14)
          WFONTCNPJ=TRANSFORM(WFONTCNPJ,"@R 99.999.999/9999-99")
          SKIP
          LOOP
        EndIf
        If LEFT(LINHA,2)="02"
          WTPPESS=SUBSTR(LINHA,3,1)
          WCPFCNPJ=SUBSTR(LINHA,4,14)
          WNOME=SUBSTR(LINHA,18,50)
          WCOTISTA=SUBSTR(LINHA,70,15)
          If WTPPESS="F"
            WCPFCNPJ=STRZERO(VAL(WCPFCNPJ),11)
            WCPFCNPJ=TRANSFORM(WCPFCNPJ,"@R 999.999.999-99")
          Else
            WCPFCNPJ=STRZERO(VAL(WCPFCNPJ),14)
            WCPFCNPJ=TRANSFORM(WCPFCNPJ,"@R 99.999.999/9999-99")
          EndIf
          SKIP
          LOOP
        EndIf
        If (LEFT(LINHA,2)="03" .OR. LEFT(LINHA,2)="04" .OR. LEFT(LINHA,2)="06" .OR. LEFT(LINHA,2)="07" .OR. LEFT(LINHA,2)="09") .And. WTPPESS="J"
          SKIP
          LOOP
        EndIf

        If WCAB=0
          cabinf()
        EndIf

        If WTPPESS="F"
          If LEFT(LINHA,2)="03"
            contl=contl+0.5
            oprn:cmsay(contl,1,"3. RENDIMENTOS TRIBUTÁVEIS NA DECLARAÇÃO DE AJUSTE ANUAL (Valores em Reais)",oFontV8B,,nRGB(0,0,0),,0)
            contl=contl+0.4
            caja(contl,1,contl+0.5,20,oPrn,0,nrgb(0,0,0),oPen1)
            caja(contl,10,contl+0.5,10.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,15,contl+0.5,15.01,oPrn,0,nRGB(0,0,0),oPen1)
            oprn:cmsay(contl,1.3,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,11.3,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,15.2,"IMPOSTO RETIDO NA FONTE",oFontV8B,,nRGB(0,0,0),,0)
            contl=contl+0.5
            DO WHILE LEFT(LINHA,2)="03" .And. !EOF()
              caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,10,contl+0.5,10.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,15,contl+0.5,15.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
              oprn:cmsay(contl,1.3,substr(linha,3,50),oFontv8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,11,Transform(val(substr(linha,53,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,16.2,Transform(val(substr(linha,69,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
              contl=contl+0.5
              SKIP
              If CONTL>28
                ENDPAGE
                CABINF()
              EndIf
            ENDDO
          EndIf
          If LEFT(LINHA,2)="04"
            contl=contl+0.5
            oprn:cmsay(contl,1,"4. RENDIMENTOS ISENTOS (Valores em Reais)",oFontV8B,,nRGB(0,0,0),,0)
            contl=contl+0.4
            caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
            caja(contl,08,contl+0.7,08.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,12,contl+0.7,12.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,16,contl+0.7,16.01,oPrn,0,nRGB(0,0,0),oPen1)
            oprn:cmsay(contl,1.3,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,08.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,12.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,16.2,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,08.2,"ANO CALENDARIO ANTERIOR",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,12.2,"ANO CALENDARIO ATUAL",oFontV6B,,nRGB(0,0,0),,0)
            contl=contl+0.7
            DO WHILE LEFT(LINHA,2)="04" .And. !EOF()
              caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,08,contl+0.5,08.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,12,contl+0.5,12.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,16,contl+0.5,16.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
              oprn:cmsay(contl,1.3,substr(linha,3,31),oFontt8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,08.7,Transform(val(substr(linha,53,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) //oFontC8
              oprn:cmsay(contl,12.7,Transform(val(substr(linha,69,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) //oFontC8
              oprn:cmsay(contl,16.7,Transform(val(substr(linha,85,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) //oFontC8
              caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
              contl=contl+0.5
              SKIP
              If CONTL>28
                ENDPAGE
                CABINF()
              EndIf
            ENDDO
          EndIf
        EndIf

        If LEFT(LINHA,2)="05"
          contl=contl+0.5
          If WTPPESS="F"
            oprn:cmsay(contl,1,"5. RENDIMENTOS SUJEITOS A TRIBUTAÇÃO EXCLUSIVA (Valores em Reais): VIDE QUADRO 8",oFontV8B,,nRGB(0,0,0),,0)
          Else
            oprn:cmsay(contl,1,"3. RENDIMENTOS SUJEITOS A TRIBUTAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
          EndIf
          contl=contl+0.4
          caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
          caja(contl,08,contl+0.7,08.01,oPrn,0,nRGB(0,0,0),oPen1)
          caja(contl,12,contl+0.7,12.01,oPrn,0,nRGB(0,0,0),oPen1)
          caja(contl,16,contl+0.7,16.01,oPrn,0,nRGB(0,0,0),oPen1)
          oprn:cmsay(contl,1.3,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,08.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,12.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,16.2,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,08.2,"ANO CALENDARIO ANTERIOR",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,12.2,"ANO CALENDARIO ATUAL",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,16.2,"LÍQUIDOS",oFontV8B,,nRGB(0,0,0),,0)
          contl=contl+0.7
          DO WHILE LEFT(LINHA,2)="05" .And. !EOF()
            caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,08,contl+0.5,08.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,12,contl+0.5,12.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,16,contl+0.5,16.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
            oprn:cmsay(contl,1.3,substr(linha,3,50),oFontt8,,nRGB(0,0,0),,0)  //oFontv8
            oprn:cmsay(contl,08.7,Transform(val(substr(linha,53,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) // oFontC8
            oprn:cmsay(contl,12.7,Transform(val(substr(linha,69,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) // oFontC8
            oprn:cmsay(contl,16.7,Transform(val(substr(linha,85,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0) // oFontC8
            caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
            contl=contl+0.5
            SKIP
            If CONTL>28
              ENDPAGE
              CABINF()
            EndIf
          ENDDO
        EndIf

        If WTPPESS="F"
          If LEFT(LINHA,2)="06"
            contl=contl+0.5
            oprn:cmsay(contl,1,"6. SALDO EM CONTAS CORRENTES E EM VGBL (Valores em Reais)",oFontV8B,,nRGB(0,0,0),,0)
            contl=contl+0.4
            caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
            caja(contl,10,contl+0.7,10.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,15,contl+0.7,15.01,oPrn,0,nRGB(0,0,0),oPen1)
            oprn:cmsay(contl,1.3,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,10.3,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,15.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,10.3,"ANO CALENDARIO ANTERIOR",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,15.2,"ANO CALENDARIO ATUAL",oFontV6B,,nRGB(0,0,0),,0)
            contl=contl+0.7
            DO WHILE LEFT(LINHA,2)="06" .And. !EOF()
              caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,10,contl+0.5,10.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,15,contl+0.5,15.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
              oprn:cmsay(contl,1.3,substr(linha,3,50),oFontv8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,11,Transform(val(substr(linha,53,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,16.2,Transform(val(substr(linha,69,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
              contl=contl+0.5
              SKIP
              If CONTL>28
                ENDPAGE
                CABINF()
              EndIf
            ENDDO
          EndIf

          If LEFT(LINHA,2)="07"
            contl=contl+0.5
            oprn:cmsay(contl,1,"7. CRÉDITOS EM TRANSITO (Valores em Reais)",oFontV8B,,nRGB(0,0,0),,0)
            contl=contl+0.4
            caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
            caja(contl,10,contl+0.7,10.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,15,contl+0.7,15.01,oPrn,0,nRGB(0,0,0),oPen1)
            oprn:cmsay(contl,1.3,"ESPECIFICAÇÃO",oFontV8B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,10.3,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl,15.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,10.3,"ANO CALENDARIO ANTERIOR",oFontV6B,,nRGB(0,0,0),,0)
            oprn:cmsay(contl+0.3,15.2,"ANO CALENDARIO ATUAL",oFontV6B,,nRGB(0,0,0),,0)
            contl=contl+0.7
            DO WHILE LEFT(LINHA,2)="07" .And. !EOF()
              caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,10,contl+0.5,10.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,15,contl+0.5,15.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
              oprn:cmsay(contl,1.3,substr(linha,3,50),oFontv8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,11,Transform(val(substr(linha,53,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,16.2,Transform(val(substr(linha,69,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
              caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
              contl=contl+0.5
              SKIP
              If CONTL>28
                ENDPAGE
                CABINF()
              EndIf
            ENDDO
          EndIf
        EndIf

        If LEFT(LINHA,2)="08"
          contl=contl+0.5
          If WTPPESS="F"
            oprn:cmsay(contl,1,"8. INFORMAÇÕES COMPLEMENTARES (Valores em Reais)",oFontV8B,,nRGB(0,0,0),,0)
          Else
            oprn:cmsay(contl,1,"4. INFORMAÇÕES COMPLEMENTARES",oFontV8B,,nRGB(0,0,0),,0)
          EndIf
          contl=contl+0.4
          caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
          caja(contl,09,contl+0.7,09.01,oPrn,0,nRGB(0,0,0),oPen1)
          caja(contl,13,contl+0.7,13.01,oPrn,0,nRGB(0,0,0),oPen1)
          caja(contl,17,contl+0.7,17.01,oPrn,0,nRGB(0,0,0),oPen1)
          oprn:cmsay(contl,1.3,"PRODUTO",oFontV8B,,nRGB(0,0,0),,0)
          oPrn:cmsay(CONTL,6.8,"CNPJ",oFontV8B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,09.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,13.2,"SALDOS EM 31/12",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl,17.2,"RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,09.2,"ANO CALENDARIO ANTERIOR",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,13.2,"ANO CALENDARIO ATUAL",oFontV6B,,nRGB(0,0,0),,0)
          oprn:cmsay(contl+0.3,17.2,"LÍQUIDOS",oFontV8B,,nRGB(0,0,0),,0)
          contl=contl+0.7
          DO WHILE LEFT(LINHA,2)="08" .And. !EOF()
            caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,09,contl+0.5,09.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,13,contl+0.5,13.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,17,contl+0.5,17.01,oPrn,0,nRGB(0,0,0),oPen1)
            caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
            If SUBSTR(LINHA,3,5)<>"TOTAL"
              oprn:cmsay(contl+0.1,1.3,substr(linha,3,30),oFontv,,nRGB(0,0,0),,0) // tava 50
              oprn:cmsay(contl+0.1,6.3,TRANSFORM(substr(linha,53,14),"@R 99.999.999/9999-99"),oFontv,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,10.0,Transform(val(substr(linha,89,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,14.0,Transform(val(substr(linha,127,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,17,Transform(val(substr(linha,143,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
            Else
              oprn:cmsay(contl+0.1,1.3,substr(linha,3,50),oFontv,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,10.0,Transform(val(substr(linha,89,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,14.0,Transform(val(substr(linha,127,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,17,Transform(val(substr(linha,143,16))/100,"@ZE 999,999,999.99"),oFontt8,,nRGB(0,0,0),,0)
            EndIf
            caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
            contl=contl+0.5
            SKIP
            If CONTL>28
              ENDPAGE
              CABINF()
            EndIf
          ENDDO
        EndIf

        If WTPPESS="J"
          If LEFT(LINHA,2)="10"
            DO WHILE !EOF() .And. LEFT(LINHA,2)="10"
              If CONTL>24
                ENDPAGE
                CABINF()
              EndIf
              contl=contl+0.5
              oprn:cmsay(contl,1,"Produto: "+SUBSTR(LINHA,17,50),oFontV8B,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,10,"CNPJ: "+TRANSFORM(substr(linha,3,14),"@R 99.999.999/9999-99"),oFontV8B,,nRGB(0,0,0),,0)
              contl=contl+0.4
              caja(contl,1,contl+0.7,20,oPrn,0,nrgb(0,0,0),oPen1)
              caja(contl,10,contl+0.7,10.01,oPrn,0,nRGB(0,0,0),oPen1)
              caja(contl,15,contl+0.7,15.01,oPrn,0,nRGB(0,0,0),oPen1)
              oprn:cmsay(contl,1.3,"MES",oFontV8B,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,10.3,"RENDIMENTO BRUTO",oFontV8B,,nRGB(0,0,0),,0)
              oprn:cmsay(contl,15.2,"IMPOSTO DE RENDA RETIDO",oFontV8B,,nRGB(0,0,0),,0)
              contl=contl+0.7
              CLUBEA=SUBSTR(LINHA,3,14)
              DO WHILE LEFT(LINHA,2)="10" .And. CLUBEA=SUBSTR(LINHA,3,14) .And. !EOF()
                caja(contl,1,contl+0.5,1.01,oPrn,0,nRGB(0,0,0),oPen1)
                caja(contl,10,contl+0.5,10.01,oPrn,0,nRGB(0,0,0),oPen1)
                caja(contl,15,contl+0.5,15.01,oPrn,0,nRGB(0,0,0),oPen1)
                caja(contl,20,contl+0.5,20.01,oPrn,0,nRGB(0,0,0),oPen1)
                oprn:cmsay(contl,1.3,substr(linha,67,10),oFontv8,,nRGB(0,0,0),,0)
                oprn:cmsay(contl,11,Transform(val(substr(linha,77,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
                oprn:cmsay(contl,16.2,Transform(val(substr(linha,93,16))/100,"@ZE 999,999,999.99"),oFontC8,,nRGB(0,0,0),,0)
                caja(contl+0.5,1,contl+0.51,20,oPrn,0,nRGB(0,0,0),oPen1)
                contl=contl+0.5
                SKIP
                If CONTL>28
                  ENDPAGE
                  CABINF()
                  If LEFT(LINHA,2)="10"
                    Contl=contl+0.5
                    caja(contl,1,contl+0.01,20,oPrn,0,nRGB(0,0,0),oPen1)
                  EndIf
                EndIf
              ENDDO
            ENDDO
          EndIf
        EndIf
        If (LEFT(LINHA,2)="03" .OR. LEFT(LINHA,2)="04" .OR. LEFT(LINHA,2)="06" .OR. LEFT(LINHA,2)="07" .OR. LEFT(LINHA,2)="09") .And. WTPPESS="J"
          LOOP
        EndIf
        oprn:cmsay(contl,1,"Aprovado pela IN SRF N.698 de 2006.",oFontV8B,,nRGB(0,0,0),,0)
        contl=contl+0.5
        oprn:cmsay(contl,1,"ESTE INFORME DE RENDIMENTOS SUBSTITUI QUALQUER OUTRO ENVIADO ANTERIORMENTE.",oFontV8B,,nRGB(0,0,0),,0)
        ENDPAGE
      ENDDO
      ENDPRINT
      SELE 36
      USE
      Erase &ARQTXTO
      SELE 35
      USE
      SELE 32
      USE
      SELE 31
      USE
      Erase &NARQ4
      Erase &NARQ6
      Erase &NARQ5
      SELE 30
    return 

    Procedure CABINF()
      PAGE
      WCLIE=VAL(WCOTISTA)
      oprn:cmsay(0.5,0.5,"Cliente: "+ALLTRIM(TRANSFORM(WCLIE,"999999999")),oFontV8,,nRGB(255,255,255),,0) //Código do Cliente
      oprn:cmsay(0.5,7.0,"Assessor - ",oFontV8,,nRGB(255,255,255),,0) //Assessor
      caja(1.0,1,2.5,20,oPrn,0,nrgb(0,0,0)) // primeiro quadro
      caja(1.0,10,2.5,10.01,oPrn,0,nrgb(0,0,0),oPen1) //COLUNA DIVISORIA
      saybitmap(1.1,1.35,2.6,1.35,WEBLOCAL+"LOGOREC.bmp",oPrn)
      oPrn:cmsay(0.5,4.0,"INFORME DE RENDIMENTOS FINANCEIROS E POSIÇÃO PATRIMONIAL",oFontV10B,,nRGB(0,0,0),,0)
      oprn:cmsay(1.5,4.5,"MINISTÉRIO DA FAZENDA",oFontV8B,,nRGB(0,0,0),,0)
      oprn:cmsay(1.9,4.0,"SECRETARIA DA RECEITA FEDERAL",oFontV8B,,nRGB(0,0,0),,0)
      oprn:cmsay(1.1,10.5,"INFORME DE RENDIMENTOS FINANCEIROS",oFontV10B,,nRGB(0,0,0),,0)
      oprn:cmsay(1.5,10.5,"ANO CALENDÁRIO DE "+LEFT(WANOCALATU,4),oFontV10B,,nRGB(0,0,0),,0)
      oprn:cmsay(2.8,1,"1-IDENTIFICAÇÃO DA FONTE PAGADORA",oFontV8B,,nRGB(0,0,0),,0)
      caja(3.2,1,4.0,20,oPrn,0,nRGB(0,0,0),oPen1) //SEGUNDO quadro
      caja(3.2,6.5,4.0,6.51,oPrn,0,nRGB(0,0,0),oPen1) //LINHA DIVISORIA
      oprn:cmsay(3.2,1.3,"CNPJ",oFontV6B,,nRGB(0,0,0),,0)
      oprn:cmsay(3.2,7.0,"NOME EMPRESARIAL",oFontV6B,,nRGB(0,0,0),,0)
      oprn:cmsay(3.5,1.3,wfontcnpj,oFontV8,,nRGB(0,0,0),,0)
      oprn:cmsay(3.5,7.0,wfontnome,oFontV8,,nRGB(0,0,0),,0)
      caja(4.7,1,5.5,20,oPrn,0,nrgb(0,0,0),oPen1) //TERCEIRO quadro
      caja(4.7,6.5,5.5,6.51,oPrn,0,nRGB(0,0,0),oPen1) //LINHA DIVISORIA
      If WTPPESS="F"
        oprn:cmsay(4.3,1,"2. PESSOA FÍSICA BENEFICIARIA DOS RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(4.7,1.3,"CPF",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(4.7,7.0,"NOME",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(1.9,10.5,"IMPOSTO DE RENDA - PESSOA FISICA",oFontV10B,,nRGB(0,0,0),,0)
      Else
        oprn:cmsay(4.3,1,"2. PESSOA JURIDICA BENEFICIARIA DOS RENDIMENTOS",oFontV8B,,nRGB(0,0,0),,0)
        oprn:cmsay(4.7,1.3,"CNPJ",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(4.7,7.0,"RAZAO SOCIAL",oFontV6B,,nRGB(0,0,0),,0)
        oprn:cmsay(1.9,10.5,"IMPOSTO DE RENDA - PESSOA JURIDICA",oFontV10B,,nRGB(0,0,0),,0)
      EndIf
      //DADOS TERCEIRO QUADRO
      oprn:cmsay(5.0,1.3,WCPFCNPJ,oFontv8,,nRGB(0,0,0),,0)
      oprn:cmsay(5.0,7.0,WNOME,oFontv8,,nRGB(0,0,0),,0)
      contl=5.6
      WCAB=1
    return 

    Function decodeBase64(base64)
      DM:=TOleAuto():New("Microsoft.XMLDOM")
      EL = DM:createElement("tmp")
      EL:DataType = "bin.base64"
      EL:Text = base64
      decodificado=EL:NodeTypedValue
      EL:=nil
    return  decodificado

    Procedure LEXML(wbusc,wonde)
      WBUSC1="<"+ALLTRIM(WBUSC)+">"
      WBUSC2="</"+ALLTRIM(WBUSC)+">"
      If AT(WBUSC1,WONDE)=0 .OR. AT(WBUSC2,WONDE)=0
        return  ""
      EndIf
      WRET=SUBSTR(WONDE,(AT(WBUSC1,WONDE)+LEN(WBUSC1)),LEN(WONDE))
      WONDE=WRET
      WRET=LEFT(WONDE,(AT(WBUSC2,WONDE)-1))
    return  WRET

    Procedure RETNUL(QUAL,QUAL1)
      If VALTYPE(oRs:fields(QUAL):value)="U"
        If QUAL1="D"
          return  CTOD("")
        ElseIf QUAL1="N"
          return  0
        ElseIf QUAL1="C"
          return  ""
        EndIf
      Else
        return  (oRs:fields(QUAL):value)
      EndIf
    return 

    Procedure RETNUL1(QUAL,QUAL1)
      If VALTYPE(oRs1:fields(QUAL):value)="U"
        If QUAL1="D"
          return  CTOD("")
        ElseIf QUAL1="N"
          return  0
        ElseIf QUAL1="C"
          return  ""
        EndIf
      Else
        return  (oRs1:fields(QUAL):value)
      EndIf
    return 

    Procedure INTCAM(PARAM)
      SEENVIAEMAIL=PARAM
      SELE 30
      GO TOP
      CURSORWAIT()
      WEBEVIEWER=.F.
      DO WHILE .NOT. EOF()
        SysRefresh()
        If !WINIMAGE // Integra Sictur Winimage
          SKIP
          LOOP
        EndIf
        ARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        If !FILE(ARQ)
          SKIP
          LOOP
        EndIf
        SELE 30
        ARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        WEBLOCAL1=ALLTRIM(GeteNV("Tmp"))+"\BOLSASNET\"
        VERDIRETORIO(WEBLOCAL1)
        WEBLOCAL2=ALLTRIM(PASTA)+"ARQF\"
        VERDIRETORIO(WEBLOCAL2)
        WEBLOCAL3=ALLTRIM(PASTA)+"ECONOMIAECUSTOS\"
        VERDIRETORIO(WEBLOCAL3)
        WEBLOCAL4=alltrim(PASTA)+"DOCTOSV\"
        VERDIRETORIO(WEBLOCAL4)
        WEBVIEWER=WEBLOCAL4
        SELE 30
        WEBCUSTO=CUSTOMIL
        WEBCOR=CORRETORA
        WEBASSUNTO="INTEGRAÇÃO SICTUR/CSNU"
        WEBNOMECOR=OEMTOANSI(NOMECOR)
        WEBBOLSA=BOLSA
        WEBMAIL=EMAIL
        WEBCORPO="INTEGRAÇÃO SICTUR/CSNU "
        MCONTA=M->FROM
        WEBSERV=ALLTRIM(SMTP)
        WEBLOCAL=ALLTRIM(PASTA)
        WEBFORCA=FORCAEMSTP
        WEBLOGIN=LOGEMAIL
        WEBREL=EMAIL
        WEBEVIEWER=.F.
        WEBCSNU=.F.
        DECLARE VLABEL[ADIR(WEBLOCAL+"*.zip")]
        ADIR(WEBLOCAL+"*.zip",vlabel)
        If LEN(VLABEL)=0
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.ARJ")]
          ADIR(WEBLOCAL+"*.ARJ",vlabel)
        EndIf
        I=0
        FOR I=1 TO LEN(VLABEL)
          SysRefresh()
          CURSORWAIT()
          NARQ=WEBLOCAL+VLABEL[I]
          COMANDO=WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+WEBLOCAL+'"'
          If UPPER(LEFT(VLABEL[I],5))="CSNU_"
            WEBCSNU=.T.
          EndIf
          If UPPER(LEFT(VLABEL[I],6))="MOVCOR"
            COMANDO=WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+WEBVIEWER+'"'
            WEBEVIEWER=.T.
            NARQ7=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[I])
            LZCOPYFILE(NARQ,NARQ7)
          EndIf
          WAITRUN(COMANDO,0)
          Erase &NARQ
        NEXT I
        SELE 30
        DECLARE VLABEL1[ADIR(WEBLOCAL1+"*.*")]
        ADIR(WEBLOCAL1+"*.*",VLABEL1)
        FOR O=1 TO LEN(VLABEL1)
          SysRefresh()
          ARQDE=WEBLOCAL1+ALLTRIM(VLABEL1[O])
          Erase &ARQDE
        NEXT O
        WDADOS=WEBLOCAL1+"ENVIADO.TXT"
        If !FILE(WEBLOCAL+"CONEX.DAD") .And. !WEBEVIEWER
          If SEENVIAEMAIL="S"
            ENVIAEMAIL("wander@exatus.net","SICTUR /BolsasNet","Não configurado conexão ","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          Else
            MSGINFO("Não configurada a Conexão para integrar com o Sictur !",WSIST)
          EndIf
          Erase &ARQ
          SELE 30
          return 
        EndIf
        If WEBEVIEWER
          If !FILE(WEBLOCAL+"CONEXV.DAD")
            If SEENVIAEMAIL="S"
              ENVIAEMAIL("wander@exatus.net","SICTUR /BolsasNet","Não configurado conexão ","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            Else
              MSGINFO("Não configurada a Conexão para integrar com o Sictur !",WSIST)
            EndIf
            Erase &ARQ
            SELE 30
            return 
          EndIf
          ARQDE=WEBLOCAL+"CONEXV.DAD"
        ElseIf WEBCSNU
          ARQDE=WEBLOCAL+"CONEX.LIS"
        Else
          ARQDE=WEBLOCAL+"CONEX.DAD"
        EndIf
        RESTORE FROM &ARQDE ADDITIVE
        CONEXON=trim(Decrypt(CONEXON,"leopardo"))
        CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
        CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
        CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
        CONEXINTO=90
        CONEXCORRE=CONEXCORRE
        If !WEBEVIEWER
          DECLARE VLABEL[ADIR(WEBLOCAL+"*.TXT")]
          ADIR(WEBLOCAL+"*.TXT",vlabel)
          If CONEXCORRE=30961 .OR. CONEXCORRE=38031
            DECLARE VLABEL[ADIR(WEBLOCAL+"*.XML")]
            ADIR(WEBLOCAL+"*.XML",vlabel)
          EndIf
          If CONEXCORRE=27143 .And. WEBCSNU
            DECLARE VLABEL[ADIR(WEBLOCAL+"*.XML")]
            ADIR(WEBLOCAL+"*.XML",vlabel)
          EndIf
          I=0
          DECLARE TX1[1]
          TX1[1]="BolsasNet - SICTUR "+WEBCOR
          LOGFILE(WEBLOCAL+"ROTINAS.TXT",TX1)
          TX1[1]=WEBASSUNTO
          LOGFILE(WDADOS,TX1)
          If AUTOMATICO
            DECLARE TX1[1]
            TX1[1]="Executando SICTUR "+WEBCOR+" "+WEBLOCAL
            LOGFILE(WCORRENTE+"\ROTBOLSA.TXT",TX1)
          EndIf
          FEZALGO=.F.
          If LEN(VLABEL)#0
            If WEBCSNU
              INTCSNU()
              If SEENVIAEMAIL="S"
                If !EMPTY(WEBMAIL)
                  ENVIAEMAIL(WEBMAIL,"Integrado CSNU","Integrado LISTA CSNU Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
                EndIf
                ENVIAEMAIL("wander@exatus.net","Integrado LISTA CSNU","Integrado LISTA CSNU Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
              Else
                MSGINFO("Integrado lista CSNU",wsist)
              EndIf
            Else
              INTSICTUR()
            EndIf
          EndIf
          Erase &ARQ
        ElseIf WEBEVIEWER
          If INTSISFAIR()
            SELE 30
            CURSORARROW()
            If SEENVIAEMAIL="S"
              If !EMPTY(WEBMAIL)
                ENVIAEMAIL(WEBMAIL,"Integrado SISFAIR","Integrado SISFAIR Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
              EndIf
              ENVIAEMAIL("wander@exatus.net","Integrado SISFAIR","Integrado SISFAIR Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            Else
              MSGINFO("Integrado SISFAIR",wsist)
            EndIf
          Else
            SELE 30
            CURSORARROW()
            If SEENVIAEMAIL="S"
              If !EMPTY(WEBMAIL)
                ENVIAEMAIL(WEBMAIL,"Problema na integração SISFAIR","Problema na integração SISFAIR Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
              EndIf
              ENVIAEMAIL("wander@exatus.net","Problema na integração SISFAIR","Problema na integração SISFAIR Em:"+dtoc(Date())+" "+Time(),"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            Else
              MSGINFO("Não integrou SISFAIR",wsist)
            EndIf
          EndIf
        EndIf
        SELE 30
        ARQ=ALLTRIM(PASTA)+"FEITO.FEI"
        If FILE(ARQ)
          Erase &ARQ
        EndIf
        SKIP
      ENDDO
      DECLARE TX1[1]
      TX1[1]="Sictur - Fim do serviço "
      LOGFILE(wcorrente+"\ROTBOLSA.TXT",TX1)
      cursorarrow()
    return 

    Procedure INTCSNU()
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
        return 
      EndIf
      SET PACKETSIZE TO 25
      SELE 35
      If !USEREDE("listas",.F.,10,"listas","M")
        SQLDISCONNECT(CONEXINTO)
        return 
      EndIf
      FOR OI=1 TO LEN(VLABEL)
        NARQ6=ALLTRIM(WEBLOCAL)+ALLTRIM(VLABEL[OI])
        NARQ7=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
        LZCOPYFILE(NARQ6,NARQ7)
        NARQ=NARQ6
        lelayout(15)
        Erase &NARQ6
      NEXT OI
      DECLARE TX1[1]
      TX1[1]="lista CSNU - Fim do serviço"
      LOGFILE(wcorrente+"\ROTBOLSA.TXT",TX1)
      cursorarrow()
      SELE 35
      USE
      SQLDISCONNECT(CONEXINTO)
    return 

    Procedure INTSICTUR()
      DO WHILE .T.
        SysRefresh()
        M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
        NARQ3=WEBLOCAL1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".DAT"
        NARQ5=WEBLOCAL1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".CSV"
        NARQ9=WEBLOCAL1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".XLS"
        NARQ7=""
        If !FILE(NARQ3)
          EXIT
        EndIf
      ENDDO
      MAT_STRU={}
      AADD(MAT_STRU,{"NOMECLI","C",120,0})
      AADD(MAT_STRU,{"ENDCLI","C",60,0})
      AADD(MAT_STRU,{"NUMERO","C",10,0})
      AADD(MAT_STRU,{"COMPL","C",20,0})
      AADD(MAT_STRU,{"CEP","C",10,0})
      AADD(MAT_STRU,{"CIDADE","C",30,0})
      AADD(MAT_STRU,{"UF","C",2,0})
      AADD(MAT_STRU,{"CGCCPF","C",14,0})
      AADD(MAT_STRU,{"TPPES","C",1,0})
      AADD(MAT_STRU,{"NASC","C",10,0})
      AADD(MAT_STRU,{"EMAIL","C",200,0})
      AADD(MAT_STRU,{"DDD","C",3,0})
      AADD(MAT_STRU,{"FONE","C",10,0})
      AADD(MAT_STRU,{"DDDC","C",3,0})
      AADD(MAT_STRU,{"CELUL","C",10,0})
      AADD(MAT_STRU,{"CV","C",1,0})
      AADD(MAT_STRU,{"MOEDA","C",3,0})
      AADD(MAT_STRU,{"QUANTM","C",21,0})
      AADD(MAT_STRU,{"TAXA","C",21,0})
      AADD(MAT_STRU,{"REAIS","C",21,0})
      AADD(MAT_STRU,{"IOF","C",21,0})
      AADD(MAT_STRU,{"OUTRAS","M",10,0})
      AADD(MAT_STRU,{"PAIS","C",4,0})
      AADD(MAT_STRU,{"RECEBEDOR","C",120,0})
      AADD(MAT_STRU,{"NAT1","C",5,0})
      AADD(MAT_STRU,{"NAT2","C",2,0})
      AADD(MAT_STRU,{"NAT3","C",1,0})
      AADD(MAT_STRU,{"NAT4","C",2,0})
      AADD(MAT_STRU,{"NAT5","C",2,0})
      AADD(MAT_STRU,{"INTEGROU","C",1,0})
      AADD(MAT_STRU,{"ATUALIZOU","C",1,0})
      AADD(MAT_STRU,{"ENVBC","C",1,0})
      AADD(MAT_STRU,{"NRSICTUR","N",15,0})
      AADD(MAT_STRU,{"FILIAL","N",3,0})
      AADD(MAT_STRU,{"CARTAO","C",1,0})
      aAdd(mat_stru,{"BCO","C",10,0})
      aAdd(mat_stru,{"NOMEBCO","C",40,0})
      aAdd(mat_stru,{"AGENCIA","C",30,0})
      aAdd(mat_stru,{"CTA","C",30,0})
      aAdd(mat_stru,{"TPCTA","C",5,0})
      aAdd(mat_stru,{"CTAESTR","C",8,0})
      aAdd(mat_stru,{"CTAREAIS","C",8,0})
      aAdd(mat_stru,{"OPERADOR","C",8,0})
      aAdd(mat_stru,{"LIQUIDA","C",10,0})
      aAdd(mat_stru,{"REFER","C",20,0})
      aAdd(mat_stru,{"MOTIVO","C",250,0})
      aAdd(mat_stru,{"PAGTOEXT","C",1,0})
      aAdd(mat_stru,{"FORMA","C",2,0})
      aAdd(mat_stru,{"TPCART","C",30,0})
      aAdd(mat_stru,{"NRCART","C",30,0})
      aAdd(mat_stru,{"EMPVTM","C",5,0})
      aAdd(mat_stru,{"TIPOCART","C",1,0})
      aAdd(mat_stru,{"MAE","C",150,0})
      aAdd(mat_stru,{"NRQUEST","C",3,0})
      aAdd(mat_stru,{"RESPCART","C",100,0})
      aAdd(mat_stru,{"GEMCART","C",1,0})
      aAdd(mat_stru,{"DOCTO","C",30,0})
      aAdd(mat_stru,{"EMDOCTO","C",30,0})
      // NOVOS CAMPOS
      aAdd(mat_stru,{"VINCULO","C",2,0})
      aAdd(mat_stru,{"ENDPAGRE","C",50,0})
      aAdd(mat_stru,{"CIDPAGRE","C",30,0})
      aAdd(mat_stru,{"SWIFTIBA","C",1,0})
      aAdd(mat_stru,{"CODSWIFI","C",20,0})
      aAdd(mat_stru,{"NOMEBCO1","C",110,0})
      aAdd(mat_stru,{"ISWIFTIB","C",1,0})
      aAdd(mat_stru,{"ICODSWIFI","C",20,0})
      aAdd(mat_stru,{"NOMEBCOI","C",110,0})
      aAdd(mat_stru,{"DESPESAS","C",21,0})
      aAdd(mat_stru,{"TREAIS","C",21,0})
      aAdd(mat_stru,{"PAISPAGR","C",5,0})
      aAdd(mat_stru,{"ENDBCOREC","C",50,0})
      aAdd(mat_stru,{"IBANSN","C",1,0})
      aAdd(mat_stru,{"CDIBAM","C",50,0})
      aAdd(mat_stru,{"CTASWIFT","C",50,0})
      aAdd(mat_stru,{"IR","C",1,0})
      aAdd(mat_stru,{"PORCIR","C",16,0})
      aAdd(mat_stru,{"IRREC","C",1,0})
      aAdd(mat_stru,{"CDDARF","C",4,0})
      aAdd(mat_stru,{"REAJUS","C",1,0})
      aAdd(mat_stru,{"TXFISCAL","C",21,0})
      aAdd(mat_stru,{"BASEIR","C",21,0})
      DBCREATE(NARQ3,mat_stru,"DBFCDX")
      SysRefresh()
      mat_stru={}
      aAdd(mat_stru,{"LINHA","C",1200,0})
      DBCREATE(NARQ5,MAT_STRU,"DBFCDX")
      RELEASE mat_stru
      CURSORWAIT()
      ZAGTOEXT=""
      SELE 33
      If !USEREDE(NARQ3,.T.,10,"TEMP")
        SELE 30
        return 
      EndIf
      SELE 31
      If !USEREDE(NARQ5,.F.,10,"CSV")
        SELE 33
        USE
        SELE 30
        return 
      EndIf
      If CONEXCORRE#30961 .And. CONEXCORRE#38031
        FOR OI=1 TO LEN(VLABEL)
          If UPPER(VLABEL[OI])="ARQ_MORE.TXT" .OR. UPPER(VLABEL[OI])="ROTINAS.TXT"
            LOOP
          EndIf
          SysRefresh()
          CURSORWAIT()
          SELE 31
          NARQ6=ALLTRIM(WEBLOCAL)+ALLTRIM(VLABEL[OI])
          APPEND FROM &NARQ6 SDF
          NARQ7=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
          LZCOPYFILE(NARQ6,NARQ7)
          Erase &NARQ6
        NEXT OI
      EndIf
      SELE 31
      GO TOP
      ZEFER=""
      ZOCTO=""
      If FILE(WEBLOCAL+"ARQ_MORE.TXT")
        INTMOREREC()
      ElseIf CONEXCORRE=30961 .OR. CONEXCORRE=38031  // 38031=foco
        INTCHANGE()
      ElseIf CONEXCORRE=57506
        INTMOREPAG()
      ElseIf CONEXCORRE=29428
        INTCONSEG()
      Else
        INTOUTROS()
      EndIf
      MANDASICTUR()
    return 

    Procedure INTMOREREC()
      If !FILE(NARQ7)
        return 
      EndIf
      MDADFECHA=MEMOREAD(weblocal+"arq_more.txt")
      DO WHILE !EOF()
        SysRefresh()
        ZV="C"
        ZAIS=""
        ZOEDA=SUBSTR(MDADFECHA,AT("MOEDA...: ",MDADFECHA)+10,3)
        ZAXA=VAL(Strtran(ALLTRIM(SUBSTR(MDADFECHA,AT("TAXA....: ",MDADFECHA)+10,20)),",","."))
        NATUR=SUBSTR(MDADFECHA,AT("NATUREZA: ",MDADFECHA)+10,12)
        ZILIAL=VAL(SUBSTR(MDADFECHA,AT("FILIAL..: ",MDADFECHA)+10,12))
        ZTAESTR=SUBSTR(MDADFECHA,AT("CONTA ME: ",MDADFECHA)+10,3)
        ZTAREAIS=SUBSTR(MDADFECHA,AT("CONTA R$: ",MDADFECHA)+10,3)
        ZPERADOR=SUBSTR(MDADFECHA,AT("OPERADOR: ",MDADFECHA)+10,4)
        MD1=ALLTRIM(Strtran(LINHA,"|",";"))+"; ; ; ; ; ; ; ; ; ; ; ; ; ; ;"
        NADA=LELINHA(MD1)
        NADA=LELINHA(MD1)
        ZECEBEDOR=SINAL(ALLTRIM(LELINHA(MD1)))
        ZPAIS=SINAL(ALLTRIM(LELINHA(MD1)))
        ZMAIL=LOWER(ALLTRIM(LELINHA(MD1)))
        ZPCTA=SINAL(ALLTRIM(LELINHA(MD1)))
        ZCO=ALLTRIM(LELINHA(MD1))
        ZOMEBCO=SINAL(ALLTRIM(LELINHA(MD1)))
        ZGENCIA=ALLTRIM(LELINHA(MD1))
        ZTA=ALLTRIM(LELINHA(MD1))
        ZFONE=ALLTRIM(LELINHA(MD1))
        ZOMECLI=ALLTRIM(LELINHA(MD1))
        ZGCCPF=ALLTRIM(lelinha(MD1))
        ZEAIS=VAL(ALLTRIM(LELINHA(MD1)))
        NADA=LELINHA(MD1)
        ZEFER=ALLTRIM(LELINHA(MD1))
        If EMPTY(ZOMECLI) .OR. TRIM(ZOMECLI)="Sender"
          SKIP
          LOOP
        EndIf
        ZNDCLI="NAO IDENTIFICADO - MORE"
        ZUMERO=""
        ZOMPL=""
        ZEP="00000-000"
        ZIDADE=""
        ZF=""
        ZPPES="F"
        ZASC=""
        ZDD=""
        ZONE=""
        ZDDC=""
        ZELUL=""
        ZOF=ROUND(((ZEAIS*1.003814)*0.0038),2)
        If ZOF<0.01
          ZOF=0.01
        EndIf
        ZEAIS=ZEAIS+ZOF
        ZUANTM=ROUND((ZEAIS/ZAXA),2)
        If ZUANTM=0 .OR. ZAXA=0 .OR. ZEAIS=0
          SKIP
          LOOP
        EndIf
        ZUANTM=TRANSFORM(ZUANTM,"9999999999999999.99")
        ZAXA=TRANSFORM(ZAXA,"9999999.99999999")
        ZEAIS=TRANSFORM(ZEAIS,"9999999999999999.99")
        ZOF=TRANSFORM(ZOF,"9999999999999999.99")
        NATUR=VAL(NATUR)
        If NATUR=0
          NATUR=370040200290
        EndIf
        NATUR=STRZERO(NATUR,12)
        ZAT1=LEFT(NATUR,5)
        ZAT2=SUBSTR(NATUR,6,2)
        ZAT3=SUBSTR(NATUR,8,1)
        ZAT4=SUBSTR(NATUR,9,2)
        ZAT5=SUBSTR(NATUR,11,2)
        If val(zat3)<>0
          zat3="0"
        EndIf
        If val(zat4)<>2
          zat4="02"
        EndIf
        ZIQUIDA=dtoc(date())
        ZUTRAS="RETENCAO DO IOF DE 0,38 POR FORCA DO DECRETO NR.6339 DE 03/01/08 O COMPRADOR CONFIRMA,"+CHR(13)+CHR(10)+"DEPOSITO REALIZADO NO BANCO: "+TRIM(UPPER(ZOMEBCO))+"AG: "+ZGENCIA+" CTA: "+ZTA+CHR(13)+CHR(10)+"REF. MORE: "+ZEFER
        If ZPPES="F"
          ZGCCPF=STRZERO(VAL(ZGCCPF),11)
        ElseIf ZPPES="J"
          ZGCCPF=STRZERO(VAL(ZGCCPF),14)
        EndIf
        ZRSICTUR=0
        ZONSULTA=""
        ZNTEGROU=" "
        ZTUALIZOU=" "
        ZOTIVO=" "
        If VAL(ZGCCPF)=0 .or. val(ZGENCIA)=0 .or. val(ZGENCIA)>9999
          If VAL(ZGCCPF)=0
            ZOTIVO="CPF/CNPJ nao informado"
          ElseIf val(zgencia)=0
            zotivo="Agencia Nao Informada"
          ElseIf val(zgencia)>9999
            zotivo="Agencia Invalida"
          EndIf
          ZNVBC=""
        Else
          ZNVBC="R"
        EndIf
        ZARTAO="N"
        SELE 33
        ADIREG(0)
        PUTREC()
        SELE 31
        SKIP
      ENDDO
      SELE 31
      USE
      Erase &NARQ5
      SELE 33
      GO TOP
    return 

    Procedure INTOUTROS()
      DO WHILE !EOF()
        SysRefresh()
        MD1=ALLTRIM(LINHA)+"; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ;"
        ZOMECLI=UPPER(SINAL(ALLTRIM(LELINHA(MD1))))
        If EMPTY(ZOMECLI) .OR. TRIM(ZOMECLI)="NOME CLIENTE" .OR. TRIM(ZOMECLI)="NOME DO CLIENTE"
          SKIP
          LOOP
        EndIf
        ZESPESAS=""
        ZIQUIDA=""
        ZINCULO=""
        ZNDPAGRE=""
        ZIDPAGRE=""
        ZWIFTIBA=""
        ZODSWIFI=""
        ZOMEBCO1=""
        ZSWIFTIB=""
        ZCODSWIFI=""
        ZOMEBCOI=""
        ZESPESAS=""
        ZREAIS=""
        ZAISPAGR=""
        ZNDBCOREC=""
        ZBANSN=""
        ZDIBAM=""
        ZTASWIFT=""
        ZORCIR=""
        ZRREC=""
        ZDDARF=""
        ZEAJUS=""
        ZXFISCAL=""
        ZASEIR=""
        ZORMA=""
        ZR=""
        ZNDCLI=UPPER(SINAL(ALLTRIM(LELINHA(MD1))))
        ZUMERO=ALLTRIM(LELINHA(MD1))
        If val(zumero)>99999
          zumero="0"
        EndIf
        ZOMPL=SINAL(ALLTRIM(LELINHA(MD1)))
        ZEP=ALLTRIM(LELINHA(MD1))
        ZIDADE=SINAL(ALLTRIM(LELINHA(MD1)))
        ZF=ALLTRIM(LELINHA(MD1))
        ZGCCPF=ALLTRIM(lelinha(MD1))
        ZGCCPF=Strtran(Strtran(Strtran(ZGCCPF,".",""),"-",""),"/","")
        ZPPES=ALLTRIM(LELINHA(MD1))
        ZASC=ALLTRIM(LELINHA(MD1))
        ZMAIL=LOWER(ALLTRIM(LELINHA(MD1)))
        ZDD=ALLTRIM(LELINHA(MD1))
        ZONE=ALLTRIM(LELINHA(MD1))
        ZDDC=ALLTRIM(LELINHA(MD1))
        ZELUL=ALLTRIM(LELINHA(MD1))
        ZV=ALLTRIM(LELINHA(MD1))
        ZOEDA=ALLTRIM(LELINHA(MD1))
        ZUANTM=Strtran(ALLTRIM(LELINHA(MD1)),".","")
        ZUANTM=Strtran(ZUANTM,",",".")
        ZAXA=ALLTRIM(LELINHA(MD1))
        ZAXA=Strtran(ZAXA,".","")
        ZAXA=Strtran(ZAXA,",",".")
        ZEAIS=ALLTRIM(LELINHA(MD1))
        ZEAIS=Strtran(ZEAIS,".","")
        ZEAIS=Strtran(ZEAIS,",",".")
        ZNVBC="R"
        ZOTIVO=""
        cBuf=""
        If CONEXCORRE=57287 .OR. CONEXCORRE=27143 .OR. (CONEXCORRE=57434 .And. VAL(ZEAIS)=0)
          If conexcorre<>99999
            cBuf=conscpf(zgccpf,CTOD(zasc))
          Else
            cbuf=" REGULAR"
          EndIf
          If AT("REGULAR",cBuf)=0 .OR. AT("REGULARIZ",cBuf)<>0
            ZNVBC="I"
            ZOTIVO=cBuf
          EndIf
          If AT("NAO CONSUL",cBuf)<>0
            ZNVBC="N"
            ZOTIVO="NAO CONSULTOU"
          EndIf
        EndIf
        ZAIS=ALLTRIM(LELINHA(MD1))
        If ZAIS="BRA"
          ZAIS="1058"
        ElseIf ZAIS="US"
          ZAIS="2496"
        Else
          ZAIS="0000"
        EndIf
        ZECEBEDOR=SINAL(ALLTRIM(LELINHA(MD1)))
        ZAT1=ALLTRIM(LELINHA(MD1))
        ZAT2=ALLTRIM(LELINHA(MD1))
        ZAT3=ALLTRIM(LELINHA(MD1))
        ZAT4=ALLTRIM(LELINHA(MD1))
        ZAT5=ALLTRIM(LELINHA(MD1))
        ZOF=ALLTRIM(LELINHA(MD1))
        ZOF=Strtran(ZOF,".","")
        ZOF=Strtran(ZOF,",",".")
        ZUTRAS=SINAL(alltrim(LELINHA(MD1)))
        ZILIAL=VAL(ALLTRIM(LELINHA(MD1)))
        If ZILIAL=0 .or. zilial>999
          ZILIAL=1
        EndIf
        ZTAESTR=ALLTRIM(LELINHA(MD1))
        ZTAREAIS=ALLTRIM(LELINHA(MD1))
        ZPERADOR=ALLTRIM(LELINHA(MD1))
        ZORMA=ALLTRIM(LELINHA(MD1))
        ZOMEBCO=ALLTRIM(LELINHA(MD1))
        ZCO=ALLTRIM(LELINHA(MD1))
        ZGENCIA=ALLTRIM(LELINHA(MD1))
        ZTA=ALLTRIM(LELINHA(MD1))
        ZPCTA=ALLTRIM(LELINHA(MD1))
        ZPCART=ALLTRIM(LELINHA(MD1))
        ZRCART=ALLTRIM(LELINHA(MD1))
        ZMPVTM=ALLTRIM(LELINHA(MD1))
        ZIPOCART=ALLTRIM(LELINHA(MD1))
        ZAE=ALLTRIM(LELINHA(MD1))
        ZRQUEST=ALLTRIM(LELINHA(MD1))
        ZESPCART=ALLTRIM(LELINHA(MD1))
        ZEMCART=ALLTRIM(LELINHA(MD1))
        ZOCTO=ALLTRIM(LELINHA(MD1))
        ZMDOCTO=ALLTRIM(LELINHA(MD1))
        If EMPTY(ZUTRAS)
          If DATE()<CTOD("03/05/2016")
            ZUTRAS="RETENCAO DO IOF DE 0,38 POR FORCA DO DECRETO NR.6339 DE 03/01/08 O COMPRADOR CONFIRMA,"+CHR(13)+CHR(10)+"NESTE ATO, O RECEBIMENTO DO MONTANTE DA MOEDA ESTRANGEIRA ACIMA CARACTERIZADA"+CHR(13)+CHR(10)
          Else
            ZUTRAS="RETENCAO DO IOF DE 1,10 POR FORCA DO DECRETO NR.8731 DE 30/04/2016 O COMPRADOR CONFIRMA,"+CHR(13)+CHR(10)+"NESTE ATO, O RECEBIMENTO DO MONTANTE DA MOEDA ESTRANGEIRA ACIMA CARACTERIZADA"+CHR(13)+CHR(10)
          EndIf
        EndIf
        ZPLIQUID=val(ZIQUIDA)
        If CTOD(ZIQUIDA)<DATE()
          ZIQUIDA=DTOC(DATE())
        EndIf
        ZRSICTUR=0
        ZONSULTA=""
        ZNTEGROU=" "
        ZTUALIZOU=" "
        ZARTAO="N"
        If zilial>999 .OR. ZILIAL=0
          zilial=1
        EndIf
        If ZPPES="F"
          ZGCCPF=STRZERO(VAL(ZGCCPF),11)
        ElseIf ZPPES="J"
          ZGCCPF=STRZERO(VAL(ZGCCPF),14)
        EndIf
        SELE 33
        ADIREG(0)
        PUTREC()
        SELE 31
        SKIP
      ENDDO
      SELE 31
      USE
      Erase &NARQ5
      SELE 33
      GO TOP
    return 

    Procedure INTCONSEG()
      DO WHILE !EOF()
        SysRefresh()
        MD1=ALLTRIM(LINHA)+"; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ; ;"
        ZOMECLI=UPPER(SINAL(ALLTRIM(LELINHA(MD1))))
        If EMPTY(ZOMECLI) .OR. TRIM(ZOMECLI)="NOME CLIENTE"
          SKIP
          LOOP
        EndIf
        ZNDCLI=UPPER(SINAL(ALLTRIM(LELINHA(MD1))))
        ZUMERO=ALLTRIM(LELINHA(MD1))
        If val(zumero)>99999
          zumero="0"
        EndIf
        ZOMPL=SINAL(ALLTRIM(LELINHA(MD1)))
        ZEP=ALLTRIM(LELINHA(MD1))
        ZEP=LEFT(ZEP,5)+"-"+RIGHT(ZEP,3)
        ZIDADE=SINAL(ALLTRIM(LELINHA(MD1)))
        ZF=ALLTRIM(LELINHA(MD1))
        ZGCCPF=ALLTRIM(lelinha(MD1))
        ZPPES=ALLTRIM(LELINHA(MD1))
        ZASC=ALLTRIM(LELINHA(MD1))
        ZMAIL=ALLTRIM(LELINHA(MD1))
        ZDD=ALLTRIM(LELINHA(MD1))
        ZONE=ALLTRIM(LELINHA(MD1))
        ZONE=LEFT(ZONE,4)+"-"+RIGHT(ZONE,4)
        ZDDC=ALLTRIM(LELINHA(MD1))
        ZELUL=ALLTRIM(LELINHA(MD1))
        ZELUL=LEFT(ZELUL,5)+"-"+RIGHT(ZELUL,4)
        ZV=ALLTRIM(LELINHA(MD1))
        If VAL(ZV)=0
          If EMPTY(ZV)
            ZV="V"
          EndIf
        EndIf
        ZOEDA=ALLTRIM(LELINHA(MD1))
        ZUANTM=Strtran(lelinha(md1),",",".")
        ZAXA=ALLTRIM(LELINHA(MD1))
        ZAXA=Strtran(ZAXA,",",".")
        ZEAIS=ALLTRIM(LELINHA(MD1))
        ZEAIS=Strtran(ZEAIS,",",".")
        If VAL(ZUANTM)=0 .OR. VAL(ZAXA)=0 .OR. VAL(ZEAIS)=0
          SKIP
          LOOP
        EndIf
        ZAIS=ALLTRIM(LELINHA(MD1))
        ZECEBEDOR=SINAL(ALLTRIM(LELINHA(MD1)))
        ZAT1=ALLTRIM(LELINHA(MD1))
        ZAT2=ALLTRIM(LELINHA(MD1))
        ZAT3=ALLTRIM(LELINHA(MD1))
        ZAT4=ALLTRIM(LELINHA(MD1))
        ZAT5=ALLTRIM(LELINHA(MD1))
        ZOF=ALLTRIM(LELINHA(MD1))
        ZOF=Strtran(ZOF,",",".")
        ZUTRAS=SINAL(alltrim(LELINHA(MD1)))
        If CONEXCORRE<>27143
          ZIQUIDA=SUBSTR(ZUTRAS,12,10)
          ZIQUIDA=Strtran(ZIQUIDA,".","/")
        EndIf
        ZILIAL=VAL(ALLTRIM(LELINHA(MD1)))
        If ZILIAL=0 .or. zilial>999
          ZILIAL=1
        EndIf
        ZTAESTR=ALLTRIM(LELINHA(MD1))
        ZTAREAIS=ALLTRIM(LELINHA(MD1))
        IIF(VAL(ZTAESTR)=0,ZTAESTR:="193","")
        IIF(VAL(ZTAREAIS)=0,ZTAREAIS:="86","")
        ZPERADOR=ALLTRIM(LELINHA(MD1))
        If VAL(ZPERADOR=0)
          ZPERADOR="2"
        EndIf
        ZFORMAP=ALLTRIM(LELINHA(MD1))
        ZCO=ALLTRIM(LELINHA(MD1))
        ZOMEBCO=ALLTRIM(LELINHA(MD1))
        ZGENCIA=ALLTRIM(LELINHA(MD1))
        ZTA=ALLTRIM(LELINHA(MD1))
        ZPCTA=ALLTRIM(LELINHA(MD1))
        // CAMPOS NOVOS DIFERENTE DA GETMONEY
        ZINCULO=ALLTRIM(LELINHA(MD1))
        ZNDPAGRE=ALLTRIM(LELINHA(MD1))
        ZIDPAGRE=ALLTRIM(LELINHA(MD1))
        ZWIFTIBA=ALLTRIM(LELINHA(MD1))
        ZODSWIFI=ALLTRIM(LELINHA(MD1))
        ZOMEBCO1=ALLTRIM(LELINHA(MD1))
        ZSWIFTIB=ALLTRIM(LELINHA(MD1))
        ZCODSWIFI=ALLTRIM(LELINHA(MD1))
        ZOMEBCOI=ALLTRIM(LELINHA(MD1))
        ZESPESAS=ALLTRIM(LELINHA(MD1))
        ZESPESAS=Strtran(ZESPESAS,",",".")
        ZIQUIDO=ALLTRIM(LELINHA(MD1))
        ZIQUIDO=Strtran(ZIQUIDO,",",".")
        ZREAIS=ZIQUIDO
        ZAISPAGR=ALLTRIM(LELINHA(MD1))
        ZNDBCOREC=ALLTRIM(LELINHA(MD1))
        ZBANSN=ALLTRIM(LELINHA(MD1))
        ZDIBAM=ALLTRIM(LELINHA(MD1))
        ZTASWIFT=ALLTRIM(LELINHA(MD1))
        ZR=ALLTRIM(LELINHA(MD1))
        ZORCIR=ALLTRIM(LELINHA(MD1))
        ZRREC=ALLTRIM(LELINHA(MD1))
        ZDDARF=ALLTRIM(LELINHA(MD1))
        ZEAJUS=ALLTRIM(LELINHA(MD1))
        ZXFISCAL=ALLTRIM(LELINHA(MD1))
        ZASEIR=ALLTRIM(LELINHA(MD1))
        ZOTIVO=""
        ZRSICTUR=0
        ZONSULTA=""
        ZNTEGROU=" "
        ZTUALIZOU=" "
        ZNVBC="R"
        ZARTAO="N"
        If zilial>999 .OR. ZILIAL=0
          zilial=1
        EndIf
        If ZPPES="F"
          ZGCCPF=STRZERO(VAL(ZGCCPF),11)
        ElseIf ZPPES="J"
          ZGCCPF=STRZERO(VAL(ZGCCPF),14)
        EndIf
        SELE 33
        ADIREG(0)
        PUTREC()
        SELE 31
        SKIP
      ENDDO
      SELE 31
      USE
      Erase &NARQ5
      SELE 33
      GO TOP
    return 

    Procedure INTMOREPAG()
      DO WHILE !EOF()
        SysRefresh()
        ZV="V"
        MD1=ALLTRIM(Strtran(LINHA,"|",";"))+"; ; ; ; ; ; ; ; ; ; ; ; ; ; ;"
        NADA=LELINHA(MD1)
        ZSEQ=LELINHA(MD1)
        ZEFER=ALLTRIM(LELINHA(MD1))
        ZUANTM=VAL(LELINHA(MD1))
        ZUANTM=TRANSFORM(ZUANTM,"9999999999999999.99")
        ZAXA=VAL(LELINHA(MD1))
        ZAXA=TRANSFORM(ZAXA,"9999999.99999999")
        ZOEDA=LELINHA(MD1)
        ZOEDA="220"
        ZEAIS=VAL(LELINHA(MD1))
        ZEAIS=TRANSFORM(ZEAIS,"9999999999999999.99")
        NADA=LELINHA(MD1)
        ZOMECLI=SINAL(ALLTRIM(LELINHA(MD1)))+" "+SINAL(ALLTRIM(LELINHA(MD1)))
        ZOCTO=LELINHA(MD1)
        ZNDCLI=SINAL(LELINHA(MD1))
        ZUMERO=""
        ZOMPL=""
        ZEP="00000-000"
        ZIDADE=SINAL(LELINHA(MD1))
        ZF=""
        ZPPES="F"
        ZASC=""
        NADA=LELINHA(MD1)
        ZDD=LEFT(NADA,2)
        ZONE=SUBSTR(NADA,3,15)
        ZDDC=""
        ZELUL=""
        ZECEBEDOR=SINAL(ALLTRIM(LELINHA(MD1)))+" "+SINAL(ALLTRIM(LELINHA(MD1)))
        ZUTRAS="REFERENCIA: "+ZEFER+CHR(13)+CHR(10)+"ENDERECO RECEBEDOR: "+SINAL(ALLTRIM(LELINHA(MD1)))+" "+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10)
        ZAIS=LELINHA(MD1)
        ZUTRAS=ZUTRAS+"FONE RECEBEDOR: "+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10)
        ZUTRAS=ZUTRAS+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10) //MENSAGEM
        ZUTRAS=ZUTRAS+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10) //BANCO
        ZUTRAS=ZUTRAS+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10) //AGENCIA
        ZUTRAS=ZUTRAS+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10) //CONTA
        ZUTRAS=ZUTRAS+SINAL(ALLTRIM(LELINHA(MD1)))+CHR(13)+CHR(10) //TIPO DE CONTA
        ZGCCPF=Strtran(ALLTRIM(lelinha(MD1)),"-","")
        If EMPTY(ZOMECLI)
          SKIP
          LOOP
        EndIf
        ZILIAL=VAL(LELINHA(MD1))
        If ZILIAL<1 .or. zilial>999
          ZILIAL=1
        EndIf
        ZOF=VAL(LELINHA(MD1))
        If ZOF<0.01
          ZOF=0.01
        EndIf
        ZPERADOR=LELINHA(MD1)
        If VAL(ZPERADOR)=0
          ZPERADOR="2"
        EndIf
        ZOF=TRANSFORM(ZOF,"9999999999999999.99")
        NATUR=STRZERO(370040010090,12)
        ZAT1=LEFT(NATUR,5)
        ZAT2=SUBSTR(NATUR,6,2)
        ZAT3=SUBSTR(NATUR,8,1)
        ZAT4=SUBSTR(NATUR,9,2)
        ZAT5=SUBSTR(NATUR,11,2)
        If val(zat3)<>0
          zat3="0"
        EndIf
        If val(zat4)<>0
          zat4="00"
        EndIf
        If CONEXCORRE<>27143
          ZIQUIDA=dtoc(date())
        EndIf
        ZGCCPF=STRZERO(VAL(ZGCCPF),11)
        ZTAESTR="51"
        ZTAREAIS="51"
        ZPAIS=ZAIS
        ZMAIL=""
        ZPCTA=""
        ZCO=""
        ZOMEBCO=""
        ZGENCIA=""
        ZTA=""
        ZRSICTUR=0
        ZONSULTA=""
        ZNTEGROU=" "
        ZTUALIZOU=" "
        ZOTIVO=" "
        If VAL(ZGCCPF)=0
          ZOTIVO="CPF/CNPJ nao informado"
          ZNVBC=""
        Else
          ZNVBC="R"
        EndIf
        ZARTAO="N"
        SELE 33
        ADIREG(0)
        PUTREC()
        SELE 31
        SKIP
      ENDDO
      SELE 31
      USE
      Erase &NARQ5
      SELE 33
      GO TOP
    return 

    Procedure MANDASICTUR()
      CURSORWAIT()
      If CONEXCORRE<>57434
        SQL CONNECT ON CONEXON;
          PORT CONEXPORT;
          DATABASE CONEXDTBAS;
          USER CONEXUSER;
          PASSWORD CONEXSENHA;
          OPTIONS SQL_NO_WARNING;
          LIB 'MySQL';
          INTO CONEXINTO

        If SQL_ErrorNO() > 0
          return 
        EndIf
        SET PACKETSIZE TO 25
        SELE 35
        If CONEXCORRE=57287 .OR. CONEXCORRE=27143
          WMOVTO=strzero(CONEXCORRE,5)+"_PROPOSTAS"
        Else
          WMOVTO=strzero(CONEXCORRE,5)+"_MOVIMENTO"
        EndIf
        WCLIENTES=STRZERO(CONEXCORRE,5)+"_CLIENTES"
        If !USEREDE(WCLIENTES,.F.,10,"CLIENTES","M")
          SELE 33
          USE
          Erase &NARQ5
          SELE 30
          return 
        EndIf
        SELE 33
        GO TOP
        ATCLIENTE()
        COMMIT
        SELE 36
        If !USEREDE(WMOVTO,.F.,10,"MOVTO","M")
          SELE 33
          USE
          Erase &NARQ5
          SELE 30
          return 
        EndIf
        SELE 33
        GO TOP
        CONTT=0
        DO WHILE !EOF()
          If (ENVBC<>"R" .And. ENVBC<>"N") .or. VAL(QUANTM)=0 .OR. VAL(TAXA)=0 .OR. VAL(REAIS)=0 .OR. INTEGROU="S" .OR. NRSICTUR#0
            SKIP
            LOOP
          EndIf
          SELE 35
          If TEMP->TPPES="E"
            SQL FILTER TO "DOCUMENTO='"+ALLTRIM(TEMP->CGCCPF)+"'"
          Else
            SQL FILTER TO "CNPJCPF='"+ALLTRIM(TEMP->CGCCPF)+"'"
          EndIf
          GO TOP
          If TEMP->TPPES="E"
            If ALLTRIM(DOCUMENTO)<>ALLTRIM(TEMP->CGCCPF)
              SELE 33
              SKIP
              LOOP
            EndIf
          EndIf
          If TEMP->TPPES$"FJ"
            If VAL(CNPJCPF)<>VAL(TEMP->CGCCPF)
              SELE 33
              SKIP
              LOOP
            EndIf
          EndIf
          ZCODCLI=CODIGO
          ZPER=OPERADOR
          If CONEXCORRE=57434 .And. ZPER=0
            ZPER=2
          EndIf
          If CONEXCORRE=57287 .And. ZPER=0
            ZPER=2
          EndIf
          SELE 36
          ATSIC()
          If CONEXCORRE<>30961
            COMMIT
          EndIf
          SELE 33
          REPLACE INTEGROU WITH "S"
          REPLACE NRSICTUR WITH MOVTO->CODIGO
          If CONEXCORRE=57506
            REPLACE PAGTOEXT WITH "1"
          EndIf
          NRINT=NRSICTUR
          If CONEXCORRE<>30961
            COMMIT
          EndIf
          If INTEGROU<>"S" .OR. NRSICTUR=0
            SKIP
            LOOP
          EndIf
          If ATUALIZOU="S"
            SKIP
            LOOP
          EndIf
          If CONEXCORRE#30961 .And. CONEXCORRE#38031 .And. CONEXCORRE#57287 .And. CONEXCORRE#27143
            SELE 36
            COMMIT
            SELE 33
            If SEENVIAEMAIL="S"
              ATSALDO()
            EndIf
          Else
            ATUALIZOU="S"
          EndIf
          If ATUALIZOU="S"
            If FILE(WEBLOCAL+"ARQ_MORE.TXT")
              If SEENVIAEMAIL="S"
                ATSALDOSTR()
              EndIf
            EndIf
          EndIf
          If CONEXCORRE=30961
            CONTT=CONTT+1
            If CONTT>1000
              CONTT=0
              COMMIT
            EndIf
          EndIf
          SELE 33
          SKIP
        ENDDO
      EndIf
      ARQ_XLS=WEBLOCAL3+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+"CUSTO.XLS"
      NARQ6=WEBLOCAL3+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+"RETORNO.TXT"
      GERAXLS(2,"Relatório Integrações "+WEBCOR+" "+WEBASSUNTO)
      NARQ7=WEBLOCAL+"ARQ_MORE.TXT"
      NARQ8=""
      SELE 36
      USE
      SELE 35
      USE
      If FILE(WEBLOCAL+"ARQ_MORE.TXT")
        GERARQSTR()
      EndIf
      SQLDISCONNECT(CONEXINTO)
      SELE 30
      CURSORARROW()
      If SEENVIAEMAIL="S"
        //  ENVIAEMAIL("wander@exatus.net","BolsasNet -Sictur"+WEBCOR,"Em anexo relatorio",WDADOS,ARQ_XLS,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
        If !EMPTY(WEBMAIL)
          ENVIAEMAIL(WEBMAIL,"BolsasNet -Sictur "+WEBCOR,"Em anexo relatorio",WDADOS,ARQ_XLS,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          If FILE(NARQ7) .OR. FILE(NARQ8)
            ENVIAEMAIL(WEBMAIL,"BolsasNet -Sictur - Retorno More/STR "+WEBCOR,"Em anexo relatorio",NARQ6,NARQ8,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
            ENVIAEMAIL("wander@exatus.net","BolsasNet -Sictur Retorno More/STR "+WEBCOR,"Em anexo relatorio",NARQ6,NARQ8,"exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
          EndIf
        EndIf
      EndIf
      Erase &NARQ6
      Erase &NARQ7
    return 

    Procedure ATCLIENTE()
      CONTT=0
      DO WHILE !EOF()
        If VAL(TEMP->CGCCPF)=0 .OR. ENVBC<>"R"
          SKIP
        LOOP
      EndIf
      SELE 35
      SQL FILTER TO "CNPJCPF='"+ALLTRIM(TEMP->CGCCPF)+"'"
      GO TOP
      MCEP=ALLTRIM(Strtran(TEMP->CEP,"-",""))
      MCEP=LEFT(MCEP,5)+"-"+RIGHT(MCEP,3)
      If VAL(CNPJCPF)=VAL(TEMP->CGCCPF) .And. REGLOCK(10)
        BEGIN TRANSACTION
        If CEP<>MCEP .And. !empty(mcep)
          REPLACE CEP WITH MCEP
        EndIf
        If LOGRADOURO<>TEMP->ENDCLI .And. !empty(temp->endcli)
          REPLACE LOGRADOURO WITH (TEMP->ENDCLI)
        EndIf
        If NUMERO<>VAL(TEMP->NUMERO) .And. val(temp->numero)#0
          REPLACE NUMERO WITH VAL(TEMP->NUMERO)
        EndIf
        If TEMP->COMPL<>COMPL .And. !empty(temp->compl)
          REPLACE COMPL WITH TEMP->COMPL
        EndIf
        If CIDADE<>TEMP->CIDADE .And. !empty(temp->cidade)
          REPLACE CIDADE WITH (TEMP->CIDADE)
        EndIf
        If UF<>(TEMP->UF) .And. !empty(temp->uf)
          REPLACE UF WITH (TEMP->UF)
        EndIf
        If DOCUMENTO<>TEMP->DOCTO
          REPLACE DOCUMENTO WITH ALLTRIM(TEMP->DOCTO)
        EndIf
        If ctod(TEMP->NASC)<>DTNASC .And. (conexcorre=57287 .or. conexcorre=27143)
          REPLACE DTNASC WITH CTOD(TEMP->NASC)
        EndIf
        If FILIALCLI<>TEMP->FILIAL .And. (conexcorre=57287 .or. conexcorre=27143)
          REPLACE FILIALCLI WITH TEMP->FILIAL
        EndIf
        If DOCUMENTO<>TEMP->DOCTO
          REPLACE DOCUMENTO WITH TEMP->DOCTO
        EndIf
      END TRANSACTION
    EndIf
      If VAL(CNPJCPF)#VAL(TEMP->CGCCPF)
        BEGIN TRANSACTION
        ADIREG(0)
        REPLACE NOME WITH (TEMP->NOMECLI)
        REPLACE CEP WITH MCEP
        REPLACE DOCUMENTO WITH TEMP->DOCTO
        REPLACE LOGRADOURO WITH (TEMP->ENDCLI)
        REPLACE NUMERO WITH VAL(TEMP->NUMERO)
        REPLACE COMPL WITH TEMP->COMPL
        REPLACE BAIRRO WITH "-"
        REPLACE CIDADE WITH (TEMP->CIDADE)
        REPLACE UF WITH (TEMP->UF)
        REPLACE TIPO WITH TEMP->TPPES
        REPLACE CNPJCPF WITH TEMP->CGCCPF
        REPLACE DOCUMENTO WITH TEMP->DOCTO
        If TEMP->TPPES="E"
          REPLACE DOCUMENTO WITH TEMP->CGCCPF
        EndIf
        REPLACE DDD WITH TEMP->DDD
        MFONE=Strtran(TEMP->FONE,"-","")
        MFONE=Strtran(MFONE,".","")+SPACE(15)
        MFONE=LEFT(MFONE,4)+"-"+SUBSTR(MFONE,5,4)
        REPLACE TELEFONE WITH MFONE
        REPLACE DDDCEL WITH TEMP->DDDC
        MFONE=Strtran(TEMP->CELUL,"-","")
        MFONE=ALLTRIM(Strtran(MFONE,".","")+SPACE(15))
        If LEN(MFONE)=9
          MFONE=LEFT(MFONE,5)+"-"+SUBSTR(MFONE,6,4)
        Else
          MFONE=LEFT(MFONE,4)+"-"+SUBSTR(MFONE,5,4)
        EndIf
        REPLACE CELULAR WITH MFONE
        REPLACE EMAIL WITH TEMP->EMAIL
        If TEMP->TPPES="F" .OR. TEMP->TPPES="E"
          REPLACE NATCLI WITH 95
          If TEMP->TPPES="E"
            REPLACE NATCLI WITH 99
          EndIf
        Else
          REPLACE NATCLI WITH 50
        EndIf
        If !EMPTY(TEMP->NASC)
          ZASC=TEMP->NASC
          If CONEXCORRE<>27143
            ZASC=LEFT(TEMP->NASC,4)+"/"+SUBSTR(TEMP->NASC,6,2)+"/"+SUBSTR(TEMP->NASC,9,2)
          EndIf
          If CONEXCORRE=57434
            ZASC=TEMP->NASC
          EndIf
          If ctod(ZASC)<>CTOD("")
            REPLACE DTNASC WITH CTOD(ZASC)
          EndIf
        EndIf
        REPLACE LIMITE WITH "D"
        REPLACE BLOQUEADO WITH "N"
        If CONEXCORRE=27143
          REPLACE MENSAL WITH 10000.00
          REPLACE TRIMESTRAL WITH 30000.00
          REPLACE ANUAL WITH 120000.00
        Else
          REPLACE MENSAL WITH 3000.00
          REPLACE TRIMESTRAL WITH 9000.00
          REPLACE ANUAL WITH 36000.00
        EndIf
        REPLACE FILIALCLI WITH TEMP->FILIAL
        replace naturalidade with "-"
        REPLACE INATIVO WITH "N"
        REPLACE USERINC WITH "WORKER-CSV-TXT"
        REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
        REPLACE CONTATO WITH "  "
        If CONEXCORRE=30961 .OR. CONEXCORRE=38031 .OR. CONEXCORRE=57287 .OR. CONEXCORRE=27143
          REPLACE PAIS WITH VAL(TEMP->PAIS)
        Else
          REPLACE PAIS WITH 2496
        EndIf
        If VAL(TEMP->CGCCPF)#0
          REPLACE PAIS WITH 1058
        EndIf
        REPLACE OPERADOR WITH VAL(TEMP->OPERADOR)
        If CONEXCORRE=57434
          REPLACE OPERADOR WITH 2
        EndIf
        If CONEXCORRE=57287
          REPLACE OPERADOR WITH 1106
        EndIf
        REPLACE ADIG WITH "N"
        If VAL(TEMP->PAIS)#0 .And. (CONEXCORRE=57287 .OR. CONEXCORRE=27143)
          REPLACE PAIS WITH VAL(TEMP->PAIS)
        EndIf
      END TRANSACTION
      CONTT=CONTT+1
      If CONTT>1000
        COMMIT
        CONTT=0
      EndIf
    EndIf
    SELE 33
    SKIP
    ENDDO
  return 

  Procedure GERARQSTR()
    If CONEXCORRE=57506
      VENDO=49
      WFIL=1
    Else
      return 
    EndIf
    SELE 35
    MOVCONT=STRZERO(CONEXCORRE,5)+"_MOVCONTABIL"
    If !USEREDE(MOVCONT,.F.,10,"MOVTO","M")
      SELE 35
      USE
      return 
    EndIf
    DT1="'"+STRZERO(YEAR(DATE()),4)+"-"+STRZERO(MONTH(DATE()),2)+"-"+STRZERO(DAY(DATE()),2)+"'"
    WFILT="ATIVO="+str(VENDO)+" AND datal="+dt1+" AND CONC='C' AND periodo=0 AND DC='S' AND OPERACAO>0"
    SQL FILTER TO WFILT
    GO TOP
    MCORPO=MEMOREAD(WEBLOCAL+"CORPO.BAS")
    MDADOS=MEMOREAD(WEBLOCAL+"DADOS.BAS")
    NARQ8=ALLTRIM(WEBLOCAL3)+"STR"+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+".ENV"
    Erase &NARQ8
    If CONEXCORRE=57506
      STRCNPJCOR="60746948"
      STRNOMECOR="S.HAYATA CORRETORA DE CAMBIO S/A"
    Else
      SELE 35
      USE
      return 
    EndIf
    GO TOP
    wdados=""
    DO WHILE !EOF()
      If MOVTO->DC<>"S"
        SKIP
      LOOP
    EndIf
    STRINTERNO=STRZERO(MOVTO->OPERACAO,10)
    If MOVTO->OPERACAO=0
      STRINTERNO=DTOS(DATA)+STRZERO(RECNO(),7)
    EndIf
    STRAGENCIA=STRZERO(VAL(SUBSTR(MOVTO->HIST3,5,4)),4)
    STRTPCTA=SUBSTR(MOVTO->HIST3,10,2)
    STRNOME=ALLTRIM(SINAL(MOVTO->FAVORECIDO))
    STRVALOR=ALLTRIM(TRANSFORM(MOVTO->VALOR,"@E 999999999.99"))
    STRHIST=ALLTRIM(SINAL(MOVTO->HIST1))
    STRCTA=ALLTRIM(SUBSTR(MOVTO->HIST3,13,13))
    STRMOTIVO="68"  // COLOCAR 1 QUANDO FOR INTERBANCARIO
    STRBCO=LEFT(MOVTO->HIST3,3)
    WNRC="SELECT ISPB FROM BANCOS WHERE CODIGO="+LEFT(MOVTO->HIST3,3)
    MZ=SQLArrayAssoc(WNRC)
    wcodif=MZ[1,1]
    STRCNPJBCO=STRZERO(VAL(WCODIF),8)
    WNRC="SELECT TPPESSOA FROM "+STRZERO(CONEXCORRE,5)+"_MOVIMENTO WHERE CODIGO="+STRZERO(MOVTO->OPERACAO,11)
    MZ=SQLArrayAssoc(WNRC)
    wcodif=MZ[1,1]
    STRTPPESSOA=ALLTRIM(WCODIF)
    If STRTPPESSOA="J"
      WNRC="SELECT CGCCLI FROM "+STRZERO(CONEXCORRE,5)+"_MOVIMENTO WHERE CODIGO="+STRZERO(MOVTO->OPERACAO,11)
    ElseIf STRTPPESSOA="F"
      WNRC="SELECT CPFCLI FROM "+STRZERO(CONEXCORRE,5)+"_MOVIMENTO WHERE CODIGO="+STRZERO(MOVTO->OPERACAO,11)
    EndIf
    MZ=SQLArrayAssoc(WNRC)
    wcodif=MZ[1,1]
    STRCNPJCPF=ALLTRIM(WCODIF)
    If VAL(STRINTERNO)>0 .And. VAL(STRAGENCIA)>0 .And. STRTPCTA$"CCCDPGPP" .And. MOVTO->VALOR>0 .And. !EMPTY(STRNOME) .And. !EMPTY(STRCTA) .And. at("INTERBANCARIO",MOVTO->HIST1)=0 .And. VAL(STRINTERNO)>0 .And. STRTPPESSOA$"FJ" .And. VAL(STRCNPJCPF)>0 .And. REGLOCK(10)
      BEGIN TRANSACTION
      REPLACE PERIODO WITH 1
    END TRANSACTION
    COMMIT
    WDADOS=WDADOS+PREENCHEXML(MDADOS)
  EndIf
    SKIP
    ENDDO
    If !EMPTY(WDADOS)
      WXML=PREENCHEXML(MCORPO)
      MEMOWRIT(NARQ8,WXML)
    EndIf
    SELE 35
    USE
  return 

  Procedure PREENCHEXML(LINH)
    CURSORWAIT()
    while AT(";;",LINH)#0
      SysRefresh()
      IN=AT(";;",LINH)
      FI=AT(";:",LINH)
      TAM=LEN(LINH)
      TPC=ALLTRIM(SUBSTR(LINH,IN+2,FI-2-IN))
      TPC=Strtran(TPC,";","")
      TPC=Strtran(TPC,":","")
      If EMPTY(TPC)
        return  LINH
      EndIf
      If VALTYPE(&TPC)="U"
        return  LINH
      EndIf
      VARI1=""
      If VALTYPE(&TPC)="C"
        VARI1=&TPC
      ElseIf VALTYPE(&TPC)="D"
        VARI2=DTOS(&TPC)
        VARI1=LEFT(VARI2,4)+"-"+SUBSTR(VARI2,5,2)+"-"+RIGHT(VARI2,2)
      ElseIf VALTYPE(&TPC)="N"
        VARI1=ALLTRIM(STR(&TPC))
        If VAL(VARI1)=0
          VARI1=""
        EndIf
      EndIf
      LINH=TRIM(LEFT(LINH,IN-1)+OEMTOANSI(VARI1)+SUBSTR(LINH,FI+2,TAM))
    ENDDO
    If AT(CHR(13),LINH)=0
      LINH=LINH+CHR(13)+CHR(10)
    EndIf
  return  LINH

  Function GRAV(ARQU,REGISTR)
    FOR ZI=1 TO LEN(REGISTR)
      CONTAD=SUBSTR(REGISTR,ZI,1)
      If CONTAD=CHR(177) .OR. CONTAD=CHR(9)
        CONTAD=CHR(13)+CHR(10)+CONTAD+CHR(13)+CHR(10)
      EndIf
      FWRITE (ARQ, CONTAD, LEN (CONTAD))
      If FERROR () != 0
        MSGINFO ("Erro na Gravação do Arquivo","")
        FCLOSE (ARQ)
      EndIf
    NEXT ZI
  return  .T.

  Procedure ATSALDO()
    mnar1=Strtran(WEBLOCAL1,"/","\")+"aguarde.txt"
    mnar2=Strtran(WEBLOCAL1,"/","\")+"fazer.bat"
    mnar3=Strtran(WEBLOCAL1,"/","\")+"log1.txt"
    mnar4=Strtran(WEBLOCAL1,"/","\")+"log2.txt"
    Erase &mnar1
    Erase &mnar2
    Erase &mnar3
    Erase &mnar4
    ATUA=.F.
    CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+"/wget.exe -T120 -o"+mnar3+" -O"+mnar4
    wcta=STRZERO(VAL(TEMP->BCO),3)+STRZERO(VAL(TEMP->AGENCIA),4)+LEFT(ALLTRIM(TEMP->TPCTA),2)+ALLTRIM(TEMP->CTA)
    If SEENVIAEMAIL="S"
      CMACRO=Strtran(CMACRO,"/","\")+' "https://sictur.exatus.net/wsatualizasaldoh.asp?codbacen='+strzero(CONEXCORRE,5)+"&cta="+wcta+"&nrinterno="+strzero(NRINT,12)+'" --no-check-certificate'+CHR(13)+CHR(10)+"DEL "+mnar1+CHR(13)+CHR(10)+"DEL "+MNAR2+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    Else // PARA TESTE COM A CORRETORA 777777 precisa modificar o wsatualizasaldoh.asp
      //  CMACRO=Strtran(CMACRO,"/","\")+' "https://sictur.exatus.net/wsatualizasaldoh.asp?codbacen='+strzero(CONEXCORRE,5)+"&cta="+wcta+"&nrinterno="+strzero(NRINT,12)+'" --no-check-certificate'+CHR(13)+CHR(10)+"DEL "+mnar1+CHR(13)+CHR(10)+"DEL "+MNAR2+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
      return 
      //  CMACRO=Strtran(CMACRO,"/","\")+' "https://sictur.exatus.net/wsatualizasaldoh.asp?codbacen='+strzero(77777,5)+"&cta="+wcta+"&nrinterno="+strzero(NRINT,12)+'" --no-check-certificate'+CHR(13)+CHR(10)+"DEL "+mnar1+CHR(13)+CHR(10)+"DEL "+MNAR2+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    EndIf
    MEMOWRIT(MNAR2,CMACRO)
    MEMOWRIT(MNAR1,CMACRO)
    WAITRUN(MNAR2,0)
    DO WHILE FILE(MNAR1) .And. FILE(MNAR2)
    ENDDO
    V_ERRO=MEMOREAD(MNAR4)
    Erase &mnar1
    Erase &mnar2
    Erase &mnar3
    Erase &mnar4
    If AT("{success: true, msg: 'Saldos atualizados.'}",V_ERRO)#0
      ATUA=.T.
    Else
      DECLARE TX1[1]
      TX1[1]=V_ERRO
      LOGFILE(WDADOS,TX1)
    EndIf
    SELE 33
    If ATUA
      REPLACE ATUALIZOU WITH "S"
    Else
      REPLACE MOTIVO WITH V_ERRO
    EndIf
  return 

  Procedure ATSALDOSTR()
    mnar1=Strtran(WEBLOCAL1,"/","\")+"aguarde.txt"
    mnar2=Strtran(WEBLOCAL1,"/","\")+"fazer.bat"
    mnar3=Strtran(WEBLOCAL1,"/","\")+"log1.txt"
    mnar4=Strtran(WEBLOCAL1,"/","\")+"log2.txt"
    Erase &mnar1
    Erase &mnar2
    Erase &mnar3
    Erase &mnar4
    ATUA=.F.
    CMACRO="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+"/wget.exe -T120 -o"+mnar3+" -O"+mnar4
    dtbx=strzero(day(date()),2)+"/"+strzero(month(date()),2)+"/"+strzero(year(date()),4)
    If SEENVIAEMAIL="S"
      CMACRO=Strtran(CMACRO,"/","\")+' "https://sictur.exatus.net/wsbancariobaixaoper.asp?codbacen='+strzero(CONEXCORRE,5)+"&nrlancamento="+strzero(NRSICTUR,15)+"&databaixa="+dtbx+'" --no-check-certificate'+CHR(13)+CHR(10)+"DEL "+mnar1+CHR(13)+CHR(10)+"DEL "+MNAR2+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    EndIf
    MEMOWRIT(MNAR2,CMACRO)
    MEMOWRIT(MNAR1,CMACRO)
    WAITRUN(MNAR2,0)
    DO WHILE FILE(MNAR1) .And. FILE(MNAR2)
    ENDDO
    V_ERRO=MEMOREAD(MNAR4)
    Erase &mnar1
    Erase &mnar2
    Erase &mnar3
    Erase &mnar4
    SELE 33
  return 

  Procedure ATSIC()
    BEGIN TRANSACTION
    ADIREG(0)
    If (CONEXCORRE<>57287 .And. CONEXCORRE<>27143)
      REPLACE SIMPLIf WITH "N"
      REPLACE ENVBC WITH .T.
      REPLACE TAXAORG WITH VAL(TEMP->TAXA)
      REPLACE VALUEDATE WITH CTOD(TEMP->LIQUIDA)
      REPLACE DTENTMOEDA WITH CTOD("")
      REPLACE LIQUIDADO WITH " "
      If CONEXCORRE=57506
        REPLACE PAGTOEXT WITH "1"
      EndIf
      REPLACE CONTAVTM WITH VAL(TEMP->CTAESTR)
      REPLACE NAT1 WITH VAL(TEMP->NAT1)
      REPLACE NAT2 WITH VAL(TEMP->NAT2)
      REPLACE NAT3 WITH VAL(TEMP->NAT3)
      REPLACE NAT4 WITH VAL(TEMP->NAT4)
      REPLACE NAT5 WITH VAL(TEMP->NAT5)
      REPLACE LIQUIDACAO WITH CTOD(TEMP->LIQUIDA)
      If LIQUIDACAO<DATE()
        REPLACE LIQUIDACAO WITH DATE()
      EndIf
      REPLACE DTENTMOEDA WITH DATE()
      REPLACE CAIXA WITH OPERADOR
      REPLACE NOMEPAG WITH TEMP->RECEBEDOR
      REPLACE PAISPAG WITH TEMP->PAIS
      REPLACE VINCULO WITH "20"
      REPLACE IR WITH '0'
      REPLACE INTEGRADO WITH "N"
      REPLACE ADIG WITH "N"
      REPLACE ACIC WITH "S"
      REPLACE GEROUACIC WITH "N"
      REPLACE DESATIVADA WITH "N"
      REPLACE CONFIRMADA WITH "S"
      REPLACE DI WITH 0
      REPLACE TIPOOPCAM WITH "IB"
      REPLACE IR WITH STRZERO(VAL(TEMP->IR),1)
      REPLACE PORCIR WITH VAL(TEMP->PORCIR)
      REPLACE IRREC WITH VAL(TEMP->IRREC)
      REPLACE CDDARF WITH VAL(TEMP->CDDARF)
      REPLACE REAJUS WITH STRZERO(VAL(TEMP->REAJUS),1)
      REPLACE TXFISCAL WITH VAL(TEMP->TXFISCAL)
      REPLACE BASEIR WITH VAL(TEMP->BASEIR)
      REPLACE VTM WITH "0"
      REPLACE TIPOVTM WITH "0"
      If val(TEMP->NRCART)>0
        REPLACE VTM WITH "1"
        REPLACE EMPVTM WITH TEMP->EMPVTM
        REPLACE TIPOVTM WITH "1"
        REPLACE NRCARTAO WITH TEMP->NRCART
        REPLACE NR_QUESTAO WITH VAL(TEMP->NRQUEST)
        REPLACE RESPOSTA WITH TEMP->RESPCART
        REPLACE NOMEMAE WITH TEMP->MAE
        REPLACE GENERO WITH TEMP->GEMCART
        REPLACE NRRG WITH TEMP->DOCTO
      EndIf
    Else
      REPLACE VTM WITH "0"
      REPLACE TIPOVTM WITH "0"
      If val(TEMP->NRCART)>0
        REPLACE VTM WITH "1"
        REPLACE EMPVTM WITH TEMP->EMPVTM
        REPLACE TIPOVTM WITH "1"
        REPLACE NRCARTAO WITH TEMP->NRCART
        REPLACE NR_QUESTAO WITH VAL(TEMP->NRQUEST)
        REPLACE RESPOSTA WITH TEMP->RESPCART
        REPLACE NOMEMAE WITH TEMP->MAE
        REPLACE GENERO WITH TEMP->GEMCART
        REPLACE NRRG WITH TEMP->DOCTO
      EndIf
    EndIf
    If CONEXCORRE=99999
      REPLACE CBIOF WITH '0'
    Else
      REPLACE CBIOF WITH '1'
    EndIf
    REPLACE PAISVIAGEM WITH 0
    If TEMP->CV="C"
      REPLACE TIPO WITH '3'
    Else
      REPLACE TIPO WITH '4'
    EndIf
    If VAL(TEMP->CV)#0
      REPLACE TIPO WITH TEMP->CV
    EndIf
    REPLACE CLIENTE WITH ZCODCLI
    REPLACE TPPESSOA WITH TEMP->TPPES
    If TEMP->TPPES="F"
      REPLACE CPFCLI WITH TEMP->CGCCPF
      REPLACE CGCCLI WITH ""
    ElseIf TEMP->TPPES="J"
      REPLACE CGCCLI WITH TEMP->CGCCPF
      REPLACE CPFCLI WITH ""
    ElseIf TEMP->TPPES="E"
      REPLACE CGCCLI WITH ""
      REPLACE CPFCLI WITH "ISENTO"
      REPLACE DOCUMENTO WITH TEMP->CGCCPF
    EndIf
    REPLACE NOMECLI WITH TEMP->NOMECLI
    REPLACE ENDCLI1 WITH ALLTRIM(TEMP->ENDCLI)+" "+ALLTRIM(TEMP->NUMERO)+" "+ALLTRIM(TEMP->COMPL)
    REPLACE ENDCLI2 WITH ALLTRIM(TEMP->CIDADE)+" "+ALLTRIM(TEMP->UF)
    If TEMP->FILIAL#0
      REPLACE FILIAL WITH TEMP->FILIAL
    Else
      REPLACE FILIAL WITH 1
    EndIf
    REPLACE DATA WITH DATE()
    If TEMP->MOEDA="USD"
      REPLACE MOEDA WITH 220
    ElseIf TEMP->MOEDA="EUR"
      REPLACE MOEDA WITH 978
    ElseIf TEMP->MOEDA="ARS"
      REPLACE MOEDA WITH 706
    Else
      REPLACE MOEDA WITH VAL(TEMP->MOEDA)
    EndIf
    REPLACE ESTRANGEIR WITH VAL(TEMP->QUANTM)
    REPLACE REAIS WITH VAL(TEMP->REAIS)
    REPLACE TAXA WITH VAL(TEMP->TAXA)
    REPLACE CONTA WITH VAL(TEMP->CTAREAIS)
    REPLACE IOF WITH VAL(TEMP->IOF)
    REPLACE DESPESAS WITH VAL(TEMP->DESPESAS)
    If TIPO$"246"
      If VAL(CBIOF)=0
        REPLACE VLROPER WITH (REAIS+DESPESAS)
      Else
        REPLACE VLROPER WITH (REAIS+IOF+DESPESAS)
      EndIf
    Else
      If VAL(CBIOF)=0
        REPLACE VLROPER WITH (REAIS-DESPESAS)
      Else
        REPLACE VLROPER WITH (REAIS-IOF-DESPESAS)
      EndIf
    EndIf
    REPLACE OUTRAS WITH TEMP->OUTRAS
    REPLACE FORMAENTRE WITH 50
    If CONEXCORRE=57434
      REPLACE FORMAENTRE WITH 55
    EndIf
    REPLACE OPERADOR WITH VAL(TEMP->OPERADOR)
    If CONEXCORRE=57434 .And. OPERADOR=0
      REPLACE OPERADOR WITH 2
    EndIf
    If CONEXCORRE=57287 .And. OPERADOR=0
      REPLACE OPERADOR WITH 1106
      REPLACE CODUSUARIO WITH 1106
    EndIf
    REPLACE USERINC WITH "WORKER - SICTUR"
    REPLACE INCLUIDO WITH DTOC(DATE())+" "+TIME()
    REPLACE TPLIQUID with VAL(TEMP->FORMA)
    If CONEXCORRE<>30961
      REPLACE CODUSUARIO WITH VAL(TEMP->OPERADOR)
    EndIf
    If CONEXCORRE=29428
      REPLACE SIMPLIf WITH "S"
      REPLACE FORMAENTRE WITH 65
      REPLACE VINCULO WITH TEMP->VINCULO
      REPLACE ENDPAG WITH TEMP->ENDPAGRE
      REPLACE CIDPAG WITH TEMP->CIDPAGRE
      REPLACE TIPO1PAG WITH TEMP->SWIFTIBA
      REPLACE CODIGO1PAG WITH TEMP->CODSWIFI
      REPLACE BANCO1PAG WITH TEMP->NOMEBCO1
      REPLACE TIPO2PAG WITH TEMP->ISWIFTIB
      REPLACE CODIGO2PAG WITH TEMP->ICODSWIFI
      REPLACE BANCO2PAG WITH TEMP->NOMEBCOI
      REPLACE DESPESAS WITH VAL(TEMP->DESPESAS)
      If TIPO$"246"
        REPLACE VLROPER WITH VLROPER+VAL(TEMP->DESPESAS)
      Else
        REPLACE VLROPER WITH VLROPER-VAL(TEMP->DESPESAS)
      EndIf
      REPLACE VLROPER WITH VAL(TEMP->TREAIS)
      REPLACE PAISBANCO1PAG WITH TEMP->PAISPAGR
      REPLACE ENDBANCO1PAG WITH TEMP->ENDBCOREC
      REPLACE IBANPAG WITH TEMP->IBANSN
      If IBANPAG="S"
        REPLACE CODIBANPAG WITH TEMP->CDIBAM
      Else
        REPLACE CODIBANPAG WITH TEMP->CTASWIFT
      EndIf
    EndIf
    If CONEXCORRE=30961 .OR. CONEXCORRE=38031
      REPLACE DATA WITH CTOD(TEMP->LIQUIDA)
      REPLACE LIQUIDACAO WITH CTOD(TEMP->LIQUIDA)
    Else
      If (CONEXCORRE<>57287 .And. CONEXCORRE<>27143)
        DELE
      EndIf
    EndIf
  END TRANSACTION
return 

Procedure PUXASERVIDORES()
  DEFINE DIALOG oDlgg FROM 1,1 TO 22,70 TITLE "Servidores" BRUSH oBrushTela STYLE WS_CAPTION
  @ 0.4,24 BITMAP oBmp1 RESOURCE "PROCURA" ADJUST NOBORDER OF oDlgg SIZE 80,110
  @   0.44,0.5 SAY " " BOX RAISED SIZE 185,75
  WDE=1
  WATE=72987
  @ 00.87,1 say "Pagina de: " FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @ 01.74,1 say "Pagina até:" FONT oFontsay COLOR CLR_BLACK,nRGB(248,248,248)
  @ 01,10 GET WDE FONT oFontget PICTURE "99999" SIZE 45,11
  @ 02,10 GET WATE FONT oFontget PICTUR "99999" SIZE 45,11
  previa=0
  @ 07,10 SBUTTON PROMPT "&Iniciar" RESOURCE "PROC" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) ACTION (PREVIA:=2,oDlgg:End()) FONT oFontbut SIZE 40,35
  @ 07,19 SBUTTON PROMPT "&Cancela" RESOURCE "SAIR2" TEXT ON_BOTTOM COLOR 0,RGB(255,255,255) ACTION (PREVIA:=0,oDlgg:End()) FONT oFontbut SIZE 40,35
  oDlgg:lhelpicon:=.f.
  ACTIVATE DIALOG oDlgg CENTERED
  If PREVIA=0
    return 
  EndIf
  locat=wcorrente+"\pesq\"
  If !VERDIRETORIO(locat)
    MSGINFO("Não foi possivel criar o diretorio "+locat,WSIST)
    QUIT
  EndIf
  DARQ=locat+"FZ.BAT"
  ARQ100=FCREATE(DARQ)
  BSAVES=FWRITE(ARQ100,"@ECHO OFF"+CHR(13)+CHR(10),11)
  cmacro="@ECHO OFF"+CHR(13)+CHR(10)
  darq1=locat+"AGUARDE.TXT"
  MEMOWRIT(locat+"AGUARDE.TXT",cmacro)
  FOR IO=WDE TO WATE
    SysRefresh()
    WNOMEARQ=LOCAT+"PG_"+STRZERO(IO,6)+".TXT"
    If FILE(wnomearq)
      Erase &wnomearq
    EndIf
    COMANDO=WCORRENTE+"\wget.exe -T10 -olog.txt -O"+WNOMEARQ+" "+`"http://www.portaldatransparencia.gov.br/servidores/Servidor-ListaServidores.asp?bogus=1&Pagina='+ALLTRIM(TRANSFORM(IO,"999999"))+'" -U --user-agent=MOZILLA'+CHR(13)+CHR(10)
    BSAVE=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
  NEXT IO
  COMANDO="DEL "+DARQ1+CHR(13)+CHR(10)+"DEL "+DARQ+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
  BSAVES=FWRITE(ARQ100,COMANDO,LEN(COMANDO))
  FCLOSE(ARQ100)
  CURSORWAIT()
  WAITRUN(DARQ,0)
  CURSORWAIT()
  DO WHILE FILE(DARQ1) .And. FILE(DARQ)
    SysRefresh()
  ENDDO
  CURSORWAIT()
  Erase &DARQ
  DARQ=LOCAT+"SERV.DB"
  MAT_STRU={}
  AADD(mat_stru,{"A01CPF","C",06,0})   // Código na Exatus
  AADD(mat_stru,{"a01nome","C",50,0})   // Nome
  AADD(MAT_STRU,{"A01ORGAO","C",60,0})  // Endereço
  DBCREATE(DARQ,MAT_STRU)
  SELE 1
  If !USEREDE(DARQ,.F.,10,"SERVIDOR")
    return 
  EndIf
  DECLARE VLABEL[ADIR(LOCAT+"PG_*.TXT")]
  ADIR(LOCAT+"PG_*.TXT",vlabel)
  WTOT=0
  FOR I=1 TO LEN(VLABEL)
    SysRefresh()
    NARQ=LOCAT+VLABEL[I]
    LIDO1=MEMOREAD(NARQ)
    LIDO=SUBSTR(LIDO1,AT(">***",LIDO1),LEN(LIDO1))
    DO WHILE AT(">***",LIDO)#0
      WCPF=Strtran(SUBSTR(LIDO,AT(">***",LIDO)+5,7),".","")
      WNOME=SUBSTR(LIDO,AT("IdServidor",LIDO),150)
      WNOME=SUBSTR(WNOME,AT(">",WNOME)+1,150)
      WNOME=ALLTRIM(LEFT(WNOME,AT("</a>",WNOME)-1))
      LIDO=SUBSTR(LIDO,AT(WNOME,LIDO)+LEN(WNOME),LEN(LIDO))
      WORGAO=SUBSTR(LIDO,AT('<td style="text-align: left;">',LIDO)+30,100)
      WORGAO=ALLTRIM(LEFT(WORGAO,AT("</td>",WORGAO)-1))
      ADIREG(0)
      REPLACE A01CPF WITH WCPF
      REPLACE A01NOME WITH WNOME
      REPLACE A01ORGAO WITH WORGAO
      WTOT=WTOT+1
    ENDDO
  NEXT I
  SELE 1
  USE
  MSGINFO("TERMINOU "+STRZERO(WTOT,7),WSIST)
return 

Procedure AJUSTABASE()
  // Cadastro de Tabelas *
  mat_stru = {}
  AADD(mat_stru,{"DOC","C",11,0})
  AADD(mat_stru,{"NASC","C",6,0})
  AADD(mat_stru,{"NOME","C",31,0})
  AADD(mat_stru,{"END","C",23,0})
  AADD(mat_stru,{"NUM","C",6,0})
  AADD(mat_stru,{"COMPL","C",15,0})
  AADD(mat_stru,{"BAIRRO","C",15,0})
  AADD(mat_stru,{"CEP","C",8,0})
  AADD(mat_stru,{"CIDADE","C",29,0})
  AADD(mat_stru,{"UF","C",02,0})
  AADD(mat_stru,{"DDD","C",4,0})
  AADD(mat_stru,{"FONE","C",8,0})
  AADD(mat_stru,{"ATIV","C",3,0})
  AADD(mat_stru,{"RD","N",19,5})
  AADD(mat_stru,{"SEXO","C",1,0})
  AADD(mat_stru,{"ESTCIV","C",1,0})
  AADD(mat_stru,{"CMI","C",1,0})
  AADD(mat_stru,{"FX","C",1,0})
  AADD(mat_stru,{"TEMFONE","C",1,0})
  AADD(mat_stru,{"PERIODO","N",4,0})
  AADD(mat_stru,{"ALEATO","N",10,0})
  AADD(mat_stru,{"PCPF","C",6,0})
  DBCREATE("C:\T5\NAT_BF.DBF",mat_stru)
  SELE 1
  If !USEREDE("C:\T5\NAT_BF.DBF",.F.,10,"EMP")
    MsgStop("NAT_BF.DBF não está  disponível", WSIST)
    CLOSE DATABASE
    return 
  EndIf
  SELE 2
  If !USEREDE("C:\T5\FISICOES.DBF",.F.,10,"FIS")
    MsgStop("FISICOES.DBF não está  disponível", WSIST)
    CLOSE DATABASE
    return 
  EndIf
  SELE 3
  If !USEREDE("C:\T5\SERV.DBF",.F.,10,"SERV")
    MsgStop("SERV.DBF não está  disponível", WSIST)
    CLOSE DATABASE
    return 
  EndIf
  SELE 2
  GO TOP
  DO WHILE !EOF()
    If EMPTY(NOME)
      SKIP
      LOOP
    EndIf
    SELE 1
    ADIREG(0)
    REPLACE DOC WITH encrypt(FIS->DOC,"leopardo")
    REPLACE NASC WITH encrypt(FIS->NASC,"leopardo")
    REPLACE NOME WITH encrypt(FIS->NOME,"leopardo")
    REPLACE END WITH encrypt(FIS->END,"leopardo")
    REPLACE NUM WITH encrypt(FIS->NUM,"leopardo")
    REPLACE COMPL WITH encrypt(FIS->COMPL,"leopardo")
    REPLACE BAIRRO WITH encrypt(FIS->BAIRRO,"leopardo")
    REPLACE CEP WITH encrypt(FIS->CEP,"leopardo")
    REPLACE CIDADE WITH encrypt(FIS->CIDADE,"leopardo")
    REPLACE UF WITH FIS->UF
    REPLACE DDD WITH encrypt(FIS->DDD,"leopardo")
    REPLACE FONE WITH encrypt(FIS->FONE,"leopardo")
    REPLACE ATIV WITH FIS->ATIV
    REPLACE RD WITH FIS->RD
    REPLACE SEXO WITH FIS->SEXO
    REPLACE ESTCIV WITH FIS->ESTCIV
    REPLACE CMI WITH FIS->CMI
    REPLACE FX WITH FIS->FX
    REPLACE TEMFONE WITH FIS->TEMFONE
    REPLACE PERIODO WITH 0
    REPLACE ALEATO WITH HB_RANDOM(1,10000000)
    REPLACE PCPF WITH SUBSTR(FIS->DOC,4,6)
    SELE 2
    SKIP
  ENDDO
  SELE 1
  INDEX ON val(PCPF) tag pcpf TO c:\t5\T1.ndx
  SELE 3
  GO TOP
  DO WHILE !EOF()
    SELE 1
    SEEK val(SERV->A01CPF)
    DO WHILE VAL(serv->A01CPF)=VAL(PCPF)
      If LEFT(DECRYPT(NOME,"leopardo"),10) = LEFT(NOME,10)
        REGLOCK(10)
        REPLACE PERIODO WITH 9999
      EndIf
      SKIP
    ENDDO
    SELE 3
    SKIP
  ENDDO
  CLOSE DATABASE
  MSGINFO("TERMINOU !",WSIST)
return 

Procedure CRIADB(NMARQ)
  Erase &NMARQ
  mat_stru={}
  aAdd(mat_stru,{"a10datam","D",08,0})
  aAdd(mat_stru,{"a10docto","C",06,0})
  aAdd(mat_stru,{"a10codrd","C",05,0})
  aAdd(mat_stru,{"a10digrd","C",01,0})
  aAdd(mat_stru,{"a10codcd","C",10,0})
  aAdd(mat_stru,{"a10digcd","C",01,0})
  aAdd(mat_stru,{"a10codrc","C",05,0})
  aAdd(mat_stru,{"a10digrc","C",01,0})
  aAdd(mat_stru,{"a10codcc","C",10,0})
  aAdd(mat_stru,{"a10digcc","C",01,0})
  aAdd(mat_stru,{"a10valor","N",19,2})
  aAdd(mat_stru,{"a10hist1","C",40,0})
  aAdd(mat_stru,{"a10hist2","C",40,0})
  aAdd(mat_stru,{"a10hist3","C",40,0})
  aAdd(mat_stru,{"a10flag1","C",01,0})
  aAdd(mat_stru,{"a10filia","N",05,0})
  dbcreate(NMARQ,mat_stru)
  release mat_stru
return 

Procedure INTCHANGE()
  FOR OI=1 TO LEN(VLABEL)
    SysRefresh()
    SELE 31
    NARQ6=ALLTRIM(WEBLOCAL)+ALLTRIM(VLABEL[OI])
    NARQ7=ALLTRIM(WEBLOCAL2)+DTOS(DATE())+LEFT(TIME(),2)+SUBSTR(TIME(),4,2)+ALLTRIM(VLABEL[OI])
    LZCOPYFILE(NARQ6,NARQ7)
    Cxml=MEMOREAD(NARQ6)
    If at("<?xml  version=",cXML)=0
      If LEFT(Cxml,1)="F" .OR. LEFT(Cxml,1)="J" .OR. LEFT(Cxml,1)="E"
        INTCC()
        Erase &NARQ6
      EndIf
      LOOP
    EndIf
    SysRefresh()
    CONTT=0
    DO WHILE .T.
      If empty(LEXML("ideDeclarado",Cxml))
        EXIT
      EndIf
      DECLARADO=LEXML("ideDeclarado",Cxml)
      mvto=left(Cxml,at("</evtMovOpFin>",Cxml))
      mvto=substr(mvto,at("<anoMesCaixa>",mvto)-13,len(mvto))
      Cxml=substr(Cxml,at("<ideDeclarado>",Cxml)+14,len(Cxml))
      If !EMPTY(LEXML("ideDeclarado",Cxml))
        Cxml=substr(Cxml,at("<ideDeclarado>",Cxml)-14,len(Cxml))
      Else
        CXML=""
      EndIf
      ZOMECLI=(UPPER(alltrim(LEXML("NomeDeclarado",DECLARADO))))
      ZGCCPF=LEXML("NIDeclarado",DECLARADO)
      ZNDERECO=(UPPER(alltrim(LEXML("EnderecoLivre",DECLARADO))))
      IIF(empty(ZNDERECO),ZNDERECO:="NAO INFORMADO",)
      ZPAIS=LEXML("Pais",DECLARADO)
      ZPDOC=LEXML("tpNI",DECLARADO)
      CONTT=CONTT+1
      If CONTT>1000
        COMMIT
        CONTT=0
      EndIf
      DO WHILE .T.
        If empty(LEXML("anoMesCaixa",mvto))
          EXIT
        EndIf
        wref=LEXML("anoMesCaixa",mvto)
        WVAL=VAL(Strtran(LEXML("totCompras",MVTO),",","."))
        If WVAL#0
          GEREF("V","32016",wval)
        EndIf
        WVAL=VAL(Strtran(LEXML("totVendas",MVTO),",","."))
        If WVAL#0
          GEREF("C","32016",wval)
        EndIf
        WVAL=VAL(Strtran(LEXML("totTransferencias",mvto),",","."))
        If WVAL#0
          GEREF("V","67500",wval)
        EndIf
        MVTO=substr(mvto,at("<anoMesCaixa>",mvto)+13,len(mvto))
        If AT("<anoMesCaixa>",mvto)>0
          mvto=substr(mvto,at("<anoMesCaixa>",mvto)-13,len(mvto))
        Else
          MVTO=""
        EndIf
      ENDDO
    ENDDO
    Erase &NARQ6
  NEXT OI
  cursorarrow()
  SELE 31
  USE
  SELE 33
  GO TOP
return 

Procedure LEIAXML(wbusc,wonde)
  WBUSC1="<"+ALLTRIM(WBUSC)
  WBUSC2="</"+ALLTRIM(WBUSC)
  If AT(WBUSC1,WONDE)=0 .OR. AT(WBUSC2,WONDE)=0
    return  ""
  EndIf
  WRET=SUBSTR(WONDE,(AT(WBUSC1,WONDE)+LEN(WBUSC1)),LEN(WONDE))
  WONDE=WRET
  WRET=LEFT(WONDE,(AT(WBUSC2,WONDE)-1))
  WRET=Strtran(WRET,CHR(10),"")
  WRET=Strtran(WRET,CHR(13),"")
return  WRET

Procedure GEREF(ZCV,ZAT11,WVAL1)
  If WVAL1<0
    WVAL=WVAL1*-1
  Else
    WVAL=WVAL1
  EndIf
  SysRefresh()
  ZV=ZCV
  IIF(ZPAIS="",ZAIS:="1058",ZAIS:="1058")
  ZOEDA="220"
  ZAXA=1
  NATUR=ZAT11+"0000000"
  ZILIAL=1
  ZTAESTR="00"
  ZTAREAIS="00"
  ZPERADOR="0"
  ZECEBEDOR=""
  ZMAIL=""
  ZPCTA=""
  ZCO=""
  ZOMEBCO=""
  ZGENCIA=""
  ZTA=""
  ZFONE=""
  ZEAIS=WVAL
  ZEFER=""
  ZNDCLI=ZNDERECO
  ZUMERO=""
  ZOMPL=""
  ZEP="00000-000"
  ZIDADE=""
  ZF=""
  ZR=""
  ZORCIR=""
  ZRREC=""
  ZDDARF=""
  ZEAJUS=""
  ZXFISCAL=""
  ZASEIR=""
  If ZPDOC="1"
    ZPPES="F"
  ElseIf ZPDOC="2"
    ZPPES="J"
  Else
    ZPPES="E"
  EndIf
  ZASC=""
  ZDD=""
  ZONE=""
  ZDDC=""
  ZELUL=""
  ZOF=0
  ZEAIS=ZEAIS+ZOF
  ZUANTM=ROUND((ZEAIS/ZAXA),2)
  ZUANTM=TRANSFORM(ZUANTM,"9999999999999999.99")
  ZAXA=TRANSFORM(ZAXA,"9999999.99999999")
  ZEAIS=TRANSFORM(ZEAIS,"9999999999999999.99")
  ZOF=TRANSFORM(ZOF,"9999999999999999.99")
  NATUR=VAL(NATUR)
  NATUR=STRZERO(NATUR,12)
  ZAT1=LEFT(NATUR,5)
  ZAT2=SUBSTR(NATUR,6,2)
  ZAT3=SUBSTR(NATUR,8,1)
  ZAT4=SUBSTR(NATUR,9,2)
  ZAT5=SUBSTR(NATUR,11,2)
  ZIQUIDA="01/"+right(wref,2)+"/"+left(wref,4)
  ZUTRAS="IMPORTADO DO XML CHANGE"
  If ZPPES="F"
    ZGCCPF=STRZERO(VAL(ZGCCPF),11)
  ElseIf ZPPES="J"
    ZGCCPF=STRZERO(VAL(ZGCCPF),14)
  EndIf
  ZRSICTUR=0
  ZONSULTA=""
  ZNTEGROU=" "
  ZTUALIZOU=" "
  ZOTIVO=" "
  ZNVBC="R"
  ZARTAO="N"
  ZINCULO=""
  ZNDPAGRE=""
  ZIDPAGRE=""
  ZWIFTIBA=""
  ZODSWIFI=""
  ZOMEBCO1=""
  ZSWIFTIB=""
  ZCODSWIFI=""
  ZOMEBCOI=""
  ZESPESAS=""
  ZREAIS=""
  ZAISPAGR=""
  ZNDBCOREC=""
  ZBANSN=""
  ZDIBAM=""
  ZTASWIFT=""
  SELE 33
  If CONEXCORRE<>30961
    ADIREG(0)
    PUTREC()
  Else
    APPEND BLANK
    If RLOCK()
      REPLACE NOMECLI WITH ZOMECLI,ENDCLI WITH ZNDCLI,CGCCPF WITH ZGCCPF,TPPES WITH ZPPES,CV WITH ZV,MOEDA WITH ZOEDA,QUANTM WITH ZUANTM,TAXA WITH ZAXA,REAIS WITH ZEAIS,OUTRAS WITH ZUTRAS,PAIS WITH ZAIS,NAT1 WITH ZAT1,ENVBC WITH ZNVBC,FILIAL WITH ZILIAL,CARTAO WITH ZARTAO,LIQUIDA WITH ZIQUIDA,CTAESTR WITH ZTAESTR,CTAREAIS WITH ZTAREAIS,OPERADOR WITH ZPERADOR
    EndIf
  EndIf
return 

Procedure INTCC()
  SELE 34
  USE
  MAT_STRU={}
  AADD(MAT_STRU,{"TPPESSOA","C",1,0})
  AADD(MAT_STRU,{"CNPJCPF","C",14,0})
  AADD(MAT_STRU,{"PASSAPORTE","C",30,0})
  AADD(MAT_STRU,{"NOMECLI","C",80,0})
  AADD(MAT_STRU,{"ENDER","C",44,0})
  AADD(MAT_STRU,{"NUMERO","C",5,0})
  AADD(MAT_STRU,{"COMPL","C",10,0})
  AADD(MAT_STRU,{"BAIRRO","C",50,0})
  AADD(MAT_STRU,{"CEP","C",9,0})
  AADD(MAT_STRU,{"CIDADE","C",20,0})
  AADD(MAT_STRU,{"UF","C",2,0})
  AADD(MAT_STRU,{"TITCTA","C",1,0})
  AADD(MAT_STRU,{"SUBCTA","C",3,0})
  AADD(MAT_STRU,{"TIPOCTA","C",7,0})
  AADD(MAT_STRU,{"NRCTA","C",50,0})
  AADD(MAT_STRU,{"RELDECL","C",1,0})
  AADD(MAT_STRU,{"NRTITULARES","C",1,0})
  If CONEXCORRE=38031
    AADD(MAT_STRU,{"MESANO","C",6,0})
  Else
    AADD(MAT_STRU,{"MESANO","C",7,0})
  EndIf
  AADD(MAT_STRU,{"TOTCRED","C",19,0})
  AADD(MAT_STRU,{"TOTDEB","C",19,0})
  AADD(MAT_STRU,{"TOTCREDMT","C",19,0})
  AADD(MAT_STRU,{"TOTDEBMT","C",19,0})
  AADD(MAT_STRU,{"TPFATCA","C",8,0})
  AADD(MAT_STRU,{"TOTPAGTOS","C",19,0})
  AADD(MAT_STRU,{"SALDODEZ","C",19,0})
  AADD(MAT_STRU,{"PAISISOFATCA","C",2,0})
  AADD(MAT_STRU,{"NIF","C",30,0})
  AADD(MAT_STRU,{"PAISRES","C",2,0})
  DBCREATE(NARQ9,mat_stru,"DBFCDX")
  SysRefresh()
  mat_stru={}
  SELE 34
  If !USEREDE(NARQ9,.T.,10,"TEMPCC")
    SELE 30
    return 
  EndIf
  APPEND FROM &NARQ6 SDF
  GO TOP
  If EOF()
    SELE 34
    USE
    Erase &NARQ9
    return 
  EndIf

  SQL CONNECT ON CONEXON;
    PORT CONEXPORT;
    DATABASE CONEXDTBAS;
    USER CONEXUSER;
    PASSWORD CONEXSENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXINTO

  If SQL_ErrorNO() > 0
    return 
  EndIf
  SET PACKETSIZE TO 25
  SELE 35
  WMOVTO="CCEFINA"
  If !USEREDE(WMOVTO,.F.,10,"MOVTO","M")
    SELE 34
    USE
    Erase &NARQ9
    SELE 30
    return 
  EndIf
  SELE 34
  GO TOP
  DO WHILE !EOF()
    DECLARE _VAR1[FCOUNT()]
    AFIELDS(_VAR1)
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      &_var4 = &_var5
    NEXT t
    RELEASE _VAR1
    SELE 35
    ZIPOREG="CC"
    ZNDERECO=TRIM(TEMPCC->ENDER)+","+TRIM(TEMPCC->NUMERO)+" "+TRIM(TEMPCC->COMPL)+" "+TRIM(TEMPCC->BAIRRO)+"-"+TRIM(TEMPCC->CIDADE)+"-"+TEMPCC->UF+" "+TEMPCC->CEP
    ZOTCRED=VAL(TEMPCC->TOTCRED)
    ZOTDEB=VAL(TEMPCC->TOTDEB)
    If ZOTCRED<0
      ZOTCRED=ZOTCRED*-1
    EndIf
    If ZOTDEB<0
      ZOTDEB=ZOTDEB*-1
    EndIf
    ZRTITULARES=TEMPCC->NRTITULARE
    If VAL(ZRTITULARES)<1
      ZRTITULARES="1"
    EndIf
    ZESANO=RIGHT(ZESANO,4)+LEFT(ZESANO,2)
    ZOTCREDMT=VAL(TEMPCC->TOTCREDMT)
    ZOTDEBMT=VAL(TEMPCC->TOTDEBMT)
    If ZOTCREDMT<0
      ZOTCREDMT=ZOTCREDMT*-1
    EndIf
    If ZOTDEBMT<0
      ZOTDEBMT=ZOTDEBMT*-1
    EndIf
    ZOTPAGTOS=VAL(TEMPCC->TOTPAGTOS)
    If ZOTPAGTOS<0
      ZOTPAGTOS=ZOTPAGTOS*-1
    EndIf
    ZALDODEZ=VAL(TEMPCC->SALDODEZ)
    ZOTCOMPRAS=0
    ZOTVENDAS=0
    ZOTTRANSF=0
    ZAISISOFATCA=TEMPCC->PAISISOFAT
    If conexcorre=38031
      ZIPOCTA="OECD605"
      ZUBCTA="199"
    EndIf
    SELE 35
    If (ZOTCRED)#0 .or. (ZOTDEB)#0 .or. (ZOTCREDMT)#0 .or. (ZOTDEBMT)#0 .or. (ZOTPAGTOS)#0 .or. (ZALDODEZ)#0
      ADIREG(0)
      PUTREC()
    EndIf
    SELE 34
    SKIP
  ENDDO
  SELE 34
  USE
  SELE 35
  use
  SQLDISCONNECT(CONEXINTO)
  Erase &NARQ9
return 

Procedure TXFINAUD(SEENVIAEMAIL)
  HORAS=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
  If HORAS<1930 .And. SEENVIAEMAIL="S"
    return 
  EndIf
  FNAME="C:\EXATUS\FINALU.TXT"
  If !FILE(FNAME)
    return 
  EndIf
  CONEXON="mysqlcorretoras.exatus.net"
  CONEXUSER="txfinald"
  CONEXSENHA="B#=Se)udGxWrfk"
  CONEXPORT=3306
  CONEXDTBAS="db_turismo"
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
    return 
  EndIf
  SET PACKETSIZE TO 25
  SQLSETCONNECTION(CONEXINTO)
  SELE 35
  If !USEREDE("empresas",.F.,10,"empresas","M")
    SQLDISCONNECT(CONEXINTO)
    SELE 30
    return 
  EndIf
  GO TOP
  do while !eof()
    If FINAUD=CTOD("")
      SKIP
      LOOP
    EndIf
    WHASH=ALLTRIM(EMPRESAS->CODBACEN)+"_"
    ATUATX(empresas->codbacen,SEENVIAEMAIL)
    SQLSETCONNECTION(CONEXINTO)
    SELE 35
    SKIP
  ENDDO
  SELE 35
  USE
  SQLDISCONNECT(CONEXINTO)
  SELE 30
return 

Procedure ATUATX(qbacen,SEENVIAEMAIL)
  NTABEL=STRZERO(VAL(EMPRESAS->CODBACEN),5)+"_CPSALDO"
  NTABEL1=STRZERO(VAL(EMPRESAS->CODBACEN),5)+"_MOVIMENTO"
  CONEXDTBAS=ALLTRIM(EMPRESAS->BANCODADOS)
  CONEXUSER=ALLTRIM(EMPRESAS->USUARIODB)
  CONEXSENHA=ALLTRIM(EMPRESAS->SENHADB)
  CONEXON=ALLTRIM(EMPRESAS->SERVERDB)
  SQL CONNECT ON CONEXON;
    PORT CONEXPORT;
    DATABASE CONEXDTBAS;
    USER CONEXUSER;
    PASSWORD CONEXSENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEX1INTO1


  If SQL_ErrorNO() > 0
    SQLDISCONNECT(CONEX1INTO1)
    return 
  EndIf
  WNRC="SELECT MAX(DATA) FROM "+NTABEL
  mz=SQLArrayAssoc(WNRC)
  MZX=MZ
  MZX1=RIGHT(MZX[1,1],2)+"/"+SUBSTR(MZX[1,1],6,2)+"/"+LEFT(MZX[1,1],4)
  MZX2=""
  MZX3=MZX[1][1]
  MZ1=Strtran(MZ[1,1],"-","")
  MZ='"'+MZ[1,1]+'"'
  If DTOS(EMPRESAS->FINAUD)<>MZ1 .or. SEENVIAEMAIL="N"
    SELE 35
    If REGLOCK(10) .And. SEENVIAEMAIL="S"
      REPLACE FINAUD WITH CTOD(SUBSTR(MZX[1][1],9,2)+"/"+SUBSTR(MZX[1][1],6,2)+"/"+LEFT(MZX[1][1],4))
    EndIf

    // alterar aqui para gerar de outras datas.

    //MZX1="12/01/2021"
    //MZX2=""
    //MZX3="2021-01-12"
    //MZ="'2021-01-12'"

    // MSGINFO('TO AQUI')

    SQLSETCONNECTION(CONEX1INTO1)
    DO WHILE .T.
      SysRefresh()
      M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
      ARQ_PRN=LOCALNOT1+"TMP"+SUBS(STR(M->EX_T+1000,4),2)
      ARQ_PRN1:=ARQ_PRN+".DAT"
      ARQ_PRN2:=ARQ_PRN+".KEY"
      FNAME=LOCALNOT1+"RD_MOEDA.CSV"
      If !FILE(ARQ_PRN1)
        EXIT
      EndIf
    ENDDO
    MAT_STRU={}
    Aadd(mat_stru,{"tipo","C",1,0})
    Aadd(mat_stru,{"ESTRANGEIR","N",19,2})
    AADD(MAT_STRU,{"LIQUIDA","C",10,0})
    AADD(MAT_STRU,{"OPERACAO","N",10,0})
    Aadd(mat_stru,{"MOEDAISO","C",3,0})
    Aadd(mat_stru,{"POSICAO","C",1,0})
    Aadd(mat_stru,{"DATA","D",8,0})
    Aadd(mat_stru,{"COSIF","C",13,0})
    DBCREATE(ARQ_PRN1,mat_stru,"DBFNTX")
    SELE 36
    USE &ARQ_PRN1 EXCLUSIVE
    INDEX ON MOEDAISO+TIPO TO &ARQ_PRN2
    gravatx("SELECT tipo,moedas.simbolo,SUM(estrangeir) as estrangeir,liquidacao FROM "+NTABEL1+" LEFT JOIN MOEDAS ON "+NTABEL1+".MOEDA=MOEDAS.CODIGO WHERE "+NTABEL1+".sql_deleted='F' and liquidacao>"+MZ+" AND DATA="+MZ+' AND (tipo="1" OR tipo="3") GROUP BY '+NTABEL1+".moeda,liquidacao")
    gravatx("SELECT tipo,moedas.simbolo,SUM(estrangeir) as estrangeir,liquidacao FROM "+NTABEL1+" LEFT JOIN MOEDAS ON "+NTABEL1+".MOEDA=MOEDAS.CODIGO WHERE "+NTABEL1+".sql_deleted='F' and liquidacao>"+MZ+" AND DATA="+MZ+' AND (tipo="2" OR tipo="4") GROUP BY '+NTABEL1+".moeda,liquidacao")
    gravatx("SELECT tipo,moedas.simbolo,SUM(estrangeir) as estrangeir,liquidacao FROM "+NTABEL1+" LEFT JOIN MOEDAS ON "+NTABEL1+".MOEDA=MOEDAS.CODIGO WHERE "+NTABEL1+".sql_deleted='F' and liquidacao>"+MZ+" AND DATA="+MZ+' AND tipo="5" GROUP BY '+NTABEL1+".moeda,liquidacao")
    gravatx("SELECT tipo,moedas.simbolo,SUM(estrangeir) as estrangeir,liquidacao FROM "+NTABEL1+" LEFT JOIN MOEDAS ON "+NTABEL1+".MOEDA=MOEDAS.CODIGO WHERE "+NTABEL1+".sql_deleted='F' and liquidacao>"+MZ+" AND DATA="+MZ+' AND tipo="6" GROUP BY '+NTABEL1+".moeda,liquidacao")
    WNRC=Strtran(Strtran("SELECT tipo,simbolo,SUM(estrangeir) AS estrangeir, liquidacao FROM (SELECT 'S' AS tipo,moedas.simbolo,QUANTIDADE AS estrangeir,'L' AS liquidacao FROM 99999_cpsaldo LEFT JOIN MOEDAS ON 99999_cpsaldo.MOEDA=MOEDAS.CODIGO WHERE 99999_cpsaldo.sql_deleted='F' AND DATA='8888-88-88' AND 99999_cpsaldo.moeda<>790 UNION ALL SELECT 'S' AS tipo,moedas.simbolo,99999_transffilial.quantidade AS estrangeir,'L' AS liquidacao FROM 99999_transffilial LEFT JOIN moedas ON 99999_transffilial.moeda=moedas.codigo WHERE 99999_transffilial.moeda<>790 AND 99999_transffilial.data<='8888-88-88' AND (99999_transffilial.dataconfirmada='' OR 99999_transffilial.dataconfirmada='0000-00-00'  OR 99999_transffilial.dataconfirmada='0001-01-01' OR 99999_transffilial.dataconfirmada>'8888-88-88')) AS saldototal GROUP BY simbolo ORDER BY simbolo","99999",STRZERO(VAL(EMPRESAS->CODBACEN),5)),'8888-88-88',MZX3)
    gravatx(WNRC)
    SELE 36                                                                                                                                                                                                                                                                                          \
    GO TOP
    DO WHILE !EOF()
      If (TIPO="1" .OR. TIPO="3" .OR. TIPO="5" .OR. TIPO="S") .And. ESTRANGEIR>0
        REPLACE POSICAO WITH "C"
      Else
        REPLACE POSICAO WITH "V"
      EndIf
      If TIPO<>"S"
        WVAL=ESTRANGEIR
        WMOE=MOEDAISO
        WREG=RECNO()
        WPOS=POSICAO
        SEEK WMOE+"S"
        If !FOUND()
          ADIREG(0)
          REPLACE TIPO WITH "S"
          REPLACE MOEDAISO WITH WMOE
          REPLACE ESTRANGEIR WITH 0
          REPLACE LIQUIDA WITH "L"
          REPLACE POSICAO WITH "C"
        EndIf
        If WPOS="C"
          REPLACE ESTRANGEIR WITH ESTRANGEIR-WVAL
        Else
          REPLACE ESTRANGEIR WITH ESTRANGEIR+WVAL
        EndIf
        GO WREG
      EndIf
      SKIP
    ENDDO
    GO TOP
    DO WHILE !EOF()
      If TIPO="S"
        If ESTRANGEIR<0
          REPLACE ESTRANGEIR WITH ESTRANGEIR*-1
          REPLACE POSICAO WITH "V"
        Else
          REPLACE POSICAO WITH "C"
        EndIf
      EndIf
      If LIQUIDA="L" .And. POSICAO="C"
        REPLACE COSIf WITH "1.1.5.40.00-9"
      ElseIf LIQUIDA="L" .And. POSICAO<>"C"
        REPLACE COSIf WITH "4.9.2.05.20-6"
      ElseIf TIPO="1" .OR. TIPO="3"
        REPLACE COSIf WITH "1.8.2.06.30-8"
      ElseIf TIPO="2" .OR. TIPO="4"
        REPLACE COSIf WITH "4.9.2.05.20-6"
      ElseIf TIPO="5"
        REPLACE COSIf WITH "1.8.2.06.40-1"
      ElseIf TIPO="6"
        REPLACE COSIf WITH "4.9.2.05.30-9"
      EndIf
      SKIP
    ENDDO
    SELE 36
    GO TOP
    ARQ=FCREATE(FNAME)
    ZLINHA="Tipo;Operação;Descrição; Quantidade ;Posição;Data;ContaCosif;Data de vencimento;;;;;"+CHR(13)+CHR(10)
    FWRITE(ARQ,ZLINHA,LEN(ZLINHA))
    DO WHILE !EOF()
      ZLINHA="Moeda Estrangeira;"+STRZERO(RECNO(),4)+";"+MOEDAISO+";"+Strtran(ALLTRIM(STR(ESTRANGEIR,19,2)),".",",")+";"+IIF(POSICAO="C","COMPRADA","VENDIDA")+";"+MZX1+";"+TRIM(COSIF)+";"+IIF(LIQUIDA<>"L",RIGHT(LIQUIDA,2)+"/"+SUBSTR(LIQUIDA,6,2)+"/"+LEFT(LIQUIDA,4),"")+";;;;;;"+CHR(13)+CHR(10)
      FWRITE(ARQ,ZLINHA,LEN(ZLINHA))
      SKIP
    ENDDO
    FCLOSE(ARQ)
    SELE 36
    USE
    Erase &ARQ_PRN1
    Erase &ARQ_PRN2
    comando=WCORRENTE+"\BRAZIP.EXE AH "+LOCALNOT1+UPPER(WHASH)+Strtran(MZX2,"/","")+".zip "+FNAME+CHR(13)+CHR(10)
    MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+"DEL "+LOCALNOT1+"AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    MEMOWRIT(LOCALNOT1+"FZ.BAT",MZLINHA)
    MEMOWRIT(LOCALNOT1+"AGUARDE.TXT","rotina TAXAS Finald"+MZLINHA)
    CURSORWAIT()
    WAITRUN(LOCALNOT1+"FZ.BAT",0)
    CURSORWAIT()
    DO WHILE FILE(LOCALNOT1+"AGUARDE.TXT")
      SysRefresh()
    ENDDO
    LZCOPYFILE(LOCALNOT1+UPPER(WHASH)+Strtran(MZX2,"/","")+".ZIP","\SITES\EXATUS\GRAFICO\"+UPPER(WHASH)+Strtran(MZX2,"/")+".ZIP")
    Erase &FNAME
    FNAME=LOCALNOT1+"FZ.BAT"
    Erase &FNAME
    FNAME=LOCALNOT1+UPPER(WHASH)+Strtran(MZX2,"/","")+".ZIP"
    Erase &FNAME
    If SEENVIAEMAIL="S"
      If VAL(qbacen)=30673
        ENVIAEMAIL("angela@conexionvalores.com.br,wander@exatus.net","Arquivo Finald","Gerado arquivo com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"\SITES\EXATUS\GRAFICO\"+UPPER(WHASH)+Strtran(MZX2,"/")+".ZIP","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      ElseIf VAL(qbacen)=57434
        ENVIAEMAIL("financeiro@exim.com.br,renata@exim.com.br,wander@exatus.net","Arquivo Finald","Gerado arquivo com Sucesso em : "+DTOC(DATE())+" as  "+TIME(),"\SITES\EXATUS\GRAFICO\"+UPPER(WHASH)+Strtran(MZX2,"/")+".ZIP","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
      EndIf
    Else
      MSGINFO("Gerado "+"08/04/2020",wsist)
    EndIf
  EndIf
  SQLDISCONNECT(CONEX1INTO1)
  SELE 35
return 

Procedure GRAVATX(wnrc)
  TP12=SQLArrayAssoc(WNRC)
  FOR WCO=1 TO LEN(TP12)
    SELE 36
    ADIREG(0)
    REPLACE TIPO WITH tp12[WCO][1]
    REPLACE MOEDAISO WITH tp12[WCO][2]
    REPLACE ESTRANGEIR WITH val(tp12[WCO][3])
    REPLACE LIQUIDA WITH tp12[WCO][4]
  NEXT WCO
return 

Procedure LISTA(param)
  SEENVIAEMAIL=PARAM
  If !FILE("CONEX.LIS")
    return 
  EndIf
  LISTASPUXA() // alterar aqui para puxar novamente as atualizacoes
  /// alterar listaspuxa para puxar as outras
  DECLARE VLABEL[ADIR(LOCALnot3+"LISTA*.*")]
  ADIR(localnot3+"LISTA*.*",vlabel)
  FOR IW=1 TO LEN(VLABEL)
    NARQ=UPPER(LOCALNOT3+VLABEL[IW])
    NOMAARQUIVO:=DIRECTORY(NARQ)
    If NOMAARQUIVO[1][2]<25000
      Erase &NARQ
    EndIf
  NEXT IW
  LEGRAVALISTA()
  If param="S"
    ENVIAEMAIL("bene@exatus.net","Listas","Listas Atualizadas","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    ENVIAEMAIL("wanderlei@exatus.net","Listas","Listas Atualizadas","","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    DECLARE TX1[1]
    tx1[1]="Senhas Sictur user 1"
    LOGFILE(WCORRENTE+"\Exrotina.TXT",TX1)
    SENHASICTUR()
  Else
    msginfo("terminou",wsist)
  EndIf
return 

Procedure LEGRAVALISTA()
  DECLARE VLABEL[ADIR(LOCALnot3+"LISTA*.*")]
  ADIR(localnot3+"LISTA*.*",vlabel)
  RESTORE FROM CONEX.LIS ADDITIVE
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
    return 
  EndIf
  SET PACKETSIZE TO 25
  SELE 35
  If !USEREDE("listas",.F.,10,"listas","M")
    SELE 7
    return 
  EndIf
  FOR IW=1 TO LEN(VLABEL)
    SysRefresh()
    NARQ=UPPER(LOCALNOT3+VLABEL[IW])
    reg=val(substr(vlabel[iW],6,3))
    SELE 7
    go reg
    lelayout(atua)
  NEXT IW
  DECLARE TX1[1]
  TX1[1]="listas - Fim do serviço"
  LOGFILE(wcorrente+"\ROTBOLSA.TXT",TX1)
  cursorarrow()
  SELE 35
  USE
  SQLDISCONNECT(CONEXINTO)
return 

Procedure LELAYOUT(qtipo)
  // Arquivos originalmente zipados nao serao executados se estiverem descompactados
  NOMAARQUIVO:=DIRECTORY(NARQ)
  If NOMAARQUIVO[1][2]<25000
    return 
  EndIf
  If QTIPO<13
    SQL EXECUTE "delete from listas where nrlista="+strzero(qtipo,3)
    SQL EXECUTE "COMMIT"
  EndIf
  SELE 35
  DECLARE _VAR1[FCOUNT()]
  AFIELDS(_VAR1)
  GO BOTT
  SKIP
  FOR t = 1 TO FCOUNT()
    SysRefresh()
    _var4 = "Z"+SUBSTR(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    _var5 = _var1 [t]
    &_var4 = &_var5
  NEXT t
  RELEASE _VAR1
  ZRLISTA=QTIPO
  If AT("ZIP",UPPER(NARQ))#0 .And. ZRLISTA=6
    cmacro="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+LOCALnot3+'"'+CHR(13)+CHR(10)+"DEL AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    MEMOWRIT("AGUARDE.TXT","rotina Listas"+CMACRO)
    MEMOWRIT("FZNOT.BAT",CMACRO)
    CMACRO="FZNOT.BAT"
    waitrun(CMACRO,0)
    CURSORWAIT()
    DO while file("AGUARDE.TXT")
    ENDDO
    Erase &NARQ
    DECLARE VLABEL1[ADIR(LOCALnot3+"REL_AREAS*.*")]
    ADIR(localnot3+"REL_AREAS*.*",vlabel1)
    VLABEL1:=ASORT(VLABEL1)
    NARQ=Strtran(NARQ,".ZIP",".TXT")
    FRENAME(LOCALNOT3+VLABEL1[1],NARQ)
  EndIf
  If AT("ZIP",UPPER(NARQ))#0 .And. (ZRLISTA=9 .or. zrlista=10)
    cmacro="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+LOCALnot3+'"'+CHR(13)+CHR(10)+"DEL AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    MEMOWRIT("AGUARDE.TXT","rotina Listas"+CMACRO)
    MEMOWRIT("FZNOT.BAT",CMACRO)
    CMACRO="FZNOT.BAT"
    waitrun(CMACRO,0)
    CURSORWAIT()
    DO while file("AGUARDE.TXT")
    ENDDO
    Erase &NARQ
    DECLARE VLABEL1[ADIR(LOCALnot3+"*_AFASTAMEN*.*")]
    ADIR(localnot3+"*_afastamen*.*",vlabel1)
    FOR ZI=1 TO LEN(VLABEL1)
      WNARQ1=LOCALNOT3+VLABEL1[ZI]
      Erase &WNARQ1
    NEXT ZI
    DECLARE VLABEL1[ADIR(LOCALnot3+"*_HONOR*.*")]
    ADIR(localnot3+"*_HONOR*.*",vlabel1)
    FOR ZI=1 TO LEN(VLABEL1)
      WNARQ1=LOCALNOT3+VLABEL1[ZI]
      Erase &WNARQ1
    NEXT ZI
    DECLARE VLABEL1[ADIR(LOCALnot3+"*_OBSERV*.*")]
    ADIR(localnot3+"*_OBSERV*.*",vlabel1)
    FOR ZI=1 TO LEN(VLABEL1)
      WNARQ1=LOCALNOT3+VLABEL1[ZI]
      Erase &WNARQ1
    NEXT ZI
    DECLARE VLABEL1[ADIR(LOCALnot3+"*_REMUNERA*.*")]
    ADIR(localnot3+"*_REMUNERA*.*",vlabel1)
    FOR ZI=1 TO LEN(VLABEL1)
      WNARQ1=LOCALNOT3+VLABEL1[ZI]
      Erase &WNARQ1
    NEXT ZI
    DECLARE VLABEL1[ADIR(LOCALnot3+"*_CADASTRO*.*")]
    ADIR(localnot3+"*_CADASTRO*.*",vlabel1)
    NARQ=Strtran(NARQ,".ZIP",".TXT")
    FRENAME(LOCALNOT3+VLABEL1[1],NARQ)
  EndIf
  If AT("ZIP",UPPER(NARQ))#0 .And. ZRLISTA=12
    cmacro="@ECHO OFF"+CHR(13)+CHR(10)+WCORRENTE+'\brazip.exe EH "'+NARQ+'" "'+LOCALnot3+'"'+CHR(13)+CHR(10)+"DEL AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
    MEMOWRIT("AGUARDE.TXT","rotina Listas"+CMACRO)
    MEMOWRIT("FZNOT.BAT",CMACRO)
    CMACRO="FZNOT.BAT"
    waitrun(CMACRO,0)
    CURSORWAIT()
    DO while file("AGUARDE.TXT")
    ENDDO
    Erase &NARQ
  EndIf
  NOMAARQUIVO:=DIRECTORY(NARQ)
  If ZRLISTA<>12 .And. NOMAARQUIVO[1][2]<25000
    return 
  EndIf
  If ZRLISTA<>9 .And. ZRLISTA<>12 .And. zrlista<>10  .And. zrlista<>14
    Cxml1=MEMOREAD(narq)
  EndIf
  If QTIPO=1
    WDT=LEXML("Publish_Date",CXML1)
    ZtPublicacao=CTOD(SUBSTR(WDT,4,2)+"/"+SUBSTR(WDT,1,2)+"/"+RIGHT(WDT,4))
    Zomelista=TROCA->PARA
    zender=troca->ENVIAFTP
    Cxml1=substr(Cxml1,at("<sdnEntry",Cxml1),len(Cxml1))
    leitxml=left(Cxml1,at("</sdnEntry",Cxml1)+11)
    WCONTAD=0
    BEGIN TRANSACTION
    DO WHILE AT("</sdnEntry",Cxml1)#0
    leitxml=left(cXML1,at("</sdnEntry",Cxml1)+11)
    Cxml1=substr(Cxml1,at("</sdnEntry",Cxml1)+11,len(Cxml1))
    Zid=lexml("uid",leitxml)
    ZastName=upper(lexml("lastName",leitxml))
    Zype=lexml("sdnType",leitxml)
    Zotes=trim(leitxml)
    ZaisNasc=upper(lexml("country",leitxml))
    Zidade=upper(lexml("city",leitxml))
    Zirstname=upper(lexml("firstName",leitxml))
    Zndereco=upper(lexml("address1",leitxml))
    ZiddleName=upper(lexml("middleName",leitxml))
    zholename=UPPER(Strtran(UPPER(ALLTRIM(ZIRSTNAME)+" "+ALLTRIM(ZiddleName)+" "+ALLTRIM(ZastName)),"  "," "))
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    If EMPTY(Zidade)
      Zidade=UPPER(lexml("program",leitxml))
    EndIf
    ZostalCode=lexml("postalCode",leitxml)
    Zuncao=UPPER(lexml("title",leitxml))
    HOLENAME()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
  EndIf
    ENDDO
  END TRANSACTION


  ElseIf QTIPO=2
  WDT=SUBSTR(CXML1,at("<WHOLE Date=",CXML1)+13,10)
  ZtPublicacao=CTOD(WDT)
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  Cxml1=substr(Cxml1,at("<ENTITY Id=",Cxml1),len(Cxml1))
  leitxml=left(Cxml1,at("</ENTITY>",Cxml1)+9)
  WCONTAD=0
  BEGIN TRANSACTION
  DO WHILE AT("<ENTITY Id=",Cxml1)#0
    leitxml=left(cXML1,at("</ENTITY>",Cxml1)+9)
    Cxml1=substr(Cxml1,at("</ENTITY>",Cxml1)+9,len(Cxml1))
    Zid=strzero(val(substr(leitxml,at("<NAME Id=",leitxml)+10,10)),10)
    ZastName=Strtran(upper(lexml("LASTNAME",leitxml)),"-"," ")
    Zype=lexml("GENDER",leitxml)
    If Zype="M"
      Zype="Masculino"
    ElseIf Zype="F"
      Zype="Feminino"
    Else
      Zype="Juridica"
    EndIf
    Zotes=trim(leitxml)
    ZaisNasc=upper(lexml("PLACE",leitxml))
    Zidade=upper(lexml("CITY",leitxml))
    Zirstname=Strtran(upper(lexml("FIRSTNAME",leitxml)),"-"," ")
    ZiddleName=Strtran(upper(lexml("MIDDLENAME",leitxml)),"-"," ")
    zholename=UPPER(Strtran(upper(lexml("WHOLENAME",leitxml)),"-"," "))
    Zuncao=UPPER(lexml("FUNCTION",leitxml))
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    HOLENAME()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION
  ElseIf QTIPO=3
  WDT=SUBSTR(CXML1,at("dateGenerated=",CXML1)+15,10)
  ZtPublicacao=CTOD(SUBSTR(WDT,9,2)+"/"+SUBSTR(WDT,6,2)+"/"+LEFT(WDT,4))
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  Cxml1=substr(Cxml1,at("<INDIVIDUAL>",Cxml1),len(Cxml1))
  leitxml=left(Cxml1,at("</INDIVIDUAL>",Cxml1)+13)
  WCONTAD=0
  BEGIN TRANSACTION
  DO WHILE AT("</INDIVIDUAL>",Cxml1)#0
    leitxml=left(cXML1,at("</INDIVIDUAL>",Cxml1)+13)
    Cxml1=substr(Cxml1,at("</INDIVIDUAL>",Cxml1)+13,len(Cxml1))
    Zid=lexml("DATAID",leitxml)
    ZastName=upper(lexml("FOURTH_NAME",leitxml))
    If EMPTY(ZastName)
      ZastName=upper(lexml("THIRD_NAME",leitxml))
    EndIf
    Zirstname=upper(lexml("FIRST_NAME",leitxml))
    ZiddleName=upper(lexml("SECOND_NAME",leitxml))
    zholename=UPPER(Strtran(Strtran(upper(ALLTRIM(ZIRSTNAME)+" "+ALLTRIM(ZiddleName)+" "+ALLTRIM(lexml("THIRD_NAME",leitxml))+" "+ALLTRIM(lexml("FOURTH_NAME",leitxml))),"  "," "),"-"," "))
    Zotes=trim(leitxml)
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    HOLENAME()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION
  ElseIf QTIPO=4 .OR. QTIPO=5
  WCONTAD=0
  BEGIN TRANSACTION
  DO WHILE AT(chr(13),Cxml1)#0
    leitxml=left(cXML1,at(chr(13),Cxml1)+1)
    Cxml1=substr(Cxml1,at(chr(13),Cxml1)+1,len(Cxml1))
    Zomelista=TROCA->PARA
    zender=troca->ENVIAFTP
    Zid=strzero(val(cxml1),4)
    If VAL(cxml1)=0 .or. VAL(cxml1)<>INT(val(cxml1))
      LOOP
    EndIf
    ZastName=""
    Zirstname=""
    ZiddleName=""
    zholename=SUBSTR(CXML1,AT(CHR(9),CXML1)+1,100)
    zholename=UPPER(left(zholename,AT(CHR(9),zholename)-1))
    Zotes=trim(leitxml)
    HOLENAME()
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION

  ElseIf QTIPO=6
  WCONTAD=0
  BEGIN TRANSACTION
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  DO WHILE AT(chr(13),Cxml1)#0
    leitxml=left(cXML1,at(chr(10),Cxml1)+1)
    leitxml1=leitxml
    Cxml1=substr(Cxml1,at(chr(10),Cxml1)+1,len(Cxml1))
    If LEFT(LEITXML,8)<>"<tr><td>"
      LOOP
    EndIf
    LEITXML=SUBSTR(LEITXML,AT("</td><td>",leitxml)+8,LEN(LEITXML))
    LEITXML=SUBSTR(LEITXML,AT("</td><td>",leitxml)+8,LEN(LEITXML))
    LEITXML=SUBSTR(LEITXML,AT("</td><td>",leitxml)+8,LEN(LEITXML))
    LEITXML=SUBSTR(LEITXML,AT("</td><td>",leitxml)+8,LEN(LEITXML))
    ZHOLENAME=ALLTRIM(Strtran(Strtran(UPPER(SINAL(LEFT(LEITXML,AT("</td><td>",LEITXML)))),"<",""),">",""))
    ZPF=SUBSTR(LEITXML,AT("</td><td>",leitxml)+8,LEN(LEITXML))
    ZPF=LEFT(ZPF,AT("</td><td>",ZPF))
    ZPF=Strtran(ZPF," ","")
    ZPF=Strtran(ZPF,".","")
    ZPF=Strtran(ZPF,"-","")
    ZPF=Strtran(ZPF,"/","")
    ZPF=Strtran(ZPF,"<","")
    ZPF=alltrim(Strtran(ZPF,">",""))
    ZastName=""
    Zirstname=""
    ZiddleName=""
    Zotes=trim(leitxml1)
    HOLENAME()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      REPLACE WHOLENAME WITH ZHOLENAME
      REPLACE CPF WITH ZPF
      REPLACE NRLISTA WITH 6
      REPLACE NOMELISTA WITH ZOMELISTA
      REPLACE GENDER WITH ZENDER
      REPLACE INCEM WITH Strtran(DTOM(date()),"'","")+" "+time()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION
  ElseIf QTIPO=7


  ElseIf QTIPO=8
  WCONTAD=0
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  BEGIN TRANSACTION
  DO WHILE AT(chr(13),Cxml1)#0
    leitxml=left(cXML1,at(chr(13),Cxml1)+1)
    leitxml1=leitxml
    Cxml1=substr(Cxml1,at(chr(13),Cxml1)+1,len(Cxml1))
    ZPF=substr(LEITXML1,3,14)
    LEITXML=SUBSTR(LEITXML,19,LEN(LEITXML))
    COLU=AT(";",leitxml)
    zholename=SINAL(SUBSTR(LEITXML,1,COLU-1))
    ZastName=""
    Zirstname=""
    ZiddleName=""
    Zotes=trim(leitxml1)
    HOLENAME()
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    WCONTAD=WCONTAD+1
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      replace nrlista with 8
      replace WholeName with zholename
      replace cpf with zpf
      replace nomelista with "PEP"
      replace Notes with zotes
      REPLACE INCEM WITH Strtran(DTOM(date()),"'","")+" "+time()
    EndIf
    If WCONTAD>1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    SQL EXECUTE "COMMIT"
    EndIf
  ENDDO
  END TRANSACTION

  ElseIf QTIPO=9 .or. qtipo=10
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  DO WHILE .T.
    SysRefresh()
    M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
    NARQ3=LOCALNOT1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".DAT"
    If !FILE(NARQ3)
      EXIT
    EndIf
  ENDDO
  MAT_STRU={}
  AADD(MAT_STRU,{"LINHA","C",900,0})
  DBCREATE(NARQ3,MAT_STRU)
  SELE 36
  If !USEREDE(NARQ3,.T.,10,"TMP")
    SELE 35
    return 
  EndIf
  APPEND FROM &NARQ SDF
  GO TOP
  WCONTAD=0
  BEGIN TRANSACTION
  DO WHILE !EOF()
    leitxml=ALLTRIM(LINHA)
    leitxml1=LINHA
    Zirstname=""
    ZiddleName=""
    zholename=SUBSTR(LEITXML,AT(CHR(9),LEITXML)+1,150)
    zholename=UPPER(left(zholename,AT(CHR(9),zholename)-1))
    LEITXML=SUBSTR(LEITXML,at(chr(9),leitxml)+1,800)
    LEITXML=SUBSTR(LEITXML,at(chr(9),leitxml)+1,800)
    ZPF=LEFT(alltrim(LEITXML),14)
    Zotes=trim(leitxml1)
    SELE 35
    HOLENAME()
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    SELE 36
    SKIP
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION
  SELE 36
  USE
  Erase &NARQ3

  ElseIf QTIPO=11
  WCONTAD=0
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  BEGIN TRANSACTION
  DO WHILE AT(chr(10),Cxml1)#0
    leitxml=sinal(left(cXML1,at(chr(10),Cxml1)+1))
    leitxml1=leitxml
    Cxml1=substr(Cxml1,at(chr(10),Cxml1)+1,len(Cxml1))
    COLU=AT(";",leitxml)
    LEITXML=SUBSTR(LEITXML,COLU+1,LEN(LEITXML))
    COLU=AT(";",leitxml)
    LEITXML=SUBSTR(LEITXML,COLU+1,LEN(LEITXML))
    COLU=AT(";",leitxml)
    LEITXML=SUBSTR(LEITXML,COLU+1,LEN(LEITXML))
    COLU=AT(";",leitxml)
    zholename=(SUBSTR(LEITXML,1,COLU-1))
    zholename=Strtran(zholename,'"',"")
    LEITXML=SUBSTR(LEITXML,COLU+1,LEN(LEITXML))
    ZastName=""
    Zirstname=""
    ZiddleName=""
    Zotes=trim(leitxml1)
    HOLENAME()
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      PUTREC()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  ENDDO
  END TRANSACTION

  ElseIf QTIPO=13
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  oText:=TTxtfile():NEW(NARQ)
  CONTAD=oText:RecCount()
  WCONTAD=0
  If VALTYPE(CONTAD)<>"N"
    CONTAD=1
  EndIf
  oText:Skip()
  BEGIN TRANSACTION
  FOR WZI=2 TO CONTAD
    SysRefresh()
    Zirstname=""
    ZiddleName=""
    WLI=ALLTRIM(oText:ReadLine())+"; ; ;"
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    ZPF=Strtran(LEFT(WLI,COLU-1),'"','')
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    zholename=Strtran(Strtran(LEFT(WLI,COLU-1)," ",""),'"','')
    oText:Skip()
    SELE 35
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    WNRC="SELECT wholename FROM listas WHERE cpf='"+ZPF+"'"
    mz=SQLArrayAssoc(WNRC)
    If LEN(mz)#0
      ZHOLENAME=""
    EndIf
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      replace nrlista with 13
      replace WholeName with zholename
      replace cpf with zpf
      replace nomelista with "AUXILIO EMERGENCIAL"
      REPLACE INCEM WITH Strtran(DTOM(date()),"'","")+" "+time()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1500
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  NEXT WZI
  END TRANSACTION
  SQL EXECUTE "COMMIT"
  oText:Close()
  Erase &NARQ

  ElseIf QTIPO=14
  Zomelista=TROCA->PARA
  zender=troca->ENVIAFTP
  oText:=TTxtfile():NEW(NARQ)
  CONTAD=oText:RecCount()
  WCONTAD=0
  If VALTYPE(CONTAD)<>"N"
    CONTAD=1
  EndIf
  BEGIN TRANSACTION
  oText:Skip()
  FOR WZI=2 TO CONTAD
    SysRefresh()
    WLI=ALLTRIM(oText:ReadLine())+"; ; ;"
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    ZPF=Strtran(LEFT(WLI,COLU-1),'"','')
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    WLI=SUBSTR(WLI,COLU+1,900)
    COLU=AT(";",WLI)
    zholename=Strtran(Strtran(LEFT(WLI,COLU-1)," ",""),'"','')
    oText:Skip()
    SELE 35
    zncem=Strtran(DTOM(date()),"'","")+" "+time()
    WNRC="SELECT wholename FROM listas WHERE cpf='"+ZPF+"'"
    mz=SQLArrayAssoc(WNRC)
    If LEN(mz)#0
      ZHOLENAME=""
    EndIf
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      replace nrlista with 14
      replace WholeName with zholename
      replace cpf with zpf
      replace nomelista with "BOLSA FAMILIA"
      REPLACE INCEM WITH Strtran(DTOM(date()),"'","")+" "+time()
    EndIf
    WCONTAD=WCONTAD+1
    If WCONTAD=1500
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    EndIf
  NEXT WZI
  END TRANSACTION
  SQL EXECUTE "COMMIT"
  oText:Close()
  Erase &NARQ
  EndIf
  If QTIPO=12
    WCONTAD=0
    Zomelista=TROCA->PARA
    zender=troca->ENVIAFTP
    DO WHILE .T.
      SysRefresh()
      M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
      NARQ3=LOCALNOT1+"TMP"+SUBSTR(STR(M->EX_T+1000,4),2)+".DAT"
      If !FILE(NARQ3)
        EXIT
      EndIf
    ENDDO
    MAT_STRU={}
    AADD(MAT_STRU,{"LINHA","C",900,0})
    DBCREATE(NARQ3,MAT_STRU)
    SELE 36
    If !USEREDE(NARQ3,.T.,10,"TMP")
      SELE 35
      return 
    EndIf
    DECLARE VLABEL[ADIR(LOCALNOT3+"CONSULTA*.*")]
    ADIR(LOCALNOT3+"CONSULTA*.*",vlabel)
    VLABEL:=ASORT(VLABEL)
    IC=0
    FOR IC=1 TO LEN(VLABEL)
      SysRefresh()
      NARQ=LOCALNOT3+VLABEL[IC]
      SELE 36
      ZAP
      CONTEM=MEMOREAD(NARQ)
      CONTEM=Strtran(CONTEM,CHR(10),CHR(13)+CHR(10))
      MEMOWRIT(NARQ,CONTEM)
      APPEND FROM &NARQ SDF
      GO TOP
      DO WHILE !EOF()
        If AT("NÃO ELEITO",LINHA)#0
          SKIP
        LOOP
      EndIf
      leitxml:=zotes:=TRIM(linha)
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      LEITXML=SUBSTR(LEITXML,at(";",leitxml)+1,LEN(LEITXML))
      COLU=AT(";",leitxml)
      zholename=(SUBSTR(LEITXML,1,COLU-1))
      zholename=Strtran(Strtran(zholename,'"',"")," ","")
      ZastName=""
      Zirstname=""
      ZiddleName=""
      zncem=Strtran(DTOM(date()),"'","")+" "+time()
      If !EMPTY(ZHOLENAME)
        SELE 35
        ADIREG(0)
        PUTREC()
        SELE 36
      EndIf
      WCONTAD=WCONTAD+1
      If WCONTAD=1000
      END TRANSACTION
      BEGIN TRANSACTION
      WCONTAD=0
    EndIf
      skip
      ENDDO
      Erase &NARQ
    NEXT IC
  END TRANSACTION
  SELE 36
  USE
  Erase &NARQ3
  SELE 35
  EndIf
  If QTIPO=15
    SELE 35
    WDT=LEXML("LISTED_ON",CXML1)
    ZtPublicacao=CTOD(SUBSTR(WDT,4,2)+"/"+SUBSTR(WDT,1,2)+"/"+RIGHT(WDT,4))
    Zomelista=TRIM(TROCA->PARA)
    zender=troca->ENVIAFTP
    WCONTAD=0
    BEGIN TRANSACTION
    DO WHILE AT("</INDIVIDUAL>",Cxml1)#0
    leitxml=left(cXML1,at("</INDIVIDUAL>",Cxml1)+13)
    Cxml1=substr(Cxml1,at("</INDIVIDUAL>",Cxml1)+13,len(Cxml1))
    zholename=lexml("FIRST_NAME",leitxml)+LEXML("SECOND_NAME",leitxml)+lexml("THIRD_NAME",LEITXML)
    ZastName=""
    Zirstname=""
    ZiddleName=""
    Zotes=""
    HOLENAME()
    WNRC="SELECT wholename FROM listas WHERE WholeName='"+zholename+"'"
    mz=SQLArrayAssoc(WNRC)
    If LEN(mz)#0
      ZHOLENAME=""
    EndIf
    WCONTAD=WCONTAD+1
    If !EMPTY(ZHOLENAME)
      ADIREG(0)
      replace nrlista with 15
      replace WholeName with zholename
      replace GENDER with "S"
      replace nomelista with Zomelista
      REPLACE INCEM WITH Strtran(DTOM(date()),"'","")+" "+time()
    EndIf
    If WCONTAD>1000
    END TRANSACTION
    BEGIN TRANSACTION
    WCONTAD=0
    SQL EXECUTE "COMMIT"
  EndIf
    ENDDO
  END TRANSACTION
  EndIf
return 

Procedure LISTASPUXA()
  M->LOCAL=LOCALNOT3
  DECLARE VLABEL[ADIR(M->LOCAL+"LISTA*.*")]
  ADIR(M->LOCAL+"LISTA*.*",vlabel)
  VLABEL:=ASORT(VLABEL)
  I=0
  FOR I=1 TO LEN(VLABEL)
    SysRefresh()
    NARQ=LOCALNOT3+VLABEL[I]
    Erase &NARQ
  NEXT I
  SELE 7
  USE
  If !USEREDE(LOCNOT3+"TROCA.DBF",.F.,10,"TROCA")
    return 
  EndIf
  SELE 7
  GO TOP
  CURSORWAIT()
  DO WHILE !EOF()
    If EMPTY(DE) .OR. LEFT(DE,1)="."
      SKIP
      LOOP
    EndIf
    DE1=DE
    ANO=YEAR(DATE())
    MES=MONTH(DATE())
    FOR ZI=1 TO 4
      M->EX_T=(VAL(SUBSTR(TIME(),4,2))*10)+VAL(SUBSTR(TIME(),7,2))
      ARQOBS=LOCALNOT3+"LISTA"+STRZERO(recno(),3)+"_"+SUBSTR(STR(M->EX_T+1000,4),2)+"."+ALLTRIM(CLIENTE)
      If ATUA=9
        If MES=1
          ANO=ANO-1
          MES=12
        Else
          MES=MES-1
        EndIf
        DE1=Strtran(DE,"MM",STRZERO(MES,2))
        DE1=Strtran(DE1,"AAAA",STRZERO(ANO,4))
      EndIf
      oPg := CreateObject("Microsoft.XMLHTTP")
      oPg:Open("GET",ALLTRIM(DE1),.f.)
      oPg:Send()
      MEMOWRIT(ARQOBS,oPg:responsebody)
      oPg:=nil
      CURSORWAIT()
      If ATUA<>9
        EXIT
      EndIf
      NOMAARQUIVO:=DIRECTORY(ARQOBS)
      If NOMAARQUIVO[1][2]<10000
        Erase &ARQOBS
      Else
        EXIT
      EndIf
    NEXT ZI
    Erase FZNOT.BAT
    SELE 7
    SKIP
  ENDDO
  CURSORARROW()
return 

Procedure TESTSOAP()
  wxml='<?xml version="1.0" encoding="utf-8"?> <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">   <soap:Body>     <PessoaFisicaNFe xmlns="SOAWebServices">       <Credenciais><Email>wanderlei@exatus.net</Email> <Senha>HJkubh4$</Senha>  </Credenciais>   <Documento>06666901894</Documento>       <DataNascimento>07/07/1967</DataNascimento>  </PessoaFisicaNFe>  </soap:Body> </soap:Envelope> '
  objHTML:= TOleAuto():New("MSXML2.ServerXMLHTTP.6.0")
  objHTML:open("POST","http://www.soawebservices.com.br/webservices/producao/cdc/cdc.asmx")
  objHTML:setRequestHeader("Host","www.soawebservices.com.br")
  objHTML:setRequestHeader("Content-Type","text/xml; charset=utf-8")
  //objHTML:setRequestHeader("SOAPAction","SOAWebServices/Pefin")
  objHTML:setRequestHeader("SOAPAction","SOAWebServices/PessoaFisicaNFe")
  objHtml:setRequestHeader("Content-Length:",strzero(len(wxml),4))
  objHTML:send(WXML)
  objretorno=objHTML:status
  WXML1=objHTML:Responsebody
  memowrit("c:\sic\msgbc\m.htm",wxml1)
  msginfo(wxml1)
  objHTML:=nil
return 

Procedure enviaaviso(wmens,wcpf1)
  wcpf=alltrim(wcpf1)
  wxml='{ "app_id": "6b19c118-15ac-4159-b6e4-2b646f8d44e1",  "headings": {"en": "AssinaNet"},  "contents": {"en": "'+wmens+'"},  "filters": [{"field": "tag", "key": "cpf", "relation": "=", "value": "'+wcpf+'"}]}'
  objHTML:= TOleAuto():New("MSXML2.ServerXMLHTTP.6.0")
  objHTML:open("POST","https://onesignal.com/api/v1/notifications")
  objHTML:setRequestHeader("Host","https://onesignal.com/api/v1/notifications")
  objHTML:setRequestHeader("Content-Type","application/json; charset=utf-8")
  objHTML:setRequestHeader("authorization","Basic YTU5NGEyNjEtNzUzZS00NjJlLTg4ZjQtNmRlNjc0MWM1OWNj")
  //objHtml:setRequestHeader("Content-Length:",strzero(len(wxml),4))
  TRY
    objHTML:send(WXML)
    objretorno=objHTML:status
    WXML1=objHTML:Responsebody
  CATCH oerro
  END TRY
  objHTML:=nil
return 

Function HOLENAME()
  zholename=sinal1(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(Strtran(UPPER(zholename)," ",""),"-",""),",",""),".",""),"&",""),'"',""),"'",""),";",""))
return 

Procedure SINAL1(WMTEXTO)
  If VALTYPE(WMTEXTO)<>"C"
    return  ALLTRIM(WMTEXTO)
  EndIf
  WA="#$Á&ÉÍÓÚáéíóúçÇãõÃÕàèìòùÀÈÌÒÙÂÊÎÔÛâêîôûäëïöüÄËÏÖÜºªÑ<>,.&;"
  WB=" DAEEIOUaeioucCaoAOaeiouAEIOUAEIOUaeiouaeiouAEIOU"+SPACE(300)
  FOR SII=1 TO LEN(WA)
    DO WHILE .T.
      LETRA=SUBSTR(WA,SII,1)
      TROCA=SUBSTR(WB,SII,1)
      ACHOU=AT(LETRA,WMTEXTO)
      If ACHOU>0
        WMTEXTO=LEFT(WMTEXTO,ACHOU-1)+TROCA+RIGHT(WMTEXTO,LEN(WMTEXTO)-ACHOU)
      Else
        EXIT
      EndIf
    ENDDO
  NEXT SII
  WMTEXTO=ALLTRIM(WMTEXTO)
  FOR SII=1 TO LEN(WMTEXTO)
    If ASC(SUBSTR(WMTEXTO,SII,1))>122 .OR. (ASC(SUBSTR(WMTEXTO,SII,1))>57 .And. ASC(SUBSTR(WMTEXTO,SII,1))<65) .OR. (ASC(SUBSTR(WMTEXTO,SII,1))>90 .And. ASC(SUBSTR(WMTEXTO,SII,1))<97)
      WMTEXTO=LEFT(WMTEXTO,SII-1)+SUBSTR(WMTEXTO,SII+1,LEN(WMTEXTO))
    EndIf
  NEXT SII
return  WMTEXTO

Procedure CALCULAGET(QDT)
  CONEX1ON="localhost"
  CONEX1DTBAS="turmalina"
  WDE:=QDT
  If !AUTOMATICO
    CONEX1USER="calcgetmoney"
    CONEX1SENHA="E(fFNsm#K;whiJ"
  Else
    CONEX1USER="calcgetmoney"
    CONEX1SENHA="E(fFNsm#K;whiJ"
  EndIf

  CONEX1PORT=3306
  SQL CONNECT ON CONEX1ON;
    PORT CONEX1PORT;
    DATABASE CONEX1DTBAS;
    USER CONEX1USER;
    PASSWORD CONEX1SENHA;
    OPTIONS SQL_NO_WARNING;
    LIB 'MySQL';
    INTO CONEXEX

  If SQL_ErrorNO() > 0
    return 
  EndIf

  SELE 41
  If !USEREDE("27143_spread",.F.,10,"SPREAD","M")
    CLOSE DATABASE
    return 
  EndIf

  SELE 40
  If !USEREDE("27143_movimento",.F.,10,"MOVTO","M")
    CLOSE DATABASE
    return 
  EndIf
  WDE1=STRZERO(YEAR(WDE),4)+"-"+STRZERO(MONTH(WDE),2)+"-"+STRZERO(DAY(WDE),2)
  SELE 40
  SQL FILTER TO "DATA='"+WDE1+"' AND NAT1<>99000 AND COMISSAO=0"
  GO TOP
  DO WHILE !EOF()
    SELE 41
    GO TOP
    LOCATE FOR OPERADOR=MOVTO->OPERADOR .And. MOEDA=MOVTO->MOEDA
    If EOF()
      SELE 40
      If REGLOCK(10)
        REPLACE COMISSAO WITH -.01
      EndIf
      SELE 40
      SKIP
      LOOP
    EndIf
    SELE 40
    If TIPO$"135"
      ZIPO="C"
    Else
      ZIPO="V"
    EndIf
    WVARI=""
    If NAT1=32009 .OR. NAT1=32016 .OR. NAT1=32023
      If FORMAENTRE=50
        WVARI="SPREAD->SMOEDA"+ZIPO
        WVARIOF="SPREAD->IOFM"
      Else
        WVARI="SPREAD->SCARTAO"+ZIPO
        WVARIOF="SPREAD->IOFC"
      EndIf
    Else
      WVARI="SPREAD->SREM"+ZIPO
      WVARIOF="SPREAD->IOFR"
    EndIf
    TXCOML=TXCUSTO
    If TXCOML=0
      TXCOML=TAXA
    EndIf
    QUANTID=ESTRANGEIR
    WSPREAD=&WVARI
    TXCOBE=TXCOML*(1+(WSPREAD/100))
    If ZIPO="C"
      TXCOBE=TXCOML*(1-(WSPREAD/100))
    EndIf
    TXCOBE=ArredUp(TXCOBE,3)
    WPIS=4.65
    WREAIS=VLROPER
    If SPREAD->PLATAF="N"
      If ZIPO="V"
        RESULT=(((Wreais/quantid)-txcobe)*quantid)-despesas
        If &WVARIOF="S"
          RESULT=(((Wreais/quantid)-txcobe)*quantid)-despesas-IOF
        EndIf
      Else
        RESULT=((txcobe-(Wreais/quantid))*quantid)-IOF-despesas
      EndIf
      PIS=RESULT*((100-WPIS)/100)
      If RESULT<0
        PIS=RESULT+(RESULT*(WPIS/100))
      EndIf
      PORREM=SPREAD->PORCREMV
      REMU=PIS*(PORREM/100)
      If REGLOCK(10)
        REPLACE COMISSAO WITH REMU
        REPLACE RESULTGER WITH WSPREAD*1000
      EndIf
    ElseIf SPREAD->PLATAF="S"
      TXCOBE=0
      RESULT=WREAIS-despesas-IOF
      PIS=RESULT*((100-WPIS)/100)
      If ZIPO="V"
        PORREM=SPREAD->PORCREMV
      Else
        PORREM=SPREAD->PORCREMC
      EndIf
      If SPREAD->IOFPLATAF="S"
        REMU=PIS*(PORREM/100)
      Else
        REMU=RESULT*(PORREM/100)
      EndIf
      If REGLOCK(10)
        REPLACE COMISSAO WITH REMU
      EndIf
    EndIf
    SKIP
  ENDDO
  SELE 41
  USE
  SELE 40
  USE
  SQLDISCONNECT(CONEXEX)
  If !AUTOMATICO
    MSGINFO("Terminou Get Money",wsist)
  EndIf
return 

Procedure ArredUp(wtempv,wcasas1)
  WTEMPV1=wtempv*val("1"+replicate("0",wcasas1))
  IIF(WTEMPV1<>INT(WTEMPV1),WTEMPV1:=INT(WTEMPV1+1),WTEMPV1:=WTEMPV1)
return  (WTEMPV1/val("1"+replicate("0",wcasas1)))

Procedure VERHUB()
  If dow(date())=7 .or. dow(date())=1
    return 
  EndIf
  CDT=DTOS(DATE())+".HUB"
  COMANDO=WCORRENTE+"\wget.exe -T60 -olog.txt -O"+CDT+' "http://ws.hubdodesenvolvedor.com.br/v2/saldo/?info&token=8216570RjmqHVALJo14834768" --no-check-certificate  -U --user-agent=MOZILLA'
  MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+wcorrente+"\AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
  MEMOWRIT(WCORRENTE+"\fzhub.BAT",mzlinha)
  MEMOWRIT(WCORRENTE+"\AGUARDE.TXT","rotina 19"+MZLINHA)
  WAITRUN(WCORRENTE+"\fzhub.bat",0)
  DO WHILE FILE(WCORRENTE+"\AGUARDE.TXT")
    SysRefresh()
  ENDDO
  NARQ=WCORRENTE+"\fzhub.bat"
  Erase &NARQ
  v_erro:=MEMOREAD("LOG.TXT")
  v_erro1:=memoread(cdt)
  SEENVIAEMAIL="S"
  Erase log.txt
  Erase &cdt
  If SEENVIAEMAIL="S"
    ENVIAEMAIL("bene@exatus.net;wanderlei@exatus.net","Saldo consultas CPF/CNPJ",v_erro1,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
  EndIf
return 

Procedure VERHUB1()
  If dow(date())=7 .or. dow(date())=1
    return 
  EndIf
  CDT=DTOS(DATE())+".HU1"
  COMANDO=WCORRENTE+"\wget.exe -T60 -olog.txt -O"+CDT+' "http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf=11157746896&data=20/10/1970&token=8216570RjmqHVALJo14834768&ignore_db" --no-check-certificate  -U --user-agent=MOZILLA'
  MZLINHA="@ECHO OFF"+CHR(13)+CHR(10)+COMANDO+CHR(13)+CHR(10)+"DEL "+wcorrente+"\AGUARDE.TXT"+CHR(13)+CHR(10)+"EXIT"+CHR(13)+CHR(10)
  MEMOWRIT(WCORRENTE+"\fzhub.BAT",mzlinha)
  MEMOWRIT(WCORRENTE+"\AGUARDE.TXT","rotina 19"+MZLINHA)
  WAITRUN(WCORRENTE+"\fzhub.bat",0)
  NARQ=WCORRENTE+"\fzhub.bat"
  Erase &NARQ
  NARQ=WCORRENTE+"\aguarde.txt"
  Erase &NARQ
  v_erro:=MEMOREAD("LOG.TXT")
  v_erro1:=memoread(cdt)
  SEENVIAEMAIL="S"
  Erase log.txt
  Erase &cdt
  If SEENVIAEMAIL="S"
    If AT("numero_de_cpf",v_erro1)=0
      ENVIAEMAIL("bene@exatus.net;wanderlei@exatus.net","Erro na Consulta do CPF teste",v_erro1,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    EndIf
  EndIf
return 

Procedure CAM0021()
  DECLARE VLABEL[ADIR("C:\T5\*.XML")]
  ADIR("C:\T5\*.XML",vlabel)
  ZI=0
  FOR ZI=1 TO LEN(VLABEL)
    NFILE1="C:\T5\"+VLABEL[ZI]
    v_Erro:=MEMOREAD(NFILE1)
    If ASC(LEFT(V_ERRO,1))=0
      v_erro:=Strtran(MEMOREAD(NFILE1),chr(0),"")
      Erase &NFILE1
      objStreamxml:=TOleAuto():New("ADODB.Stream")
      objStreamxml:Open()
      objStreamxml:charset="utf-8"
      objStreamxml:WriteText(v_erro)
      objStreamxml:SaveToFile(NFILE1,2)
      objStreamxml:Close()
      objStreamxml:=nil
    EndIf
  NEXT ZI
  MSGINFO("TERMINOU")

Procedure QB()
  ARQDE="c:\NOTAS\MSBANK\BMF\CONEX.DAD"
  RESTORE FROM &ARQDE ADDITIVE
  CONEXON=trim(Decrypt(CONEXON,"leopardo"))
  CONEXDTBAS=trim(decrypt(CONEXDTBAS,"leopardo"))
  CONEXUSER=trim(decrypt(CONEXUSER,"leopardo"))
  CONEXSENHA=trim(decrypt(CONEXSENHA,"leopardo"))
  CONEXINTO=90
  CONEXCORRE=CONEXCORRE
  SELE 33
  USE C:\temp\TMP304.DBF ALIAS TEMP EXCLUSIVE
  MANDASICTUR()
  MSGINFO("TERMINOU")
  CLOSE DATABASE
return 

Procedure conscpf(wbcpf,wbnasc)
  WBCPF=Strtran(WBCPF,".","")
  WBCPF=Strtran(WBCPF,"=","")
  WBCPF=Strtran(WBCPF,"'","")
  WBCPF=Strtran(WBCPF,'"',"")
  WBCPF=ALLTRIM(Strtran(WBCPF,"-",""))
  WBCPF=STRZERO(VAL(WBCPF),11)
  cBuf:=""
  If ISINTERNET() .And. VAL(WBCPF)#0
    oPg := CreateObject("Microsoft.XMLHTTP")
    xComando := "http://ws.hubdodesenvolvedor.com.br/v2/cpf/?cpf="+WBCPF+"&data="+DTOC(WBNASC)+"&token=8216570RjmqHVALJo14834768&ignore_db"
    oPg:Open("GET",xComando,.f.)
    TRY
      oPg:Send()
      cBuf := Strtran(oPg:responseBody,"\","")
      cBuf := Strtran(cBuf,chr(13),"")
      cBuf := Strtran(cBuf,chr(10),"")
      cBuf := Strtran(cBuf,chr(9),"")
      cBuf := Strtran(cBuf,chr(8),"")
      cBuf := Strtran(cBuf,"}z{|}","")
      cBuf := Strtran(cBuf,'"',"")
      cBuf := Strtran(cBuf,'{',"")
      cBuf := Strtran(cBuf,'}',"")
      cBuf := sinal(cbuf)
      DECLARE TX1[1]
      TX1[1]=cBuf
      LOGFILE(WEBLOCAL+"CONSULTASRECEITA.CON",TX1)
    CATCH oErr
      cBuf:="NAO CONSULTOU"
    END TRY
    oPg:=nil
  EndIf
return  cBuf



Procedure ACERTACONTROL()
  If !USEREDE("C:\T0\CONTROLE.DBF",.T.,10,"EMP")
    MsgStop("O arquivo de EMPRESAS não está  disponível", WSIST)
    CLOSE DATABASE
    return 
  EndIf
  PACK
  Sx_MemoPack( 32, {|| QQOut(".")}, LastRec() / 10 )
  MSGINFO("TERMINOU")
return 


Procedure ETIQU()
  NARQ8="C:\T4\ETIQ.DBF"
  If !FILE(narq8)
    aStructure:={}
    aAdd( aStructure, { "CODIGO"    ,"N" ,   9, 0 } )
    aAdd( aStructure, { "WCEP"      ,"C" ,   8, 0 } )
    aAdd( aStructure, { "NOME"      ,"C" ,  60, 0 } )
    aAdd( aStructure, { "DATAENV"   ,"D" ,   8, 0 } )
    aAdd( aStructure, { "HORAENV"   ,"C" ,  20, 0 } )
    aAdd( aStructure, { "OQUE"      ,"C" , 254, 0 } )
    aAdd( aStructure, { "ENDERECO"  , "C",  30, 0 } )
    aAdd( aStructure, { "NUMERO"    , "C",   5, 0 } )
    aAdd( aStructure, { "COMPL"     , "C",  10, 0 } )
    aAdd( aStructure, { "BAIRRO"    , "C",  18, 0 } )
    aAdd( aStructure, { "CIDADE"    , "C",  28, 0 } )
    aAdd( aStructure, { "UF"        , "C",   2, 0 } )
    aAdd( aStructure, { "CEP"       , "C",   9, 0 } )
    aAdd( aStructure, { "FOLHAS"    , "N",   5, 0 } )
    aAdd( aStructure, { "AUDIT"     , "L",   1, 0 } )
    aAdd( aStructure, { "ENDE1","C",100,0})
    aAdd( aStructure, { "ENDE2","C",100,0})
    aAdd( aStructure, { "FAZER"     , "L",   1, 0 } )
    aAdd( aStructure, { "EMAIL"     , "L",   1, 0 } )
    aAdd( aStructure, { "ASSESSOR"  , "N",   5, 0 } )
    aAdd( aStructure, { "Enviado"   , "L",   1, 0 } )
    aAdd( aStructure, { "frentev"   , "L",   1, 0 } )
    aAdd( aStructure, { "EmailC"     , "C",1000, 0 } )
    aAdd( aStructure, { "BAIXA"     , "L",   1, 0 } )
    aAdd( aStructure, { "IMPRESSO"  , "L",   1, 0 } )
    aAdd( aStructure, { "LOGMAIL"   , "C", 250, 0 } )
    dbCreate(narq8,aStructure,"DBFNTX")
    RELEASE aStructure
  EndIf
  SELE 37
  If !USEREDE(NARQ8,.T.,10,"ETIQ1")
    MSGINFO("deu erro na abertura")
    return 
  EndIf
  ZAP
  YARQ="C:\T4\ETIQ.CSV"
  oText:=TTxtfile():NEW(YARQ)
  CONTAD=oText:RecCount()
  If VALTYPE(CONTAD)<>"N"
    CONTAD=1
  EndIf
  zcodigo=0
  FOR WZI=1 TO CONTAD
    ZCODIGO=zcodigo+1
    SysRefresh()
    WLI=ALLTRIM(oText:ReadLine())+"; ; ; ; ;"
    COLU=AT(";",WLI)
    Zclube=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zcpfcnpj=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zNOME=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zFJ=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zCOTA=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zendereco=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zcidade=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zuf=LEFT(WLI,COLU-1)
    WLI=SUBSTR(WLI,COLU+1,300)
    COLU=AT(";",WLI)
    zcep=LEFT(WLI,COLU-1)
    zcep=Strtran(zcep,"-","")
    ZCEP=STRZERO(VAL(zcep),8)
    If !empty(znome)
      adireg(0)
      replace codigo with zcodigo
      replace wcep with zcep
      replace nome with znome
      replace ende1 with zendereco
      replace ende2 with left(zcep,5)+"-"+substr(zcep,6,3)+" "+zcidade+" - "+zuf
    EndIf
    oText:Skip()
  NEXT WZI
  oText:Close()
  oReg := TReg32():Create(HKEY_CURRENT_USER,"SOFTWARE\Dane Prairie Systems\Win2PDF")
  ZARQ1="C:\T4\ETIQAT10.PDF"
  Erase &ZARQ1
  oReg:Set("PDFFilename",ZARQ1)
  oReg:Close()
  PRINT oPrn NAME "Envelope" TO TEMPDF1
  DEFINE FONT oFont1 NAME "verdana" SIZE 0,11 OF oPrn
  DEFINE FONT oFont2 NAME "verdana" SIZE 0,12 BOLD OF oPrn
  DEFINE FONT oFont3 NAME "verdana" SIZE 0,12 OF oPrn
  DEFINE FONT oFont4 NAME "verdana" SIZE 0,30 BOLD ITALIC OF oPrn
  DEFINE FONT oFont5 NAME "verdana" SIZE 0,12 BOLD ITALIC OF oPrn
  DEFINE font oFont7 NAME "verdana" SIZE 0,09 ITALIC BOLD OF oPrn
  DEFINE FONT oFont6 NAME "verdana" SIZE 0,08 ITALIC OF oPrn
  DEFINE FONT oFont3 NAME "verdana" SIZE 0,12 ITALIC OF oPrn
  nRowStep = oPrn:nVertRes() / 66
  nColStep = oPrn:nHorzRes() / 80
  oPrn:SetLandscape()
  GO TOP
  DO WHILE !EOF()
    MCODIGO=STRZERO(CODIGO,9)
    MCODIGO1=STRZERO((CODIGO+3)*2,9)
    PAGE
    CALCCEP(ENDE2,10.1,8.0)
    oprn:cmsay(11.0,8.0,LEFT(NOME,50),oFont2,,nRGB(0,0,0),,0)
    oprn:cmsay(11.5,8.0,ENDE1,oFont1,,nRGB(0,0,0),,0)
    oprn:cmsay(12.0,8.0,ENDE2,ofont1,,nRGB(0,0,0),,0)
    @ nRowStep*30.5,nColStep*30 CODE128 strzero(codigo,9) OF oPrn SIZE 0.8 MODE "A"
    oprn:cmsay(15.5,08.0,"Para devolução: Rua Surubim, 373 7.Andar - São Paulo - SP CEP 04571-050",oFont6,,nRGB(0,0,0),,0)
    ENDPAGE
    skip
    //     If codigo>10
    //       exit
    //     EndIf
  ENDDO
  oPrn:SetPortrait()
  ENDPRINT
  DO WHILE .T.
    CURSORWAIT()
    If (nHandle:=FOpen(ZARQ1,FO_READ+FO_EXCLUSIVE))!=-1
      fClose(nHandle)
      EXIT
    Else
      fClose(nHandle)
    EndIf
  ENDDO
  MSGINFO("GERADO")
return 

Procedure SENHASICTUR()
  CONEXON="mysqlcorretoras.exatus.net"
  CONEXUSER="WorkerAccess"
  CONEXSENHA="W0rk3r@cc355"
  CONEXDTBAS="db_turismo"
  CONEXPORT=3306
  CONEXINTOO=50
  CURSORWAIT()
  ULT=""
  ALTERADOS=""
  DO WHILE .T.
    SQL CONNECT ON CONEXON;
      PORT CONEXPORT;
      DATABASE CONEXDTBAS;
      USER CONEXUSER;
      PASSWORD CONEXSENHA;
      OPTIONS SQL_NO_WARNING;
      LIB 'MySQL';
      INTO CONEXINTOO

    If SQL_ErrorNO() > 0
      return 
    EndIf

    SET PACKETSIZE TO 25
    SQLSETCONNECTION(CONEXINTOO)
    SELE 35
    If !USEREDE("empresas",.F.,10,"GCORRETOR","M")
      SQLDISCONNECT(CONEXINTOO)
      return 
    EndIf
    GO TOP
    If VAL(ULT)=0
      ULT=CODBACEN
    Else
      LOCATE FOR ULT=CODBACEN
      SKIP
    EndIf
    If EOF()
      SQLDISCONNECT(CONEXINTOO)
      EXIT
    EndIf
    ULT=CODBACEN
    CURSORWAIT()
    CONEX1ON=alltrim(serverdb)
    CONEX1DTBAS=alltrim(bancodados)
    CONEX1USER=alltrim(usuariodb)
    CONEX1SENHA=alltrim(senhadb)
    CONEX1PORT=3306
    CONEX1INTOO=60
    CONEX1CORRE=codbacen
    SQLDISCONNECT(CONEXINTOO)
    SQL DISCONNECT ALL
    ALTSENHASICTUR()
  ENDDO
  SQLDISCONNECT(CONEXINTOO)
  SQL DISCONNECT ALL
  CLOSE DATABASE
  If !EMPTY(ALTERADOS)
    //  ENVIAEMAIL("bene@exatus.net","Sictur - Novas Senhas",alterados,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    ENVIAEMAIL("wander@exatus.net","Sictur - Novas Senhas",alterados,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
    //  ENVIAEMAIL("paulovicente@exatus.net","Sictur - Novas Senhas",alterados,"","","exatus@exatus.net","ckutzfiu",cIPServer,"exatus@exatus.net",.F.)
  EndIf
  abrearqs()
  CURSORARROW()
return 

Procedure ALTSENHASICTUR()
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
    return 
  EndIf
  SET PACKETSIZE TO 25
  WSENHA1=gErasenha(2)
  wsenha=strzero(val(CONEX1CORRE),5)+"u$d"+wsenha1
  wfunc="UPDATE usuarios set tentativas=0,bloqueado=0,expira="+dtom(date()+60)+","+'senha=md5("'+wsenha+'") where codigo=1'
  ALTERADOS=ALTERADOS+"<br>// "+STRZERO(val(CONEX1CORRE),5)+" senha: "+wsenha1+"</br>"
  SQL EXECUTE wfunc
  SQL EXECUTE "COMMIT"
  SQLDISCONNECT(CONEX1INTOO)
  SQLSETCONNECTION(CONEXINTOO)
return 

Function GEraseNHA(s_TAM)
  S_senha:= ""
  Do While Len(s_Senha) < s_Tam-1
    RND=HB_RANDOM(Strtran(time(),":",""))
    s_LetraMaiuscula = Chr(Int((65 - 90 + 1) * Rnd + 90))
    s_letraMinuscula = Chr(Int((97 - 122 + 1) * Rnd + 122))
    s_CaracterEspecial = Chr(Int((33 - 47 + 1) * Rnd + 47))
    If s_CaracterEspecial='"'.or. s_CaracterEspecial="'"
      s_CaracterEspecial="."
    EndIf
    s_Numero = Chr(Int((48 - 57 + 1) * Rnd + 57))
    s_opcao = Int((2 * Rnd))
    s_opcaoletra = Int((3 * Rnd))
    iif(s_opcao = 1,(iif(s_opcaoletra = 1,s_senha:=s_senha + s_letraMinuscula,s_senha:=s_senha + s_Letramaiuscula)),s_senha:= s_senha + s_Numero)
  ENDDO
  s_senha=s_senha+s_CaracterEspecial
return  S_senha

Function EnviaCDO(host_,from_,from_name_,para_,to_name_,titulo_,corpo_,anexo_,usersmtp_,senhasmtp_,portasmtp_,anexo1_)
  myMail:= CreateObject("CDO.Message")
  myMail:Subject:=titulo_
  myMail:From:=FROM_
  myMail:HTMLBody:=corpo_
  myMail:HTMLBodyPart:ContentTransferEncoding:="8bit"
  //myMail:Organization="Exatus Net Ltda"
  If !empty(anexo_) .And. file(anexo_)
    myMail:AddAttachment(anexo_)
  EndIf
  If !empty(anexo1_) .And. file(anexo1_)
    myMail:AddAttachment(anexo1_)
  EndIf
  If AT("amazonaws",LOWER(cserv))=0
    SSL_=.F.
    cPort:=587
  Else
    SSL_=.T.
    cPort:=465
  EndIf
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/sendusing",2)
  //	  'Name or IP of remote SMTP server
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/smtpserver",host_)
  //	  'Server port
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport",cport)
  //	  'requires authentication
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate",1)
  //	  'username
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/sendusername",usersmtp_)
  //	  'password
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/sendpassword",senhasmtp_)
  //	'TimeOut
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout",30)
  //	'startTLS
  //       myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = true
  myMail:Configuration:Fields:Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl",SSL_)
  myMail:Configuration:Fields:Update()

  cPara1=ALLTRIM(LOWER(PARA_))+";"
  cPara=""
  DO WHILE .T.
    ACHE=AT(";",cPara1)
    ACHE1=LEN(cPara1)
    If ACHE=0
      EXIT
    EndIf
    wmmmail=alltrim(substr(cPara1,1,ACHE-1))
    myMail:To=wmmmail
    cPara1=substr(cPara1,ACHE+1,ACHE1)
    DECLARE TX1[1]
    TX1=[1]="enviado"
    TRY
      myMail:Send()
    CATCH oErr
      DECLARE TX1[1]
      TX1[1]=" ERRO: PARA: "+para_+"  DE: "+from_
      tx2[1]=" ERRO: PARA: "+para_+"  DE: "+from_
      TX2[2]=.F.
      LOGFILE("c:\exatus\nao_enviado.ret",TX1)
      tx1[1]=""
    END TRY
  ENDDO
  myMail:=nil
return 


Procedure limpaplanner()
  SELE 1
  use c:\t\planner\cliente.dbf EXCLUSIVE
  index on codigo to t
  SELE 2
  use c:\t\planner\mcliente.dbf
  go top
  do while !eof()
    zcod=codigo
    SELE 1
    seek zcod
    If found()
      replace eimpre with .f.
    EndIf
    SELE 2
    skip
  enddo
  MSGINFO("ALTERADO")
  USE
return 

Procedure CONSIND()
  MCPF=SPACE(11)
  MDTNASC=DATE()
  MCPF=OUTROS("CPF",MCPF,"99999999999")
  MDTNASC=OUTROS("NASCIMENTO",MDTNASC,"@D")
  MNOME=SPACE(100)
  MNOME=OUTROS("NOME",MNOME,"@!")
  MARQ2=""
  BUSCATEMP("08493",MNOME,MCPF,MDTNASC,"0000","C:\T\TREVISO\")
  MSGINFO("GERADO "+MARQ2)
return 


Procedure BUSCATEMP(wCDBACEN,MWNOME,WBCPF,WBNASC,qualcli,wpastapdf)
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
      marq2=WPASTAPDF+Strtran(ALLTRIM(wevidencias),"/","\")
      memowrit(MARQ2,opg1:responsebody)
      oPg1:=nil
      WXML1=""
    EndIf
  Else
    MEMOWRIT("C:\EXATUS\FLAG.TXT",wxmlB3)
  EndIf
  objHTML:=nil
return 

Procedure INTSISFAIR()
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
    return  .F.
  EndIf
  SET PACKETSIZE TO 25
  DECLARE VLABEL[ADIR(WEBVIEWER+"*.DBF")]
  ADIR(WEBVIEWER+"*.DBF",vlabel)
  I=0
  FOR I=1 TO LEN(VLABEL)
    WDE=alltrim(WEBVIEWER+VLABEL[I])
    WPARA=LOWER(VLABEL[I])
    WPARA=alltrim(LEFT(WPARA,LEN(WPARA)-4))
    SELE 35
    If !USEREDE(WDE,.F.,10,"DE")
      SELE 35
      USE
      Erase &WDE
      SELE 30
      return  .F.
    EndIf
    SELE 36
    If !USEREDE("sql_tabelas",.F.,10,"PARA","M")
      SELE 35
      USE
      Erase &WDE
      SELE 30
      return  .F.
    EndIf
    GO TOP
    LOCATE FOR alltrim(nome_tab)=wpara
    If EOF()
      SELE 35
      USE
      SELE 36
      USE
      LOOP
    EndIf
    WINDICE="Z"+SUBS(INDICE,2,LEN(ALLTRIM(INDICE))-1)
    WLOCALIZA=ALLTRIM(INDICE)
    SELE 36
    USE
    If !USEREDE(WPARA,.F.,10,"PARA","M")
      SELE 35
      USE
      Erase &WDE
      SELE 30
      return  .F.
    EndIf
    MSGRUN("Verificando arquivo: "+wpara,wsist,{ ||SISFAIR_ARQ()})
    SELE 35
    USE
    Erase &WDE
  NEXT I
  If LEN(VLABEL)>0
    MSGRUN("Ajustando Clausulas",wsist,{ ||intclaustexto()})
    // JUNTANDO DI movdoc2
    SELE 35
    If !USEREDE("mov_di",.F.,10,"DE","M")
      SELE 35
      USE
      SELE 30
    Else
      SELE 36
      USE
      If !USEREDE("movdoc2",.F.,10,"PARA","M")
        SELE 35
        USE
        SELE 36
        USE
        SELE 30
      Else
        SELE 35
        GO TOP
        WCONT=0
        BEGIN TRANSACTION
        DO WHILE !EOF()
        SELE 36
        WFILT="BOLETO="+strzero(DE->BOLETO,10)+" and DI='"+DE->DI+"' and CONTROLE="+STRZERO(DE->CONTROLE,10)
        SQL FILTER TO WFILT
        GO TOP
        If EOF() .OR. BOF()
          APPEND BLANK
          REPLACE CORRETORA WITH DE->CORRETORA
          REPLACE CLIENTE WITH DE->CLIENTE
          REPLACE TIPO WITH DE->TIPO
          REPLACE BOLETO WITH DE->BOLETO
          REPLACE DI WITH DE->DI
          REPLACE EMISSAO WITH DE->EMISSAO
          REPLACE FOB WITH DE->FOB
          REPLACE FRETE WITH DE->FRETE
          REPLACE SEGURO WITH DE->SEGURO
          REPLACE CONTROLE WITH DE->CONTROLE
          REPLACE VENCIMENTO WITH DE->VENCIMENTO
        EndIf
        SQL FILTER TO
        WCONT=WCONT+1
        If WCONT>1000
        END TRANSACTION
        COMMIT
        BEGIN TRANSACTION
        WCONT=0
      EndIf
        SELE 35
        SKIP
        ENDDO
      END TRANSACTION
      COMMIT
    EndIf
  EndIf
  EndIf
  SELE 36
  USE
  SELE 35
  USE
  SELE 30
  SQLDISCONNECT(CONEXINTO)
return  .T.

Function SISFAIR_ARQ()
  If wpara="claus" .or. wpara="moedas" .or. wpara="cadnat"
    WFILT="TRUNCATE TABLE "+wpara
    SQL EXECUTE wfilt
    SQL EXECUTE "COMMIT"
    APPEND FROM &WDE
    If wpara="claus"
      WFILT="UPDATE claus SET DESCRICAO=REPLACE(TRIM(DESCRICAO),CHAR(9),'')"
    ElseIf wpara="moedas"
      WFILT="UPDATE moedas SET NOME_SING=REPLACE(TRIM(NOME_SING),CHAR(9),''),NOME_PLUR=REPLACE(TRIM(NOME_PLUR),CHAR(9),''),CENT_SING=REPLACE(TRIM(CENT_SING),CHAR(9),''),NOME_PLUR=REPLACE(TRIM(NOME_PLUR),CHAR(9),'')"
    ElseIf wpara="cadnat"
      WFILT="UPDATE cadnat SET DESC1_OPER=REPLACE(TRIM(DESC1_OPER),CHAR(9),'')"
    EndIf
    SQL EXECUTE wfilt
    SQL EXECUTE "COMMIT"
    return 
  EndIf
  SELE 35
  zext=""
  zod_iso=""
  zncem=""
  WCONT=0
  BEGIN TRANSACTION
  GO TOP
  DO WHILE !EOF()
    DECLARE _var1 [FCOUNT()]
    AFIELDS(_var1)
    FOR t = 1 TO FCOUNT()
      SysRefresh()
      _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
      _var5 = _var1 [t]
      If valtype(&_var5)="C"
        &_var4=alltrim(&_var5)
      Else
        &_var4 = &_var5
      EndIf
    NEXT t
    SELE 36
    WFILT=&WLOCALIZA
    SQL FILTER TO WFILT
    GO TOP
    If EOF() .OR. BOF()
      APPEND BLANK
      PUTREC_S()
    EndIf
    SQL FILTER TO
    WCONT=WCONT+1
    If WCONT>1000
    END TRANSACTION
    COMMIT
    BEGIN TRANSACTION
    WCONT=0
    EndIf
    SELE 35
    SKIP
  ENDDO
  END TRANSACTION
  COMMIT
  Erase &WDE
return 

Procedure intclaustexto()
  SELE 35
  If !USEREDE("claus",.F.,10,"claus","M")
    SELE 35
    USE
    return  .F.
  EndIf
  GO TOP
  ANTES=CODBAN+CODCLAU
  CLAUSULA=""
  BEGIN TRANSACTION
  WCONT=0
  DO WHILE !EOF()
    ATUAL=CODBAN+CODCLAU
    If ATUAL<>ANTES
      WFILT="UPDATE claus set TEXT='"+CLAUSULA+"' WHERE CODBAN='"+LEFT(ANTES,5)+"' AND CODCLAU='"+RIGHT(ANTES,4)+"' AND CODLIN=1"
      SQL EXECUTE wfilt
      ANTES=ATUAL
      CLAUSULA=""
      WCONT=WCONT+1
      If WCONT>1000
      END TRANSACTION
      SQL EXECUTE "COMMIT"
      BEGIN TRANSACTION
      WCONT=0
    EndIf
    EndIf
    CLAUSULA=Strtran(CLAUSULA+" "+Strtran(ALLTRIM(DESCRICAO),CHR(9)," "),"  "," ")
    SKIP
  ENDDO
  If !EMPTY(CLAUSULA)
    ANTES=ATUAL
    WFILT="UPDATE claus set TEXT='"+CLAUSULA+"' WHERE CODBAN='"+LEFT(ANTES,5)+"' AND CODCLAU='"+RIGHT(ANTES,4)+"' AND CODLIN=1"
    SQL EXECUTE wfilt
  EndIf
  END TRANSACTION
  SQL EXECUTE "COMMIT"
return 


Function PUTREC_S()
  If !REGLOCK()
    MSGINFO("Registro sendo utilizado por outro Usuário","Tente Novamente")
    return  .T.
  EndIf
  DECLARE _var1 [FCOUNT()]
  AFIELDS(_var1)
  FOR t = 1 TO FCOUNT()
    _var4 = "Z"+SUBS(_var1 [t],2,LEN(ALLTRIM(_var1 [t]))-1)
    If UPPER(_var4)="ZNCEM"
      LOOP
    EndIf
    _var5 = _var1 [t]
    If VALTYPE(&_var5)<>VALTYPE(&_var4)
      If VALTYPE(&_VAR4)="D"
        &_var4=STRZERO(YEAR(&_var4),4)+"-"+STRZERO(MONTH(&_var4),2)+"-"+STRZERO(DAY(&_var4),2)
      EndIf
    EndIf
    REPLACE &_var5 WITH &_var4
  NEXT t
  UNLOCK
return (.T.)

