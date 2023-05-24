#include "FiveWin.ch"
STATIC oWnd, oTimer

FUNCTION Main()
SET TALK OFF
SET DATE FRENCH
SET BELL OFF
SET STAT OFF
SET SCORE OFF
SET WRAP ON
SET CURSOR OFF
SET DELETED ON
SET EPOCH TO 1999
SET DECIMAL TO 2
SET CENTURY ON
SET EXCLUSIVE OFF
SET MULTIPLE OFF
FAZER=.F.
CONTINUA=.F.
LOCALINTERVALO=300
DEFINE BRUSH oBrush STYLE NULL
DEFINE WINDOW oWnd FROM 0,40 to 40,120 TITLE "Ver Aplicativos" BRUSH oBrush MENU BuildMenu()
verifica()
DEFINE TIMER oTimer INTERVAL (LOCALINTERVALO*1000) ACTION (Verifica(),ROTINAS()) OF oWnd
ACTIVATE TIMER oTimer
ACTIVATE WINDOW oWnd ON INIT oWnd:minimize()
RETURN

FUNCTION buildmenu()
LOCAL oMenu
MENU oMenu
  MENUITEM "&Sair" MESSAGE "Finaliza o Sistema" ACTION (oWnd:end())
ENDMENU
RETURN( oMenu )

FUNCTION Verifica()
Local Cart,Work
if !file("SMTPSEN.BAT")
  NARQ="SMTPSEND.EXE -@%1"+chr(13)+chr(10)+"exit"+chr(13)+chr(10)
  MEMOWRIT("SMTPSEN.BAT",NARQ)
ENDIF
MCONTA="exatus@exatus.net"
Cart:=FindWindow(0,"BolsasNet")
Work:=FindWindow(0,"CambioNet")
//PostMessage(FindWindow(0,"NetCall Suporte"),WM_CLOSE,0,0) //Encerra uma aplicação
IF Work=0
   ShellExecute( 0, 'Open', 'WORKER.EXE','B','C:\EXATUS')
ENDIF
IF Cart=0
   ShellExecute( 0, 'Open', 'WORKER.EXE','W','C:\EXATUS')
ENDIF
SysRefresh()
RETURN

PROCEDURE ROTINAS()
if fazer
  return
endif
FAZER=.T.
TEMPO=VAL(LEFT(TIME(),2)+SUBSTR(TIME(),4,2))
IF TEMPO<0730 .OR. TEMPO>2000 .or. DOW(date())=7 .OR. DOW(DATE())=1
  FAZER=.F.
  RETURN
ENDIF
SET DELE ON
MARQ1="c:\sites\exatus\grafico\dia\HORARIW.csv" // BolsasNet
MARQ2="c:\sites\exatus\grafico\dia\horario.csv"
IF FILE(MARQ1)
  v_erro:=MEMOREAD(MARQ1)
  WWHORA=LEFT(V_ERRO,8)
  WWDATA=RIGHT(V_ERRO,10)
  IF CTOD(WWDATA)=DATE()
    IF !EMPTY(WWHORA)
      HOR=WWHORA
      HORA=TIME()
      DIF=VAL(SUBSTR(TIMEDIFF(HOR,HORA),4,2))
      DIF=TIMEDIFF(HOR,HORA)
      DIF1=TIMEASSECONDS(DIF)
      IF DIF1>3600
        *  AQUI É HORARIO DA CARTEIRA
        mconta="exatus@exatus.net"
        ENVIAEMAIL("wsc19@iclaro.com.br","BolsasNet Parado","BolsasNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("josebeneditofilho@iclaro.com.br","BolsasNet Parado","BolsasNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("wanderlei@exatus.net","BolsasNet Parado","BolsasNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("paulovicente@exatus.net","BolsasNet Parado","BolsasNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("bene@exatus.net","BolsasNet Parado","BolsasNet","","",TRIM(LOWER(MCONTA)))
        ERASE &MARQ1
      ENDIF
    ENDIF
  ENDIF
ENDIF
IF FILE(MARQ2)
  v_erro:=MEMOREAD(MARQ2)
  WWHORA=LEFT(V_ERRO,8)
  WWDATA=RIGHT(V_ERRO,10)
  IF CTOD(WWDATA)=DATE()
    IF !EMPTY(WWHORA)
      HOR=WWHORA
      HORA=TIME()
      DIF=VAL(SUBSTR(TIMEDIFF(HOR,HORA),4,2))
      DIF=TIMEDIFF(HOR,HORA)
      DIF1=TIMEASSECONDS(DIF)
      IF DIF1>3600
        *  AQUI É HORARIO DO WORKER
        MCONTA="exatus@exatus.net"
        ENVIAEMAIL("wsc19@iclaro.com.br","BolsasNet - CambioNet Parado","BolsasNet - CambioNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("josebeneditofilho@iclaro.com.br","BolsasNet - CambioNet Parado","BolsasNet - CambioNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("wanderlei@exatus.net","BolsasNet - CambioNet Parado","BolsasNet CambioNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("paulovicente@exatus.net","BolsasNet - CambioNet Parado","BolsasNet - CambioNet Parado","","",TRIM(LOWER(MCONTA)))
        ENVIAEMAIL("bene@exatus.net","BolsasNet - CambioNet Parado","Bolsasnet - CambioNet Parado","","",TRIM(LOWER(MCONTA)))
        ERASE &MARQ2
      ENDIF
    ENDIF
  ENDIF
ENDIF
FAZER=.F.
SysRefresh()
RETURN

FUNCTION TimeDiff( cStartTime, cEndTime )
   RETURN TimeAsString(IF(cEndTime < cStartTime, 86400 , 0) +;
          TimeAsSeconds(cEndTime) - TimeAsSeconds(cStartTime))

FUNCTION TimeAsString( nSeconds )
   RETURN StrZero(INT(Mod(nSeconds / 3600, 24)), 2, 0) + " " +;
          StrZero(INT(Mod(nSeconds / 60, 60)), 2, 0) + " " +;
          StrZero(INT(Mod(nSeconds, 60)), 2, 0)

FUNCTION TimeAsSeconds( cTime )
   RETURN VAL(cTime) * 3600 + VAL(SUBSTR(cTime, 4)) * 60 +;
          VAL(SUBSTR(cTime, 7))

PROCEDURE ENVIAEMAIL(MQUEMRECEBE,MTITULO,MTEXTO,MANEXO,MANEXO1,MCONTA)
LOCALSERVER="smtp.exatus.net"
CIPSERVER=LOCALSERVER
LOCALNOT1="\EXATUS\NOTTEMP\"
cPara=mquemrecebe
cAssun=mtitulo
cAnexo=manexo
cDe=mconta
cServ=cIPServer
cBody=mtexto+chr(13)+chr(10)
cAnexo1=manexo1
while .T.
  M->EX_T=(VAL(SUBS(TIME(),4,2))*10)+VAL(SUBS(TIME(),7,2))
  ARQ_EMAIL=LOCALNOT1+"EM"+SUBS(STR(M->EX_T+1000,4),2)
  ARQ_EMAIL1=ARQ_EMAIL+".ENV"
  ARQ_EMAIL2=ARQ_EMAIL+".RET"
  ARQ_EMAIL3=ARQ_EMAIL+".TXT"
  IF !FILE(ARQ_EMAIL1)
    EXIT
  ENDIF
ENDDO
CURSORWAIT()
cPara1=ALLTRIM(LOWER(mquemrecebe))+";"
cPara=""
DO WHILE .T.
  ACHE=AT(";",cPara1)
  ACHE1=LEN(cPara1)
  IF ACHE=0
    EXIT
  ENDIF
  cPara=cPara+"-t"+ALLTRIM(substr(cPara1,1,ACHE-1))+chr(13)+chr(10)
  cPara1=substr(cPara1,ACHE+1,ACHE1)
ENDDO
v_file="-fexatus@exatus.net"+CHR(13)+CHR(10)+ALLTRIM(LOWER(cPara))+CHR(13)+CHR(10)+ "-s"+ALLTRIM(cAssun)+CHR(13)+CHR(10)+"-hsmtp.exatus.net"+CHR(13)+CHR(10)+"-luexatus@exatus.net"+chr(13)+chr(10)+"-lpckutzfiu"+chr(13)+chr(10)
MEMOWRIT(ARQ_EMAIL1,v_file)
CMACRO="SMTPSEN.BAT "+ARQ_EMAIL1
CURSORWAIT()
WAITRUN(CMACRO,0)
RETURN .T.
