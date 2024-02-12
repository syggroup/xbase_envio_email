/*
 *   $Id: envia_email.prg 2147 2012-04-07 02:15:31Z leonardo $
 */

#include "hbblat.ch"
#include "common.ch"
#include "hwgui.ch"

#pragma /w2
#pragma /es2

#IfnDef __XHARBOUR__
   #include "hbcompat.ch"
   REQUEST HB_GT_GUI_DEFAULT
#endif

STATIC Thisform
STATIC nTempo := 0

FUNCTION Main( cFILE )
LOCAL nRet, cANEXOS:=''
LOCAL oBlat := HBBlat():New()

Local aFiles  := {}
Local cServer :=''
Local nPort   :=0
Local ccFrom  :=''
Local aTo     :=''
Local aCC     :=''
Local aBCC    :=''
Local ccBody  :=''
Local cSubject:=''
Local cUser   :=''
Local cPass   :=''
Local lCONF   :=.f.
Local lMOSTRA :=.f.
Local lHTML   :=.f.

IF !FILE(cFILE)
   Return(.f.)
ENDIF
TRY
   USE (cFILE) ALIAS EMAIL READONLY
catch
   Return(.f.)
END
SELE EMAIL
DBGOTOP()
cServer =ALLTRIM(EMAIL->SERVER)
nPort   =EMAIL->PORTA
ccFrom  =ALLTRIM(EMAIL->FROM)
aTo     =ALLTRIM(EMAIL->ATO)
aCC     =ALLTRIM(EMAIL->ACC)
aBCC    =ALLTRIM(EMAIL->ABCC)
ccBody  =ALLTRIM(EMAIL->MSG)
cSubject=ALLTRIM(EMAIL->SUBJECT)
cUser   =ALLTRIM(EMAIL->USER)
cPass   =ALLTRIM(EMAIL->PASS)
lCONF   =EMAIL->CONF
lMOSTRA =EMAIL->MOSTRA
lHTML   =EMAIL->HTML

DBGOTOP()
DO WHILE .NOT. EOF()
   cANEXOS=EMAIL->FILES

   IF !EMPTY(cANEXOS)
      AADD(aFiles, alltrim(cANEXOS))
   ENDIF
   
   SELE EMAIL
   DBSKIP()
ENDDO
DBCLOSEALL()
FERASE(cFILE)

IF cServer=Nil .OR. nPort=Nil .OR. ccFrom=Nil .OR. aTo=Nil .OR. aCC=Nil
   Return(.f.)
ENDIF

WITH OBJECT oBlat
   :cFrom                   := ccFrom
   :cTo                     := aTo
   :cUserAUTH               := cUser
   :cPasswordAUTH           := cPass
   //:cUserPOP3               := cUser
   :cHostname               := cServer
   if !EMPTY(aCC)
     :cCC                   := aCC   // com copia
   endif
   if !EMPTY(aBCC)
      :cBCC                  := aBCC  // com copia oculta
   endif
   //:cReplyTo // responder para ?
   //:cServerPOP3             := cServer  // servidor SMTP
   :cServerSMTP             := cServer  // servidor SMTP
   :nPortSMTP               := nPort    // porta do servidor SMTP
   :cSubject                := cSubject // asunto da mensagem
   :cBody                   := ccBody   // corpo da mensagem
   :lHtml                   := lHTML
   IF lCONF=.T.
      :lRequestReturnReceipt:= TRUE
      :lRequestDisposition  := TRUE   // confirmação de leitura
   ENDIF
   //:lRequestReturnReceipt   := TRUE   //
   IF len(aFiles) > 0
      :aAttachBinFiles      := aFiles   // arquivos anexos
   ENDIF
   IF lMOSTRA=.T.
      :cLogFile             := "log_email.txt"  // habilita um LOG
   ENDIF
   //:lDebug                := TRUE     // habilita o debug
   //:lLogOverwrite         := TRUE     // quando o LOG é usado deve sobre-escrever no arquivo LOG
END
nRet := oBlat:Send()
IF lMOSTRA=.T.
   if nRet # 0
      MY_MSGINFO("Erro ao Tentar Enviar o Email: " +oBlat:BlatErrorString(),10)
   ELSE
      ferase("log_email.txt")
      MY_MSGINFO("Email Enviado com Sucesso",5)
   endif
ENDIF
Return(.t.)

*******************************
FUNCTION MY_MSGINFO(cTIT,cTempo)
*******************************
Local oSAY, oButton1, oDlg
Local oTime_MSG

IF nTempo <> 0
   nTempo :=0
ENDIF

IF cTempo=Nil
   cTempo=30
ENDIF

INIT DIALOG oDlg TITLE "Aviso do Sistema" NOEXIT NOEXITESC ;
AT 0,0 SIZE 400,105 ;
ICON HIcon():AddResource(1001) ;
STYLE DS_CENTER +WS_VISIBLE
Thisform := oDlg

@ 10,15 SAY oSAY CAPTION cTIT SIZE 400,20;
STYLE SS_CENTER;
FONT HFont():Add( '',0,-11,400,,,)

@ 150,40 BUTTONEX oButton1 Caption "&OK" ON CLICK {|| (Thisform:oTime_MSG:end()), Thisform:Close() };
SIZE 60,22 STYLE WS_TABSTOP + DS_CENTER ;
FONT HFont():Add( '',0,-12,400,,,);
TOOLTIP "Clique Aqui Para Proseguir"

SET TIMER oTIME_MSG OF oDlg ID 9007 VALUE 1000 ACTION {|| ATUALIZA_TIMER(cTempo) }

ACTIVATE DIALOG oDlg

Return(.t.)

**************************************
STATIC FUNCTION ATUALIZA_TIMER(cTempo)
**************************************
nTempo++
IF nTempo=cTempo
   Thisform:oTime_MSG:end()
   Thisform:Close()
ENDIF
//HWG_DOEVENTS()
Thisform:oButton1:SETTEXT("&OK-"+alltrim(str(cTempo-nTempo)) )
Thisform:oButton1:REFRESH()
RETURN(.T.)

