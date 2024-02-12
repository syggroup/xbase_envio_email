//#pragma /w2
//#pragma /es2


PREPARA_ENVIA_EMAIL({'C:\Users\usuario\Desktop\CTE_FG\52120808542246000105570010000006891999993104-ProcCTe.xml'},;
                    'TESTE DE ENVIO DE XML',;
                    'masterservicevrb@gmail.com',;
                    'em anexo xml',;
                    'smtp.gmail.com',;
                    'cte.nfe.informais@gmail.com',;
                    'cte.nfe.informais@gmail.com',;
                    '06892990000104',;
                    '465',;
                    'marcio@coopmontenegro.com.br',;
                    'marcio@coopmontenegro.com.br',.f.,.t.,.f.)

FUNCTION MAIN()
LOCAL cSMTP   :='smtp.gmail.com'+SPACE(46), nPORTA:=465
LOCAL cUSER   :='sygecom@gmail.com'+SPACE(43) //SPACE(60)
LOCAL cSENHA  :=SPACE(60)
LOCAL cDEST   :='leonardodemachado@hotmail.com'+SPACE(37)
LOCAL cASSUNTO:='Teste de envio de EMAIL com CDO'+SPACE(29)
LOCAL cMSG    :='Mensagem no corpo do email     '+SPACE(29)
LOCAL aFiles  :={}

SetUnhandledExceptionFilter( @GpfHandler() ) // testar simulando um GPF forçado no sistema

REQUEST HB_CODEPAGE_PTISO
REQUEST HB_LANG_PT

HB_SETCODEPAGE( 'PTISO' )
HB_LANGSELECT( 'PT' )

SETMODE(25,80)

CLS
@ 1,1 SAY "Verificando Internet, aguarde..."

IF Inetestaconectada()=.F.
   @ 1,1 SAY "Sem acesso a Internet, favor revisar"
   RETURN NIL
ENDIF
@ 1,1 SAY SPACE(32)
@ 1,1 SAY "Internet OK"

@ 2,1 SAY "CONFIGURACOES:"
@ 3,1 SAY "SMTP....:" GET cSMTP
@ 4,1 SAY "PORTA...:" GET nPORTA
@ 5,1 SAY "USUARIO.:" GET cUSER
@ 6,1 SAY "SENHA...:" GET cSENHA

@ 8 ,1 SAY "DADOS PARA ENVIO:"
@ 9 ,1 SAY "EMAIL...:" GET cDEST
@ 10,1 SAY "ASSUNTO.:" GET cASSUNTO
@ 11,1 SAY "MENSAGEM:" GET cMSG
READ

IF EMPTY(cSENHA)
   ALERT("Informe a senha")
   RETURN NIL
ENDIF

@ 13,1 SAY "Enviando email,aguarde..."

AADD(aFiles,CAMINHO_EXE()+'\email.rc')

IF CONFIG_MAIL(aFiles,cASSUNTO,cDEST,cMSG,cSMTP,cUSER,cUSER,cSENHA,nPORTA,'','',.F.,.T.,.F.)
   ALERT("E-mail enviado com Sucesso")
ELSE
   ALERT("Erro ao enviar email")
ENDIF
@ 14,1 SAY "Fim..."

HB_GCALL() // limpa a memoria
__Quit()

RETURN(.T.)


****************************************************************************************************************************
FUNCTION CONFIG_MAIL(aFiles,cSubject,cDEST,cMsg,cServerIp,cFrom,cUser,cPass,nPORTSMTP,aCC,aBCC,lEMAIL_CONF,lSSL_EMAIL,lHTML)
****************************************************************************************************************************
//link de suporte para duvidas sobre o uso do CDO
//http://msdn.microsoft.com/en-us/library/ms873053(v=EXCHG.65).aspx
local lRet := .F.
local oCfg, oError
local lAut := .T.

TRY
  oCfg := CREATEOBJECT( "CDO.Configuration" )
CATCH oError
    ALERT( "Não Foi possível Enviar o e-Mail!"  + HB_OsNewLine() + ;
             "Error: "     + Transform(oError:GenCode,   nil) + HB_OsNewLine() + ;
             "SubC: "      + Transform(oError:SubCode,   nil) + HB_OsNewLine() + ;
             "OSCode: "    + Transform(oError:OsCode,    nil) + HB_OsNewLine() + ;
             "SubSystem: " + Transform(oError:SubSystem, nil) + HB_OsNewLine() + ;
             "Mensaje: "   + oError:Description )

END

TRY
    WITH OBJECT oCfg:Fields
         :Item( "http://schemas.microsoft.com/cdo/configuration/smtpserver"             ):Value := ALLTRIM(cServerIp)
         :Item( "http://schemas.microsoft.com/cdo/configuration/smtpserverport"         ):Value := nPORTSMTP
         :Item( "http://schemas.microsoft.com/cdo/configuration/sendusing"              ):Value := 2
         :Item( "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"       ):Value := lAut
         :Item( "http://schemas.microsoft.com/cdo/configuration/smtpusessl"             ):Value := lSSL_EMAIL
         :Item( "http://schemas.microsoft.com/cdo/configuration/sendusername"           ):Value := alltrim(cUser)
         :Item( "http://schemas.microsoft.com/cdo/configuration/sendpassword"           ):Value := alltrim(cPass)
         :Item( "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"  ):Value := 30
	 		  :Update()
    END WITH
    lRet := .t.
CATCH oError
    ALERT("Erro ao tentar enviar o E-mail, favor revisar")
END
if lRet
   lRet := Envia_Mail(oCfg,cFrom,cDEST,aFiles,cSubject,cMsg,aCC,aBCC,lEMAIL_CONF,lHTML)
endif
oCfg:=nil
Return(lRet)

*************************************************************************************
FUNCTION ENVIA_MAIL(oCfg,cFrom,cDEST,aFiles,cSubject,cMsg,aCC,aBCC,lEMAIL_CONF,lHTML)
*************************************************************************************
local aTo := {},i,x
local lRet := .f.
local nEle, oError, oMsg

aTo      := { alltrim(cDEST) } //--> PARA
nEle := 1

for i:=1 to len(aTo)
    TRY
      oMsg := CREATEOBJECT ( "CDO.Message" )
        WITH OBJECT oMsg
             :Configuration = oCfg
             //:BodyPart:Charset := "iso-8859-2" // "iso-8859-1" "utf-8"
             :From = cFrom
             :To = aTo[i]
             :Cc = aCC
             :BCC = aBCC
             :Subject = cSubject
             if lHTML
                :HTMLBody  = cMsg
             else
                :TextBody = cMsg
             endif
             For x := 1 To Len( aFiles )
                 :AddAttachment(AllTrim(aFiles[x]))
             Next
             IF lEMAIL_CONF=.T.
                :Fields( "urn:schemas:mailheader:disposition-notification-to" ):Value := cFrom
                :Fields:update()
             ENDIF
             :Send()
        END WITH
        lRet := .t.
    CATCH oError
        ALERT( "Não Foi possível Enviar o e-Mail!"  + HB_OsNewLine() + ;
                 "Error: "     + Transform(oError:GenCode,   nil) + HB_OsNewLine() + ;
                 "SubC: "      + Transform(oError:SubCode,   nil) + HB_OsNewLine() + ;
                 "OSCode: "    + Transform(oError:OsCode,    nil) + HB_OsNewLine() + ;
                 "SubSystem: " + Transform(oError:SubSystem, nil) + HB_OsNewLine() + ;
                 "Mensaje: "   + oError:Description )
        lRet := .f.
END
next
oMsg:=Nil
Return lRet


********************************************************************************
***********VEREFICA O NOME DO EXECUTAVEL E O CAMINHO DO MESMO*******************
*NomeExecutavel()    // verefica o nome
*NomeExecutavel(.t.) // verefica o caminho
********************************************************************************
FUNCTION NOMEEXECUTAVEL(lPath)
LOCAL nPos, cRet
If Empty(lpath)
   nPos:= RAT("\", hb_argv(0))
   cRet:= substr(hb_argv(0), nPos+1)
else
   cRet:= hb_argv(0)
endif
Return cRet
********************
*Retorna o caminho do EXE
FUNCTION CAMINHO_EXE
Return(Substr(Nomeexecutavel(.t.),1,(len(Nomeexecutavel(.t.))- len(Nomeexecutavel()))-1 ))

********************************************************************************
***********FIM DA ROTINA DE VEREFICAÇÃO DE EXECUTAL*****************************
********************************************************************************
********************************************************************************
***************INICIO DO TESTE DE CONEXÃO DE INTERNET***************************
********************************************************************************
FUNCTION INETESTACONECTADA(cAddress)
LOCAL aHosts
InetInit()
IF cAddress = NIL
   cAddress := "www.registro.br"
ENDIF
aHosts := InetGetHosts( cAddress )
IF aHosts == NIL .or. len(aHosts)=0
   InetCleanup()
   RETURN .f.
endif
InetCleanup()
RETURN(.t.)

#include "hbexcept.ch"

*******************
FUNCTION GPFHANDLER( Exception )
*******************
local cMsg, nCode, oError
IF Exception <> NIL
   nCode := Exception:ExceptionRecord:ExceptionCode
   SWITCH nCode
      CASE EXCEPTION_ACCESS_VIOLATION
           cMsg := "EXCEPTION_ACCESS_VIOLATION - O thread tentou ler/escrever num endereço virtual ao qual não tinha acesso."
           EXIT

      CASE EXCEPTION_DATATYPE_MISALIGNMENT
           cMsg := "EXCEPTION_DATATYPE_MISALIGNMENT - O thread tentou ler/escrever dados desalinhados em hardware que não oferece alinhamento. Por exemplo, valores de 16 bits precisam ser alinhados em limites de 2 bytes; valores de 32 bits em limites de 4 bytes, etc. "
           EXIT

      CASE EXCEPTION_ARRAY_BOUNDS_EXCEEDED

           cMsg := "EXCEPTION_ARRAY_BOUNDS_EXCEEDED - O thread tentou acessar um elemento de array fora dos limites e o hardware possibilita a checagem de limites."
           EXIT

      CASE EXCEPTION_FLT_DENORMAL_OPERAND
           cMsg := "EXCEPTION_FLT_DENORMAL_OPERAND - Um dos operandos numa operação de ponto flutuante está desnormatizado. Um valor desnormatizado é um que seja pequeno demais para poder ser representado no formato de ponto flutuante padrão."
           EXIT

      CASE EXCEPTION_FLT_DIVIDE_BY_ZERO
           cMsg := "EXCEPTION_FLT_DIVIDE_BY_ZERO - O thread tentou dividir um valor em ponto flutuante por um divisor em ponto flutuante igual a zero."
           EXIT

      CASE EXCEPTION_FLT_INEXACT_RESULT
           cMsg := "EXCEPTION_FLT_INEXACT_RESULT - O resultado de uma operação de ponto flutuante não pode ser representado como uma fração decimal exata."
           EXIT

      CASE EXCEPTION_FLT_INVALID_OPERATION
           cMsg := "EXCEPTION_FLT_INVALID_OPERATION - Qualquer operação de ponto flutuante não incluída na lista."
           EXIT

      CASE EXCEPTION_FLT_OVERFLOW
           cMsg := "EXCEPTION_FLT_OVERFLOW - O expoente de uma operação de ponto flutuante é maior que a magnitude permitida pelo tipo correspondente."
           EXIT

      CASE EXCEPTION_FLT_STACK_CHECK
           cMsg := 'EXCEPTION_FLT_STACK_CHECK - A pilha ficou desalinhada ("estourou" ou "ficou abaixo") como resultado de uma operação de ponto flutuante.'
           EXIT

      CASE EXCEPTION_FLT_UNDERFLOW
           cMsg := "EXCEPTION_FLT_UNDERFLOW - O expoente de uma operação de ponto flutuante é menor que a magnitude permitida pelo tipo correspondente."
           EXIT

      CASE EXCEPTION_INT_DIVIDE_BY_ZERO
           cMsg := "EXCEPTION_INT_DIVIDE_BY_ZERO - O thread tentou dividir um valor inteiro por um divisor inteiro igual a zero."
           EXIT

      CASE EXCEPTION_INT_OVERFLOW
           cMsg := "EXCEPTION_INT_OVERFLOW - O resultado de uma operação com inteiros causou uma transposição (carry) além do bit mais significativo do resultado."
           EXIT

      CASE EXCEPTION_PRIV_INSTRUCTION
           cMsg := "EXCEPTION_PRIV_INSTRUCTION - O thread tentou executar uma instrução cuja operação não é permitida no modo de máquina atual."
           EXIT

      CASE EXCEPTION_IN_PAGE_ERROR
           cMsg := "EXCEPTION_IN_PAGE_ERROR - O thread tentou acessar uma página que não estava presente e o sistema não foi capaz de carregar a página. Esta exceção pode ocorrer, por exemplo, se uma conexão de rede é perdida durante a execução do programa via rede."
           EXIT

      CASE EXCEPTION_ILLEGAL_INSTRUCTION
           cMsg := "EXCEPTION_ILLEGAL_INSTRUCTION - O thread tentou executar uma instrução inválida."
           EXIT

      CASE EXCEPTION_NONCONTINUABLE_EXCEPTION
           cMsg := "EXCEPTION_NONCONTINUABLE_EXCEPTION - O thread tentou continuar a execução após a ocorrência de uma exceção irrecuperável."
           EXIT

      CASE EXCEPTION_STACK_OVERFLOW
           cMsg := "EXCEPTION_STACK_OVERFLOW - O thread esgotou sua pilha (estouro de pilha)."
           EXIT

      CASE EXCEPTION_INVALID_DISPOSITION
           cMsg := "EXCEPTION_INVALID_DISPOSITION - Um manipulador (handle) de exceções retornou uma disposição inválida para o tratador de exceções. Uma exceção deste tipo nunca deveria ser encontrada em linguagens de médio/alto nível."
           EXIT

      CASE EXCEPTION_GUARD_PAGE
           cMsg := "CASE EXCEPTION_GUARD_PAGE"
           EXIT

      CASE EXCEPTION_INVALID_HANDLE
           cMsg := "EXCEPTION_INVALID_HANDLE"
           EXIT

      CASE EXCEPTION_SINGLE_STEP
           cMsg := "EXCEPTION_SINGLE_STEP Um interceptador de passos ou outro mecanismo de instrução isolada sinalizou que uma instrução foi executada."
           EXIT

      CASE EXCEPTION_BREAKPOINT
           cMsg := "EXCEPTION_BREAKPOINT - Foi encontrado um ponto de parada (breakpoint)."
           EXIT

      DEFAULT
         cMsg := "UNKNOWN EXCEPTION (" + cStr( Exception:ExceptionRecord:ExceptionCode ) + ")"
   END
ENDIF

oError := ErrorNew( "GPFHANDLER", 0, 0, ProcName(), "Erro de GPF", { cMsg, Exception, nCode }, Procfile(), Procname(), procline() )

alert( DefError( oError ), "Error" )

RETURN EXCEPTION_EXECUTE_HANDLER

