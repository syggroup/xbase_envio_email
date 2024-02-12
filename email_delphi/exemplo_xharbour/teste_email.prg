FUNCTION MAIN()
LOCAL cSEPARADOR:='', cFILES:='', aFILES:={}, X:=0

LOCAL nPORTSMTP:=587
LOCAL cServerIP:='smtp.gmail.com'   
LOCAL cUSER:='seuemail@gmail.com'
LOCAL cFROM:='seuemail@gmail.com'
LOCAL cPASS:='sua_senha'
LOCAL cMSG:='Descricao da mensagem', cASUNTO:='Descrição do assunto'
LOCAL cBCC:='email_destino_copia_oculta@dominio.com.br'
LOCAL cCC:='email_destino_copia@dominio.com.br'
LOCAL cQUEM:='destino@nomedominio.com.br'

LOCAL lSSL_EMAIL:=.T.  //.F.=NÃO USA SSL(ANTIGA PORTA 25) e .T.=USA SSL 
LOCAL lTLS_EMAIL:=.T.  //.F.=NÃO USA TLS e .T.=USA TSL( EX: OFICCE365 ) 
LOCAL lEMAIL_CONF:=.F. // .F.=NÃO ENVIA CONFIRMAÇÃO DE LEITURA e .T.=ENVIA CONFIRMAÇÃO DE LEITURA
LOCAL lHTML:=.T.       // .F.=MODO TEXTO e .T.=MODO 

setmode(25,80)

IF VALTYPE(lEMAIL_CONF) = 'L'
   lEMAIL_CONF:= IF(lEMAIL_CONF, "TRUE", "FALSE")
ENDIF

IF VALTYPE(lSSL_EMAIL) = 'L'
   lSSL_EMAIL:= IF(lSSL_EMAIL, "TRUE", "FALSE")
ENDIF

IF VALTYPE(lSSL_EMAIL) = 'L'
   lTLS_EMAIL:= IF(lTLS_EMAIL, "TRUE", "FALSE")
ENDIF

IF LEN(aFILES) > 0
   FOR X:=1 TO LEN(aFILES)
      cFILES:= cFILES + cSEPARADOR + aFILES[X]
      cSEPARADOR:= ';'
   NEXT
ELSE
   cFILES:=''
ENDIF

IF DLLCALL(CAMINHO_EXE()+'\sygmail.dll', 32, 'LZ_SEND_EMAIL_SMTP', STR(nPORTSMTP), cServerIP, cUSER, cPASS, lSSL_EMAIL, lTLS_EMAIL, cFROM, cMSG, cASUNTO, cFILES, cBCC, cQUEM, cCC, lEMAIL_CONF, IF(lHTML,'TRUE','FALSE')) != 1
   ALERT('Erro ao conectar ao email, favor revisar')
   RETURN .F.
ELSE
   ALERT('enviado com sucesso')
   RETURN .F.
ENDIF

RETURN(.T.)

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
