Function Main()


Monta_DBF_Email("c:\temp.dbf", "smtp.servidor.com.br", 25, "meu_email@servidor.com.br", "destino@servidor.com.br", "copia@servidor.com.br", "Assunto do email", "meuemailuser@servidor.com.br", "senha", {"c:\arq1.txt","arq2.txt"}, "copiaOculta@servidor.com.br", .F.,.T.)

Run "envia_email c:\temp.dbf"

Return

******************************************************************************************************************************************
Function Monta_DBF_Email(cFILE, cServerIP, vPORTSMTP, cFrom, aQuem, aCC, cMsg, cSUBJECT, cUSER, cPASS, aFiles, aBCC, vEMAIL_CONF,vMOSTRA )
******************************************************************************************************************************************
LOCAL vARQ:={}
aadd( vARQ, {"SERVER",     "C",190 ,0} )
aadd( vARQ, {"PORTA",      "N", 10 ,0} )
aadd( vARQ, {"FROM",       "C",190 ,0} )
aadd( vARQ, {"ATO",        "C",190 ,0} )
aadd( vARQ, {"ACC",        "C",190 ,0} )
aadd( vARQ, {"ABCC",       "C",190 ,0} )
aadd( vARQ, {"MSG",        "M", 10 ,0} )
aadd( vARQ, {"SUBJECT",    "M", 10 ,0} )
aadd( vARQ, {"FILES",      "C",190 ,0} )
aadd( vARQ, {"USER",       "C",190 ,0} )
aadd( vARQ, {"PASS",       "C", 90 ,0} )
aadd( vARQ, {"USUARIO",    "C", 30 ,0} )
aadd( vARQ, {"CONF",       "L",  1 ,0} )
aadd( vARQ, {"MOSTRA",     "L",  1 ,0} )
DBcreate(cFILE, vARQ )

USE (cFILE) alias "EMAIL" shared new

SELE EMAIL
APPEND BLANK
Replace SERVER WITH cServerIP,;
PORTA          WITH vPORTSMTP,;
FROM           WITH cFrom,;
ATO            WITH aQuem,;
ACC            WITH aCC,;
ABCC           WITH aBCC,;
MSG            WITH ALLTRIM(cMsg),;
SUBJECT        WITH cSUBJECT,;
USER           WITH cUSER,;
PASS           WITH cPass,;
USUARIO        WITH "MEUNOME",;
CONF           WITH vEMAIL_CONF,;
MOSTRA         WITH vMOSTRA
IF LEN(aFiles) > 0
   For x=1 TO LEN(aFiles)
      APPEND BLANK
      REPLACE FILES WITH aFiles[x]
   Next
ENDIF
SELE EMAIL
USE
Return
