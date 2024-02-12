unit unitemail;

{$mode objfpc}{$H+}
{$codepage utf8}

interface

uses
  Classes, SysUtils, pop3Send, ssl_openssl, laz_synapse,
  Dbf,  blcksock, FileUtil, Synacode, windows,
   { Alterado por  SIlvane Patrícia em  15/05/2017 - My suite 5521 }
   {General db Unit} sqldb,
   {add Conecction with SQLServer} pqconnection;

type
   RTagLinha = RECORD
   TAG:STRING;
   POS_BEG:INTEGER;
   POS_END:INTEGER;


  END;


  RCabecalho = RECORD
   _TO:STRING;
   _FROM:STRING;
   _DATE:STRING;
   _SUBJECT:STRING;
  END;

var
  POP3: TPOP3Send;
  TLINHA: RTagLinha;
  TCABECALHO:RCabecalho;
  ALINHA: ARRAY OF RTagLinha;
  AMIME: ARRAY OF STRING;
  AMIMEHEADER: ARRAY OF STRING;
  AFILENAME: ARRAY OF STRING;
  MSG_BLOCO:STRING;
  MSG_GRAVANDO:BOOLEAN;
  //Silvane
  FConn: TSQLConnector;
  FQuery: TSQLQuery;
  FTran: TSQLTransaction;

FUNCTION GET_EMAIL_POP(nPORTA:INTEGER; cSERVER, cUSUARIO, cSENHA, cAutoTLS, cFullSSL, cAuthType, cTESTE: PCHAR;cTipoConexaoDB,cServidorDB,cNomedoBanco,cUsuarioDB,cSenhaDB,nPortaDB:String):BOOLEAN;

implementation

FUNCTION CONECTA(cTipoConexaoDB:String;cServidorDB:String;cNomedoBanco:String;cUsuarioDB:String;cSenhaDB:String;nPortaDB:String):BOOLEAN;
BEGIN
        {Dll que está em produção ---11/08/2017 - Silvane }
        FConn:=TSQLConnector.Create(nil);
        FQuery:=TSQLQuery.Create(nil);
        FTran:=TSQLTransaction.Create(nil);
        FConn.Transaction:=FTran;
        FQuery.DataBase:=FConn;

        FConn.ConnectorType:=PCHAR(cTipoConexaoDB);
        FConn.HostName:=PCHAR(cServidorDB);
        Fconn.DatabaseName:=PCHAR(cNomedoBanco);
        FConn.UserName:=PCHAR(cUsuarioDB);
        FConn.Password:=PCHAR(cSenhaDB);
        FConn.Params.Add(PCHAR(nPortaDB));
        FConn.Transaction:=FTran;
        RESULT:=TRUE ;
END;

//SALVA EMAIL EM DBF
FUNCTION SALVA_EMAIL(cSERVER:String;NID:INTEGER):BOOLEAN;
BEGIN
        FQuery.SQL.Clear;
        FQuery.SQL.Text:=' INSERT INTO cag_eml('+
                       ' seqeml, dateml, frmeml, msgeml, popeml, titeml)'+
                       ' VALUES (:seqeml,:dateml,:frmeml, :msgeml,:popeml,:titeml)';

      FQuery.Params.ParamByName('seqeml').AsInteger:=NID;
      FQuery.Params.ParamByName('dateml').AsDate:=DATE();
      FQuery.Params.ParamByName('frmeml').AsString:=COPY(TCABECALHO._FROM, 6,(LENGTH(TCABECALHO._FROM)));
      IF LENGTH(AMIME) > 0 THEN BEGIN
         AMIME[0]:= STRINGREPLACE(AMIME[0], '<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
         FQuery.Params.ParamByName('msgeml').AsString:= AMIME[0];
      END ELSE BEGIN
          FQuery.Params.ParamByName('msgeml').AsString:= MSG_BLOCO;
      END;
      FQuery.Params.ParamByName('titeml').AsString:=COPY(TCABECALHO._SUBJECT, 9, (LENGTH(TCABECALHO._SUBJECT)));
      FQuery.Params.ParamByName('popeml').AsString:=cSERVER;

      FTran.StartTransaction;
      FQuery.ExecSQL;
      FQuery.Close;
      FTran.Commit;

   //LIMPA RECORD
   TCABECALHO._DATE:='';
   TCABECALHO._FROM:='';
   TCABECALHO._TO:='';
   TCABECALHO._SUBJECT:='';

   RESULT:= TRUE;
END;

//SALVAR ANEXOS EM DBF
FUNCTION SALVA_ANEXO(NID:INTEGER):BOOLEAN;

VAR
   X:INTEGER;
BEGIN

   FOR X := 1 TO LENGTH(AMIME)-1 DO BEGIN

      //DECODIFICAÇÃO QUOTED-PRINTABLE
      IF (POS('QUOTED', UPPERCASE(AMIMEHEADER[X])) > 0) THEN BEGIN
         AMIME[X]:= STRINGREPLACE(AMIME[X], '=<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
         AMIME[X]:= STRINGREPLACE(AMIME[X], '<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
         AMIME[X]:= Synacode.DecodeQuotedPrintable(AMIME[X]);
         AMIME[X]:= STRINGREPLACE(AMIME[X], '''', '"',[RFREPLACEALL,RFIGNORECASE]);
         AMIME[X]:= STRINGREPLACE(AMIME[X], '="', '</SYG>',[RFREPLACEALL,RFIGNORECASE]);
         AMIME[X]:= STRINGREPLACE(AMIME[X], '=', '',[RFREPLACEALL,RFIGNORECASE]);
         AMIME[X]:= STRINGREPLACE(AMIME[X], '</SYG>', '="',[RFREPLACEALL,RFIGNORECASE]);
       END;


      //Se a extensão do arquivo for XML

      IF (POS( '.XML', UPPERCASE(AFILENAME[X])) > 0) THEN BEGIN
        //Valido se não é um binário...ocorreu em meus testes
        IF (POS(UPPERCASE('<'), AMIME[X]) >0) THEN BEGIN
           AMIME[X]:= COPY(AMIME[X],1,(LENGTH(AMIME[X])));
        END ELSE BEGIN
           AMIME[X]:= Synacode.DecodeBase64(AMIME[X]);
        END;
      END ELSE BEGIN
         AMIME[X]:= Synacode.EncodeBase64(AMIME[X]);
      END;
      {
      //ERRO AO ABRIR .PDF SALVO NO DBF COM XHARBOUR
      IF (POS('.PDF', UPPERCASE(AMIMEHEADER[X])) > 0) AND (POS('.BASE64', UPPERCASE(AMIMEHEADER[X])) = 0) THEN BEGIN
           AMIME[X]:= Synacode.EncodeBase64(AMIME[X]);
      END;
      //ERRO AO ABRIR .RTF SALVO NO DBF COM XHARBOUR
      IF (POS('.RTF', UPPERCASE(AMIMEHEADER[X])) > 0) AND (POS('.BASE64', UPPERCASE(AMIMEHEADER[X])) = 0) THEN BEGIN
           AMIME[X]:= Synacode.EncodeBase64(AMIME[X]);
      END;
      }

      IF Trim(AMIME[X])='' THEN BEGIN
         RESULT:= FALSE;
         EXIT;
      END ELSE BEGIN
       //SOLICITADO POR LEONARDO EM 14/09/2017
        FQuery.SQL.Clear;
        FQuery.SQL.Text:=' set standard_conforming_strings to "off" ';
        FQuery.SQL.Text:=' set DateStyle to "iso, dmy"';
        FQuery.SQL.Text:=' set statement_timeout to "0"';
        FTran.StartTransaction;
        FQuery.ExecSQL;
        FQuery.Close;
        FTran.Commit;


        FQuery.SQL.Clear;
        FQuery.SQL.Text:=' INSERT INTO cag_aml('+
                         ' seqaml, nomaml,arqaml)'+
                         ' VALUES (:seqaml,:nomaml,:arqaml)';

        FQuery.Params.ParamByName('seqaml').AsInteger:=NID;
        FQuery.Params.ParamByName('nomaml').AsString:=AFILENAME[X];
        FQuery.Params.ParamByName('arqaml').AsString:=AMIME[X];

        FTran.StartTransaction;
        FQuery.ExecSQL;
        FQuery.Close;
        FTran.Commit;
        RESULT:= TRUE;
      END;
   END;


END;

//CONFIGURA OS BLOCOS DE TEXTOS E CAMPOS NESCESSÁRIOS
FUNCTION CONFIGURA_BLOCO_LINHA():BOOLEAN;
BEGIN

    //DATA E HORA
    TLINHA.TAG:='Date: ';
    TLINHA.POS_BEG:=0;
    TLINHA.POS_END:=0;
    SETLENGTH(ALINHA, LENGTH(ALINHA)+1);
    ALINHA[HIGH(ALINHA)] := TLINHA;

    //TO
    TLINHA.TAG:='To: ';
    TLINHA.POS_BEG:=0;
    TLINHA.POS_END:=0;
    SETLENGTH(ALINHA, LENGTH(ALINHA)+1);
    ALINHA[HIGH(ALINHA)] := TLINHA;

    //From:
    TLINHA.TAG:='From: ';
    TLINHA.POS_BEG:=0;
    TLINHA.POS_END:=0;
    SETLENGTH(ALINHA, LENGTH(ALINHA)+1);
    ALINHA[HIGH(ALINHA)] := TLINHA;

    //Subject:
    TLINHA.TAG:='Subject: ';
    TLINHA.POS_BEG:=0;
    TLINHA.POS_END:=0;
    SETLENGTH(ALINHA, LENGTH(ALINHA)+1);
    ALINHA[HIGH(ALINHA)] := TLINHA;

    RESULT:= TRUE;
END;

//PEGA O CONTEUDO DO EMAIL EM PARTES
FUNCTION PEGAR_CONTEUDO_EMAIL(ARQUIVO: TStringList):BOOLEAN;

VAR
   X:INTEGER;
   Y:INTEGER;
   XTEMP1:STRING;
   BOUNDARY_ABERTO:STRING;
   TEMP:STRING;
   MIXED:STRING;
   RELATED:STRING;
   ALTERNATIVE:STRING;
   TEMPHEADER:STRING;
   FILENAME:STRING;
   FILENAME_TEMP:STRING;
   MARCA_TEMP:STRING;
   ARQUIVOX_TEMP:STRING;
   //IN_DEBUG:STRING;

BEGIN
  TCABECALHO._DATE:='';
  TCABECALHO._FROM:='';
  TCABECALHO._TO:='';
  TCABECALHO._SUBJECT:='';
  XTEMP1:='';
  BOUNDARY_ABERTO:='';
  TEMPHEADER:='';
  TEMP:='';
  MIXED:='';
  ALTERNATIVE:='';
  RELATED:='';
  SETLENGTH(AMIME, 0);
  SETLENGTH(AMIMEHEADER, 0);
  TEMPHEADER:='';
  FILENAME:='';
  SETLENGTH(AFILENAME, 0);
  MSG_GRAVANDO:= FALSE;
  MSG_BLOCO:='';
  FILENAME_TEMP:='';
  MARCA_TEMP:='';
  ARQUIVOX_TEMP:='';
  //IN_DEBUG:='';

  FOR X := 1 TO ARQUIVO.Count -1 DO BEGIN

       // DEBUG
       //IN_DEBUG:= IN_DEBUG + ARQUIVO[x];

       //PEGA O CABECALHO
       FOR Y := 0 TO LENGTH(ALINHA) -1 DO BEGIN
         IF (POS( ALINHA[Y].TAG, ARQUIVO[X]) > 0)
         THEN BEGIN

            XTEMP1:= COPY(ARQUIVO[X], (POS( ALINHA[Y].TAG, ARQUIVO[X]) +
            ALINHA[Y].POS_BEG), (LENGTH(ARQUIVO[X]) - ALINHA[Y].POS_END) );

            IF ((Y = 0) AND (TCABECALHO._DATE = ''))
            THEN BEGIN
               TCABECALHO._DATE:=XTEMP1;
            end;
            IF ((Y = 1) AND (TCABECALHO._TO = ''))
            THEN BEGIN
               TCABECALHO._TO:=XTEMP1;
            end;
            IF ((Y = 2) AND (TCABECALHO._FROM = ''))
            THEN BEGIN
               TCABECALHO._FROM:=XTEMP1;
            end;
            IF ((Y = 3) AND (TCABECALHO._SUBJECT = ''))
            THEN BEGIN
               TCABECALHO._SUBJECT:=XTEMP1;
            end;
         END;
       END;

      //PEGA MENSAGEM QUANDO ESTA NÃO VEM COM MIME TYPE
      IF (POS( 'boundary=', ARQUIVO[X]) > 0) THEN BEGIN
         MSG_GRAVANDO:= FALSE;
         MSG_BLOCO:='X';
      END;
      IF MSG_GRAVANDO THEN BEGIN
            MSG_BLOCO := MSG_BLOCO + ' <br> ' + ARQUIVO[X];
      END;
      IF ((TRIM(ARQUIVO[X]) = '') AND (MSG_BLOCO = '')) THEN BEGIN
         MSG_GRAVANDO:= TRUE;
      END;

      ARQUIVOX_TEMP:= ARQUIVO[X];
      ARQUIVO[X]:= UPPERCASE(ARQUIVO[X]);

      //PEGA FILENAME DO ARQUIVO E COLOCAR EM UMA ARRAY GLOBAL
      IF (POS('FILENAME=',ARQUIVO[X]) > 0) AND (FILENAME = '') OR (POS('FILENAME*0=',ARQUIVO[X]) > 0) AND (FILENAME = '')
      THEN BEGIN

         IF (POS('FILENAME=',ARQUIVO[X]) > 0)
         THEN BEGIN

            FILENAME:= COPY(ARQUIVO[X], (POS('FILENAME',ARQUIVO[X])),
            (LENGTH(ARQUIVO[X])));
            FILENAME:= STRINGREPLACE(FILENAME, '"', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '=', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, 'FILENAME', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '''' , '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ' ', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ':', '', [RFREPLACEALL,RFIGNORECASE]);

         END ELSE BEGIN

            //SÓ PARA O THUNDER BIRD
            FILENAME:= COPY(ARQUIVO[X], (POS('FILENAME',ARQUIVO[X])),
            (LENGTH(ARQUIVO[X])));
            FILENAME:= STRINGREPLACE(FILENAME, '"', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '=', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, 'FILENAME', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '''' , '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ' ', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ':', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= COPY(FILENAME, 3, LENGTH(FILENAME));

            IF (POS(';', FILENAME) > 0) THEN BEGIN
                 FILENAME:= STRINGREPLACE(FILENAME, ';', '', [RFREPLACEALL,RFIGNORECASE]);
                 FILENAME_TEMP:= COPY(ARQUIVO[X+1], (POS('"',ARQUIVO[X])), (LENGTH(ARQUIVO[X])));
                 FILENAME_TEMP:= COPY(FILENAME_TEMP, 1, (POS('"',ARQUIVO[X])));
                 FILENAME_TEMP:= STRINGREPLACE(FILENAME_TEMP, '"', '', [RFREPLACEALL,RFIGNORECASE]);
                 FILENAME:= FILENAME + FILENAME_TEMP;
            END;

         END;

      END;

      //QUANDO A TAG FILENAME NÃO EXISTE PROCURA A TAG NAME
      IF (POS('NAME=',ARQUIVO[X]) > 0) AND (FILENAME = '')
      THEN BEGIN

         IF (POS('NAME=',ARQUIVO[X]) > 0)
         THEN BEGIN

            IF ((POS('.XML',ARQUIVO[X])) > 0) AND (FILENAME = '') THEN BEGIN
               FILENAME:= COPY(ARQUIVO[X], (POS('NAME',ARQUIVO[X])), (POS('.XML',ARQUIVO[X]))+2);
            END;

            IF ((POS('.PDF',ARQUIVO[X])) > 0) AND (FILENAME = '') THEN BEGIN
               FILENAME:= COPY(ARQUIVO[X], (POS('NAME',ARQUIVO[X])), (POS('.PDF',ARQUIVO[X]))+2);
            END;

             IF ((POS('CONTENT-',ARQUIVO[X])) > 0) AND (FILENAME = '') THEN BEGIN
               FILENAME:= COPY(ARQUIVO[X], (POS('NAME',ARQUIVO[X])), (POS('CONTENT-',ARQUIVO[X]))-1);
            END;

            FILENAME:= STRINGREPLACE(FILENAME, '"', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '=', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, 'NAME', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, '''' , '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ' ', '', [RFREPLACEALL,RFIGNORECASE]);
            FILENAME:= STRINGREPLACE(FILENAME, ':', '', [RFREPLACEALL,RFIGNORECASE]);

         END;

      END;

      ARQUIVO[X]:= ARQUIVOX_TEMP;
      FILENAME:= LOWERCASE(FILENAME);

      IF (POS(MIXED, ARQUIVO[X]) = 0) AND (POS(ALTERNATIVE, ARQUIVO[X]) = 0) AND (POS(RELATED, ARQUIVO[X]) = 0)
      THEN BEGIN

         TEMP:= TEMP + ARQUIVO[X] + MARCA_TEMP;

         IF POS('</SYG>', TEMPHEADER) = 0 THEN BEGIN

            IF ARQUIVO[X] = '' THEN BEGIN

               TEMPHEADER := TEMPHEADER + '</SYG>';
               IF (POS('QUOTED', UPPERCASE(TEMPHEADER)) > 0) THEN BEGIN
                  MARCA_TEMP:=  '<SYG-Q>';
               END;

            END ELSE BEGIN
               TEMPHEADER := TEMPHEADER + ARQUIVO[X];
            END;

         END;
      END;

     //PEGA MARCACOES
     IF MIXED = ''  THEN BEGIN
        IF (POS( 'multipart/mixed', ARQUIVO[X]) > 0)
        THEN BEGIN
           IF (POS( 'boundary=', ARQUIVO[X]) > 0) THEN BEGIN
              MIXED:= COPY(ARQUIVO[X], (POS( 'boundary=', ARQUIVO[X])+10), 20 );
           END ELSE BEGIN
              MIXED:= COPY(ARQUIVO[X+1], (POS('boundary=',ARQUIVO[X+1])+10),20);
           END;
        END ELSE BEGIN  //Adicionado por Silvane Patrícia em 04/11/2016  - My suite 4580 e 4617
          IF (POS('Multipart/mixed', ARQUIVO[X]) > 0)
          THEN BEGIN
               IF (POS( 'boundary=', ARQUIVO[X]) > 0)
             THEN BEGIN
                 MIXED:= COPY(ARQUIVO[X],(POS('boundary=',ARQUIVO[X])+10),20);
              END ELSE BEGIN
                 MIXED:=COPY(ARQUIVO[X+1],(POS('boundary=',ARQUIVO[X+1])+10),20);
              END;
          END;
        END;
     END;
     IF ALTERNATIVE = ''  THEN BEGIN
        IF (POS( 'multipart/alternative', ARQUIVO[X]) > 0)
        THEN BEGIN
           IF (POS( 'boundary=', ARQUIVO[X]) > 0)
           THEN BEGIN
              ALTERNATIVE:= COPY(ARQUIVO[X],(POS('boundary=',ARQUIVO[X])+10),20);
           END ELSE BEGIN
              ALTERNATIVE:=COPY(ARQUIVO[X+1],(POS('boundary=',ARQUIVO[X+1])+10),20);
           END;
        END;
     END;
     IF RELATED = ''  THEN BEGIN
        IF (POS( 'multipart/related', ARQUIVO[X]) > 0)
        THEN BEGIN
           IF (POS( 'boundary=', ARQUIVO[X]) > 0)
           THEN BEGIN
              RELATED:= COPY(ARQUIVO[X],(POS('boundary=',ARQUIVO[X])+10),20);
           END ELSE BEGIN
              RELATED:=COPY(ARQUIVO[X+1],(POS('boundary=',ARQUIVO[X+1])+10),20);
           END;
        END;
     END;

     IF (POS(MIXED, ARQUIVO[X]) > 0) AND (POS('boundary=', ARQUIVO[X]) = 0)
     THEN BEGIN
        IF BOUNDARY_ABERTO = MIXED THEN BEGIN
           TEMP := STRINGREPLACE(TEMP,COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1), '',[rfReplaceAll,rfIgnoreCase]);

           SETLENGTH(AMIME, LENGTH(AMIME)+1);
           AMIME[HIGH(AMIME)]:= TEMP;
           TEMPHEADER:= COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1);
           TEMPHEADER:= STRINGREPLACE(TEMPHEADER, '<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
           SETLENGTH(AMIMEHEADER, LENGTH(AMIMEHEADER)+1);
           AMIMEHEADER[HIGH(AMIMEHEADER)]:= TEMPHEADER;
           SETLENGTH(AFILENAME, LENGTH(AFILENAME)+1);
           AFILENAME[HIGH(AFILENAME)]:= FILENAME;

           FILENAME:='';
           TEMP:='';
           TEMPHEADER:='';
           MARCA_TEMP:=  '';

        END ELSE BEGIN
           BOUNDARY_ABERTO:= MIXED;
           TEMP:='';
           TEMPHEADER:='';
           FILENAME:='';
        END;

     END;

     IF (POS(ALTERNATIVE, ARQUIVO[X]) > 0) AND (POS('boundary=', ARQUIVO[X]) = 0)
     THEN BEGIN
        IF BOUNDARY_ABERTO = ALTERNATIVE THEN BEGIN

           TEMP := STRINGREPLACE(TEMP,COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1), '',[rfReplaceAll,rfIgnoreCase]);

           SETLENGTH(AMIME, LENGTH(AMIME)+1);
           AMIME[HIGH(AMIME)]:= TEMP;
           TEMPHEADER:= COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1);
           TEMPHEADER:= STRINGREPLACE(TEMPHEADER, '<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
           SETLENGTH(AMIMEHEADER, LENGTH(AMIMEHEADER)+1);
           AMIMEHEADER[HIGH(AMIMEHEADER)]:= TEMPHEADER;
           SETLENGTH(AFILENAME, LENGTH(AFILENAME)+1);
           AFILENAME[HIGH(AFILENAME)]:= FILENAME;

           FILENAME:='';
           TEMP:='';
           TEMPHEADER:='';
           MARCA_TEMP:=  '';

        END ELSE BEGIN
           BOUNDARY_ABERTO:= ALTERNATIVE;
           TEMP:='';
           TEMPHEADER:='';
           FILENAME:='';
        END;
     END;

     IF (POS(RELATED, ARQUIVO[X]) > 0) AND (POS('boundary=', ARQUIVO[X]) = 0)
     THEN BEGIN
        IF BOUNDARY_ABERTO = RELATED THEN BEGIN

           TEMP := STRINGREPLACE(TEMP,COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1), '',[rfReplaceAll,rfIgnoreCase]);

           SETLENGTH(AMIME, LENGTH(AMIME)+1);
           AMIME[HIGH(AMIME)]:= TEMP;
           TEMPHEADER:= COPY(TEMPHEADER, 0, POS( '</SYG>', TEMPHEADER)-1);
           TEMPHEADER:= STRINGREPLACE(TEMPHEADER, '<SYG-Q>', '',[RFREPLACEALL,RFIGNORECASE]);
           SETLENGTH(AMIMEHEADER, LENGTH(AMIMEHEADER)+1);
           AMIMEHEADER[HIGH(AMIMEHEADER)]:= TEMPHEADER;
           SETLENGTH(AFILENAME, LENGTH(AFILENAME)+1);
           AFILENAME[HIGH(AFILENAME)]:= FILENAME;

           FILENAME:='';
           TEMP:='';
           TEMPHEADER:='';
           MARCA_TEMP:=  '';

        END ELSE BEGIN
           BOUNDARY_ABERTO:= RELATED;
           TEMP:='';
           TEMPHEADER:='';
           FILENAME:='';
        END;
     END;

  END;

  // DEBUG
  //MESSAGEBOX(0,PCHAR(IN_DEBUG),PCHAR(''),MB_OK);

  RESULT:= TRUE;
END;

//LER TODOS OS EMAILS DA CAIXA DE ENTRADA
FUNCTION LER_EMAIL(cSERVER:String):BOOLEAN;
VAR
    ARQUIVO: TStringList;
    I:INTEGER;
    NID:INTEGER;
BEGIN

   POP3.Stat();

   //TESTA SE RETORNOU ALGUM EMAIL
   IF POP3.STATCOUNT > 0 THEN BEGIN
      FOR I := 1 TO POP3.StatCount DO BEGIN
         POP3.Retr(I);
         ARQUIVO:= POP3.FullResult;

         //Adicionado por Silvane Patrícia
         FQuery.SQL.Text:='SELECT seqeml FROM cag_eml ORDER BY sr_recno DESC LIMIT 1';
         FTran.StartTransaction;
         FQuery.Open;
         NID:=FQuery.Fields[0].AsInteger + 1;
         FQuery.Close;
         FTran.Commit;
         PEGAR_CONTEUDO_EMAIL(ARQUIVO);

         SALVA_EMAIL(cSERVER,NID);
         SALVA_ANEXO(NID);
         POP3.Dele(I);
      END;
   END;
   RESULT:= TRUE;
END;

//FUNÇÃO LER EMAILS
FUNCTION GET_EMAIL_POP(nPORTA:INTEGER; cSERVER, cUSUARIO, cSENHA, cAutoTLS, cFullSSL, cAuthType, cTESTE: PCHAR;cTipoConexaoDB,cServidorDB,cNomedoBanco,cUsuarioDB,cSenhaDB,nPortaDB:String):BOOLEAN;

begin

  RESULT:= TRUE;

  TRY

   POP3:= TPOP3Send.Create();

   //CONFIGURA CONEXÃO
   POP3.Username:=TRIM(STRING(cUSUARIO));
   POP3.Password:=TRIM(STRING(cSENHA));
   POP3.TargetHost:=TRIM(STRING(cSERVER));
   POP3.TargetPort:=TRIM(INTTOSTR(nPORTA));

   IF TRIM(STRING(cAuthType)) = 'POP3AuthAll' THEN BEGIN
    POP3.AuthType:= POP3AuthAll;
   END;
   IF TRIM(STRING(cAuthType)) = 'POP3AuthLogin' THEN BEGIN
     //MESSAGEBOX(0,PCHAR('Erro: tipo ' + cAuthType + ' nao suportado.'),PCHAR(''),MB_OK);
     POP3.AuthType:= POP3AuthAll;
   END;
   IF TRIM(STRING(cAuthType)) = 'POP3AuthAPOP' THEN BEGIN
     //MESSAGEBOX(0,PCHAR('Erro: tipo ' + cAuthType + ' nao suportado.'),PCHAR(''),MB_OK);
     POP3.AuthType:= POP3AuthAll;
   END;

   IF TRIM(STRING(cFullSSL)) = 'TRUE' THEN BEGIN
      POP3.FullSSL := TRUE;
   END ELSE BEGIN
      POP3.FullSSL := FALSE;
   END;
   IF TRIM(STRING(cAutoTLS)) = 'FALSE' THEN BEGIN
      POP3.FullSSL := TRUE;
   END ELSE BEGIN
      POP3.FullSSL := FALSE;
   END;

   //TESTA CONEXÃO
   IF TRIM(STRING(cTESTE)) = 'TRUE' THEN BEGIN
     IF NOT POP3.Login() THEN BEGIN
       RESULT:= FALSE;
     END;
     POP3.Logout;
     POP3.Free;
     EXIT;
   END;
   //BAIXA EMAILS
   CONECTA(cTipoConexaoDB,cServidorDB,cNomedoBanco,cUsuarioDB,cSenhaDB,nPortaDB);
   CONFIGURA_BLOCO_LINHA();
   IF NOT POP3.LOGIN() THEN BEGIN
      RESULT:= FALSE;
      POP3.Free;
      EXIT;
   END;

   IF NOT LER_EMAIL(cSERVER)  THEN BEGIN
      RESULT:= FALSE;
      POP3.Free;
      EXIT;
   END;
   POP3.LOGOUT;
   POP3.Free;
   EXIT;

  EXCEPT
     MESSAGEBOX(0,PCHAR('ERRO INESPERADO AO BAIXAR EMAIL.'),PCHAR(''),MB_OK);
     RESULT:= FALSE;
     POP3.Free;
     EXIT;
  END;

END;

END.


