unit unitsendemail;
{$mode objfpc}{$H+}
{$codepage utf8}

interface

uses
   Classes, SysUtils, FileUtil, blcksock, smtpsend, ssl_openssl,
   Mimemess, Mimepart, synautil, laz_synapse, windows;

var
   SMTP: TSMTPSend;
   STRLIST: TStringList;
   MIME: TMimeMess;
   PARTE: TMimepart;

   //FUNÇÕES PUBLICAS
   FUNCTION SEND_EMAIL_SMTP(cPORTA, cSERVER, cUSER, cSENHA, cSSL, cTLS, cFROM, cMSG, cSUBJECT, sFILES, sBCC, sQUEM, sCC, vEMAIL_CONF, lHTML:PCHAR):BOOLEAN;

implementation

//ATIVA FORMULARIO
FUNCTION SEND_EMAIL_SMTP(cPORTA, cSERVER, cUSER, cSENHA, cSSL, cTLS, cFROM, cMSG, cSUBJECT, sFILES, sBCC, sQUEM, sCC, vEMAIL_CONF, lHTML:PCHAR):BOOLEAN;

VAR
   CONFIRMA_EMAIL: BOOLEAN;
   lHTML_EMAIL:BOOLEAN;
   sCC2: STRING;
BEGIN

   RESULT:= TRUE;
   CONFIRMA_EMAIL:=FALSE;
   lHTML_EMAIL:=FALSE;

   TRY
      //CRIANDO OBJETOS
      SMTP:= TSMTPSend.CREATE;
      STRLIST:= TStringList.CREATE;
      MIME:= TMimeMess.CREATE;

      IF TRIM(lHTML) = '' THEN BEGIN
         lHTML_EMAIL:= FALSE;
      END ELSE BEGIN
         IF TRIM(lHTML) = 'FALSE' THEN BEGIN
            lHTML_EMAIL:= FALSE;
         END ELSE BEGIN
            lHTML_EMAIL:= TRUE;
         END;
      END;

      IF TRIM(cTLS) = 'FALSE' THEN BEGIN
         SMTP.AutoTLS:= FALSE;
      END ELSE BEGIN
         SMTP.AutoTLS:= TRUE;
      END;

      IF TRIM(cSSL) = 'FALSE' THEN BEGIN
         SMTP.FullSSL:= FALSE;
      END ELSE BEGIN
          SMTP.FullSSL:= TRUE;
      END;

      IF TRIM(vEMAIL_CONF) = 'FALSE' THEN BEGIN
          CONFIRMA_EMAIL:= FALSE;
      END ELSE BEGIN
          CONFIRMA_EMAIL:= TRUE;
      END;

      SMTP.TargetPort:= cPORTA;
      SMTP.TargetHost:= cSERVER;
      SMTP.UserName:= cUSER;
      SMTP.Password:= cSENHA;
      SMTP.StartTLS;

      //LOGIN
      IF NOT SMTP.LOGIN() THEN BEGIN
         RESULT:= FALSE;
         STRLIST.FREE;
         MIME.FREE;
         SMTP.FREE;
         EXIT;
      END;

      //ADICIONA MENSAGEM
      STRLIST.Add(cMSG);

      //SMTP
      Mime.Header.From := LOWERCASE(cFROM);

      //DESTINATARIO
      Mime.Header.ToList.Add(LOWERCASE(sQUEM));

      //ASSUNTO
      Mime.Header.Subject:= cSUBJECT;
{
      //COPIA-ORIGINAL
      IF NOT (TRIM(sCC)='') THEN BEGIN
         Mime.Header.CCList.Add(sCC);
      END;
}

      IF NOT (TRIM(sCC)='') THEN BEGIN
          sCC2:=sCC;
          sCC:= PCHAR(sCC + ';');
          WHILE POS(';',sCC) > 0 DO BEGIN
             IF NOT (PCHAR(COPY(sCC, 0, POS(';',sCC)-1)) = '') THEN BEGIN
                Mime.Header.CCList.Add(PCHAR(COPY(sCC, 0, POS(';',sCC)-1)));
                sleep(500);
             END;
             sCC:= PCHAR(COPY(sCC, POS(';',sCC)+1, LENGTH(sCC)));
          END;
      END;

      //COPIA OCULTA
      IF NOT (TRIM(sBCC)='') THEN BEGIN
         Mime.Header.CustomHeaders.Add('Bcc:' + LOWERCASE(sBCC));
      END;

      //CONFIRMACAO DE EMAIL
      IF CONFIRMA_EMAIL THEN BEGIN
         Mime.Header.CustomHeaders.Add('Disposition-Notification-To:' + LOWERCASE(cFROM));
      END;

      parte := Mime.AddPartMultipart('mixed',nil);
      Mime.AddPartMultipart('', Nil);

      //ENVIO HTML OU TEXTO
      IF lHTML_EMAIL THEN BEGIN
         Mime.AddPartHTML(STRLIST,parte);
      END ELSE BEGIN
         Mime.AddPartText(StrList, parte);
      END;

      //ADICIONA ANEXOS
      sFILES:= PCHAR(sFILES + ';');

      WHILE POS(';',sFILES) > 0 DO BEGIN
         IF NOT (PCHAR(COPY(sFILES, 0, POS(';',sFILES)-1)) = '') THEN BEGIN
            Mime.AddPartBinaryFromFile(PCHAR(COPY(sFILES, 0, POS(';',sFILES)-1)), parte);
            sleep(500);
         END;
         sFILES:= PCHAR(COPY(sFILES, POS(';',sFILES)+1, LENGTH(sFILES)));
      END;

      //Encode menasagem
      Mime.EncodeMessage;

      //DEBUG DA SIDA EM TEXTO
      //MESSAGEBOX(0,PCHAR(Mime.Lines.Text),PCHAR(''),MB_OK);

      //smtp
      IF NOT SMTP.MailFrom(cFROM,Length(Mime.Lines.Text)) THEN BEGIN
         //MESSAGEBOX(0,PCHAR(cFROM),PCHAR(''),MB_OK);
         RESULT:= FALSE;
         STRLIST.FREE;
         MIME.FREE;
         SMTP.FREE;
         EXIT;
      END;

      //destinatario
      IF NOT SMTP.MailTo(sQUEM) THEN BEGIN
         //MESSAGEBOX(0,PCHAR(sQUEM),PCHAR(''),MB_OK);
         RESULT:= FALSE;
         STRLIST.FREE;
         MIME.FREE;
         SMTP.FREE;
         EXIT;
      END;
{
      //COPIA NORMAL - ORIGINAL
      IF NOT SMTP.MailTo(sCC) THEN BEGIN
         //MESSAGEBOX(0,PCHAR(sCC),PCHAR(''),MB_OK);
         RESULT:= FALSE;
         STRLIST.FREE;
         MIME.FREE;
         SMTP.FREE;
         EXIT;
      END;
}

      IF NOT (TRIM(sCC2)='') THEN BEGIN
//         MESSAGEBOX(0,PCHAR('1'),PCHAR(''),MB_OK);
         sCC2:= PCHAR(sCC2 + ';');
         WHILE POS(';',sCC2) > 0 DO BEGIN
//            MESSAGEBOX(0,PCHAR('2'),PCHAR(''),MB_OK);
            IF NOT (PCHAR(COPY(sCC2, 0, POS(';',sCC2)-1)) = '') THEN BEGIN

 //              MESSAGEBOX(0,PCHAR(COPY(sCC2, 0, POS(';',sCC2)-1)),PCHAR(''),MB_OK);

               IF NOT SMTP.MailTo(PCHAR(COPY(sCC2, 0, POS(';',sCC2)-1))) THEN BEGIN
 //                 MESSAGEBOX(0,PCHAR('DEU ERRO NA COPIA'),PCHAR(''),MB_OK);
                  RESULT:= FALSE;
                  STRLIST.FREE;
                  MIME.FREE;
                  SMTP.FREE;
                  EXIT;
               END;
               sleep(500);
            END;
            sCC2:= PCHAR(COPY(sCC2, POS(';',sCC2)+1, LENGTH(sCC2)));
         END;
      END;

      //Copia Oculta
      IF NOT (TRIM(sBCC)='') THEN BEGIN
         IF NOT SMTP.MailTo(sBCC) THEN BEGIN
            //MESSAGEBOX(0,PCHAR(sBCC),PCHAR(''),MB_OK);
            RESULT:= FALSE;
            STRLIST.FREE;
            MIME.FREE;
            SMTP.FREE;
            EXIT;
         END;
      END;

      IF NOT SMTP.MailData(Mime.Lines) THEN BEGIN
         RESULT:= FALSE;
         STRLIST.FREE;
         MIME.FREE;
         SMTP.FREE;
         EXIT;
      END;

      SMTP.LOGOUT;
      STRLIST.FREE;
      MIME.FREE;
      SMTP.FREE;

   EXCEPT



      MESSAGEBOX(0,PCHAR('ERRO INESPERADO AO ENVIAR EMAIL.'),PCHAR(''),MB_OK);
      RESULT:= FALSE;
      STRLIST.FREE;
      MIME.FREE;
      SMTP.FREE;
      EXIT;
   END;

END;

END.

//fMIMEMess            : TMimeMess;
//property MIMEMess: TMimeMess read fMIMEMess;
//fMIMEMess.Header.Priority;
//fMIMEMess.Header.Clear;
//Self.MIMEMess.Header.CCList.Assign( MIMEMess.Header.CCList )
//fMIMEMess.Header.CharsetCode := fDefaultCharsetCode;
//fMIMEMess.Header.XMailer := 'Synapse - ACBrMail'
//fSMTP.MailTo(GetEmailAddr(fMIMEMess.Header.CCList.Strings[i]))
