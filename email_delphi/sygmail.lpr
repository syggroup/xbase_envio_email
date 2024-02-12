library sygmail;

  {$mode objfpc}{$H+}
  {$codepage utf8}

  uses
  //{$IFDEF UNIX}{$IFDEF UseCThreads}
  //cthreads,
  //{$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, laz_synapse, sysutils, unitemail, unitsendemail;

  //{$R *.res}

   //LER EMAILS
  function LZ_GET_EMAIL_POP(nPORTA:INTEGER; cSERVER, cUSUARIO, cSENHA, AutoTLS, FullSSL, AuthType, cTESTE: PCHAR;cTipoConexaoDB,cServidorDB,cNomedoBanco,cUsuarioDB,cSenhaDB,nPortaDB:String):boolean; stdcall;
   begin
       RESULT:= GET_EMAIL_POP(nPORTA,cSERVER, cUSUARIO, cSENHA, AutoTLS, FullSSL, AuthType, cTESTE,cTipoConexaoDB,cServidorDB,cNomedoBanco,cUsuarioDB,cSenhaDB,nPortaDB);
   end;
   exports LZ_GET_EMAIL_POP;

   //ENVIAR EMAILS
   function LZ_SEND_EMAIL_SMTP(cPORTA, cSERVER, cUSER, cSENHA, cSSL, cTLS, cFROM, cMSG, cSUBJECT, sFILES, sBCC, sQUEM, sCC, vEMAIL_CONF, lHTML:PCHAR):boolean; stdcall;
   begin
       RESULT:= SEND_EMAIL_SMTP(cPORTA, cSERVER, cUSER, cSENHA, cSSL, cTLS, cFROM, cMSG, cSUBJECT, sFILES, sBCC, sQUEM, sCC, vEMAIL_CONF, lHTML);
   end;
   exports LZ_SEND_EMAIL_SMTP;

end.

