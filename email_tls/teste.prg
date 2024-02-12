   /*** tst.prg ***/
proc main( lTLS )
      LOCAL cFrom, cPassword, cTo

      cFrom     := "<unk...@gmail.com>"
      cPassword := "123-abc-!@#"
      cTo       := "unkn...@gmail.com"

      if lTLS == NIL
         lTLS := .F.
      else
         lTLS := !empty( lTLS )
      endif

      ? "TLS:", lTLS
      ? hb_SendMail( "smtp.gmail.com" /* cServer */,;
                     465 /* nPort*/, ;
                     cFrom,;
                     cTo,;
                     NIL /* CC */,;
                     {} /* BCC */,;
                     "test: body",;
                     "test: subject",;
                     /* aFiles (attachment) */,;
                     cFrom,;
                     cPassword,;
                     ""   /* cPopServer */ ,;
                     NIL  /* nPriority */,;
                     NIL  /* lRead */,;
                     .T.  /* lTrace */,;
                     .F.  /* lPopAuth */,;
                     NIL  /* lNoAuth */,;
                     NIL  /* nTimeOut */,;
                     NIL  /* cReplyTo */,;
                     lTLS /* lTLS */ )
      ?
return
