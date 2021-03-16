<% 
'1)impostando la proprieta Buffer dell'oggetto response = true
'prima viene elaborato tutto il codice asp e poi è inviato al browser il risultato
'2) Serve per la gestione degli errori interni all'applicazione
'   se si verifica un errore è possibile cancellare tutta l'elaborazione e mandare un messaggio
'   all'utente  
' di default è settato a false
' se si usa questo è obbligatorio scriverlo prima del tag <html>
  Response.Buffer = true%>
<html>
<head>
<title>Programma che visualizza checkbox</title>
</head>
<body bgColor="GREEN">
<%' in caso di errore passa oltre. In questo modo se si verifica un errore l'elaborazione procede fino al punto in cui l'errore viene intercettato %>
<% on error resume next %>

<%' TAG FORM:
' ATTRIBUTI DEL TAG:
' NAME indica il nome del form
'method indica come vengono inviati i dati:
' se è uguale a get reperisco i dati tramite request del querystring 
' se è uguale a post reperisco i dati con request.form("[Nomeform]") (Oggetto FORM del request)
'action pagina a cui viene inviato il contenuto del form
'Nell'invio dei dati posso usare il tag input con attributo type = submit
'Posso usare il tag input con:
' attributo type = buttom.
' inserisco l'attributo onclick = document.form[0].submit();
' si tratta di un istruzione javascript
%>
<h1>GENERAZIONE DINAMICA CHECKBOX</H1>
<form name="crea_bottoni" method="get" action="provaasp.asp">
 
<input type="text"  name="numcampi"  size="3"  maxlength="2">
<input type="submit" name="submit"  value="ok">
<input type="reset"  name="reset"   value="cancella i dati">
</FORM>
<%' testo se il parametro "Control" è presente nel querystring
' vedo se ho inserito dei valori non consentiti (In questo caso un testo o un valore non compreso tra 1 e 10)
' assegno ad una variabile il valore di un parametro. Se il parametro è uguale ad un valore non consentito
' mando un messaggio di errore (scrivo un testo html usando response.write)   %>     
             
<%    if Request.QueryString("control").Count > 0 then 
          Feedback=Request.QueryString("control")              
         if feedback = "text" then
         Response.Write "Devi inserire un valore numerico"
         elseif feedback = "extrarange" then
         Response.Write "Devi inserire un valore tra 1 e 10"
      end if    
          end if    %>  
<%' Testo se la proprieta number dell'oggetto err ha un valore: se è stato generato un errore nell'applicazione          
'se ho un errore faccio tre cose:
'response.clear svuoto il buffer ( la memoria da tutta l'elaborazione fatta fino a quel punto)
'response.write mando un messaggio di errore all'utente posso fare questo perche all'inizio ho scritto response.buffer = "true"
'usando number e description (numero e descrizione dell'errore dell'oggetto err  
'response.end  termino l'esecuzione del codice e invio all'utente il messaggio di errore
%>    
<%        if err.number <> 0 then
             Response.Clear     
             Response.Write "Si è verificato un errore " & err.Description & " numero " & err.number & ".<br>Contattare l'amministratore"
             Response.End 
          end if             %> 
<%'Controllo il valore del numero dei campi se numcampi è presente nel querystring.
'assegno ad una variabile il valore del parametro 
'testo se il contenuto della variabile è di tipo numerico: se non è numerico scarico la pagina aggiungendo il querystring con control="text"
'se ho inserito un valore non compreso tra 0 e 10 scarico la pagina aggiungendo al querystring la variabile control settata ad extrarange  %>        
            
<%    if Request.QueryString("numcampi").Count > 0 then 
           fieldcount=REQUEST.QUERYStRING("numcampi")          
         if not isnumeric(fieldcount) then 
         Response.Redirect("provaasp.asp?control=text")
         elseif fieldcount = 0 or fieldcount > 10 then
         Response.Redirect("provaasp.asp?control=extrarange")
      else    
                     
%>    
<%' altrimenti se il contenuto della variabile supera i controlli genero una tabella %> 
<TABLE>        
<%    for bcounter=1 to fieldcount  %>
 
    <TR
      <TD>
<%'con un ciclo genero dinamicamente il numero di checkbox oppure option tante quante sono il numero inserito (uguale alla variabile fieldcount %>    
        <INPUT TYPE="CHECKBOX" NAME="SCELTA<%=BCOUNTER%>">scelta <%=BCOUNTER%> 
      </TD>
    </TR>
<%    NEXT                  %>

</TABLE>              

<%  end if      %>

<%  END IF      %>

</body>
</html>