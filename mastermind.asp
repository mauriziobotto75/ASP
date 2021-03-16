<%@ Language=VBScript %>
<%   
         OPTION EXPLICIT
         dim num1, num2, num3, NUM, numero_intero, i 
     
    DIM STRINGASQL
    DIM CONNESSIONE, RECORDSET 
    
  ' SET è IL COMANDO ASP VISUAL BASIC  PER SETTARE/CREARE ED IMPOSTARE UN OGGETTO
  '                ' 
   SET CONNESSIONE = SERVER.CreateObject("ADODB.connection")
  ' l'oggetto server (asp) crea un oggetto utilizzando il componente ado che sta per 
  ' activeX Data Object 
  ' ado è un componente molto ricco e la sua documentazione e su microsoft msdn
  ' (cercare ADO reference)
    connessione.ConnectionString = "PROVIDER = Microsoft.jet.oledb.4.0; DATA SOURCE =" &  SERVER.MapPath("/mastermind/numeri.mdb")
 
    
   CONNESSIONE.Open 
   
   set recordset = server.CreateObject("ADODB.recordset")
 
   
   STRINGASQL="SELECT * FROM numero" 
   num1 = 5
   num2 = 7
   num3 = 9 
   numero_intero = num1 & num2 & num3
   RECORDSET.Open STRINGASQL, CONNESSIONE 
      I = 1 
     
      
          NUM = Request.QueryString("NUMERO1") &Request.QueryString("NUMERO2") &Request.QueryString("NUMERO3")
                 
          IF NUM = numero_intero  THEN 
             Response.Write("Complimenti hai azzeccato i tre numeri")
            
              RECORDSET.AddNew 
                 recordset("Numero1") = num1
                 recordset("Numero2") = num2
                 recordset("Numero3") = num3
              RECORDSET.Update     
             
           ELSE 
             IF NUM1 = Request.QueryString("NUMERO1") THEN
                Response.Write("HAI AZZECCATO IL PRIMO NUMERO")
             ELSE IF NUM2 = Request.QueryString("NUMERO2") THEN
                Response.Write("HAI AZZECCATO IL SECONDO NUMERO")
             ELSE IF NUM3 = Request.QueryString("NUMERO3") THEN
                Response.Write("HAI AZZECCATO IL TERZO NUMERO")
             END IF             
           END IF 
        I = I + 1   
           if i > 15 AND NUM <> NUMERO_INTERO then 
              Response.Write("Game over") 
              Response.Write("NON HAI AZZECCATO NESSUN NUMERO") 
              i = 1 
           end if   
   
   
   
 
%>        
 










<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<form name="form1"   method="get"   action="mastermind.asp">
NUMERO1<input type="text"   name="numero1"  size="2" maxlength="1"><BR>
NUMERO2<INPUT TYPE="TEXT" NAME="NUMERO2"  SIZE="2" MAXLENGTH="1"><BR>
NUMERO3<INPUT TYPE="TEXT" NAME="NUMERO3"  SIZE="2" MAXLENGTH="1"><BR>
        

<input type="submit" value="invia">
</form>
<P>&nbsp;</P>

</BODY>
</HTML>
<% RECORDSET.Close 
   CONNESSIONE.Close 
%>   
              
  