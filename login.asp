<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% option expilicit  
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Untitled Document</title>
</head>

<body>
if request("azione") == "" THEN
   CALL riconoscimento("INSERIMENTO")
else
   CALL riconoscimento(request("azione")
end if
FUNCTION riconoscimento(azione))
    SELECT CASE UCASE(AZIONE)
	       CASE "INSERIMENTO" %>
                <FORM method="get" action="">
                   <table>
                   <TR><TH>INSERISCI L'EMAIL</TH>
                      <TD><INPUT type="TEXT" NAME="EMAIL" /></TD>
                   </TR>
                   <TR><TH>INSERISCI LA PASSWORD</TH>
                      <TD><INPUT type="TEXT" NAME="PAS" /></TD>
                   </TR>
                   </table>
                   <INPUT type="hidden" NAME="AZIONE" value="VERIFICA">
                   <input TYPE="SUBMIT"  VALUE="SUBMIT">
            <% CASE "VERIFICA"
			     DIM EMAIL, PAS, RS, CON, SQL
				     email = Request.QueryString("email")
					 pas = Request.QueryString("pas")
					 SET CON = SERVER.CreateObject("ADODB.CONNECTION")
					 CON.CONNECTIONSTRING = "PROVIDER = MICROSOFT.JET.OLEDB.4.0; DATA SOURCE = " & SERVER.MapPath("SITOWEB.MDB")
					 CON.OPEN 
					 SET RS = SERVER.CreateObject("ADODB.RECORDSET")
					 SQL = "SELECT * FROM UTENTI WHERE EMAIL = ' " & email & " ' AND PASSWORD = '" & PAS & " ' ""
					 RS.OPEN SQL, CON
					    IF RS.EOF THEN 
						   RESPONSE.Write("<p>Dati inseriti non corretti</p>") &quot;
						   ("<a href= ""registrazione.asp"">Registrati</a>")
						   call riconoscimento("inserimento")
						else
						   Response.Write("<p>Ciao "& rs("Nome") & "<p>")
						   Session("user") = email
						end if
						rs.close
						con.close
						set rs = nothing
						set con = nothing
				 case else
				      Response.Write("<p>Azione non riconosciuta  AZIONE = " & AZIONE & "<p>")
			   END SELECT
END FUNCTION %>
			  
                   
                   

</body>
</html>
