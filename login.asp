<%@ Language=VBScript %>

<HTML>
<HEAD>
<script language="javascript" type="text/javascript">
<!--

function Submit1_onclick() {

}

// -->
</script>
</HEAD>
<BODY>
<form method="get">
		<table>
		<tr><th>Nome utente</th><td><input type="text" name="nome_utente"></td></tr>
		<tr><th>Password</th><td><input type="password" name="password"></td></tr>
		</table>
		<input type="submit" value="invia" id="Submit1" language="javascript" onclick="return Submit1_onclick()">
</form>
<% if request("ret").count>0 then %>
	sei entrato in area riservata, devi fare il login <br/>
	<% end if %>


<% if request("nome_utente")<>"" and request("password")<>"" then

	dim conn,rs

	set conn = server.CreateObject ("ADODB.connection")
	conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0; DATA SOURCE =" & Server.MapPath("sessione.mdb")
	conn.Open

	set rs = server.CreateObject ("ADODB.recordset")
	rs.Open "select * from utenti where nome_utente like '" & ucase(request("nome_utente")) & "'",conn 

	if rs.EOF then %>
		NON SEI REGISTRATO!
	<% else
			if rs("password")<>lcase(request("password")) then %>
			HAI SBAGLIATO LA PASSWORD!
			<% else 
					session("nome_utente")=rs("nome_utente")
					session("ruolo")=rs("ruolo")
					Response.Redirect "destinazione.asp"
			end if
	end if
	
	rs.Close
	set rs=nothing
	conn.Close
	set conn=nothing
	
else %> DEVI INSERIRE NOME E PASSWORD!
		
<% end if %>
<% if request("remove").count>0 then
	session.Abandon
	Response.Write ("sessione finita")
	end if %>
</BODY>
</HTML>