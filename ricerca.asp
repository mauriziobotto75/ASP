<%@ Language=VBScript %>
<%Option Explicit%>

<HTML>
<HEAD>
<title><h1>IL RICERCONE</h1></title>
</HEAD>
<BODY>
<%
DIM azione, r, parole, con, rs


azione = Request("azione")
SELECT CASE azione
	CASE 1
			'recupera i dati del form e inizia la ricerca
			r = request("r")
			if r = "" then
				Response.Write ("non hai inserito il testo")
			else
				Response.Write("le parole da cercare sono:" & r)
				if instr(1,r," ") = 0 then
						parole=r
						'ricerca
				set con= server.CreateObject("adodb.connection")
				set rs= server.CreateObject("adodb.recordset")
    con.connectionstring = "provider = microsoft.jet.oledb.4.0;data source = "& server.MapPath("ricerca.mdb")
    con.open 
    rs.Open "select testo from testi",con
    while not rs.eof
		if instr(1,rs("testo"),parole,1) <> 0 then
		Response.Write("<p>" & rs("testo") &"</p>")
		end if
		rs.MoveNext
		wend
		rs.Close
		con.close
				else
					parole=split(r," ")
				end if
			end if
		CALL FormRicerca
	CASE ELSE
		'visualizza il form
		CALL FormRicerca
	END SELECT


SUB FormRicerca
%>
<form action="" method="post">
<INPUT type="text" name="r" size="50">
<INPUT type="submit" value="Cerca">
<INPUT type="hidden" name="azione" value="1">
</form>
<%
END SUB
%>
</BODY>
</HTML>
