<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!doctype html%>
<% option explicit
DIM Connessione, RSRegioni, RSProduttori, Stringasql, Regione
%>
<% Regione = Request.QueryString("Regione") %>
<% Set Connessione = Server.CreateObject("ADODB.Connection")
       Connessione.ConnectionString = "Provider = Microsoft.JET.OLEDB.4.0; Data source = " & Server.MapPath('/database/negozio/Vini.mdb')
	   Connessione.Open %>
       <% SET rsRegioni = Server.CreateObject("ADODB.Connection")
	   Stringasql = "Select Regione, IDRegione FROM REGIONI"
	   RSRegioni.Open Stringasql, connessione %>
       <% SET rsProduttori = Server.CreateObject("ADODB.Connection")
	   Stringasql = "Select * from Produttori WHERE IDregione LIKE '" & regione & "'"
	   RSProduttori.Open Stringasql, connessione %>
       
<html>
<head>
<meta charset="utf-8">
<title>Untitled Document</title>
</head>

<body>
<table width = "50%">
<tr>
<td>
<table width="20%" border="1" cellspacing="5" cellpadding="0" valign="TOP" bordercolor="Gray" bgcolor="White">
<tr>
   <td bgcolor="#CCCCCC">
       <div align="center"><b><font COLOR="#0000ff">REGIONE</font></b></div>
   </td>
</tr>
     <% WHILE NOT rsRegioni.EOF %>
        <tr>
          <td bgcolor= "white">
              <div align="center"><font color="#FFFFFF">
                <A href="Vini.asp?regione=<%= RSRegioni"("IDRegione")%>">
				<%= RSRegioni ("Regione")%></font></div></a>
          </td>
        </tr>
        <% RSRegioni.MoveNext
		   Wend %>
</table>
</td>
   <td valign="top">
       <table width="30%" border="1" cellspacing= "5" cellpadding= "0" bordercolor="GRAY" bgcolor="WHITE">
       <tr>
          <td bgcolor="#CCCCCC">
              <div align="Center"><b><font color="#0000FF">PRODUTTORE</font></b></div>
          </td>
       </TR>
       <% WHILE NOT rsProduttori.EOF %>
          <tr>
             <td>
                <p align="center"><%=RSProduttori("Produttore")%></p>
             </td>
          </tr>
          <% RSProduttori.MoveNext
		     WEND  %>
             </table>
   </td>
   </tr>
  </table> 
</body>
</html>
