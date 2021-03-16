<%@ Language=VBScript  %>
<%   option explicit

DIM CONNESSIONE
DIM RSMENU
DIM RSCATALOGO
DIM RSDETTAGLI
DIM STRINGASQL
DIM CAT
DIM QUAN
DIM DET

CAT = Request.QueryString ("CAT")
QUAN = Request.QueryString ("QUAN")
DET = Request.QueryString ("DET")

SET CONNESSIONE = Server.CreateObject("ADODB.Connection")
Connessione.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =" & Server.MapPath("ES04.mdb")
CONNESSIONE.Open
SET RSmenu = Server.CreateObject("ADODB.Recordset")
STRINGASQL="SELECT * FROM PRODOTTI WHERE menu LIKE '"& cat &"'" & """""
 
SET RSDETTAGLI=server.CreateObject("ADODB.Recordset")
STRINGASQL="SELECT * FROM PRODOTTI WHERE NOMEPRODOTTO like '"& det &"'" & """""
Rsdettagli.Open Stringasql,connessione
 

SET RSCATALOGO = Server.CreateObject("ADODB.Recordset")
STRINGASQL="SELECT * FROM PRODOTTI WHERE CATEGORIA LIKE '"& cat &"'" & """""
RSCATALOGO.Open stringasql, connessione

 

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<TABLE WIDTH="90%" HEIGHT="100%" CELLPADDING="0" CELLSPACING="5">
<TR>
<TD WIDTH="20%" VALIGN="TOP">
<table width="100%" border="0" cellspacing="5" cellpadding="0" bordercolor="#0000FF" bgcolor="#CCCCCC">
        <tr> 
          <td bgcolor="#CCCCCC"> 
            <div align="center"><b><font color="#0000FF">MENU</font></b></div>
          </td>
        </tr>
        <% WHILE NOT RSMENU.EOF %>
        <tr> 
          <td bgcolor="orange"> 
            <div align="center"><font color="#FFFFFF">
            <A HREF="ES04.asp?cat=<%= RSMENU("CATEGORIA")%>">
            <%= RSMENU("CATEGORIA")%></font></div></a>
          </td>
        </tr>
        <% RSMENU.MoveNext
           WEND %>
        </table>
    </TD>


<TD WIDTH="40%" VALIGN="TOP">
<table width="100%" border="0" cellspacing="5" cellpadding="0" bordercolor="#CCCCCC" bgcolor="#CCCCCC">
        <tr bgcolor="#CCCCCC"> 
          <td valign="top" colspan="2"> 
            <div align="center"><font color="#FF0000"><b><font color="#0000FF">CATALOGO</font></b></font></div>
          </td>
        </tr>
       <% WHILE NOT rscatalogo.EOF %>
        <tr> 
          <td width="46%" height="129" bgcolor="#FFFFFF"> 
            <p align="center"><%=RSCATALOGO("NOMEPRODOTTO")%> </p>
            <p align="center">Prezzo: <%=RSCATALOGO("PREZZO")%> € </p>
          </td>
          <td width="54%" height="129" bgcolor="#CCCCCC"> 
          <p align="center"><img src="immagini/<%=rscatalogo("foto")%>"></p>
   <A HREF="ES04.asp?cat=<%=cat%>&det=<%=Rscatalogo("nomeprodotto")%>">Vedi dettaglio</A></td>
        </tr>
        <%Rscatalogo.MoveNext
        WEND%>

      </table>
    </td>
    
<TD WIDTH="40%" VALIGN="TOP">
<div valign=top align="center"><table width="100%" border="0" cellspacing="5" cellpadding="0" bordercolor="#0000FF" bgcolor="#CCCCCC">
        <tr bgcolor="#CCCCCC"> 
        <td colspan="2"><div align="center"><b><font color="#0000FF">DETTAGLI</font></b></div></td>
        </tr>
        <tr> 
        
        
        <%if det = "" then
        
        
          else%>  
          <td bgcolor="white">
          <div align="center"><font color="black">
          <% Response.Write rsdettagli("nomeprodotto")%>
          </font></div></td>
          <td><div align="center">
          <img src=immagini/<%=rsdettagli("foto")%>></div></td>
        </tr>
        <tr bgcolor="white"> 
          <td colspan="2">
              <div align="center"><font color="black">
            <%=Rsdettagli("dettagli")%>
            </font></div></td>
             </tr></font>
        <tr>
        <td>
        <form method="get" action="carrello.asp">
        <input type=hidden name="prezzo" value=<%=Rsdettagli("Prezzo")%>>
		<input type=hidden name="carrello" value=<%=Rsdettagli("Codiceprodotto")%>>
		<input type=hidden name="cat" value=<%=Request.QueryString ("cat")%>>
		<input type=hidden name="det" value=<%=Request.QueryString ("det")%>>
		<input type=hidden name="dettagli" value=<%=Rsdettagli("dettagli")%>>
		Quantità:<INPUT type="text" name="quan" size="3" maxlength="3">&nbsp
		<INPUT type="submit" value="Invia al carrello">
        </form>
        </td>
        </tr>
       <%end if%>
       </table>
      </div>
      </td>
      </table>
      
    
    
</BODY>
</HTML>



<%

RSMENU.Close
SET RSMENU = NOTHING

RSCATALOGO.Close
SET RSCATALOGO = NOTHING

RSDETTAGLI.Close
SET RSDETTAGLI = NOTHING

CONNESSIONE.Close
SET CONNESSIONE = NOTHING

%>