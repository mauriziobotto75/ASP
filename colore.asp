<%@ Language=VBScript %>
<%                    %>
<% Response.Cookies("colore")  %>
<% if Response.Cookies("colore") = " " then
      Response.Write("Devi scegliere un colore")
   end if                      %>      
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor="<%if Response.Cookies <> " " then response.Cookies("colore")%>">
                   
<form name="colore1" method="get" action="recupera_colore.asp">
<p>Scegli il colore della pagina che vuoi impostare fino al 2005</P>
GIALLO<input type="radio"  name="colore"  value="giallo"><br>
ROSSO<input type="radio"  name="colore"  value="rosso"><br>
BLU<input type="radio"  name="colore"  value="blu"><br>
<input type="submit" value="scegli il colore">
</form>
<P>&nbsp;</P>
</form>
</BODY>
</HTML>
