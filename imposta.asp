<%@ Language=VBScript %>
<%Response.Cookies("user")("name") = "Maurizio"
Response.Cookies("user")("surname") = "Botto"
Response.Cookies("user").expires =#01/03/2005#
Response.Cookies("user").path = "/"
' recupero valore dei cookies
dim a
a = Request.Cookies("user")("name")
a = a& "." & Request.Cookies ("user")("surname")
Response.Write a
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
