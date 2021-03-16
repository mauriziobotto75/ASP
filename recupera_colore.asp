<%@ Language=VBScript %>
<%  Request.Cookies("colore")
    if Request.Cookies("colore") = "rosso" then
       sfondo = "rosso"
    elseif Request.Cookies("colore") = "giallo" then
       sfondo = "giallo"
    elseif Request.Cookies("colore") = "blu"  then
       sfondo = "blu"
    end if       
    Response.Redirect "colore.asp"
 %>   
                                    

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
