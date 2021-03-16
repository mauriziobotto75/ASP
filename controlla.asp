<%@ Language=VBScript %>
<% if Request.Cookies("a") = " "   then 
         
      Response.Write "Il tuo browser non accetta i cookies"
   else
      Response.Write "Il tuo browser accetta i cookies"
   end if 
%>         
 
 
 
 
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
