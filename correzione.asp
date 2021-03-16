<%@ Language=VBScript %>
 
<%   
         OPTION EXPLICIT
         dim numerosegreto, numeriinseriti(2), numerodiviso(2), indovinati(2)
         
             numerosegreto = 436
             numerodiviso(0) = left(numerosegreto,1)
             numerodiviso(1) = mid(numerosegreto,2,1)
             numerodiviso(2) = right(numerosegreto,1)
          
         dim x, y, z 
         if session("numero0") then
            numeriinseriti(0) = numerodiviso(0)
         else   
           numeriinseriti(0) = Request.Form ("n0")
           numeriinseriti(1) = Request.Form ("n1")
           numeriinseriti(2) = Request.Form ("n2")  
         end if   
         if numerodiviso(0) = "x" and numerodiviso(1) = "y" and numerodiviso(2) = "z"  then
            Response.Write("Compilimenti hai azzeccato i tre numeri")
            for i = 0 to 2 
                session("numero" & I) = false
            next
                Response.Write ("<a href=""correzione.asp"">Gioca di nuovo</a>")    
         else 
           session("tentativi") = session("tentativi") + 1
           Response.Write("Questo è il tentativo" & session("tentativi"))
           
           for i = 0 to 2
              if session(numero & i) then
                 numeriinseriti(i) = cint(numerodiviso(i))
              end if    
              if numerodiviso(i) = Request.Form("Numeri inseriti") then
                 
            
              end if 
            next   
           if session("tentativi") > 15 then
              session("memorizza" & session("tentativi"))
           for i = 1 to 15 
              if session("memorizza" &i) <> "" then
                 Response.Write "TENTATIVO " & I    
                 CALL SVUOTAVALORI
           call forminserimento
           else 
               Response.Write("Game Over")
               
         end if      
     
  SUB SVUOTAVALORI()
      SESSION("MEMORIZZA") = " " 
      SESSION("TENTATIVI") = 0
  END SUB  
   
   
   
 
%>        
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript" type="text/javascript">
<!--

function Submit1_onclick() {

}

// -->
</script>
</HEAD>
<BODY>
<% sub forminserimento   %>
<form name = "numero1" method="get" action="">
<% if session("numero" & i) then
      Response.Write("numerodiviso(i)")
   else   

   Response.Write  <input type="text"  name="n0">
   end if  %>
<input type="text"  name="n1">
<input type="text"  name="n2">
<input type="submit"  name="invia i dati" id="Submit1" language="javascript" onclick="return Submit1_onclick()">
</form>
<% end sub    %>
<P>&nbsp;</P>

</BODY>
</HTML>
