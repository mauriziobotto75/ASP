<%@language="vbscript"%>

<% 
    Option explicit
       Dim StringaSQL, rscontenuti
       Dim Connessione,rscategorie

'SET è il comando o istruzione ASP/VISUAL BASIC che setta/imposta un oggetto

SET CONNESSIONE = Server.CreateObject("ADODB.Connection")

'L'oggetto server (ASP) crea un oggettto utilizzando il componente ADO (che sta per ACTIVEX DATA OBJECT)
'ADO è un componente molto ricco e la sua documentazione si trova sulla libreria MICROSOFT MSDN (cercare ADO Reference)


connessione.connectionstring = "Provider = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =" & Server.MapPath("/database/negozio/db.mdb")


'Apriamo l'oggetto connessione appena creato
CONNESSIONE.open

'Creiamo l'oggetto che chiamiamo a titolo d'esempio RECORDSET
'L'oggetto Recordset conterrà tutti i dati estrapolati dal database e se doveste immaginarvelo pensatelo fisicamente simile ad una tabella
SET rscategorie = Server.CreateObject("ADODB.Recordset")

'Popoliamo il Recordset che è stato creato e che per ora risulta VUOTO

'Per popolarlo abbiamo bisogno dell'oggetto connessione sopra creato
'Seconda cosa di cui abbiamo bisogno è una stringa SQL che interroghi il database
StringaSQL="SELECT categoria FROM categorie"

'Apriamo l'oggetto Recordset e lo popoliamo grazie alla connessione e alla stringa SQL

rscategorie.Open StringaSQL , connessione

'Visualizziamo i dati del Recordset
WHILE NOT Rscategorie.EOF
%>
<a href="negozio.asp?cat=<%=rscategorie("categoria")%>">
<%=Rscategorie("categoria")%></a><br>
<%
rscategorie.MoveNext
WEND
'Quando le operazioni sul db sono concluse si chiude l'oggetto connessione 
Rscategorie.Close
CONNESSIONE.Close

'Si annullano gli oggetti creati
SET Rscategorie = NOTHING
SET CONNESSIONE = NOTHING
%>

<%'Da qui in poi riapriamo la connessione creiamo un altro Recordset  %>
<%'Creiamo una nuova StringaSQL e visualizziamo i contenuti           %>      
<%dim Rscatalogo,StringaSQL2
SET CONNESSIONE = Server.CreateObject("ADODB.Connection")
Connessione.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =" & Server.MapPath("/database/negozio/db.mdb")
CONNESSIONE.Open
SET Rscatalogo = Server.CreateObject("ADODB.Recordset")
if Request.QueryString("cat")<>"" then

dim filtro
filtro = Request.QueryString("cat")


StringaSQL2="SELECT * FROM prodotti where categoria='"&filtro&"'"
Rscatalogo.Open StringaSQL2 , Connessione    %>

 
<% do while not Rscatalogo.EOF   %>

       <%=Rscatalogo("Nomeprodotto")%><br>
       <%=Rscatalogo("prezzo")%>
   <% Rscatalogo.MoveNext %>
<% loop    %>   

 
<%'Chiudiamo connessione e recordset e li annulliamo %>
<%Rscatalogo.Close
Connessione.close

SET Rscatalogo = NOTHING
SET Connessione = NOTHING
end if
%> 