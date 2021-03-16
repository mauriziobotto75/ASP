<%@language="vbscript"%>

<% 
Option explicit
Dim StringaSQL
Dim Connessione,Recordset

'SET è il comando o istruzione ASP/VISUAL BASIC che setta/imposta un oggetto

SET CONNESSIONE = Server.CreateObject("ADODB.Connection")

'L'oggetto server (ASP) crea un oggettto utilizzando il componente ADO (che sta per ACTIVEX DATA OBJECT)
'ADO è un componente molto ricco e la sua documentazione si trova sulla libreria MICROSOFT MSDN (cercare ADO Reference)

Connessione.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =" & Server.MapPath("/database/db.mdb")

'Apriamo l'oggetto connessione appena creato
CONNESSIONE.Open

'Creiamo l'oggetto che chiamiamo a titolo d'esempio RECORDSET
'L'oggetto Recordset conterrà tutti i dati estrapolati dal database e se doveste immaginarvelo pensatelo fisicamente simile ad una tabella
SET RECORDSET = Server.CreateObject("ADODB.Recordset")

'Popoliamo il Recordset che è stato creato e che per ora risulta VUOTO

'Per popolarlo abbiamo bisogno dell'oggetto connessione sopra creato
'Seconda cosa di cui abbiamo bisogno è una stringa SQL che interroghi il database
StringaSQL="SELECT * FROM SEZIONI"

'Apriamo l'oggetto Recordset e lo popoliamo grazie alla connessione e alla stringa SQL

Recordset.Open StringaSQL , connessione, ,adLockPessimistic 

'Visualizziamo i dati del Recordset
 WHILE NOT Recordset.EOF
%>
    <a href="sezioni.asp?sez=<%=recordset("Nomesezione")%>">
    <%=RECORDSET("Nomesezione")%></a><br>
    <%
      Recordset.MoveNext
 wend

'Quando le operazioni sul db sono concluse si chiude l'oggetto connessione 
Recordset.Close
CONNESSIONE.Close

'Si annullano gli oggetti creati
SET RECORDSET = NOTHING
SET CONNESSIONE = NOTHING

'Da qui in poi riapriamo la connessione creiamo un altro Recordset
'Creiamo una nuova StringaSQL e visualizziamo i contenuti
dim Rscontenuti,StringaSQL2
SET CONNESSIONE = Server.CreateObject("ADODB.Connection")
Connessione.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;DATA SOURCE =" & Server.MapPath("/database/db.mdb")
CONNESSIONE.Open
SET Rscontenuti = Server.CreateObject("ADODB.Recordset")
if Request.QueryString("sez")<>"" then

   dim filtro
       filtro = Request.QueryString("sez")


       StringaSQL2="SELECT * FROM SEZIONI where Nomesezione='"&filtro&"'"
       Rscontenuti.Open StringaSQL2 , Connessione



       Response.Write rscontenuti("Contenuto")


'Chiudiamo connessione e recordset e li annulliamo
       Rscontenuti.Close
       Connessione.close

SET Rscontenuti = NOTHING
SET Connessione = NOTHING
end if
%> 