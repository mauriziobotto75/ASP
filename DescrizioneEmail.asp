<p>&lt;%@LANGUAGE=&quot;VBSCRIPT&quot; CODEPAGE=&quot;65001&quot;%&gt;<br />
  &lt;!DOCTYPE html PUBLIC &quot;-//W3C//DTD XHTML 1.0 Transitional//EN&quot; &quot;http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd&quot;&gt;<br />
  &lt;html xmlns=&quot;http://www.w3.org/1999/xhtml&quot;&gt;<br />
  &lt;head&gt;<br />
  &lt;meta http-equiv=&quot;Content-Type&quot; content=&quot;text/html; charset=utf-8&quot; /&gt;<br />
  &lt;title&gt;Untitled Document&lt;/title&gt;<br />
  &lt;/head&gt;</p>
<p>&lt;body&gt;<br />
  &lt;% Dim nome, cognome, indirizzo, email, commento, con, rs, sql, utente<br />
  Dim errore, errore1, errore2<br />
  NOME = Request.QueryString(&quot;NOME&quot;)<br />
  IF nome = &quot;&quot; then<br />
  errore1 = &quot;Devi inserire il nome&quot;<br />
  errore = errore1<br />
  End if<br />
  COGNOME = Request.QueryString(&quot;COGNOME&quot;)<br />
  IF COGNOME = &quot;&quot; THEN<br />
  ERRORE2 = &quot;Devi inserire il cognome&quot;<br />
  errore = errore2<br />
  End if<br />
  INDIRIZZO = Request.QueryString(&quot;INDIRIZZO&quot;)<br />
  EMAIL     = Request.QueryString(&quot;Email&quot;)<br />
  COMMENTO  = Request.QueryString(&quot;Commento&quot;)<br />
  set con = SERVER.CreateObject(&quot;ADODB.CONNECTION&quot;)<br />
  con.ConnectionString = &quot;Provider = Microsoft.JET.OLEDB.4.0; data source = &quot; &amp; Server.MapPath(&quot;/NEGOZIO/gestbook.mdb&quot;)<br />
  con.Open <br />
  set rs = Server.CreateObject(&quot;adodb.recordset&quot;)<br />
  IF ERRORE = &quot;&quot; THEN<br />
  sql = &quot;Select * from gestbook where 1 &lt;&gt; 1&quot;<br />
  rs.Open sql, con, , 3<br />
  rs.AddNew<br />
  rs(&quot;Nome&quot;) = NOME<br />
  rs(&quot;Cognome&quot;) = COGNOME<br />
  rs(&quot;Indirizzo&quot;) = INDIRIZZO<br />
  rs(&quot;EMAIL&quot;) = EMAIL<br />
  rs(&quot;COMMENTO&quot;) = COMMENTO <br />
  rs.Update<br />
  rs.Close<br />
  ELSE      %&gt;<br />
  <br />
  &lt;tr&gt;<br />
  &lt;td&gt;&lt;%IF ERRORE1 &lt;&gt; &quot;&quot; THEN  %&gt;<br />
  &lt;% RESPONSE.Write(ERRORE1)  %&gt;<br />
  &lt;%ELSEIF ERRORE2 &lt;&gt; &quot;&quot;  THEN<br />
  RESPONSE.Write(ERRORE2)  %&gt;<br />
  &lt;%ENDIF     %&gt;&lt;/td&gt;<br />
  &lt;/tr&gt;<br />
  &lt;% END IF %&gt;<br />
  <br />
  &lt;/body&gt;<br />
  &lt;/html&gt;</p>
