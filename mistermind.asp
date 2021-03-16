<%@ Language=VBScript %>
<%Option explicit
'CREAZIONE DELLE VARIABILI
DIM numerosegreto,numeriinseriti(2),numerodiviso(2)
DIM indovinati(2), i
' IMPOSTAZIONE DEL NUMERO DA INDOVINARE
numerosegreto = 438
' SUDDIVIDIAMO IL NUMERO SEGRETO
' E LO MEMORIZIAMO IN UN ARRAY DI INDICE 2
' PER LA DIVISIONE UTILIZZIAMO DELLE FUNZIONI
' DI VB E TRASFORMIAMO IL NUMERO CON LA 
' FUNZIONE CINT IN UN INTERO

' LEFT(STRINGA,NUMERO CARATTERI): 
'RECUPERA UN NUMERO DI CARATTERI
' DA STRINGA PARTENDO DA SINISTRA
numerodiviso(0) = cint(left(numerosegreto,1))
' mid(STRINGA,START,NUMERO CARATTERI): 
' RECUPERA UN NUMERO DI CARATTERI
' DA STRINGA PARTENDO 
' DAL CARATTERE START
numerodiviso(1) = cint(mid(numerosegreto,2,1))
' right(STRINGA,NUMERO CARATTERI): 
'RECUPERA UN NUMERO DI CARATTERI
' DA STRINGA PARTENDO DA DESTRA
numerodiviso(2) = cint(right(numerosegreto,1))
' SI ESEGUE UN CICLO DA 0 A 2
' i NEL CICLO ASSUME I VALORI DI 0,1,2
FOR i = 0 to 2
	' CONTROLLIAMO SE LA VARIABILE DI
	' SESSIONE NUMERO -0-1-2
	IF session("numero"&i) THEN
		' SE IL VALORE DI QUESTA VARIABILE
		' è TRUE ALLORA IMPOSTIAMO
		' L'INDICE (0-1-2) DELL'ARRAY
		' NUMERI INSERITI COME IL VALORE DEL
		' DELLA ARRAY NUMERODIVISO
		'(LA VARIABILE DI SESSIONE NUMERO
		' VIENE IMPOSTATA A TRUE QUANDO 
		'L'UTENTE INDOVINA IL NUMERO
		numeriinseriti(i) = cint(numerodiviso(i))
	ELSE
		' SE LA VARIABILE DI SESSIONE è FALSE
		' ALLORA RECUPERIAMO IL VALORE DAL
		' FORM, (L'UTENTE NON HA ANCORA 
		' INDOVINATA)
		numeriinseriti(i) = Request.Form ("n"&i)
	END IF
	' FACCIAMO UN COLTROLLO SUL
	'VALORE DI NUMERIINSERITI, SE IL
	' VALORE NON è NUMERICO, LO IMPOSTIAMO
	' A 0
	IF not isnumeric(numeriinseriti(i)) THEN
		numeriinseriti(i) = 0
	ELSE
		numeriinseriti(i) = cint(numeriinseriti(i))
	END IF
NEXT
'VERIFICHIAMO LA CORRISPONDENZA
'TRA I NUMERI INSERITI E I NUMERI DEL
' NUMERO SEGRETO (MEMORIZZATI IN NUMERODIVISO)
IF numerodiviso(0) = numeriinseriti(0) AND numerodiviso(1) = numeriinseriti(1) AND numerodiviso(2) = numeriinseriti(2) THEN
	' SE I NUMERI INSERITI CORRISPONDONO
	' A QUELLI SEGRETI, L'UTENTE VINCE
	' E VIENE IMFORMATO
	Response.Write ("COMPLIMENTI, HAI VINTO!!")
	' VIENE CHIAMATA LA SUB SVUOTAVALORI
	call svuotavalori
ELSE ' SE NON HO INDOVINATO TUTTI 
	' E TRE I VALORI VERIFICO SE NE HO
	' INDOVINATO UNO DEI TRE
	FOR i = 0 TO 2
		' SE numerodiviso(i) = numeriinseriti(i)
		' HANNO LO STESSO VALORE
		' AVVISO L'UTENTE CHE HA INDOVINATO
		' UN NUMERO E MEMORIZZO
		'CHE HA INDOVINATO QUEL NUMERO
		'IMPOSTANDO NELLA VARIABILE DI
		'SESSIONE NUMERO 0-1-2 A true
		IF numerodiviso(i) = numeriinseriti(i) THEN
			Response.Write ("<p>Hai indivinato")&_
			(" il numero: " & numerodiviso(i) &"</p>")
			session("numero"&i) = true
		END IF	
	NEXT
	' INCREMENTO LA VARIABILE DI SESSIONE
	' TENTATIVI DI UNO
	session("tentativi") = session("tentativi") + 1
	Response.Write ("<p>Questo è il tentativo numero"& session("tentativi")  &"</p>")
	' SE I TENTATIVI NON SUPERANO IL 
	'VALORE DI 15
	IF session("tentativi") < 15 THEN
		' MEMORIZZO I VALORI INSERITI
		'DALL'UTENTE IN UNA VARIABILE
		'DI SESSIONE MEMORIZZA 0,1,2..15
		session("memorizza"&session("tentativi")) = numeriinseriti(0) & numeriinseriti(1) & numeriinseriti(2)
		' VISULIZZO I VALORI DELLA VARIABILE
		' MEMORIZZA 0,1,2..15
		FOR i = 0 to 15
			' SE LA VARIABILE è DIVERSA DA VUOTO
			IF session("memorizza"&i) <> "" THEN
				' GLI FACCIO SCRIVERE I VALORI
				Response.Write "<p>Tentativo "& i & ":"& session("memorizza"&i) &"</p>" 
			END IF
		NEXT
		' RICHIAMO LA SUB FORMINSERIMENTO
		CALL formInserimento
	ELSE
		' SE SESSION("TENTATIVI") è UGUALE
		' O MAGGIORE DI 15
		' DICO ALL'UTENTE CHE HA FINITO
		' IL GIOCO
		Response.Write ("HAI FINITO I TENTATIVI. GAME OVER")
		'CHIAMO LA SUB svuotavalori
		call svuotavalori
	END IF
END IF
SUB svuotavalori
	' IMPOSTA TUTTI I VALORI DELLE
	' VARIABILI DI SESSIONE COME
	' SE NON SI FOSSE ANCORA GIOCATO
	FOR i = 0 TO 2
		session("numero"&i) = false
	NEXT
	FOR i = 1 to 15
		session("memorizza"&i) = ""
	NEXT
	session("tentativi") = 0
	FOR i = 0 TO 2
		session("numero"&i) = false
	NEXT
	' VISUALIZZA UN LINK PER GIOCARE DI NUOVO
	Response.Write ("<p><a href=""mistermind.asp"">Gioca di nuovo</a></p>")
END SUB
SUB formInserimento
	' FORM
	Response.Write ("<form method=""post"" action=""mistermind.asp"">")
	' CICLO DA 0 A 2
	FOR i = 0 to 2 
		' VERIFICO SE LA VARIABILE SESSION
		' NUMERO è UGUALE A TRUE,
		' SE SI VISUALIZZO IL NUMERO 
		IF session("numero"&i) THEN
			Response.Write numerodiviso(i)
		ELSE
		' ALTRIMENTI CREO UN INPUT TEXT
		' CHE PERMETTE ALL'UTENTE
		' DI INSERNE UNO LUI PER INDOVINARE
			Response.Write  ("<input type=""text"" name=""n"& i &"""/>")
		END IF
	NEXT
	Response.Write  ("<input type=""submit"" value=""indovina"" />")
	Response.Write  ("</form>")
END SUB
%>

