Attribute VB_Name = "Module1"
Option Explicit
Dim n_record As Integer
Dim nfile As Integer
Public intnumerorecord As Integer
Public intindice As Integer
'Dim sql As String
Type myrecord
    cod_art As Long
    prezzo As Currency
    sconto As Currency
    
End Type
Public campi As myrecord

Sub main()

On Error GoTo err
Dim inserisci As Boolean
nfile = FreeFile
 
 inserisci = True
    Open App.Path & "\scontotris.dat" For Random As nfile Len = Len(campi)
' Chiude prima di riaprire in una modalità diversa.Close #1
             
       Do While inserisci = True
          campi.cod_art = InputBox("Inserisci codice articolo")
          campi.prezzo = InputBox("Inserisci il prezzo")
          
          campi.sconto = InputBox("Inserisci lo sconto")
          aggiungi
          
          If MsgBox("Vuoi inserire un altro record?", vbOKCancel) = vbCancel Then
          
          inserisci = False
          End If
        Loop
        
Close nfile
 Exit Sub

err:
 MsgBox err.Description
 End Sub

Private Sub aggiungi()

intnumerorecord = LOF(nfile) / Len(campi)
Put #nfile, intnumerorecord + 1, campi

End Sub
 

