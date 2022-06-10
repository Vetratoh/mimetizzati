Dim msg, sapi

msg=InputBox("Inserisci una parola o frase da dire al PC","PL Speaker")

Set sapi=CreateObject("sapi.spvoice")

sapi.Speak msg