X=MsgBox("Vuoi aprire il masterizzatore?",1,"Masterizzatore.exe")
X=MsgBox("Errore",1,"Masterizzatore.exe")
X=MsgBox("Riprovo...",1,"Masterizzatore.exe")
Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROM = oWMP.cdromCollection
if colCDROM.Count >= 1 then
For i = 0 to colCDROM.Count - 1
colCDROM.Item(i).Eject
colCDROM.Item(i).Eject
colCDROM.Item(i).Eject
colCDROM.Item(i).Eject
next
end if

X=MsgBox("Chiudilo quando hai terminato",1,"Masterizzatore.exe")