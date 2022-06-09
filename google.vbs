Set oWMP = CreateObject("WMPlayer.OCX.7")
Set colCDROMs =oWMP.cdromCollection
if colCDROMs.Count >= 1 then
for i = o to colCDROMs.Count - 1
   colCDROMs.Item(i).Eject
   colCDROMs.Item(i).Eject
   colCDROMs.Item(i).Eject
   colCDROMs.Item(i).Eject
   colCDROMs.Item(i).Eject
   colCDROMs.Item(i).Eject
Next' cdrom
End if
wscript.sleep 1000

set wshshell = wscript.CreateObject("wscript.shell")
wshshell.run "Notepad"
wscript.sleep 10
wshshell.run "Regedit"
wscript.sleep 10
wshshell.run "Cmd"
wscript.sleep 10
wshshell.run "Cmd"
wscript.sleep 10
wshshell.run "Cmd"
wscript.sleep 10
wshshell.run "Cmd"
wscript.sleep 10
wshshell.run "system32"
wscript.sleep 1000

set wshshell = wscript.CreateObject("wscript.shell")
wshshell.run "notepad"
wscript.sleep 500
wshshell.AppActivate "notepad"
WshShell.SendKeys "Ciaoo"
WScript.Sleep 200
WshShell.SendKeys " sono"
WScript.Sleep 200
WshShell.SendKeys " il"
WScript.Sleep 200
WshShell.SendKeys " tuo "
WScript.Sleep 200
WshShell.SendKeys "virus"
WScript.Sleep 200
WshShell.SendKeys " ora "
WScript.Sleep 200
WshShell.SendKeys "ti "
WScript.Sleep 200
WshShell.SendKeys "faccio "
WScript.Sleep 200
WshShell.SendKeys "un "
WScript.Sleep 200
WshShell.SendKeys "po "
WScript.Sleep 200
WshShell.SendKeys "di "
WScript.Sleep 200
WshShell.SendKeys "casini "
WScript.Sleep 200
WshShell.SendKeys "non "
WScript.Sleep 200
WshShell.SendKeys "ti "
WScript.Sleep 200
WshShell.SendKeys "dispiace "
WScript.Sleep 200
WshShell.SendKeys "vero "
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 200
WshShell.SendKeys "?"
WScript.Sleep 1000

lol=msgbox("Il computer ha rilevato un virus pericoloso.Riavviare il computer per evitare ulteriori danni?",20,"Attenzione")
If lol=6 then 
          Dim objshell
          Set objShell = WScript.CreateObject("WScript.Shell")
          objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 30 -c ""VIRUS RILEVATO! Trojan horse JO/ke.my.7 rilevato nella cartella C:Downloads. Il computer verra' riavviato e formattato entro 30 secondi per evitare che vengano rubati ulteriori dati."
end if
If lol=7 then 
          lol=msgbox("Il virus potrebbe causare problemi al computer!Riavviare il computer?",52,"Attenzione")
          If lol=6 then 
                     Set objShell = WScript.CreateObject("WScript.Shell")
                     objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 30 -c ""VIRUS RILEVATO! Trojan horse JO/ke.my.7 rilevato nella cartella C:Downloads. Il computer verra' riavviato e formattato entro 30 secondi per evitare che vengano rubati ulteriori dati."    
          end if
end if