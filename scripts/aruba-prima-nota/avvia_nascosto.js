// DR Portal - lancia un .bat SENZA finestra.
// L'Utilita' di pianificazione, con un'attivita' che gira "solo quando
// l'utente e' connesso", apre a ogni corsa la finestra nera del prompt
// (ogni 10 minuti per il giro su richiesta). Facendo partire il .bat da qui,
// tramite wscript, la finestra non compare. Uso nelle attivita':
//   C:\Windows\System32\wscript.exe //B //Nologo "<cartella>\avvia_nascosto.js" "<cartella>\controlla_giro.bat"
// Aspetta la fine del .bat, cosi' l'attivita' risulta "in esecuzione" finche'
// dura e il limite di 2 ore dell'attivita' vale davvero. Il codice di uscita
// e' quello del .bat. Verificato il 28/10/2026 con Smart App Control acceso.
var sh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");
if (WScript.Arguments.Count() < 1) {
  WScript.Echo("uso: avvia_nascosto.js <percorso del .bat>");
  WScript.Quit(2);
}
var bat = WScript.Arguments(0);
if (!fso.FileExists(bat)) {
  WScript.Echo("file non trovato: " + bat);
  WScript.Quit(2);
}
sh.CurrentDirectory = fso.GetParentFolderName(bat);
// Doppie virgolette esterne: cmd /c toglie la prima e l'ultima, resta il
// percorso tra virgolette anche se contiene spazi.
var codice = sh.Run('cmd.exe /c ""' + bat + '""', 0, true);
WScript.Quit(codice);
