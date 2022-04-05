'This was written on 9/11... "OH GOSH!"
'"Please Don'T Hurt me>>>>>..>>>!>>" - Kait Bell
'lines 4-11 and part of 15 was Kompletely Kopied from Michael Salsbury of Stack Overflow. Thank christ for stack overflow. without it I would never know how to do this stuff.
Function getCommandOutput(theCommand)

    Dim objShell, objCmdExec
    Set objShell = CreateObject("WScript.Shell")
    Set objCmdExec = objShell.exec(theCommand)
    getCommandOutput = objCmdExec.StdOut.ReadAll

end Function
Dim ytlink, thisDir, again
ytlink = InputBox("Please insert the link to the YouTube video") & " --audio-format mp3"
Set objShell = CreateObject("WScript.Shell")
thisDir = getCommandOutput("cmd.exe /c youtube-dl -x " & ytlink)
'MsgBox thisDir 'uncomment for special "ADMIN MODE", aka "IT JUST SHOWS YOU WHAT THE PROGRAM IS DOING AND NOTHING SPECIAL MODE"
MsgBox "Check the folder to verify your file downloaded. If it didn't, check that you pasted the link correctly."
Set objShell = Nothing
