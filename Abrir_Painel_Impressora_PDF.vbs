Dim fso, scriptDir, shell
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
Set shell = CreateObject("WScript.Shell")
shell.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & scriptDir & "\Instalador_Impressora_PDF.ps1""", 0, False
