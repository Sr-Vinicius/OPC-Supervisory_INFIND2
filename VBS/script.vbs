Option Explicit

Dim objwsh
Dim objShell, objCmdExec, commandOutput
Set objShell = CreateObject("WScript.Shell")
objShell.run("C:\python27\python C:\UFU\PROJETOS\PROJETO1\VBS\ReadOPC.py")

Sub Window_OnLoad()
  'This method will be called when the application loads
  'Add your code here
End Sub

Sub Window_OnClose()
	Close Me
End Sub

Sub Window_OnImageClick()
	Msgbox "Mapa de Imagem"
End Sub

Sub sleep(Timesec)
    Set objwsh = CreateObject("WScript.Shell")
    objwsh.Run "Timeout /T " & Timesec & " /nobreak", 0, true
    Set objwsh = Nothing
End Sub

Sub GetValue(valor)
    PV.innerHTML = Cstr(valor)
End Sub

Sub UpdateValue()
    Dim strFileText
    Dim objFileToRead

    While True
      Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("opcvalue.txt",1)
      strFileText = objFileToRead.ReadAll()
      objFileToRead.Close
      Set objFileToRead = Nothing
      GetValue(strFileText)
      sleep(2)
    Wend
End Sub