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


Sub UpdateToScreen(valor)
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
      UpdateToScreen(strFileText)
      UpdateToDB(strFileText)
      sleep(2)
    Wend
End Sub


Sub UpdateToDB(valor)
    Dim Connection, Command, SQL

    SQL = "INSERT INTO POINTVALUES (VALOR, TIMESTAMP1) VALUES (" & valor & ", #" & Now & "#)"

    Set Connection = CreateObject("ADODB.Connection")
    Set Command = CreateObject("ADODB.Command")

    Connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.\database.mdb;Persist Security Info=False"

    Command.ActiveConnection = Connection
    Command.CommandText = SQL
    Command.Execute

    Set Command = Nothing

    Connection.Close
End Sub


Sub ClearDB()
    Dim Connection, Command, SQL

    SQL = "DELETE FROM POINTVALUES"

    Set Connection = CreateObject("ADODB.Connection")
    Set Command = CreateObject("ADODB.Command")

    Connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.\database.mdb;Persist Security Info=False"

    Command.ActiveConnection = Connection
    Command.CommandText = SQL
    Command.Execute

    Set Command = Nothing

    Connection.Close
End Sub


Sub Read()
    Dim Connection, RecordSet, SQL
    
    SQL = "SELECT * FROM POINTVALUES"

    Set Connection = CreateObject("ADODB.Connection")
    Set RecordSet = CreateObject("ADODB.RecordSet")

    'Abrir a conexao usando oledb
    Connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.\database.mdb;Persist Security Info=False"

    RecordSet.Open SQL, Connection

    If RecordSet.EOF Then
        MsgBox "Nenhum registro encontrado!"
    Else
        Do While Not RecordSet.EOF
            MsgBox RecordSet("Valor") & " - " & RecordSet("Timestamp1")
            RecordSet.MoveNext
        Loop
    End If

    RecordSet.Close
    Set RecordSet = Nothing

    Connection.Close
    Set Connection = Nothing

    MsgBox "Fim!"
End Sub
