Imports System

Module Program
    Sub Main(args As String())
        Dim userSelectedOption As String = ""

        Do While userSelectedOption <> "exit"
            Console.Write("Welke afbeeldingen wilt u importeren? ")
            userSelectedOption = Console.ReadLine()

            Select Case userSelectedOption
                Case "YouVersion"
                    YouVersion.ImportImages()
                Case "Dagelijkse Broodkruimels"
                    DagelijkseBroodkruimels.ImportImages()
                Case "Vers van de dag"
                    VersVanDeDag.ImportImages()
                Case "exit"
                    Exit Do
                Case Else
                    Console.WriteLine("Geen geldige invoer.")
            End Select
        Loop
    End Sub
End Module
