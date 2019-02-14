Imports System.IO
Imports System.Net
Imports System.Xml
Imports HtmlAgilityPack

Public Class YouVersion
    Private Const folderLocation As String = "YouVersion"
    Private Const XpathImg As [String] = "//img"

    Shared Sub ImportImages()
        Dim sourceFolder As String = String.Concat(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "\", folderLocation)
        Dim imageBlackList As String() = {"twitter.png", "facebook.png", "new-sun-2x.png", "logo.png", "instagram.png", "blog.png", "app-icon-nl.png", "320x320.jpg"}

        'Check if the directory exists
        If Directory.Exists(sourceFolder) Then
            Console.WriteLine("Afbeeldingen importeren vanuit {0}", sourceFolder)

            'Loop through the files
            For Each fileName In Directory.EnumerateFiles(sourceFolder)
                If fileName.Split(".").Last() <> "html" Then Continue For

                'Parse each HTML file
                Dim htmlDoc As HtmlDocument = New HtmlDocument()
                htmlDoc.Load(fileName)

                'Find the img-tags
                Dim nodes = htmlDoc.DocumentNode.SelectNodes(XpathImg).ToList()

                For Each node In nodes
                    'Get each URL
                    For Each imageSource In node.Attributes.AttributesWithName("src")
                        'Find the correct URL
                        Dim imageUrl As String = imageSource.Value
                        Dim imageFileName As String = imageUrl.Split("/").Last()

                        If imageBlackList.Contains(imageFileName) Then Continue For

                        Dim fileNameParts As String() = fileName.Split(" - ")
                        Dim mailTitle As String = ""
                        Dim bibleVerse As String = ""
                        Dim mailDate As String = fileNameParts.GetValue(0)
                        If fileNameParts.Length = 3 Then
                            mailTitle = fileNameParts.GetValue(1)
                            bibleVerse = fileNameParts.GetValue(2)
                        End If

                        Console.WriteLine(imageUrl)
                        'Get it at a large resolution
                        Dim highResolutionImageFileUrl As String = imageUrl.Replace("320x320", "1280x1280")
                        highResolutionImageFileUrl = highResolutionImageFileUrl.Replace("640x640", "1280x1280")

                        'Download and store the file in de destination folder
                        Dim Client As New WebClient
                        If bibleVerse <> "" Then
                            Client.DownloadFile(highResolutionImageFileUrl, String.Concat(sourceFolder, "\", bibleVerse, ".jpg").Replace(".html", ""))
                        Else
                            Client.DownloadFile(highResolutionImageFileUrl, String.Concat(mailDate, ".jpg").Replace(".html", ""))
                        End If
                        Client.Dispose()
                    Next
                Next

                'TODO: Remove the HTML file

            Next
        Else
            Console.WriteLine("Locatie {0} bestaat niet", sourceFolder)
        End If

        Console.WriteLine("YouVersion images are downloaded.")
    End Sub
End Class
