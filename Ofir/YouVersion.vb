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

            'Get HTML files
            Dim htmlFileNames As List(Of String) = New List(Of String)
            For Each fileName In Directory.EnumerateFiles(sourceFolder)
                If fileName.Split(".").Last() = "html" Then htmlFileNames.Add(fileName)
            Next

            'Load HTML documents
            Dim htmlDocumentCollection As Dictionary(Of String, HtmlDocument) = New Dictionary(Of String, HtmlDocument)
            For Each htmlFileName In htmlFileNames
                Dim htmlDoc As HtmlDocument = New HtmlDocument()
                htmlDoc.Load(htmlFileName)
                htmlDocumentCollection.Add(htmlFileName, htmlDoc)
            Next

            'Find all img-tag nodes
            'TODO: Check for the "votd-image" class
            Dim imageTagNodeCollection As Dictionary(Of String, List(Of HtmlNode)) = New Dictionary(Of String, List(Of HtmlNode))
            For Each htmlDocumentEntry In htmlDocumentCollection
                imageTagNodeCollection.Add(htmlDocumentEntry.Key, htmlDocumentEntry.Value.DocumentNode.SelectNodes(XpathImg).ToList())
            Next

            'Select all image URLs not on the blacklist
            Dim importCollection As Dictionary(Of String, String) = New Dictionary(Of String, String)

            For Each imageTagNodeEntry In imageTagNodeCollection
                For Each node In imageTagNodeEntry.Value
                    For Each imageSource In node.Attributes.AttributesWithName("src")
                        Dim imageUrl = imageSource.Value
                        Dim imageFileName As String = imageUrl.Split("/").Last()

                        If imageBlackList.Contains(imageFileName) Then Continue For

                        imageUrl = imageUrl.Replace("320x320", "1280x1280")
                        imageUrl = imageUrl.Replace("640x640", "1280x1280")
                        importCollection.Add(imageTagNodeEntry.Key, imageUrl)
                    Next
                Next
            Next

            'Download the file at eacht URL
            For Each importable In importCollection
                Dim fileNameParts As String() = importable.Key.Split(" - ")
                Dim mailTitle As String = ""
                Dim bibleVerse As String = ""
                Dim mailDate As String = fileNameParts.GetValue(0)
                If fileNameParts.Length = 3 Then
                    mailTitle = fileNameParts.GetValue(1)
                    bibleVerse = fileNameParts.GetValue(2)
                End If

                'Download and store the file in de destination folder
                Dim Client As New WebClient
                If bibleVerse <> "" Then
                    Console.WriteLine("Downloading {0} to {1}", importable.Value, String.Concat(sourceFolder, "\", bibleVerse, ".jpg").Replace(".html", ""))
                    Client.DownloadFile(importable.Value, String.Concat(sourceFolder, "\", bibleVerse, ".jpg").Replace(".html", ""))
                Else
                    Console.WriteLine("Downloading {0} to {1}", importable.Value, String.Concat(mailDate, ".jpg").Replace(".html", ""))
                    Client.DownloadFile(importable.Value, String.Concat(mailDate, ".jpg").Replace(".html", ""))
                End If
                Client.Dispose()
            Next

            'TODO: Remove the HTML file

        Else
            Console.WriteLine("Locatie {0} bestaat niet", sourceFolder)
        End If

        Console.WriteLine("YouVersion images are downloaded.")
    End Sub
End Class
