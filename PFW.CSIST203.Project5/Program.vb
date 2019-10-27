Public Class Program

    Shared Sub Main(args As String())

        ' create the downloader and retrieve the data at the other end
        Using downloader As New PFW.CSIST203.Project5.Downloader()
            downloader.DownloadData()
        End Using

    End Sub

End Class
