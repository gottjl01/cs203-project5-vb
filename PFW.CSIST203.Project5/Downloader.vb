Namespace PFW.CSIST203.Project5

    ''' <summary>
    ''' Downloader class that retrieves a configured list of URLs and saves them into the specified flat file(s)
    ''' </summary>
    Public Class Downloader
        Implements IDisposable

        ''' <summary>
        ''' Configuration data from the config file
        ''' </summary>
        Friend config As PFW.CSIST203.Project5.Config.DownloadConfiguration

        ''' <summary>
        ''' Default constructor that uses the standard PFW.CSIST203.Project5.Config.DownloadConfiguration section from the application configuration file
        ''' </summary>
        Public Sub New()
            Me.New("PFW.CSIST203.Project5.Config.DownloadConfiguration")
        End Sub

        ''' <summary>
        ''' Initializes the downloader object with the application configuration section specified
        ''' </summary>
        ''' <param name="configSectionName"></param>
        Public Sub New(configSectionName As String)

            ' TODO: IMPLEMENT

        End Sub

        Public Sub DownloadData()
            'TODO Implement the download 
        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ''' <summary>
        ''' Perform disposal of resources in this method
        ''' </summary>
        ''' <param name="disposing">Redundant call?</param>
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    config = Nothing
                End If
            End If
            disposedValue = True
        End Sub

        ''' <summary>
        ''' This code added by Visual Basic to correctly implement the disposable pattern.
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
        End Sub
#End Region

    End Class

End Namespace

