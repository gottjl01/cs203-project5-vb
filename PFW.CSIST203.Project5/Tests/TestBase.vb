Imports System
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace PFW.CSIST203.Project5.Tests

    ''' <summary>
    ''' Base class for all test harnesses
    ''' </summary>
    Public MustInherit Class TestBase
        Implements System.IDisposable

        ''' <summary>
        ''' Get a value indicating whether the current test object has been disposed
        ''' </summary>
        Private isDisposed As Boolean = False

        ''' <summary>
        ''' A logging object named for the current testing harness
        ''' </summary>
        Protected logger As log4net.ILog = Nothing

        ''' <summary>
        ''' The working directory where the application was original ran from
        ''' </summary>
        Protected Shared ReadOnly OriginalWorkingDirectory As String = Nothing

        ''' <summary>
        ''' Static constructor that retains the original working directory for the Setup() method
        ''' </summary>
        Shared Sub New()
            OriginalWorkingDirectory = System.IO.Directory.GetCurrentDirectory()
        End Sub

        ''' <summary>
        ''' Provides test initialization logic
        ''' </summary>
        <TestInitialize>
        Public Overridable Sub Setup()

            log4net.Config.XmlConfigurator.ConfigureAndWatch(New FileInfo(System.AppDomain.CurrentDomain.SetupInformation.ConfigurationFile))
            logger = log4net.LogManager.GetLogger(Me.GetType())
            System.IO.Directory.SetCurrentDirectory(OriginalWorkingDirectory)

        End Sub

        ''' <summary>
        ''' Provides test cleanup logic
        ''' </summary>
        <TestCleanup>
        Public Overridable Sub Dispose() Implements IDisposable.Dispose

            If Not isDisposed Then
                log4net.LogManager.Shutdown()
                Me.isDisposed = True
            End If

        End Sub

        ''' <summary>
        ''' Utility method that retrieves a method-specific logger
        ''' </summary>
        ''' <returns>A log4net logging object named for the current stack frame method</returns>
        Protected Function GetMethodLogger() As log4net.ILog
            Return GetMethodLogger(2)
        End Function

        ''' <summary>
        ''' Utility method that retrieves a method-specific logger at the specified frame
        ''' </summary>
        ''' <param name="frame">The frame level to name the logger</param>
        ''' <returns>A log4net logging object named for the specified stack frame method</returns>
        Protected Function GetMethodLogger(frame As Integer) As log4net.ILog
            Dim stackTrace = New System.Diagnostics.StackTrace()
            Return log4net.LogManager.GetLogger(stackTrace.GetFrame(frame).GetMethod().Name)
        End Function

        ''' <summary>
        ''' Helper method that asserts delegate execution was successful
        ''' </summary>
        ''' <param name="action">Delegate that should not throw an exception</param>
        ''' <param name="Message">The message to log when the delegate execution does not succeed</param>
        Protected Sub AssertDelegateSuccess(action As Action, Message As String)
            Try
                action.Invoke()
            Catch ex As Exception
                Dim log = GetMethodLogger(2)
                Dim msg = "Error during delegate execution: " + Message
                log.Error(msg, ex)
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' Helper method that asserts delegate execution was a failure
        ''' </summary>
        ''' <param name="action">The delegate action to execute</param>
        ''' <param name="exceptionType">The type of exception that should be thrown</param>
        ''' <param name="Message">A message describing what no thrown exception means</param>
        Protected Sub AssertDelegateFailure(action As Action, exceptionType As Type, Message As String)
            Assert.IsNotNull(exceptionType, "Null exception type was supplied, but this cannot be tested for")
            Try
                action.Invoke()
                Assert.Fail("The delegate did not throw the intended exception: " + exceptionType.FullName)
            Catch ex As Exception
                If (ex.GetType() <> exceptionType) Then
                    Dim msg = String.Format("Delegate threw exception of type '{0}', but expected '{1}': {2}", ex.GetType(), exceptionType, Message)
                    Dim log = GetMethodLogger(2)
                    log.Error(msg, ex)
                    Assert.Fail(msg)
                End If
            End Try
        End Sub

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method
        ''' </summary>
        ''' <returns></returns>
        Protected Function GetMethodSpecificWorkingDirectory() As String
            Return GetMethodSpecificWorkingDirectory(
                System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.SetupInformation.ConfigurationFile),
                3
            )
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method
        ''' </summary>
        ''' <param name="baseDirectory"></param>
        ''' <param name="depth"></param>
        ''' <returns></returns>
        Protected Function GetMethodSpecificWorkingDirectory(baseDirectory As String, depth As Integer) As String
            Return GetMethodSpecificWorkingDirectory(baseDirectory, depth, True)
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method
        ''' </summary>
        ''' <param name="baseDirectory"></param>
        ''' <param name="depth"></param>
        ''' <param name="setWorkingDirectory"></param>
        ''' <returns></returns>
        Protected Function GetMethodSpecificWorkingDirectory(baseDirectory As String, depth As Integer, setWorkingDirectory As Boolean) As String
            Dim stackTrace = New System.Diagnostics.StackTrace()

            ' This handles anonymous method invokation and the two frames added by using it
            Dim method = stackTrace.GetFrame(depth).GetMethod()
            If method.DeclaringType.FullName.IndexOf("<>") >= 0 Then
                method = stackTrace.GetFrame(depth + 2).GetMethod()
            End If

            ' handle any number of nested lambda functions
            If method.Name.StartsWith("_Lambda$", StringComparison.Ordinal) Then
                Dim tmpDepth = depth
                While method.Name.StartsWith("_Lambda$", StringComparison.Ordinal)
                    tmpDepth += 2
                    method = stackTrace.GetFrame(tmpDepth).GetMethod()
                End While
            End If

            Dim methodEncapsulatingClass = method.DeclaringType
            Dim methodEncapsulatingClassName = TrimStringFromEnd(methodEncapsulatingClass.Name, "Method", True) ' GetMethodSpecificWorkingDirectoryMethod
            Dim encapsulatingClass = methodEncapsulatingClass.BaseType
            Dim encapsulatingClassName = TrimStringFromEnd(encapsulatingClass.Name, "Tests", True) ' SystemTests

            Dim finalDirectory = System.IO.Path.Combine(baseDirectory, "Tests", encapsulatingClassName, methodEncapsulatingClassName, method.Name)

            ' https://social.technet.microsoft.com/Forums/windows/en-US/43945b2c-f123-46d7-9ba9-dd6abc967dd4/maximum-path-length-limitation-on-windows-is-255-or-247?forum=w7itprogeneral
            If (finalDirectory.Length >= 247) Then
                Throw New ArgumentException("Unable to create testing directory because it is too long: " + finalDirectory)
            End If

            If setWorkingDirectory Then
                System.IO.Directory.CreateDirectory(finalDirectory)
                System.IO.Directory.SetCurrentDirectory(finalDirectory)
            End If

            Return finalDirectory

        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method
        ''' </summary>
        ''' <returns></returns>
        Protected Function GetCleanWorkingTestDirectory() As String
            Return GetCleanWorkingTestDirectory(True)
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method
        ''' </summary>
        ''' <param name="cleanWorkingDirectory"></param>
        ''' <returns></returns>
        Protected Function GetCleanWorkingTestDirectory(cleanWorkingDirectory As Boolean) As String
            Return GetCleanWorkingTestDirectory(
                System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.SetupInformation.ConfigurationFile),
                3,
                cleanWorkingDirectory
            )
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by a testing method specified at the given stack frame
        ''' </summary>
        ''' <param name="depth"></param>
        ''' <param name="cleanWorkingDirectory"></param>
        ''' <returns></returns>
        Protected Function GetCleanWorkingTestDirectory(depth As Integer, cleanWorkingDirectory As Boolean) As String
            Return GetCleanWorkingTestDirectory(
                System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.SetupInformation.ConfigurationFile),
                depth,
                cleanWorkingDirectory
            )
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by the testing method specified at the given stack frame and base directory
        ''' </summary>
        ''' <param name="baseDirectory"></param>
        ''' <param name="depth"></param>
        ''' <param name="cleanWorkingDirectory"></param>
        ''' <returns></returns>
        Protected Function GetCleanWorkingTestDirectory(baseDirectory As String, depth As Integer, cleanWorkingDirectory As Boolean) As String
            Return GetCleanWorkingTestDirectory(baseDirectory, depth, cleanWorkingDirectory, True)
        End Function

        ''' <summary>
        ''' Retrieves a clean working directory for use by the testing method specified at the given stack frame and base directory
        ''' </summary>
        ''' <param name="baseDirectory"></param>
        ''' <param name="depth"></param>
        ''' <param name="cleanWorkingDirectory"></param>
        ''' <returns></returns>
        Protected Function GetCleanWorkingTestDirectory(baseDirectory As String, depth As Integer, cleanWorkingDirectory As Boolean, setWorkingDirectory As Boolean) As String
            Dim testDirectory = GetMethodSpecificWorkingDirectory(baseDirectory, depth + 1, setWorkingDirectory)
            logger.Info(String.Format("Working Directory: {0}", testDirectory))
            If (cleanWorkingDirectory AndAlso System.IO.Directory.Exists(testDirectory)) Then
                System.IO.Directory.Delete(testDirectory, True)
            End If
            System.IO.Directory.CreateDirectory(testDirectory)
            Return testDirectory
        End Function

        ''' <summary>
        ''' Utility method that takes embedded resources located at the designated location and copies them into a target directory
        ''' </summary>
        ''' <param name="embeddedResourceBase"></param>
        ''' <param name="targetDirectory"></param>
        Protected Sub CopyEmbeddedResourceBaseToDirectory(embeddedResourceBase As String, targetDirectory As String)
            Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim resources = assembly.GetManifestResourceNames()
            For Each resource In resources
                If resource.StartsWith(embeddedResourceBase) Then

                    '' perform some basic fixing
                    Dim fixedResourceName = resource.Replace("._", ".") _
                        .Replace("_", "-")

                    Dim resourceRelativePathCompontents = fixedResourceName.Replace(embeddedResourceBase, String.Empty).TrimStart("."c).Split("."c)
                    Dim resourceRelativePath = String.Join("\\", resourceRelativePathCompontents, 0, resourceRelativePathCompontents.Length - 1) _
                        + "." + resourceRelativePathCompontents(resourceRelativePathCompontents.Length - 1)

                    '' calculate the relative path of the file's output location
                    Dim targetPath = System.IO.Path.Combine(targetDirectory, resourceRelativePath)

                    Dim Directory = System.IO.Path.GetDirectoryName(targetPath)
                    System.IO.Directory.CreateDirectory(Directory)

                    '' copy the input to the output stream
                    Using output = System.IO.File.OpenWrite(targetPath)
                        Using input = assembly.GetManifestResourceStream(resource)
                            Dim b1 = CType(input.ReadByte(), Byte)
                            Dim b2 = CType(input.ReadByte(), Byte)
                            Dim b3 = CType(input.ReadByte(), Byte)

                            Dim data = New Byte() {b1, b2, b3}
                            Dim str = System.Text.Encoding.Default.GetString(data)

                            If IsASCII(str) Then
                                input.Seek(0, SeekOrigin.Begin)
                            Else ' non-ASCII headers needs to be skipped
                                input.Seek(3, SeekOrigin.Begin)
                            End If
                            input.CopyTo(output, 4096)

                        End Using
                    End Using


                End If
            Next


        End Sub

        ''' <summary>
        ''' Utility function for determining whether the supplied string falls into the ASCII character set
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        Friend Shared Function IsASCII(value As String) As Boolean
            Return System.Text.Encoding.UTF8.GetByteCount(value) = value.Length
        End Function

        ''' <summary>
        ''' Utility function for trimming a string located at the end of the supplied input string
        ''' </summary>
        ''' <param name="str">The input string to evaluate</param>
        ''' <param name="removalString">The string that shoudl be trimmed from the end if present</param>
        ''' <param name="ignoreCase">Should the match be case insensitive</param>
        ''' <returns>The trimmed string or the original string if the matched string was not found at the end</returns>
        Friend Shared Function TrimStringFromEnd(str As String, removalString As String, ignoreCase As Boolean) As String
            Dim result = If(str Is Nothing, Nothing, str.Trim())
            If Not String.IsNullOrWhiteSpace(removalString) Then
                Dim comparison = If(ignoreCase, StringComparison.CurrentCultureIgnoreCase, StringComparison.CurrentCulture)
                If Not String.IsNullOrWhiteSpace(result) AndAlso result.EndsWith(removalString, comparison) Then

                    ' look for the last occurrance of the supplied string
                    Dim index = result.LastIndexOf(removalString, comparison)

                    ' is the discovered string located at the end?
                    If index + removalString.Length = result.Length Then
                        ' remove the trailing string from the supplied input string
                        result = result.Substring(0, index).Trim()
                    End If

                End If
            End If
            Return result
        End Function

    End Class

End Namespace

