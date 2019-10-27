Imports System
Imports System.IO
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Namespace PFW.CSIST203.Project5.Tests

    ''' <summary>
    ''' A collection of tests that verify the functionality of the testing harness itself
    ''' </summary>
    Public MustInherit Class SystemTests
        Inherits TestBase

        ''' <summary>
        ''' Test harness for GetMethodLogger()
        ''' </summary>
        <TestClass>
        Public Class GetMethodLoggerMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verify the log4net logging object has the correct name for the current executing method
            ''' </summary>
            <TestMethod>
            Public Sub LoggerNameMatchesCallingMethodName()
                Dim log = GetMethodLogger()
                Dim stackTrace = New System.Diagnostics.StackTrace()
                Assert.IsNotNull(stackTrace.GetFrame(0), "Stack Trace Frame(0) information object could not be created")
                Assert.IsNotNull(stackTrace.GetFrame(0).GetMethod(), "Stack Trace Frame(0) method information could not be retrieved")
                Assert.IsTrue(Not String.IsNullOrWhiteSpace(stackTrace.GetFrame(0).GetMethod().Name), "Stack Trace Frame(0) method name not found")
                Dim methodName = stackTrace.GetFrame(0).GetMethod().Name
                Assert.AreEqual(log.Logger.Name, methodName, "Failed to match method name to logger name")
            End Sub

            ''' <summary>
            ''' Verify the log4net logging object has the correct name for the current executing method
            ''' </summary>
            <TestMethod>
            Public Sub LoggerNameMatchesForAnyFrameLevel()
                Dim log = GetMethodLogger(1)
                Dim stackTrace = New System.Diagnostics.StackTrace()
                Assert.IsNotNull(stackTrace.GetFrame(0), "Stack Trace Frame(0) information object could not be created")
                Assert.IsNotNull(stackTrace.GetFrame(0).GetMethod(), "Stack Trace Frame(0) method information could not be retrieved")
                Assert.IsTrue(Not String.IsNullOrWhiteSpace(stackTrace.GetFrame(0).GetMethod().Name), "Stack Trace Frame(0) method name not found")
                Dim methodName = stackTrace.GetFrame(0).GetMethod().Name
                Assert.AreEqual(log.Logger.Name, methodName, "Failed to match method name to logger name")
            End Sub

        End Class

        ''' <summary>
        ''' Verify the state of the testing object is correct after the setup method has been called
        ''' </summary>
        <TestClass>
        Public Class SetupMethod
            Inherits SystemTests

            <TestMethod>
            Public Sub RootLoggerWasAssigned()
                Assert.IsNotNull(Me.logger, "Failed to assign root logger object")
                Assert.AreEqual(OriginalWorkingDirectory, System.IO.Directory.GetCurrentDirectory(),
                    "The current working directory does not match the original working directory prior to test start")
            End Sub

        End Class

        ''' <summary>
        ''' Testing harness for AssertDelegateSuccess
        ''' </summary>
        <TestClass>
        Public Class AssertDelegateSuccessMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verify the AssertDelegateSuccess does not throw an exception under normal operation
            ''' </summary>
            <TestMethod>
            Public Sub TestDelegateDidNotThrowException()
                AssertDelegateSuccess(
                        Sub()
                            ' Do Nothing here
                        End Sub,
                        "Exception was thrown unexpectedly when none should have occurred during custom assert delegate method"
                    )
            End Sub

            ''' <summary>
            ''' Verify AssertDelegateSuccess indicates failure when an exception is thrown
            ''' </summary>
            <TestMethod>
            Public Sub TestDelegateDidThrowException()
                Dim exception = New System.Exception()
                Try
                    AssertDelegateSuccess(
                        Sub()
                            Throw exception
                        End Sub,
                        "This exception was expected to be thrown"
                    )
                    Assert.Fail("An exception was thrown by custom delegate assert method and did not prevent execution as expected")
                Catch ex As Exception
                    If exception IsNot ex Then
                        Assert.Fail("Unexpected exception type was thrown: " + ex.GetType().ToString())
                    End If
                End Try
            End Sub

        End Class

        ''' <summary>
        ''' Testing harness for AssertDelegateFailure
        ''' </summary>
        <TestClass>
        Public Class AssertDelegateFailureMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verify not throwing an exception causes failure
            ''' </summary>
            <TestMethod>
            Public Sub TestDelegateDidNotThrowException()
                Dim notThrown = False
                Try
                    AssertDelegateFailure(
                        Sub()
                            ' Do Not throw an exception
                        End Sub,
                        GetType(Exception),
                        "Expected behavior: no exception was thrown, but one was expected"
                        )
                    notThrown = True
                Catch ex As Exception
                    ' Nothing to specifically handle
                End Try
                If notThrown Then
                    Assert.Fail("An exception was expected to be thrown, but one was not")
                End If
            End Sub

            ''' <summary>
            ''' Verify throwing an exception exception causes success
            ''' </summary>
            <TestMethod>
            Public Sub TestDelegateDidThrowException()
                Dim exception = New Exception
                Try
                    AssertDelegateFailure(
                        Sub()
                            Throw exception
                        End Sub,
                        GetType(Exception),
                        "Expected behavior: no exception was thrown, but one was expected"
                        )
                Catch ex As Exception
                    Assert.Fail("Error when delegate failure check: " + ex.Message)
                End Try
            End Sub

        End Class

        ''' <summary>
        ''' Testing harness for GetMethodSpecificWorkingDirectory()
        ''' </summary>
        <TestClass>
        Public Class GetMethodSpecificWorkingDirectoryMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verify the name of the testing directory created is correct
            ''' </summary>
            <TestMethod>
            Public Sub IsDirectoryNameCorrect()

                Dim stackTrace = New System.Diagnostics.StackTrace()
                Dim clazz = Me.GetType()
                Dim category = clazz.BaseType
                Dim method = stackTrace.GetFrame(0).GetMethod()
                Dim expected_directoryName = System.IO.Path.Combine(
                    System.Environment.CurrentDirectory,
                    "Tests",
                    TrimStringFromEnd(category.Name, "Tests", True),
                    TrimStringFromEnd(clazz.Name, "Method", True),
                    method.Name)

                ' cananicalize the local directory path
                Dim expected = New System.Uri(expected_directoryName).LocalPath
                Dim generated_directoryName = GetMethodSpecificWorkingDirectory()
                ' cananicalize the local directory path
                Dim actual = New System.Uri(generated_directoryName).LocalPath

                Assert.AreEqual(expected, actual,
                    "The generated path is not consistent with the expected value")
                Assert.IsTrue(System.IO.Directory.Exists(actual),
                    "The default method should have created the directory")
                Assert.AreEqual(0, System.IO.Directory.GetFiles(actual).Length,
                    "The default method should have created an empty directory")
            End Sub

            ''' <summary>
            ''' Verify the name of the testing directory is correct when using within a lambda function
            ''' </summary>
            <TestMethod>
            Public Sub IsDirectoryNameCorrectWithinLambda()

                AssertDelegateSuccess(
                    Sub()
                        Dim stackTrace = New System.Diagnostics.StackTrace()
                        Dim clazz = Me.GetType()
                        Dim category = clazz.BaseType
                        Dim method = stackTrace.GetFrame(2).GetMethod()
                        Dim expected_directoryName = System.IO.Path.Combine(
                            System.Environment.CurrentDirectory,
                            "Tests",
                            TrimStringFromEnd(category.Name, "Tests", True),
                            TrimStringFromEnd(clazz.Name, "Method", True),
                            method.Name)

                        ' cananicalize the local directory path
                        Dim expected = New System.Uri(expected_directoryName).LocalPath
                        Dim generated_directoryName = GetMethodSpecificWorkingDirectory()
                        ' cananicalize the local directory path
                        Dim actual = New System.Uri(generated_directoryName).LocalPath

                        Assert.AreEqual(expected, actual,
                            "The generated path is not consistent with the expected value")
                        Assert.IsTrue(System.IO.Directory.Exists(actual),
                            "The default method should have created the directory")
                        Assert.AreEqual(0, System.IO.Directory.GetFiles(actual).Length,
                            "The default method should have created an empty directory")
                    End Sub,
                    "No exception should have been thrown")

            End Sub

            ''' <summary>
            ''' Verify the name of the testing directory is correct when using within any number of nested lambda functions
            ''' </summary>
            <TestMethod>
            Public Sub IsDirectoryNameCorrectWithinNestedLambda()

                AssertDelegateSuccess(
                    Sub()

                        Dim nested1 As Action =
                            Sub()

                                Dim stackTrace = New System.Diagnostics.StackTrace()
                                Dim clazz = Me.GetType()
                                Dim category = clazz.BaseType
                                Dim method = stackTrace.GetFrame(6).GetMethod()
                                Dim expected_directoryName = System.IO.Path.Combine(
                                    System.Environment.CurrentDirectory,
                                    "Tests",
                                    TrimStringFromEnd(category.Name, "Tests", True),
                                    TrimStringFromEnd(clazz.Name, "Method", True),
                                    method.Name)

                                ' cananicalize the local directory path
                                Dim expected = New System.Uri(expected_directoryName).LocalPath
                                Dim generated_directoryName = GetMethodSpecificWorkingDirectory()
                                ' cananicalize the local directory path
                                Dim actual = New System.Uri(generated_directoryName).LocalPath
                                Assert.AreEqual(expected, actual,
                                    "The generated path is not consistent with the expected value")
                                Assert.IsTrue(System.IO.Directory.Exists(actual),
                                    "The default method should have created the directory")
                                Assert.AreEqual(0, System.IO.Directory.GetFiles(actual).Length,
                                    "The default method should have created an empty directory")
                            End Sub

                        Dim nested2 As Action =
                            Sub()
                                nested1()
                            End Sub
                        Dim nested3 As Action =
                            Sub()
                                nested2()
                            End Sub
                        Dim nested4 As Action =
                            Sub()
                                nested3()
                            End Sub

                        ' attempt a heavily nested series of lambda functions
                        nested4()

                    End Sub,
                    "No exception should have been thrown")

            End Sub
        End Class

        ''' <summary>
        ''' Testing harness for the IsASCII() method
        ''' </summary>
        <TestClass>
        Public Class IsASCIIMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verifies that the helper method property detects all ASCII compatible characters
            ''' </summary>
            <TestMethod>
            Public Sub DetectsASCII()
                For i As Integer = 0 To 127
                    Dim SingleByte = New Byte() {CType(i, Byte)}
                    Dim SingleCharacterString = System.Text.Encoding.Default.GetString(SingleByte)
                    Assert.IsTrue(IsASCII(SingleCharacterString),
                        "Did not detect character '" & SingleCharacterString & "' with code '" & i.ToString() & "' as ASCII")
                Next
            End Sub

            ''' <summary>
            ''' Verifies that the helper method property detects all non-ASCII compatible characters
            ''' </summary>
            <TestMethod>
            Public Sub DetectsNonASCII()
                For i As Integer = 128 To 255
                    Dim SingleByte = New Byte() {CType(i, Byte)}
                    Dim SingleCharacterString = System.Text.Encoding.Default.GetString(SingleByte)
                    Assert.IsFalse(IsASCII(SingleCharacterString),
                        "Detected character '" & SingleCharacterString & "' with code '" & i.ToString() & "' as ASCII")
                Next
            End Sub

        End Class

        ''' <summary>
        ''' Testing harness for the TrimStringFromEnd() method
        ''' </summary>
        <TestClass>
        Public Class TrimStringFromEndMethod
            Inherits SystemTests

            ''' <summary>
            ''' Verify the ignore case parameter works as expected
            ''' </summary>
            <TestMethod>
            Sub TrimsEndOfStringIgnoringCase()

                Dim testingString = "A Simple Textual Example"
                Assert.AreEqual("A Simple Textual", TrimStringFromEnd(testingString, "Example", True))
                Assert.AreEqual("A Simple", TrimStringFromEnd(testingString, " textual example", True))
                Assert.AreEqual(testingString, TrimStringFromEnd(testingString, "A Simple", True))

            End Sub

            ''' <summary>
            ''' Verify the case-sensitive replace works as expected
            ''' </summary>
            <TestMethod>
            Sub TrimsEndOfStringRespectingCase()

                Dim testingString = "A Simple Textual Example"
                Assert.AreEqual("A Simple Textual", TrimStringFromEnd(testingString, "Example", False))
                Assert.AreEqual(testingString, TrimStringFromEnd(testingString, " textual example", False))
                Assert.AreEqual(testingString, TrimStringFromEnd(testingString, "A Simple", False))

            End Sub

            ''' <summary>
            ''' Verify trailing whitespace has no impact on the trimming operation
            ''' </summary>
            <TestMethod>
            Sub TrimsEndOfStringIgnoringWhitespace()

                Dim testingString = "A Simple Textual Example "
                Assert.AreEqual("A Simple Textual", TrimStringFromEnd(testingString, "Example", True))
                Assert.AreEqual("A Simple", TrimStringFromEnd(testingString, " textual example", True))
                Assert.AreEqual(testingString.Trim(), TrimStringFromEnd(testingString, "A Simple", True))

            End Sub

            ''' <summary>
            ''' Verify no exception is thrown with null or whitespace input or removal strings
            ''' </summary>
            <TestMethod>
            Sub NoExceptionOnNullOrWhitespace()
                'Dim testingString = "A Simple Textual Example"
                AssertDelegateSuccess(
                    Sub()
                        TrimStringFromEnd(String.Empty, "anything", True)
                        TrimStringFromEnd(Nothing, "something", True)
                        TrimStringFromEnd("Nothing", String.Empty, True)
                        TrimStringFromEnd("Thing", Nothing, True)
                        TrimStringFromEnd(Nothing, Nothing, True)
                        TrimStringFromEnd(String.Empty, String.Empty, True)
                    End Sub,
                    "No exception should be thrown for null or whitespace strings")
            End Sub

            ''' <summary>
            ''' Verify that the trim method does not impact the string more than once and only the final substring if it occurs more than once
            ''' </summary>
            <TestMethod>
            Sub DoesNotRemoveMoreThanOnce()
                Dim testingString = "ATestATestATest"
                Assert.AreEqual("ATestATest", TrimStringFromEnd(testingString, "ATest", True))
            End Sub

        End Class

    End Class

End Namespace
