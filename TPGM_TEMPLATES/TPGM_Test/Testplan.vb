Option Strict On
Option Explicit On

Imports AtosF
Imports Spea.VisionNet.Instructions
Imports System.IO
Imports ArmadaPack
Imports System.Collections.Generic

Module modTestplan

    Private ReadOnly CurrentDir As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd("\"c)

    Public Function Testplan() As Integer

        Const cameraId As Integer = 1
        Dim failFlag As Integer = RESULT_FAIL
        Dim testResults As New List(Of Integer)

        Const workingDirectory As String = "C:\Armada"

        Try
            Dim ftpDir As DirectoryInfo = Directory.GetParent(CurrentDir)
            Dim visionDir As String = Path.Combine(ftpDir.FullName, "Vision")
            Dim visionXml As String = Path.Combine(visionDir, "vision.xml")
            Dim visionInstance As Vision = ArmadaHelper.Deserialize(visionXml)
            Dim pixelRatio As Integer = visionInstance.PixelRatio
            Dim offsetPixel As Integer = visionInstance.BorderWidthPixels

            For Each group As TestPartsGroup In visionInstance.Groups
                Dim site As Integer = group.Site
                UseSiteWrite(site)

                Dim previousResult As Integer = SiteResultRead(site)
                If previousResult = RESULT_NONE Then
                    Continue For
                End If

                MsgPrintLog("Start operation on site " & site, NO)

                AlignExecute()

                Dim siteResult As Integer = RESULT_PASS

                For Each part As TestPart In group.TestParts
                    Dim x As Integer = part.CoordinateX
                    Dim y As Integer = part.CoordinateY
                    Dim df As String = part.DrawingReference

                    Dim errorCode As Integer = CameraMove(cameraId, x, y)
                    Dim capturedImage As Integer

                    If errorCode = RESULT_PASS Then
                        MsgPrintLog("Site " & site & " capturing picture of component " & df, NO)
                        errorCode = IA.iaImageCapture(cameraId, df, 0, capturedImage, 0)
                    Else
                        MsgPrintLog("Site " & site & " error during moving camera for component " & df, NO)
                        siteResult = RESULT_FAIL
                        SiteResultWrite(site, RESULT_FAIL)
                        failFlag = FAIL
                        Continue For
                    End If

                    Dim filename As String = part.DrawingReference & "_" & site & "_temp.bmp"
                    Dim fullname As String = Path.Combine(visionDir, filename)

                    If errorCode = RESULT_PASS Then
                        MsgPrintLog("Site " & site & " saving picture of component " & df, NO)
                        errorCode = IA.iaImageSave(capturedImage, fullname, df, 0, 0)
                    Else
                        MsgPrintLog("Site " & site & " error during capturing image for component " & df, NO)
                        siteResult = RESULT_FAIL
                        SiteResultWrite(site, RESULT_FAIL)
                        failFlag = FAIL
                        Continue For
                    End If

                    If errorCode = RESULT_PASS Then

                        MsgPrintLog("Site " & site & " test picture of component " & df & " saved successfully.", NO)

                        Dim pixelEdge As Integer = ArmadaHelper.MillimetersToPixels(part.Edge, pixelRatio, offsetPixel)
                        Dim goldenSampleFile As String = Path.Combine(visionDir, part.GoldenSampleFilePath)
                        Dim result As CompareResult = ArmadaHelper.CompareImages(workingDirectory, goldenSampleFile, fullname, pixelEdge)

                        If result.IsPass(visionInstance.SimilarityRate / 100) Then
                            MsgPrintLog("Site " & site & " Compare Result of " & df & " is " & result.PercentResult & " @FG{White}@BG{Green}PASS", NO)
                        Else
                            MsgPrintLog("Site " & site & " Compare Result " & df & " @FG{White}@BG{Red}FAIL", NO)
                            siteResult = RESULT_FAIL
                            SiteResultWrite(site, RESULT_FAIL)
                            failFlag = FAIL
                        End If

                    Else
                        MsgPrintLog("Site " & site & "Error during saving image for component " & df, NO)
                        Continue For
                    End If

                    File.Delete(fullname)
                    MsgPrintLog("----------------------------------------------------------------", NO)
                Next

                If previousResult = RESULT_FAIL Then
                    siteResult = RESULT_FAIL
                End If
                SiteResultWrite(site, siteResult)
                testResults.Add(siteResult)

                CameraPark()
            Next

            If testResults.Contains(RESULT_FAIL) Then
                failFlag = RESULT_FAIL
            Else
                failFlag = RESULT_PASS
            End If

        Catch ex As Exception
            LogPrint(ex.Message, YES)
            LogPrint(ex.StackTrace, YES)

            failFlag = FAIL
        End Try

        ' --- Test Result management

        TplanResultSet(failFlag)

        Return failFlag

    End Function


    ' ----------------------------------------------------------------------------
    '
    ' --- TESTPLAN Initialisation
    '
    ' This function is executed only one time when the test program is loaded.
    '
    Public Function TestplanInit() As Integer

        '
        ' --- INSERT YOUR CODE HERE ...
        '

        TestplanInit = 1
    End Function


    ' ----------------------------------------------------------------------------
    '
    ' --- TESTPLAN End
    '
    ' This function is executed only one time when the test program ends.
    '
    Public Function TestplanEnd() As Integer

        '
        ' --- INSERT YOUR CODE HERE ...
        '

        TestplanEnd = 1
    End Function

End Module
