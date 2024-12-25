Option Strict On
Option Explicit On

Imports AtosF
Imports Spea.VisionNet.Instructions
Imports System.IO
Imports ArmadaPack

Module modTestplan

    Private ReadOnly CurrentDir As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd("\"c)

    Public Function Testplan() As Integer

        Const cameraId As Integer = 1
        Const suffix As String = "_golden.bmp"

        Dim failFlag As Integer = PASS

        Try
            Dim ftpDir As DirectoryInfo = Directory.GetParent(CurrentDir)
            Dim visionDir As String = Path.Combine(ftpDir.FullName, "Vision")
            Dim visionXml As String = Path.Combine(visionDir, "vision.xml")
            Dim visionInstance As Vision = ArmadaHelper.Deserialize(visionXml)

            MsgPrintLog("Clear previous session if exists...", NO)
            For Each f As String In Directory.GetFiles(visionDir, "*" & suffix)
                File.Delete(f)
            Next

            For Each group As TestPartsGroup In visionInstance.Groups

                Dim site As Integer = group.Site
                UseSiteWrite(site)

                MsgPrintLog("Start operation on site " & site, NO)

                AlignExecute()

                ' move FP camera on selected components
                For Each part As TestPart In group.TestParts
                    Dim x As Integer = part.CoordinateX
                    Dim y As Integer = part.CoordinateY
                    Dim df As String = part.DrawingReference

                    Dim errorCode As Integer = CameraMove(cameraId, x, y)
                    Dim capturedImage As Integer

                    If errorCode = RESULT_PASS Then
                        MsgPrintLog("Site " & site & " capturing golden sample picture of component " & df, NO)
                        errorCode = IA.iaImageCapture(cameraId, df, 0, capturedImage, 0)
                    Else
                        MsgPrintLog("Site " & site & " error during moving camera for component " & df, NO)
                        SiteResultWrite(site, RESULT_FAIL)
                        failFlag = FAIL
                        Continue For
                    End If

                    Dim filename As String = part.DrawingReference & "_" & site & suffix
                    Dim fullname As String = Path.Combine(visionDir, filename)

                    If errorCode = RESULT_PASS Then
                        MsgPrintLog("Site " & site & " saving golden sample picture of component " & df, NO)
                        errorCode = IA.iaImageSave(capturedImage, fullname, df, 0, 0)
                    Else
                        MsgPrintLog("Site " & site & " error during capturing image for component " & df, NO)
                        SiteResultWrite(site, RESULT_FAIL)
                        failFlag = FAIL
                        Continue For
                    End If

                    If errorCode = RESULT_PASS Then
                        part.GoldenSampleFilePath = filename
                        MsgPrintLog("Site " & site & " golden sample picture of component " & df & " saved successfully.", NO)
                    Else
                        MsgPrintLog("Site " & site & " error during saving image for component " & df, NO)
                        SiteResultWrite(site, RESULT_FAIL)
                        failFlag = FAIL
                        Continue For
                    End If

                    MsgPrintLog("----------------------------------------------------------------", NO)
                Next

                CameraPark()
            Next

            ArmadaHelper.Serialize(visionXml, visionInstance)

        Catch ex As Exception
            LogPrint(ex.Message, YES)

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
