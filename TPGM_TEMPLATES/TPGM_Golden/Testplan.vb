Option Strict On
Option Explicit On

Imports Spea.VisionNet.Instructions
Imports System.IO
Imports ArmadaPack

Module modTestplan

    Private ReadOnly CurrentDir As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd("\"c)

    Public Function Testplan() As Integer

        Const cameraId As Integer = 1
        Dim gain As Double = 50
        Dim exposureTime As Integer = 7000

        Dim failFlag As Integer = PASS

        Try
            ' Dim tp As Integer = (site - 1) * config.TestPointOffset + config.BarcodeTestPoint
            ' there is no need to add offset to the test point number

            ' 1. move camera back to the initial position
            ' AtosF.CameraPark()
            ' IA.iaCameraGainWrite(rCameraID, rGain, "", 0, 0)
            ' IA.iaCameraExposureTimeWrite(rCameraID, rExposureTime, "", 0, 0)

            AtosF.AlignExecute()

            Dim ftpDir As DirectoryInfo = Directory.GetParent(CurrentDir)
            Dim visionDir As String = Path.Combine(ftpDir.FullName, "Vision")
            Dim visionXml As String = Path.Combine(visionDir, "vision.xml")
            Dim visionObj As Vision = ArmadaHelper.Deserialize(visionXml)

            ' 2. move FP camera on selected components
            For Each part As TestPart In visionObj.TestParts
                Dim x As Integer = part.CoordinateX
                Dim y As Integer = part.CoordinateY
                Dim df As String = part.DrawingReference

                Dim errorCode As Integer = AtosF.CameraMove(cameraId, x, y)
                Dim capturedImage As Integer

                If errorCode = RESULT_PASS Then
                    MsgPrintLog("Capturing golden sample picture of component " & df, NO)
                    errorCode = IA.iaImageCapture(cameraID, df, 0, capturedImage, 0)
                Else
                    MsgPrintLog("Error during moving camera for component " & df, NO)
                    Continue For
                End If

                Dim filename As String = part.DrawingReference + "_Golden.bmp"
                Dim fullname As String = Path.Combine(visionDir, filename)

                If errorCode = RESULT_PASS Then
                    MsgPrintLog("Saving golden sample picture of component " & df, NO)
                    errorCode = IA.iaImageSave(capturedImage, fullname, df, 0, 0)
                Else
                    MsgPrintLog("Error during capturing image for component " & df, NO)
                    Continue For
                End If

                If errorCode = RESULT_PASS Then
                    part.GoldenSampleFilePath = filename
                    MsgPrintLog("Golden sample picture of component " & df & " saved successfully.", NO)
                Else
                    MsgPrintLog("Error during saving image for component " & df, NO)
                    Continue For
                End If

                MsgPrintLog("----------------------------------------------------------------", NO)
            Next

            ArmadaHelper.Serialize(visionXml, visionObj)

            'If rErrorCode = 0 Then
            '    rErrorCode = IA.iaCameraOpen(rCameraID, rTestPoint, 0, 0)
            '    If rErrorCode <> 0 Then FailFlag = FAIL
            'End If
            '----- Initialization FP camera -----  
            'If rErrorCode = 0 Then
            '    Dim pErrorCode As Integer
            '    rErrorCode = IA.iaCameraGainWrite(rCameraID, rGain, "TP", 0, pErrorCode)
            '    If rErrorCode <> 0 Then FailFlag = FAIL
            'End If
            'If rErrorCode = 0 Then
            '    rErrorCode = IA.iaCameraExposureTimeWrite(rCameraID, rExposureTime, "", 0, 0)
            '    If rErrorCode <> 0 Then FailFlag = FAIL
            'End If

            AtosF.CameraPark()

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
