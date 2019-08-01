Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    Partial Friend Class MyApplication
        Private Sub MyApplication_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup
        End Sub

        Private Sub MyApplication_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs) _
            Handles Me.UnhandledException
            Dim xRes As MsgBoxResult = MsgBox("An unhandled exception has occurred in the application: " & vbCrLf &
                                              e.Exception.Message & vbCrLf &
                                              vbCrLf &
                                              "Click Yes to save a back-up, click No otherwise, or click Cancel to ignore this exception and continue.",
                                              MsgBoxStyle.YesNoCancel + MsgBoxStyle.Critical,
                                              "Unhandled Exception")
            If xRes = MsgBoxResult.Cancel Then e.ExitApplication = False
            If xRes = MsgBoxResult.Yes Then
                Dim xFN As String
                Dim xDate As Date = DateTime.Now
                With xDate
                    xFN = "\AutoSave_" & .Year & "_" & .Month & "_" & .Day & "_" & .Hour & "_" & .Minute & "_" & .Second &
                          "_" & .Millisecond & ".IBMSC"
                End With

                'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & xFN, Form1.ExceptionSave, False)
                MainWindow.ExceptionSave(Application.Info.DirectoryPath & xFN)
                MsgBox("A back-up has been saved to " & Application.Info.DirectoryPath & xFN, MsgBoxStyle.Information)
            End If
        End Sub
    End Class
End Namespace

