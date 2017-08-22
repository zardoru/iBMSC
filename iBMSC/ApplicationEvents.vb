Namespace My

    Partial Friend Class MyApplication

        Private Sub MyApplication_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown

        End Sub

        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup

        End Sub

        Private Sub MyApplication_UnhandledException(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            Dim xRes As MsgBoxResult = MsgBox("An unhandled exception has occurred in the application: " & vbCrLf & _
                                       e.Exception.Message & vbCrLf & _
                                       vbCrLf & _
                                       "Click Yes to save a back-up, click No otherwise, or click Cancel to ignore this exception and continue.", _
                                       MsgBoxStyle.YesNoCancel + MsgBoxStyle.Critical, _
                                       "Unhandled Exception")
            If xRes = MsgBoxResult.Cancel Then e.ExitApplication = False
            If xRes = MsgBoxResult.Yes Then
                Dim xFN As String
                Dim xDate As Date = DateTime.Now
                With xDate
                    xFN = "\AutoSave_" & .Year & "_" & .Month & "_" & .Day & "_" & .Hour & "_" & .Minute & "_" & .Second & "_" & .Millisecond & ".IBMSC"
                End With

                'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & xFN, Form1.ExceptionSave, False)
                MainWindow.ExceptionSave(My.Application.Info.DirectoryPath & xFN)
                MsgBox("A back-up has been saved to " & My.Application.Info.DirectoryPath & xFN, MsgBoxStyle.Information)
            End If
        End Sub
    End Class

End Namespace

