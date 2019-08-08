Partial Public Class MainWindow

    '----Preview Options


    Public PArgs() As PlayerArguments = {New PlayerArguments("<apppath>\uBMplay.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S"),
                                         New PlayerArguments("<apppath>\o2play.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S")}

    Public CurrentPlayer As Integer = 0


    Public Function PrevCodeToReal(initStr As String) As String
        Dim xFileName As String = IIf(Not PathIsValid(_fileName),
                                      IIf(_initPath = "", My.Application.Info.DirectoryPath, _initPath),
                                      ExcludeFileName(_fileName)) _
                                  & "\___TempBMS.bms"
        Dim xMeasure As Integer = MeasureAtDisplacement(Math.Abs(FocusedPanel.VerticalPosition))
        Dim xS1 As String = Replace(initStr, "<apppath>", My.Application.Info.DirectoryPath)
        Dim xS2 As String = Replace(xS1, "<measure>", xMeasure)
        Dim xS3 As String = Replace(xS2, "<filename>", xFileName)
        Return xS3
    End Function

    Private Sub PlayerMissingPrompt()
        Dim xArg As PlayerArguments = PArgs(CurrentPlayer)
        MsgBox(Strings.Messages.CannotFind.Replace("{}", PrevCodeToReal(xArg.Path)) & vbCrLf &
               Strings.Messages.PleaseRespecifyPath, MsgBoxStyle.Critical, Strings.Messages.PlayerNotFound)

        Dim xDOpen As New OpenFileDialog
        xDOpen.InitialDirectory = IIf(ExcludeFileName(PrevCodeToReal(xArg.Path)) = "",
                                      My.Application.Info.DirectoryPath,
                                      ExcludeFileName(PrevCodeToReal(xArg.Path)))
        xDOpen.FileName = PrevCodeToReal(xArg.Path)
        xDOpen.Filter = Strings.FileType.EXE & "|*.exe"
        xDOpen.DefaultExt = "exe"
        If xDOpen.ShowDialog = DialogResult.Cancel Then Exit Sub

        'pArgs(CurrentPlayer) = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>") & _
        '                                           Mid(pArgs(CurrentPlayer), InStr(pArgs(CurrentPlayer), vbCrLf))
        'xStr = Split(pArgs(CurrentPlayer), vbCrLf)
        PArgs(CurrentPlayer).Path = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>")
        xArg = PArgs(CurrentPlayer)
    End Sub

    Private Sub TBPlay_Click(sender As Object, e As EventArgs) Handles TBPlay.Click, mnPlay.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = PArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = PArgs(CurrentPlayer)
        End If

        ' az: Treat it like we cancelled the operation
        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBms()
        Dim xFileName As String = IIf(Not PathIsValid(_fileName),
                                      IIf(_initPath = "", My.Application.Info.DirectoryPath, _initPath),
                                      ExcludeFileName(_fileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, _textEncoding)

        AddToTempFileList(xFileName)
        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.AHere))
    End Sub

    Private Sub TBPlayB_Click(sender As Object, e As EventArgs) Handles TBPlayB.Click, mnPlayB.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = PArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = PArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBms()
        Dim xFileName As String = IIf(Not PathIsValid(_fileName),
                                      IIf(_initPath = "", My.Application.Info.DirectoryPath, _initPath),
                                      ExcludeFileName(_fileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, _textEncoding)

        AddToTempFileList(xFileName)

        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.ABegin))
    End Sub

    Private Sub TBStop_Click(sender As Object, e As EventArgs) Handles TBStop.Click, mnStop.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = PArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = PArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.AStop))
    End Sub
End Class
