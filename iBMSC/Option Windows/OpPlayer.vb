Public Class OpPlayer
    Dim pArg() As PlayerArguments
    'Dim ImplicitChange As Boolean = False
    Dim CurrPlayer As Integer = - 1

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        DialogResult = DialogResult.OK
        Close()

        MainWindow.pArgs = pArg.Clone
        MainWindow.CurrentPlayer = CurrPlayer

        Dispose()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        DialogResult = DialogResult.Cancel
        Close()
        Dispose()
    End Sub

    Private Sub OpPlayer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font

        Me.Text = Strings.fopPlayer.Title
        Label1.Text = Strings.fopPlayer.Path
        Label2.Text = Strings.fopPlayer.PlayFromBeginning
        Label3.Text = Strings.fopPlayer.PlayFromHere
        Label4.Text = Strings.fopPlayer.StopPlaying
        BAdd.Text = Strings.fopPlayer.Add
        BRemove.Text = Strings.fopPlayer.Remove
        Label6.Text = Strings.fopPlayer.References & vbCrLf &
                      "<apppath> = " & Strings.fopPlayer.DirectoryOfApp & vbCrLf &
                      "<measure> = " & Strings.fopPlayer.CurrMeasure & vbCrLf &
                      "<filename> = " & Strings.fopPlayer.FileName
        OK_Button.Text = Strings.OK
        Cancel_Button.Text = Strings.Cancel
        BDefault.Text = Strings.fopPlayer.RestoreDefault
    End Sub

    Private Sub LPlayer_Click(sender As Object, e As EventArgs) Handles LPlayer.Click
        If pArg Is Nothing OrElse pArg.Length = 0 Then Exit Sub

        CurrPlayer = LPlayer.SelectedIndex
        ShowInTextbox()
    End Sub

    Private Sub LPlayer_KeyDown(sender As Object, e As KeyEventArgs) Handles LPlayer.KeyDown
        LPlayer_Click(sender, New EventArgs)
    End Sub

    Private Sub BPrevAdd_Click(sender As Object, e As EventArgs) Handles BAdd.Click
        ReDim Preserve pArg(UBound(pArg) + 1)
        CurrPlayer += 1
        For i As Integer = UBound(pArg) To CurrPlayer Step - 1
            pArg(i) = pArg(i - 1)
        Next

        LPlayer.Items.Insert(CurrPlayer,
                             GetFileName(pArg(CurrPlayer - 1).Path))
        LPlayer.SelectedIndex += 1
    End Sub

    Private Sub BPrevDelete_Click(sender As Object, e As EventArgs) Handles BRemove.Click
        If LPlayer.Items.Count = 1 Then
            MsgBox(Strings.Messages.PreviewDelError, MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        For i As Integer = CurrPlayer To UBound(pArg) - 1
            pArg(i) = pArg(i + 1)
        Next
        ReDim Preserve pArg(UBound(pArg) - 1)

        'RemoveHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged
        LPlayer.Items.RemoveAt(CurrPlayer)
        'AddHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged

        LPlayer.SelectedIndex = IIf(CurrPlayer > UBound(pArg), CurrPlayer - 1, CurrPlayer)
        CurrPlayer = Math.Min(CurrPlayer, UBound(pArg))
        ShowInTextbox()
    End Sub

    Private Sub BPrevBrowse_Click(sender As Object, e As EventArgs) Handles BBrowse.Click
        Dim xDOpen As New OpenFileDialog
        xDOpen.InitialDirectory =
            IIf(Path.GetDirectoryName(Replace(TPath.Text, "<apppath>", My.Application.Info.DirectoryPath)) = "",
                My.Application.Info.DirectoryPath,
                Path.GetDirectoryName(Replace(TPath.Text, "<apppath>", My.Application.Info.DirectoryPath)))
        xDOpen.Filter = Strings.FileType.EXE & "|*.exe"
        xDOpen.DefaultExt = "exe"
        If xDOpen.ShowDialog = DialogResult.Cancel Then Exit Sub
        TPath.Text = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>")
    End Sub

    Private Sub BPrevDefault_Click(sender As Object, e As EventArgs) Handles BDefault.Click
        'ImplicitChange = True
        If MsgBox(Strings.Messages.RestoreDefaultSettings, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No _
            Then Exit Sub

        pArg = New PlayerArguments() {New PlayerArguments("<apppath>\uBMplay.exe",
                                                          "-P -N0 ""<filename>""",
                                                          "-P -N<measure> ""<filename>""",
                                                          "-S"),
                                      New PlayerArguments("<apppath>\o2play.exe",
                                                          "-P -N0 ""<filename>""",
                                                          "-P -N<measure> ""<filename>""",
                                                          "-S")}
        CurrPlayer = 0
        ResetLPlayer_ShowInTextbox()
        'ImplicitChange = False
    End Sub

    'Affect LPlayer and all textboxes
    Private Sub ResetLPlayer_ShowInTextbox()
        LPlayer.Items.Clear()
        For i = 0 To UBound(pArg)
            LPlayer.Items.Add(GetFileName(pArg(i).Path))
        Next
        'RemoveHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged
        LPlayer.SelectedIndex = CurrPlayer
        'AddHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged
        ShowInTextbox()
        'ImplicitChange = False
    End Sub

    'affect current LPlayer index value
    Private Sub LPlayerChangeCurrIndex(xStr As String)
        'RemoveHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged
        LPlayer.Items.Item(CurrPlayer) = GetFileName(xStr)
        'AddHandler LPlayer.SelectedIndexChanged, AddressOf LPlayer_SelectedIndexChanged
    End Sub

    'Affect pArgs
    Private Sub SavePArg()
        pArg(CurrPlayer).Path = TPath.Text
        pArg(CurrPlayer).aBegin = TPlayB.Text
        pArg(CurrPlayer).aHere = TPlay.Text
        pArg(CurrPlayer).aStop = TStop.Text
        'pArg(CurrPlayer) = TPath.Text & vbCrLf & _
        '                   TPlayB.Text & vbCrLf & _
        '                   TPlay.Text & vbCrLf & _
        '                   TStop.Text
    End Sub

    'affect all textboxes
    Private Sub ShowInTextbox()
        'ImplicitChange = True
        'Dim xStr() As String = Split(pArg(CurrPlayer), vbCrLf)
        'If xStr.Length <> 4 Then ReDim Preserve xStr(3)
        TPath.Text = pArg(CurrPlayer).Path
        TPlayB.Text = pArg(CurrPlayer).aBegin
        TPlay.Text = pArg(CurrPlayer).aHere
        TStop.Text = pArg(CurrPlayer).aStop
        ValidateTextBox()
        'ImplicitChange = False
    End Sub

    Private Sub ValidateTextBox()
        For Each xT As TextBox In New TextBox() {TPath, TPlayB, TPlay, TStop}
            Dim xText As String =
                    xT.Text.Replace("<apppath>", "").Replace("<measure>", "").Replace("<filename>", "").Replace("""", "")
            Dim xContainsInvalidChar = False

            For Each xC As Char In Path.GetInvalidPathChars
                If xText.IndexOf(xC) <> - 1 Then
                    xContainsInvalidChar = True
                    Exit For
                End If
            Next

            If xContainsInvalidChar Then
                xT.BackColor = Color.FromArgb(&HFFFFC0C0)
            Else
                xT.BackColor = Nothing
            End If
        Next
    End Sub

    Public Sub New(xCurrPlayer As Integer)
        InitializeComponent()

        pArg = MainWindow.pArgs.Clone
        CurrPlayer = xCurrPlayer
        ResetLPlayer_ShowInTextbox()
    End Sub

    Private Sub TPath_KeyDown(sender As Object, e As KeyEventArgs) _
        Handles TPath.KeyUp, TPlay.KeyUp, TPlayB.KeyUp, TStop.KeyUp
        SavePArg()
        If ReferenceEquals(sender, TPath) Then _
            LPlayerChangeCurrIndex(pArg(CurrPlayer).Path)
    End Sub

    Private Sub TPath_LostFocus(sender As Object, e As EventArgs) _
        Handles TPath.LostFocus, TPlay.LostFocus, TPlayB.LostFocus, TStop.LostFocus
        SavePArg()
        If ReferenceEquals(sender, TPath) Then _
            LPlayerChangeCurrIndex(pArg(CurrPlayer).Path)
        ValidateTextBox()
    End Sub

    'Private Function pArgPath(ByVal I As Integer)
    '    Return Mid(pArg(I), 1, InStr(pArg(I), vbCrLf) - 1)
    'End Function

    Private Function GetFileName(s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        Return Mid(s, IIf(fslash > bslash, fslash, bslash) + 1)
    End Function
End Class
