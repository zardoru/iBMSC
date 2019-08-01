

Partial Public Class MainWindow
    Public ReadOnly Property FirstBpm
        Get
            Return Notes(0).Value
        End Get
    End Property

    Public ReadOnly Property CurrentWavSelectedIndex As Integer
        Get
            Return LWAV.SelectedIndex + 1
        End Get
    End Property

    Public ReadOnly Property AutoincreaseWavIndex As Boolean
        Get
            Return TBWavIncrease.Checked
        End Get
    End Property

    Private Sub NewRecent(xFileName As String)
        Dim xAlreadyExists = False
        Dim i As Integer

        For i = 0 To 4
            If _recent(i) = xFileName Then
                xAlreadyExists = True
                Exit For
            End If
        Next

        If xAlreadyExists Then
            For j As Integer = i To 1 Step - 1
                _recent(j) = _recent(j - 1)
            Next
            _recent(0) = xFileName

        Else
            _recent(4) = _recent(3)
            _recent(3) = _recent(2)
            _recent(2) = _recent(1)
            _recent(1) = _recent(0)
            _recent(0) = xFileName
        End If

        SetRecent(4, _recent(4))
        SetRecent(3, _recent(3))
        SetRecent(2, _recent(2))
        SetRecent(1, _recent(1))
        SetRecent(0, _recent(0))
    End Sub

    Private Sub SetRecent(Index As Integer, Text As String)
        Text = Text.Trim

        Dim xTBOpenR As ToolStripMenuItem
        Dim xmnOpenR As ToolStripMenuItem
        Select Case Index
            Case 0 : xTBOpenR = TBOpenR0
                xmnOpenR = mnOpenR0
            Case 1 : xTBOpenR = TBOpenR1
                xmnOpenR = mnOpenR1
            Case 2 : xTBOpenR = TBOpenR2
                xmnOpenR = mnOpenR2
            Case 3 : xTBOpenR = TBOpenR3
                xmnOpenR = mnOpenR3
            Case 4 : xTBOpenR = TBOpenR4
                xmnOpenR = mnOpenR4
            Case Else : Return
        End Select

        xTBOpenR.Text = IIf(Text = "", "<" & Strings.None & ">", GetFileName(Text))
        xTBOpenR.ToolTipText = Text
        xTBOpenR.Enabled = Not Text = ""
        xmnOpenR.Text = IIf(Text = "", "<" & Strings.None & ">", GetFileName(Text))
        xmnOpenR.ToolTipText = Text
        xmnOpenR.Enabled = Not Text = ""
    End Sub

    Private Sub OpenRecent(xFileName As String)
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = - 1

        If Not My.Computer.FileSystem.FileExists(xFileName) Then
            MsgBox(Strings.Messages.CannotFind.Replace("{}", xFileName), MsgBoxStyle.Critical)
            Exit Sub
        End If
        If ClosingPopSave() Then Exit Sub

        Select Case UCase(Path.GetExtension(xFileName))
            Case ".BMS", ".BME", ".BML", ".PMS", ".TXT"
                _initPath = ExcludeFileName(xFileName)
                SetFileName(xFileName)
                ClearUndo()
                OpenBMS(My.Computer.FileSystem.ReadAllText(xFileName, _textEncoding))
                SetFileName(_fileName)
                SetIsSaved(True)
            Case ".IBMSC"
                _initPath = ExcludeFileName(xFileName)
                SetFileName("Imported_" & GetFileName(xFileName))
                OpeniBMSC(xFileName)
                SetIsSaved(False)
        End Select
    End Sub

    Private Sub TBOpenR0_Click(sender As Object, e As EventArgs) Handles TBOpenR0.Click, mnOpenR0.Click
        OpenRecent(_recent(0))
    End Sub

    Private Sub TBOpenR1_Click(sender As Object, e As EventArgs) Handles TBOpenR1.Click, mnOpenR1.Click
        OpenRecent(_recent(1))
    End Sub

    Private Sub TBOpenR2_Click(sender As Object, e As EventArgs) Handles TBOpenR2.Click, mnOpenR2.Click
        OpenRecent(_recent(2))
    End Sub

    Private Sub TBOpenR3_Click(sender As Object, e As EventArgs) Handles TBOpenR3.Click, mnOpenR3.Click
        OpenRecent(_recent(3))
    End Sub

    Private Sub TBOpenR4_Click(sender As Object, e As EventArgs) Handles TBOpenR4.Click, mnOpenR4.Click
        OpenRecent(_recent(4))
    End Sub
End Class
