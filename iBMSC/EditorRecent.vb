Partial Public Class MainWindow
    Private Sub NewRecent(ByVal xFileName As String)
        Dim xAlreadyExists As Boolean = False
        Dim xI1 As Integer

        For xI1 = 0 To 4
            If Recent(xI1) = xFileName Then
                xAlreadyExists = True
                Exit For
            End If
        Next

        If xAlreadyExists Then
            For xI2 As Integer = xI1 To 1 Step -1
                Recent(xI2) = Recent(xI2 - 1)
            Next
            Recent(0) = xFileName

        Else
            Recent(4) = Recent(3)
            Recent(3) = Recent(2)
            Recent(2) = Recent(1)
            Recent(1) = Recent(0)
            Recent(0) = xFileName
        End If

        SetRecent(4, Recent(4))
        SetRecent(3, Recent(3))
        SetRecent(2, Recent(2))
        SetRecent(1, Recent(1))
        SetRecent(0, Recent(0))
    End Sub

    Private Sub SetRecent(ByVal Index As Integer, ByVal Text As String)
        Text = Text.Trim

        Dim xTBOpenR As ToolStripMenuItem
        Dim xmnOpenR As ToolStripMenuItem
        Select Case Index
            Case 0 : xTBOpenR = TBOpenR0 : xmnOpenR = mnOpenR0
            Case 1 : xTBOpenR = TBOpenR1 : xmnOpenR = mnOpenR1
            Case 2 : xTBOpenR = TBOpenR2 : xmnOpenR = mnOpenR2
            Case 3 : xTBOpenR = TBOpenR3 : xmnOpenR = mnOpenR3
            Case 4 : xTBOpenR = TBOpenR4 : xmnOpenR = mnOpenR4
            Case Else : Return
        End Select

        xTBOpenR.Text = IIf(Text = "", "<" & Strings.None & ">", GetFileName(Text)) : xTBOpenR.ToolTipText = Text : xTBOpenR.Enabled = Not Text = ""
        xmnOpenR.Text = IIf(Text = "", "<" & Strings.None & ">", GetFileName(Text)) : xmnOpenR.ToolTipText = Text : xmnOpenR.Enabled = Not Text = ""
    End Sub

    Private Sub OpenRecent(ByVal xFileName As String)
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If Not My.Computer.FileSystem.FileExists(xFileName) Then
            MsgBox(Strings.Messages.CannotFind.Replace("{}", xFileName), MsgBoxStyle.Critical)
            Exit Sub
        End If
        If ClosingPopSave() Then Exit Sub

        Select Case UCase(Path.GetExtension(xFileName))
            Case ".BMS", ".BME", ".BML", ".PMS", ".TXT"
                InitPath = ExcludeFileName(xFileName)
                SetFileName(xFileName)
                ClearUndo()
                OpenBMS(My.Computer.FileSystem.ReadAllText(xFileName, TextEncoding))
                SetFileName(FileName)
                SetIsSaved(True)
            Case ".IBMSC"
                InitPath = ExcludeFileName(xFileName)
                SetFileName("Imported_" & GetFileName(xFileName))
                OpeniBMSC(xFileName)
                SetIsSaved(False)
        End Select
    End Sub

    Private Sub TBOpenR0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpenR0.Click, mnOpenR0.Click
        OpenRecent(Recent(0))
    End Sub
    Private Sub TBOpenR1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpenR1.Click, mnOpenR1.Click
        OpenRecent(Recent(1))
    End Sub
    Private Sub TBOpenR2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpenR2.Click, mnOpenR2.Click
        OpenRecent(Recent(2))
    End Sub
    Private Sub TBOpenR3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpenR3.Click, mnOpenR3.Click
        OpenRecent(Recent(3))
    End Sub
    Private Sub TBOpenR4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpenR4.Click, mnOpenR4.Click
        OpenRecent(Recent(4))
    End Sub

End Class
