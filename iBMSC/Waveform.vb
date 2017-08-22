Partial Public Class MainWindow
    Private Sub BWLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWLoad.Click
        Dim xDWAV As New OpenFileDialog
        xDWAV.Filter = Strings.FileType.WAV & "|*.wav"
        xDWAV.DefaultExt = "wav"
        xDWAV.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)
        'Dim xByte As Byte() = My.Computer.FileSystem.ReadAllBytes(DImportWaveForm.FileName)
        'If xByte.Length < 48 Then Exit Sub 'If xByte.Length <= 44 Then Exit Sub 

        Dim xI1 As Integer

        Dim xFileLength As Long = FileSystem.FileLen(xDWAV.FileName)
        If xFileLength < 44 Then Exit Sub

        Dim xData As IO.BinaryReader = New IO.BinaryReader(New IO.FileStream(xDWAV.FileName, IO.FileMode.Open, FileAccess.Read))

        Dim xwChunkID As Integer = xData.ReadInt32
        Dim xwChunkSize As Integer = xData.ReadInt32
        Dim xwFormat As Integer = xData.ReadInt32

        Dim xwSubchunk1ID As Integer = xData.ReadInt32
        Dim xwSubchunk1Size As Integer = xData.ReadInt32
        Dim xwAudioFormat As Short = xData.ReadInt16
        Dim xwNumChannels As Short = xData.ReadInt16
        Dim xwSampleRate As Integer = xData.ReadInt32
        Dim xwByteRate As Integer = xData.ReadInt32
        Dim xwBlockAlign As Short = xData.ReadInt16
        Dim xwBitsPerSample As Short = xData.ReadInt16

        Dim xwCharDA As Short = 0
        Dim xwCharTA As Short = 0
        For xI2 As Integer = 0 To 64
            xwCharDA = xwCharTA
            xwCharTA = xData.ReadInt16

            If xwCharDA = &H6164 And xwCharTA = &H6174 Then
                Exit For
            End If
            If xI2 = 64 Then GoTo ErrFormat
        Next

        'Dim xwSubchunk2ID As Integer = xData.ReadInt32
        Dim xwSubchunk2Size As Integer = xData.ReadInt32

        If xFileLength < xwNumChannels * xwBitsPerSample / 8 + 44 Then GoTo ErrFormat
        If Not xwChunkID = &H46464952 Then GoTo ErrFormat
        If Not xwFormat = &H45564157 Then GoTo ErrFormat
        If Not xwSubchunk1ID = &H20746D66 Then GoTo ErrFormat
        'If Not xwSubchunk1Size = 16 Then GoTo ErrFormat
        If Not xwAudioFormat = 1 Then GoTo ErrFormat

        ReDim wWavL(xwSubchunk2Size \ (xwNumChannels * xwBitsPerSample \ 8) - 1)
        ReDim wWavR(xwSubchunk2Size \ (xwNumChannels * xwBitsPerSample \ 8) - 1)

        Dim xTemp As Byte
        Dim xTemp2 As Short
        Select Case xwNumChannels + xwBitsPerSample
            Case 9      '8b mono
                For xI1 = 0 To UBound(wWavL)
                    wWavL(xI1) = CShort(xData.ReadByte) * 256 - 32768
                    wWavR(xI1) = wWavL(xI1)
                Next

            Case 10     '8b stereo
                For xI1 = 0 To UBound(wWavL)
                    wWavL(xI1) = CShort(xData.ReadByte) * 256 - 32768
                    wWavR(xI1) = CShort(xData.ReadByte) * 256 - 32768
                Next

            Case 17     '16b mono
                For xI1 = 0 To UBound(wWavL)
                    wWavL(xI1) = xData.ReadInt16
                    wWavR(xI1) = wWavL(xI1)
                Next

            Case 18     '16b stereo
                For xI1 = 0 To UBound(wWavL)
                    wWavL(xI1) = xData.ReadInt16
                    wWavR(xI1) = xData.ReadInt16
                Next

            Case 25     '24b mono
                For xI1 = 0 To UBound(wWavL)
                    xTemp = xData.ReadByte
                    wWavL(xI1) = xData.ReadInt16
                    wWavR(xI1) = wWavL(xI1)
                Next

            Case 26     '24b stereo
                For xI1 = 0 To UBound(wWavL)
                    xTemp = xData.ReadByte
                    wWavL(xI1) = xData.ReadInt16
                    xTemp = xData.ReadByte
                    wWavR(xI1) = xData.ReadInt16
                Next

            Case 33     '32b mono
                For xI1 = 0 To UBound(wWavL)
                    xTemp2 = xData.ReadInt16
                    wWavL(xI1) = xData.ReadInt16
                    wWavR(xI1) = wWavL(xI1)
                Next

            Case 34     '32b stereo
                For xI1 = 0 To UBound(wWavL)
                    xTemp2 = xData.ReadInt16
                    wWavL(xI1) = xData.ReadInt16
                    xTemp2 = xData.ReadInt16
                    wWavR(xI1) = xData.ReadInt16
                Next

            Case Else
                Erase wWavL
                Erase wWavR
                GoTo ErrFormat
        End Select

        wBitsPerSample = xwBitsPerSample
        wSampleRate = xwSampleRate
        wNumChannels = xwNumChannels

        TWFileName.Text = xDWAV.FileName
        TWFileName.Select(Len(xDWAV.FileName), 0)
        RefreshPanelAll()
ErrFormat:
        xData.Close()
    End Sub

    Private Sub BWClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWClear.Click
        Erase wWavL
        Erase wWavR
        TWFileName.Text = "(" & Strings.None & ")"
        RefreshPanelAll()
    End Sub

    Private Sub BWLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWLock.CheckedChanged
        wLock = BWLock.Checked
        TWPosition.Enabled = Not wLock
        TWPosition2.Enabled = Not wLock
        RefreshPanelAll()
    End Sub
End Class
