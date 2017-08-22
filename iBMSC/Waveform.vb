Imports CSCore.Streams
Imports CSCore.Streams.Effects
Imports CSCore
Imports CSCore.Codecs


Partial Public Class MainWindow

    '----WaveForm Options
    Dim wWavL() As Single
    Dim wWavR() As Single
    Dim wLock As Boolean = True
    Dim wSampleRate As Integer
    Dim wPosition As Double = 0
    Dim wLeft As Integer = 50
    Dim wWidth As Integer = 100
    Dim wPrecision As Integer = 1

    Private Sub BWLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWLoad.Click
        Dim xDWAV As New OpenFileDialog
        xDWAV.Filter = "Wave files (*.wav, *.ogg)" & "|*.wav;*.ogg"
        xDWAV.DefaultExt = "wav"
        xDWAV.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)

        Dim src = CSCore.Codecs.CodecFactory.Instance.GetCodec(xDWAV.FileName)

        src.ToStereo()
        Dim samples(src.Length) As Single
        src.ToSampleSource().Read(samples, 0, src.Length)

        Dim flen = (src.Length - 1) / src.WaveFormat.Channels

        ' Copy interleaved data
        ReDim wWavL(flen + 1)
        ReDim wWavR(flen + 1)
        For i As Integer = 0 To flen
            If 2 * i < src.Length Then
                wWavL(i) = samples(2 * i)
            End If
            If 2 * i + 1 < src.Length Then
                wWavR(i) = samples(2 * i + 1)
            End If
        Next

        wSampleRate = src.WaveFormat.SampleRate
        RefreshPanelAll()

        TWFileName.Text = xDWAV.FileName
        TWFileName.Select(Len(xDWAV.FileName), 0)
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
