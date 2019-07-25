﻿Imports NVorbis
Imports CSCore
Imports CSCore.Codecs
Imports CSCore.SoundOut

Module Audio
    Dim Output As WasapiOut
    Dim Source As IWaveSource

    Public Structure BMSONSliceResult
        Public Found As Boolean
        Public Start As Double
        Public Length As Double
        Public Overrides Function ToString() As String
            Return "Slice result " & Found & " - " & Start & " => " & Length
        End Function
    End Structure

    ' https://github.com/filoe/cscore/blob/master/CSCore.Test/utils/TrimmedWaveSource.cs
    Public Class TrimmedWaveSource
        Inherits WaveAggregatorBase
        Private mSource As IWaveSource
        Private mMaxLength As Integer
        Private mRead As Integer
        Public Sub New(Source As IWaveSource, MaxLength As Integer)
            MyBase.New(Source)
            mSource = Source
            mMaxLength = MaxLength
        End Sub
        Public Overrides Function Read(buffer() As Byte, offset As Integer, count As Integer) As Integer
            Dim ReadBytes = MyBase.Read(buffer, offset, count)
            mRead = mRead + ReadBytes
            If mRead > mMaxLength Then
                ReadBytes = ReadBytes - (mRead - mMaxLength)
                mRead = mMaxLength
            End If
            Return ReadBytes
        End Function
    End Class

    Public Sub Initialize()
        Output = New WasapiOut()
        CodecFactory.Instance.Register("ogg", New CodecFactoryEntry(Function(s)
                                                                        Return New NVorbisSource(s).ToWaveSource()
                                                                    End Function, ".ogg"))
    End Sub

    Public Sub Finalize()
        Output.Stop()
        Output.Dispose()
        Output = Nothing
    End Sub

    Public Function CheckFilename(ByVal filename As String) As String
        If File.Exists(filename) Then
            Return filename
        End If
        Dim ext = Path.GetExtension(filename)
        If String.Compare(ext, ".ogg") = 0 Then
            Dim wpath = Path.ChangeExtension(filename, ".wav")
            Return IIf(File.Exists(wpath), wpath, filename)
        End If
        If String.Compare(ext, ".wav") = 0 Then
            Dim opath = Path.ChangeExtension(filename, ".ogg")
            Return IIf(File.Exists(opath), opath, filename)
        End If
        Return filename
    End Function

    Public Sub Play(ByVal filename As String, ByVal slice As BMSONSliceResult)

        If Source IsNot Nothing Then
            Output.Stop()
            Source.Dispose()
            Source = Nothing
        End If

        If filename Is "" Then
            Return
        End If

        Dim fn = CheckFilename(filename)

        If Not File.Exists(fn) Then
            Return
        End If

        Source = CodecFactory.Instance.GetCodec(fn)
        If slice.Found Then
            Source.Position = Source.WaveFormat.MillisecondsToBytes(Conversion.Int(slice.Start))
            If slice.Length > 0 Then
                Source = New TrimmedWaveSource(Source, Source.WaveFormat.MillisecondsToBytes(Conversion.Int(slice.Length)))
            End If
        End If
        Output.Initialize(Source)
        Output.Play()
    End Sub

    Public Sub Play(ByVal filename As String)
        Dim slice As BMSONSliceResult
        Play(filename, slice)
    End Sub

    Public Sub StopPlaying()
        Output.Stop()
    End Sub
End Module

Class NVorbisSource
    Implements CSCore.ISampleSource
    Dim _stream As Stream
    Dim _vorbisReader As VorbisReader
    Dim _waveFormat As WaveFormat
    Dim _disposed As Boolean

    Public Sub New(stream As Stream)
        If stream Is Nothing Or Not stream.CanRead Then
            Throw New ArgumentException("stream")
        End If
        _stream = stream
        _vorbisReader = New VorbisReader(stream, Nothing)
        _waveFormat = New WaveFormat(_vorbisReader.SampleRate, 32, _vorbisReader.Channels, AudioEncoding.IeeeFloat)
    End Sub

    Public ReadOnly Property CanSeek As Boolean Implements IAudioSource.CanSeek
        Get
            Return _stream.CanSeek
        End Get
    End Property

    Public ReadOnly Property WaveFormat As WaveFormat Implements IAudioSource.WaveFormat
        Get
            Return _waveFormat
        End Get
    End Property

    Public ReadOnly Property Length As Long Implements IAudioSource.Length
        Get
            Return IIf(CanSeek, _vorbisReader.TotalTime.TotalSeconds * _waveFormat.SampleRate * _waveFormat.Channels, 0)
        End Get
    End Property

    Public Property Position As Long Implements IAudioSource.Position
        Get
            Return IIf(CanSeek, _vorbisReader.DecodedTime.TotalSeconds * _vorbisReader.SampleRate * _vorbisReader.Channels, 0)
        End Get
        Set(value As Long)
            If Not CanSeek Then
                Throw New InvalidOperationException("Can't seek this stream.")
            End If
            If value < 0 Or value >= Length Then
                Throw New ArgumentOutOfRangeException("value")
            End If
            _vorbisReader.DecodedTime = TimeSpan.FromSeconds(value / _vorbisReader.SampleRate / _vorbisReader.Channels)
        End Set
    End Property


    Public Function Read(buffer As Single(), offset As Integer, count As Integer) As Integer Implements ISampleSource.Read
        Return _vorbisReader.ReadSamples(buffer, offset, count)
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not _disposed Then
            '_vorbisReader.Dispose()
        Else
            'Throw New ObjectDisposedException("NVorbisSource")
        End If
        _disposed = True
    End Sub

End Class
