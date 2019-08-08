Imports System.Linq
Imports System.Media
Imports System.Text
Imports iBMSC.Editor

Public Structure PlayerArguments
    Public Path As String
    Public ABegin As String
    Public AHere As String
    Public AStop As String

    Public Sub New(xPath As String, xBegin As String, xHere As String, xStop As String)
        Path = xPath
        aBegin = xBegin
        aHere = xHere
        aStop = xStop
    End Sub
End Structure

Public Class MainWindow
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA"(hwnd As IntPtr, wMsg As Integer,
                                                                              wParam As Integer, lParam As Integer) _
        As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture"() As Integer
    Public MeasureLength(999) As Double
    Public MeasureBottom(999) As Double

    Public Function MeasureUpper(idx As Integer) As Double
        Return MeasureBottom(idx) + MeasureLength(idx)
    End Function


    Dim _mColumn(999) As Integer  '0 = no column, 1 = 1 column, etc.
    Public Readonly Property GreatestVPosition As Integer
        Get
            Return - Math.Min(Notes.Max(Function(x) x.VPosition) + 2000, GetMaxVPosition())
        End Get
    End Property

    'Dim SortingMethod As Integer = 1
    Public MiddleButtonMoveMethod As Integer = 0
    Dim _textEncoding As Encoding = Encoding.UTF8
    Dim _dispLang As String = ""     'Display Language
    Dim ReadOnly _recent() As String = {"", "", "", "", ""}
    Public NtInput As Boolean = True
    Public ShowFileName As Boolean = False

    Dim _beepWhileSaved As Boolean = True
    Dim _bpMx1296 As Boolean = False
    Dim _stoPx1296 As Boolean = False

    Dim _isInitializing As Boolean = True
    Public IsFirstMouseEnterOnPanel As Boolean = True

    Dim _wavMultiSelect As Boolean = True
    Dim _wavChangeLabel As Boolean = True
    Dim _beatChangeMode As Integer = 0

    'Dim FloatTolerance As Double = 0.0001R
    Dim _bmsGridLimit As Double = 1.0R


    'IO
    Dim _fileName As String = "Untitled.bms"

    Dim _initPath As String = ""
    Dim _isSaved As Boolean = True

    'Variables for Drag/Drop
    Public DragDropFilename() As String = {}
    Dim ReadOnly _supportedFileExtension() As String = {".bms", ".bme", ".bml", ".pms", ".txt", ".sm", ".ibmsc"}
    Dim ReadOnly _supportedAudioExtension() As String = {".wav", ".mp3", ".ogg"}

    Dim ReadOnly _
        _supportedImageExtension() As String =
            {".bmp", ".png", ".jpg", ".jpeg", ".gif", ".mpg", ".mpeg", ".avi", ".m1v", ".m2v", ".m4v", ".mp4", ".webm",
             ".wmv"}

    'Variables for theme
    'Dim SaveTheme As Boolean = True

    Public State As EditorState = New EditorState ' Shared state across panels

    'Variables for select tool
    Public DisableVerticalMove As Boolean = False

    'Dim KMouseDown As Integer = -1   'Mouse is clicked on which note (for moving)


    ' Dim SelectedNotes(-1) As Note        'temp notes for undo

    'Variables for write tool
    Public ReadOnly Property ShouldDrawTempNote
        Get
            Return IsWriteMode
        End Get
    End Property


    Public LnDisplayLength As Double = 0.0#

    'Variables for post effects tool


    'Variables for Full-Screen Mode
    Public IsFullscreen As Boolean = False
    Dim _previousWindowState As FormWindowState = FormWindowState.Normal
    Dim _previousWindowPosition As New Rectangle(0, 0, 0, 0)

    'Variables misc
    Dim ReadOnly _menuVPosition As Double = 0.0#
    Dim _tempResize As Integer = 0

    '----AutoSave Options
    Dim _previousAutoSavedFileName As String = ""
    Dim _autoSaveInterval As Integer = 120000

    '----ErrorCheck Options
    Public ErrorCheck As Boolean = True

    '---- Grid Options
    Public Grid As Grid = New Grid

    '---- Columns
    Public Columns As ColumnList = New ColumnList

    '----Visual Options
    Dim _vo As New visualSettings()

    Public Sub SetVo(xvo As visualSettings)
        _vo = xvo
    End Sub

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
    Dim _previewOnClick As Boolean = True
    Dim _previewErrorCheck As Boolean = False
    Dim _clickStopPreview As Boolean = True
    Dim _pTempFileNames() As String = {}

    '----Split Panel Options
    Dim ReadOnly _spLock() As Boolean = {False, False, False}
    Dim ReadOnly _spDiff() As Integer = {0, 0, 0}
    Public PanelFocus As Integer = 0 '0 = Left, 1 = Middle, 2 = Right
    Public Property FocusedPanel As EditorPanel
        Get
            Return _spMain(PanelFocus)
        End Get
        Set
            'If value Is PMainL Then PanelFocus = 0
            'If value Is PMain Then PanelFocus = 1
            'If value Is PMainR Then PanelFocus = 2
        End Set
    End Property

    Dim _currentHoveredPanel As Integer = 1

    ' az: guess these all have to do with focusing stuff?
    Public AutoFocusPanelOnMouseEnter As Boolean = False
    Public FirstClickDisabled As Boolean = True
    Public TempFirstMouseDown As Boolean = False

    Dim _spMain() As EditorPanel = {}


    Public Sub New()
        InitializeComponent()
        Initialize()
    End Sub

    ''' <summary>
    '''     Not a property as it is calculated every time.
    '''     Not sure how to make it play nice with the SelectedNotes array.
    ''' </summary>
    ''' <returns></returns>
    Friend Function GetSelectedNotes() As IEnumerable(Of Note)
        Return Notes.Skip(1).Where(Function(x) x.Selected)
    End Function

    Private Sub DecreaseCurrentWav()
        If LWAV.SelectedIndex = - 1 Then
            LWAV.SelectedIndex = 0
        Else
            Dim newIndex As Integer = LWAV.SelectedIndex - 1
            If newIndex < 0 Then newIndex = 0
            LWAV.SelectedIndices.Clear()
            LWAV.SelectedIndex = newIndex
        End If
    End Sub

    Private Sub IncreaseCurrentWav()
        If LWAV.SelectedIndex = - 1 Then
            LWAV.SelectedIndex = 0
        Else
            Dim newIndex As Integer = LWAV.SelectedIndex + 1
            If newIndex > LWAV.Items.Count - 1 Then newIndex = LWAV.Items.Count - 1
            LWAV.SelectedIndices.Clear()
            LWAV.SelectedIndex = newIndex
            ValidateWavListView()
        End If
    End Sub

    ''' <summary>
    '''     Clears the SelectedNotes array.
    '''     To be figured out how this plays with DeselectAllNotes.
    ''' </summary>
    Public Sub ClearSelectionArray()
        ' ReDim SelectedNotes(-1)
    End Sub

    ''' <summary>
    '''     Call whenever the note order is no longer guaranteed.
    '''     Say, a modification in VPosition, insertion at the end or such.
    ''' </summary>
    Public Sub ValidateNotesArray()
        Notes = Notes.OrderBy(Function(x) x.VPosition).ToArray()
        UpdatePairing()
        CalculateTotalPlayableNotes()

        Dim newMax = -GreatestVPosition + 2000
        Dim maxVpos = GetMaxVPosition()
        Dim xI2 As Integer = -Math.Min(newMax, maxVpos)

        For Each panel In _spMain
            panel.OnUpdateScroll(xI2)
        Next

    End Sub

    Public Sub PanelPreviewNoteIndex(noteIndex As Integer)
        'Play wav
        If _clickStopPreview Then PreviewNote("", True)
        'My.Computer.Audio.Stop()
        If NoteIndex > 0 And _previewOnClick AndAlso Columns.IsColumnSound(Notes(NoteIndex).ColumnIndex) Then
            Dim j As Integer = Notes(NoteIndex).Value\10000
            If j <= 0 Then j = 1
            If j >= 1296 Then j = 1295

            If Not BmsWAV(j) = "" Then ' AndAlso Path.GetExtension(hWAV(j)).ToLower = ".wav" Then
                Dim xFileLocation As String = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName)) &
                                              "\" & BmsWAV(j)
                If Not _clickStopPreview Then PreviewNote("", True)
                PreviewNote(xFileLocation, False)
            End If
        End If
    End Sub

    Friend Sub SetToolstripVisible(v As Boolean)
        ToolStripContainer1.TopToolStripPanelVisible = v
    End Sub

    ' Why here instead of panel events?
    ' MainWindow should be taking care of the overall focus state, not the individual panels.
    Private Sub PMainInMouseEnter(sender As Object, e As EventArgs)

        _currentHoveredPanel = sender.Tag
        Dim xPMainIn As Panel = sender

        If AutoFocusPanelOnMouseEnter AndAlso Focused Then
            xPMainIn.Focus()
            PanelFocus = _currentHoveredPanel
        End If

        If IsFirstMouseEnterOnPanel Then
            IsFirstMouseEnterOnPanel = False
            xPMainIn.Focus()
            PanelFocus = _currentHoveredPanel
        End If
    End Sub

    Private Sub PMainInMouseLeave(sender As Object, e As EventArgs)

        State.Mouse.CurrentHoveredNoteIndex = -1
        ClearSelectionArray()
        State.Mouse.CurrentMouseRow = -1
        State.Mouse.CurrentMouseColumn = -1
        RefreshPanelAll()
    End Sub

    Private Sub UpdateMeasureBottom()
        MeasureBottom(0) = 0.0#
        For i = 0 To 998
            MeasureBottom(i + 1) = MeasureBottom(i) + MeasureLength(i)
        Next
    End Sub

    Private Function PathIsValid(sPath As String) As Boolean
        Return File.Exists(sPath) Or Directory.Exists(sPath)
    End Function

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

    Private Sub SetFileName(xFileName As String)
        _fileName = xFileName.Trim
        _initPath = ExcludeFileName(_fileName)
        SetIsSaved(_isSaved)
    End Sub

    Private Sub SetIsSaved(isSaved As Boolean)
        'pttl.Refresh()
        'pIsSaved.Visible = Not xBool
        Dim xVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor &
                                 IIf(My.Application.Info.Version.Build = 0, "", "." & My.Application.Info.Version.Build)
        Text = IIf(isSaved, "", "*") & GetFileName(_fileName) & " - " & My.Application.Info.Title & " " & xVersion
        Me._isSaved = isSaved
    End Sub




    Private Sub PreviewNote(xFileLocation As String, bStop As Boolean)
        If bStop Then
            StopPlaying()
        End If
        Play(xFileLocation)
    End Sub

    Public Sub AddNote(note As Note, Optional xSelected As Boolean = False, Optional overWrite As Boolean = True,
                       Optional sortAndUpdatePairing As Boolean = True)

        If note.VPosition < 0 Or note.VPosition >= GetMaxVPosition() Then Exit Sub

        Dim i = 1

        If overWrite Then
            Do While i <= UBound(Notes)
                If Notes(i).VPosition = note.VPosition And
                   Notes(i).ColumnIndex = note.ColumnIndex Then
                    RemoveNote(i)
                Else
                    i += 1
                End If
            Loop
        End If

        note.Selected = note.Selected And Columns.IsEnabled(note.ColumnIndex)
        Notes = Notes.Concat({note}).ToArray()

        If sortAndUpdatePairing Then
            ValidateNotesArray()
        Else
            CalculateTotalPlayableNotes()
        End If
    End Sub

    Public Sub RemoveNote(I As Integer, Optional ByVal sortAndUpdatePairing As Boolean = True)
        State.Mouse.CurrentHoveredNoteIndex = -1 ' az: Why here???
        Dim j As Integer

        If TBWavIncrease.Checked Then
            If Notes(I).Value = LWAV.SelectedIndex * 10000 Then
                DecreaseCurrentWav()
            End If
        End If


        For j = I + 1 To UBound(Notes)
            Notes(j - 1) = Notes(j)
        Next
        ReDim Preserve Notes(UBound(Notes) - 1)
        If sortAndUpdatePairing Then
            ValidateNotesArray()
        End If
    End Sub


    Private Sub RemoveNotes(Optional ByVal sortAndUpdatePairing As Boolean = True)
        If UBound(Notes) = 0 Then Exit Sub

        State.Mouse.CurrentHoveredNoteIndex = -1
        Dim i = 1
        Dim j As Integer
        Do
            If Notes(i).Selected Then
                For j = i + 1 To UBound(Notes)
                    Notes(j - 1) = Notes(j)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
                i = 0
            End If
            i += 1
        Loop While i < UBound(Notes) + 1
        If sortAndUpdatePairing Then
            ValidateNotesArray()
        End If
        CalculateTotalPlayableNotes()
    End Sub


    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If _pTempFileNames IsNot Nothing Then
            For Each xStr As String In _pTempFileNames
                File.Delete(xStr)
            Next
        End If
        If _previousAutoSavedFileName <> "" Then File.Delete(_previousAutoSavedFileName)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not _isSaved Then
            Dim xStr As String = Strings.Messages.SaveOnExit
            If e.CloseReason = CloseReason.WindowsShutDown Then xStr = Strings.Messages.SaveOnExit1
            If e.CloseReason = CloseReason.TaskManagerClosing Then xStr = Strings.Messages.SaveOnExit2

            Dim xResult As MsgBoxResult = MsgBox(xStr, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes Then
                If ExcludeFileName(_fileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = _initPath

                    If xDSave.ShowDialog = DialogResult.Cancel Then e.Cancel = True : Exit Sub
                    SetFileName(xDSave.FileName)
                End If
                Dim xStrAll As String = SaveBms()
                My.Computer.FileSystem.WriteAllText(_fileName, xStrAll, False, _textEncoding)
                NewRecent(_fileName)
                If _beepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then e.Cancel = True
        End If

        If Not e.Cancel Then
            SaveSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml", False)
        End If
    End Sub

    Private Function FilterFileBySupported(xFile() As String, xFilter() As String) As String()
        Dim xPath(-1) As String
        For i = 0 To UBound(xFile)
            If _
                My.Computer.FileSystem.FileExists(xFile(i)) And
                Array.IndexOf(xFilter, Path.GetExtension(xFile(i))) <> -1 Then
                ReDim Preserve xPath(UBound(xPath) + 1)
                xPath(UBound(xPath)) = xFile(i)
            End If

            If My.Computer.FileSystem.DirectoryExists(xFile(i)) Then
                Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(xFile(i)).GetFiles()
                For Each xStr As FileInfo In xFileNames
                    If Array.IndexOf(xFilter, xStr.Extension) = -1 Then Continue For
                    ReDim Preserve xPath(UBound(xPath) + 1)
                    xPath(UBound(xPath)) = xStr.FullName
                Next
            End If
        Next

        Return xPath
    End Function

    Private Sub InitializeNewBms()
        THTitle.Text = ""
        THArtist.Text = ""
        THGenre.Text = ""
        THBPM.Value = 120
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0
        CHRank.SelectedIndex = 3
        THPlayLevel.Text = ""
        THSubTitle.Text = ""
        THSubArtist.Text = ""
        THStageFile.Text = ""
        THBanner.Text = ""
        THBackBMP.Text = ""
        CHDifficulty.SelectedIndex = 0
        THExRank.Text = ""
        THTotal.Text = ""
        THComment.Text = ""
        'THLnType.Text = "1"
        CHLnObj.SelectedIndex = 0

        TExpansion.Text = ""

        LBeat.Items.Clear()
        For i = 0 To 999
            MeasureLength(i) = 192.0R
            MeasureBottom(i) = i * 192.0R
            LBeat.Items.Add(Add3Zeros(i) & ": 1 ( 4 / 4 )")
        Next
    End Sub

    Private Sub InitializeOpenBms()
        CHPlayer.SelectedIndex = 0
        'THLnType.Text = ""
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()),
                                                     _supportedFileExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragLeave(sender As Object, e As EventArgs) Handles Me.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, _supportedFileExtension)
        If xPath.Length > 0 Then
            Dim xProg As New fLoadFileProgress(xPath, _isSaved)
            xProg.ShowDialog(Me)
        End If

        RefreshPanelAll()
    End Sub

    Private Sub SetFullScreen(value As Boolean)
        If value Then
            If Me.WindowState = FormWindowState.Minimized Then Exit Sub

            SuspendLayout()
            _previousWindowPosition.Location = Me.Location
            _previousWindowPosition.Size = Me.Size
            _previousWindowState = Me.WindowState

            WindowState = FormWindowState.Normal
            FormBorderStyle = FormBorderStyle.None
            WindowState = FormWindowState.Maximized
            ToolStripContainer1.TopToolStripPanelVisible = False

            ResumeLayout()
            IsFullscreen = True
        Else
            SuspendLayout()
            FormBorderStyle = FormBorderStyle.Sizable
            ToolStripContainer1.TopToolStripPanelVisible = True
            WindowState = FormWindowState.Normal

            WindowState = _previousWindowState
            If Me.WindowState = FormWindowState.Normal Then
                Location = _previousWindowPosition.Location
                Size = _previousWindowPosition.Size
            End If

            ResumeLayout()
            IsFullscreen = False
        End If
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F11
                SetFullScreen(Not IsFullscreen)
        End Select
    End Sub

    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Friend Sub ReadFile(xPath As String)
        Select Case LCase(Path.GetExtension(xPath))
            Case ".bms", ".bme", ".bml", ".pms", ".txt"
                OpenBms(My.Computer.FileSystem.ReadAllText(xPath, _textEncoding))
                ClearUndo()
                NewRecent(xPath)
                SetFileName(xPath)
                SetIsSaved(True)

            Case ".sm"
                If OpenSm(My.Computer.FileSystem.ReadAllText(xPath, _textEncoding)) Then Return
                _initPath = ExcludeFileName(xPath)
                ClearUndo()
                SetFileName("Untitled.bms")
                SetIsSaved(False)

            Case ".ibmsc"
                OpeniBmsc(xPath)
                _initPath = ExcludeFileName(xPath)
                NewRecent(xPath)
                SetFileName("Imported_" & GetFileName(xPath))
                SetIsSaved(False)

        End Select
    End Sub


    Public Function Gcd(numA As Double, numB As Double) As Double
        Dim xNMax As Double = numA
        Dim xNMin As Double = numB
        If numA < numB Then
            xNMax = numB
            xNMin = numA
        End If
        Do While xNMin >= _bmsGridLimit
            Gcd = xNMax - Math.Floor(xNMax / xNMin) * xNMin
            xNMax = xNMin
            xNMin = Gcd
        Loop
        Gcd = xNMax
    End Function

    <DllImport("user32.dll")>
    Private Shared Function LoadCursorFromFile(fileName As String) As IntPtr
    End Function

    Public Shared Function ActuallyLoadCursor(path As String) As Cursor
        Return New Cursor(LoadCursorFromFile(path))
    End Function

    Private Sub Unload() Handles MyBase.Disposed
        Audio.Finalize()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'On Error Resume Next
        TopMost = True
        SuspendLayout()
        Visible = False

        SetFileName(_fileName)

        InitializeNewBms()

        Try
            Dim xTempFileName As String = RandomFileName(".cur")
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeDown, False)
            Dim xDownCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeLeft, False)
            Dim xLeftCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            My.Computer.FileSystem.WriteAllBytes(xTempFileName, My.Resources.CursorResizeRight, False)
            Dim xRightCursor As Cursor = ActuallyLoadCursor(xTempFileName)
            File.Delete(xTempFileName)

            POWAVResizer.Cursor = xDownCursor
            POBMPResizer.Cursor = xDownCursor
            POBeatResizer.Cursor = xDownCursor
            POExpansionResizer.Cursor = xDownCursor

            POptionsResizer.Cursor = xLeftCursor

            'SpL.Cursor = xRightCursor
            'SpR.Cursor = xLeftCursor
        Catch ex As Exception

        End Try

        _spMain = {EditorPanel1}
        For Each panel In _spMain
            panel.Init(Me, _vo)
        Next
        ' PMain.BringToFront()

        Dim i As Integer

        sUndo(0) = New UndoRedo.NoOperation
        sUndo(1) = New UndoRedo.NoOperation
        sRedo(0) = New UndoRedo.NoOperation
        sRedo(1) = New UndoRedo.NoOperation
        sI = 0

        LWAV.Items.Clear()
        LBMP.Items.Clear()
        For i = 1 To 1295
            LWAV.Items.Add(C10to36(i) & ":")
            LBMP.Items.Add(C10to36(i) & ":")
        Next
        LWAV.SelectedIndex = 0
        LBMP.SelectedIndex = 0
        CHPlayer.SelectedIndex = 0


        TBLangRefresh_Click(TBLangRefresh, Nothing)
        TBThemeRefresh_Click(TBThemeRefresh, Nothing)

        POHeaderPart2.Visible = False
        POGridPart2.Visible = False
        POWaveFormPart2.Visible = False

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml") Then
            LoadSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml")
        End If
        'On Error GoTo 0
        SetIsSaved(True)

        Dim xStr() As String = Environment.GetCommandLineArgs

        If xStr.Length = 2 Then
            ReadFile(xStr(1))
            If _
                LCase(Path.GetExtension(xStr(1))) = ".ibmsc" AndAlso
                GetFileName(xStr(1)).StartsWith("AutoSave_", True, Nothing) Then GoTo 1000
        End If

        'pIsSaved.Visible = Not IsSaved
        _isInitializing = False

        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then GoTo 1000
        Dim xFiles() As FileInfo =
                My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath).GetFiles("AutoSave_*.IBMSC")
        If xFiles Is Nothing OrElse xFiles.Length = 0 Then GoTo 1000

        'Me.TopMost = True
        If _
            MsgBox(Replace(Strings.Messages.RestoreAutosavedFile, "{}", xFiles.Length),
                   MsgBoxStyle.YesNo Or MsgBoxStyle.MsgBoxSetForeground) = MsgBoxResult.Yes Then
            For Each xF As FileInfo In xFiles
                'MsgBox(xF.FullName)
                Process.Start(Application.ExecutablePath, """" & xF.FullName & """")
            Next
        End If

        For Each xF As FileInfo In xFiles
            ReDim Preserve _pTempFileNames(UBound(_pTempFileNames) + 1)
            _pTempFileNames(UBound(_pTempFileNames)) = xF.FullName
        Next

1000:
        _isInitializing = False
        PoStatusRefresh()
        ResumeLayout()

        _tempResize = WindowState
        TopMost = False

        Visible = True
    End Sub

    Private Sub UpdatePairing()
        Dim i As Integer, j As Integer

        If NtInput Then
            For i = 0 To UBound(Notes)
                Notes(i).HasError = False
                Notes(i).LNPair = 0
                If Notes(i).Length < 0 Then Notes(i).Length = 0
            Next

            For i = 1 To UBound(Notes)
                If Notes(i).Length <> 0 Then
                    For j = i + 1 To UBound(Notes)
                        If Notes(j).VPosition > Notes(i).VPosition + Notes(i).Length Then Exit For
                        If Notes(j).ColumnIndex = Notes(i).ColumnIndex Then Notes(j).HasError = True
                    Next
                Else
                    For j = i + 1 To UBound(Notes)
                        If Notes(j).VPosition > Notes(i).VPosition Then Exit For
                        If Notes(j).ColumnIndex = Notes(i).ColumnIndex Then Notes(j).HasError = True
                    Next

                    If Notes(i).Value \ 10000 = LnObj AndAlso Not Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then
                        For j = i - 1 To 1 Step -1
                            If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                            If Notes(j).Hidden Then Continue For

                            If Notes(j).Length <> 0 OrElse Notes(j).Value \ 10000 = LnObj Then
                                Notes(i).HasError = True
                            Else
                                Notes(i).LNPair = j
                                Notes(j).LNPair = i
                            End If
                            Exit For
                        Next
                        If j = 0 Then
                            Notes(i).HasError = True
                        End If
                    End If
                End If
            Next

        Else
            For i = 0 To UBound(Notes)
                Notes(i).HasError = False
                Notes(i).LNPair = 0
            Next

            For i = 1 To UBound(Notes)

                If Notes(i).LongNote Then
                    'LongNote: If overlapping a note, then error.
                    '          Else if already matched by a LongNote below, then match it.
                    '          Otherwise match anything above.
                    '              If ShortNote above then error on above.
                    '          If nothing above then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        If Notes(j).VPosition = Notes(i).VPosition Then
                            Notes(i).HasError = True
                            GoTo EndSearch
                        ElseIf Notes(j).LongNote And Notes(j).LNPair = i Then
                            Notes(i).LNPair = j
                            GoTo EndSearch
                        Else
                            Exit For
                        End If
                    Next

                    For j = i + 1 To UBound(Notes)
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        Notes(i).LNPair = j
                        Notes(j).LNPair = i
                        If Not Notes(j).LongNote AndAlso Notes(j).Value \ 10000 <> LnObj Then
                            Notes(j).HasError = True
                        End If
                        Exit For
                    Next

                    If j = UBound(Notes) + 1 Then
                        Notes(i).HasError = True
                    End If
EndSearch:

                ElseIf Notes(i).Value \ 10000 = LnObj And
                       Not Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then
                    'LnObj: Match anything below.
                    '           If matching a LongNote not matching back, then error on below.
                    '           If overlapping a note, then error.
                    '           If mathcing a LnObj below, then error on below.
                    '       If nothing below, then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(i).ColumnIndex <> Notes(j).ColumnIndex Then Continue For
                        If Notes(j).LNPair <> 0 And Notes(j).LNPair <> i Then
                            Notes(j).HasError = True
                        End If
                        Notes(i).LNPair = j
                        Notes(j).LNPair = i
                        If Notes(i).VPosition = Notes(j).VPosition Then
                            Notes(i).HasError = True
                        End If
                        If Notes(j).Value \ 10000 = LnObj Then
                            Notes(j).HasError = True
                        End If
                        Exit For
                    Next

                    If j = 0 Then
                        Notes(i).HasError = True
                    End If

                Else
                    'ShortNote: If overlapping a note, then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(j).VPosition < Notes(i).VPosition Then Exit For
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        Notes(i).HasError = True
                        Exit For
                    Next

                End If
            Next


        End If
    End Sub

    ' az: Handle zoom in/out. Should work with any of the three splitters.
    Private Sub PMain_Scroll(sender As Object, e As MouseEventArgs) Handles _
            EditorPanel1.MouseWheel

        If Not My.Computer.Keyboard.CtrlKeyDown Then Exit Sub
        Dim dv = Math.Round(CGHeight2.Value + e.Delta / 120)
        CGHeight2.Value = Math.Min(CGHeight2.Maximum, Math.Max(CGHeight2.Minimum, dv))
        CGHeight.Value = CGHeight2.Value / 4
    End Sub

    Public Sub ExceptionSave(path As String)
        SaveiBmsc(path)
    End Sub

    ''' <summary>
    '''     True if pressed cancel. False elsewise.
    ''' </summary>
    ''' <returns>True if pressed cancel. False elsewise.</returns>

    Private Function ClosingPopSave() As Boolean
        If Not _isSaved Then
            Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveOnExit,
                                                 MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes Then
                If ExcludeFileName(_fileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = _initPath

                    If xDSave.ShowDialog = DialogResult.Cancel Then Return True
                    SetFileName(xDSave.FileName)
                End If
                Dim xStrAll As String = SaveBms()
                My.Computer.FileSystem.WriteAllText(_fileName, xStrAll, False, _textEncoding)
                NewRecent(_fileName)
                If _beepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then Return True
        End If
        Return False
    End Function

    Private Sub TBNew_Click(sender As Object, e As EventArgs) Handles TBNew.Click, mnNew.Click

        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        ClearUndo()
        InitializeNewBms()

        ReDim Notes(0)
        ReDim _mColumn(999)
        ReDim BmsWAV(1295)
        ReDim BmsBMP(1295)
        ReDim BmsBPM(1295)    'x10000
        ReDim BmsSTOP(1295)
        ReDim BmsSCROLL(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        Notes(0) = New Note With {
            .ColumnIndex = ColumnType.BPM,
            .VPosition = -1,
            .Value = 1200000
        }

        THBPM.Value = 120

        LWAV.Items.Clear()
        LBMP.Items.Clear()
        Dim i As Integer
        For i = 1 To 1295
            LWAV.Items.Add(C10to36(i) & ": " & BmsWAV(i))
            LBMP.Items.Add(C10to36(i) & ": " & BmsBMP(i))
        Next
        LWAV.SelectedIndex = 0
        LBMP.SelectedIndex = 0

        SetFileName("Untitled.bms")
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        CalculateTotalPlayableNotes()

        RefreshPanelAll()
        PoStatusRefresh()
    End Sub


    Private Sub TBOpen_ButtonClick(sender As Object, e As EventArgs) Handles TBOpen.ButtonClick, mnOpen.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog With {
                .Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt",
                .DefaultExt = "bms",
                .InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
                }

        If xDOpen.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDOpen.FileName)
        OpenBms(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, _textEncoding))
        ClearUndo()
        SetFileName(xDOpen.FileName)
        NewRecent(_fileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBImportIBMSC_Click(sender As Object, e As EventArgs) Handles TBImportIBMSC.Click, mnImportIBMSC.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Return

        Dim xDOpen As New OpenFileDialog With {
                .Filter = Strings.FileType.IBMSC & "|*.ibmsc",
                .DefaultExt = "ibmsc",
                .InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
                }

        If xDOpen.ShowDialog = DialogResult.Cancel Then Return
        _initPath = ExcludeFileName(xDOpen.FileName)
        SetFileName("Imported_" & GetFileName(xDOpen.FileName))
        OpeniBmsc(xDOpen.FileName)
        NewRecent(xDOpen.FileName)
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBImportSM_Click(sender As Object, e As EventArgs) Handles TBImportSM.Click, mnImportSM.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog With {
                .Filter = Strings.FileType.SM & "|*.sm",
                .DefaultExt = "sm",
                .InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
                }

        If xDOpen.ShowDialog = DialogResult.Cancel Then Exit Sub
        If OpenSm(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, _textEncoding)) Then Exit Sub
        _initPath = ExcludeFileName(xDOpen.FileName)
        SetFileName("Untitled.bms")
        ClearUndo()
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBSave_ButtonClick(sender As Object, e As EventArgs) Handles TBSave.ButtonClick, mnSave.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1

        If ExcludeFileName(_fileName) = "" Then
            Dim xDSave As New SaveFileDialog With {
                    .Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                              Strings.FileType.BMS & "|*.bms|" &
                              Strings.FileType.BME & "|*.bme|" &
                              Strings.FileType.BML & "|*.bml|" &
                              Strings.FileType.PMS & "|*.pms|" &
                              Strings.FileType.TXT & "|*.txt|" &
                              Strings.FileType._all & "|*.*",
                    .DefaultExt = "bms",
                    .InitialDirectory = _initPath
                    }

            If xDSave.ShowDialog = DialogResult.Cancel Then Exit Sub
            _initPath = ExcludeFileName(xDSave.FileName)
            SetFileName(xDSave.FileName)
        End If
        Dim xStrAll As String = SaveBms()
        My.Computer.FileSystem.WriteAllText(_fileName, xStrAll, False, _textEncoding)
        NewRecent(_fileName)
        SetFileName(_fileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
        If _beepWhileSaved Then Beep()
    End Sub

    Private Sub TBSaveAs_Click(sender As Object, e As EventArgs) Handles TBSaveAs.Click, mnSaveAs.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1

        Dim xDSave As New SaveFileDialog With {
                .Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                          Strings.FileType.BMS & "|*.bms|" &
                          Strings.FileType.BME & "|*.bme|" &
                          Strings.FileType.BML & "|*.bml|" &
                          Strings.FileType.PMS & "|*.pms|" &
                          Strings.FileType.TXT & "|*.txt|" &
                          Strings.FileType._all & "|*.*",
                .DefaultExt = "bms",
                .InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
                }

        If xDSave.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDSave.FileName)
        SetFileName(xDSave.FileName)
        Dim xStrAll As String = SaveBms()
        My.Computer.FileSystem.WriteAllText(_fileName, xStrAll, False, _textEncoding)
        NewRecent(_fileName)
        SetFileName(_fileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
        If _beepWhileSaved Then Beep()
    End Sub

    Private Sub TBExport_Click(sender As Object, e As EventArgs) Handles TBExport.Click, mnExport.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1

        Dim xDSave As New SaveFileDialog With {
                .Filter = Strings.FileType.IBMSC & "|*.ibmsc",
                .DefaultExt = "ibmsc",
                .InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
                }
        If xDSave.ShowDialog = DialogResult.Cancel Then Exit Sub

        SaveiBmsc(xDSave.FileName)
        'My.Computer.FileSystem.WriteAllText(xDSave.FileName, xStrAll, False, TextEncoding)
        NewRecent(_fileName)
        If _beepWhileSaved Then Beep()
    End Sub


    'Private Sub PanelVerticalScrollChanged(sender As Object,
    '                                       e As EventArgs) Handles 
    '                                                               LeftPanelScroll.ValueChanged
    '    Dim iI As Integer = sender.Tag

    '    ' az: We got a wheel event when we're zooming in/out
    '    If My.Computer.Keyboard.CtrlKeyDown Then
    '        sender.Value = FocusedPanel.LastVerticalScroll ' Undo the scroll
    '        Exit Sub
    '    End If

    '    Dim currentPanel = _spMain(iI)
    '    If iI = PanelFocus And
    '       Not State.Mouse.LastMouseDownLocation = New Point(-1, -1) And
    '       Not FocusedPanel.LastVerticalScroll = -1 Then
    '        State.Mouse.LastMouseDownLocation.Y += (FocusedPanel.LastVerticalScroll - sender.Value) * Grid.HeightScale
    '    End If


    '    If _spLock((iI + 1) Mod 3) Then
    '        Dim verticalScroll As Integer = currentPanel.VerticalScroll + _spDiff(iI)
    '        If verticalScroll > 0 Then verticalScroll = 0
    '        If verticalScroll < MainPanelScroll.Minimum Then verticalScroll = MainPanelScroll.Minimum
    '        Select Case iI
    '            Case 0 : MainPanelScroll.Value = verticalScroll
    '            Case 1 : RightPanelScroll.Value = verticalScroll
    '            Case 2 : LeftPanelScroll.Value = verticalScroll
    '        End Select
    '    End If

    '    If _spLock((iI + 2) Mod 3) Then
    '        Dim verticalScroll As Integer = currentPanel.VerticalScroll - _spDiff((iI + 2) Mod 3)
    '        If verticalScroll > 0 Then verticalScroll = 0
    '        If verticalScroll < MainPanelScroll.Minimum Then verticalScroll = MainPanelScroll.Minimum
    '        Select Case iI
    '            Case 0 : RightPanelScroll.Value = verticalScroll
    '            Case 1 : LeftPanelScroll.Value = verticalScroll
    '            Case 2 : MainPanelScroll.Value = verticalScroll
    '        End Select
    '    End If

    '    _spDiff(iI) = _spMain((iI + 1) Mod 3).VerticalScroll - currentPanel.VerticalScroll
    '    _spDiff((iI + 2) Mod 3) = currentPanel.VerticalScroll - _spMain((iI + 2) Mod 3).VerticalScroll
    'End Sub

    Private Sub cVSLock_CheckedChanged(sender As Object, e As EventArgs) _
        Handles cVSLockL.CheckedChanged, cVSLock.CheckedChanged, cVSLockR.CheckedChanged
        Dim iI As Integer = sender.Tag
        _spLock(iI) = sender.Checked
        If Not _spLock(iI) Then Return

        Dim currentPanel = _spMain(iI)
        _spDiff(iI) = _spMain((iI + 1) Mod 3).VerticalPosition - currentPanel.VerticalPosition
        _spDiff((iI + 2) Mod 3) = currentPanel.VerticalPosition - _spMain((iI + 2) Mod 3).VerticalPosition
    End Sub

    Private Sub HsValueChanged(sender As Object, e As EventArgs)

        If Not State.Mouse.LastMouseDownLocation = New Point(-1, -1) Then
            State.Mouse.LastMouseDownLocation.X += (FocusedPanel.LastHorizontalScroll - sender.Value) * Grid.WidthScale
        End If
    End Sub

    Private Sub TBSelect_Click(sender As Object, e As EventArgs) Handles TBSelect.Click, mnSelect.Click
        TBSelect.Checked = True
        TBWrite.Checked = False
        TBTimeSelect.Checked = False
        mnSelect.Checked = True
        mnWrite.Checked = False
        mnTimeSelect.Checked = False

        FStatus2.Visible = False
        FStatus.Visible = True

        State.Mouse.CurrentMouseColumn = -1
        State.Mouse.CurrentMouseRow = -1
        LnDisplayLength = 0

        State.TimeSelect.StartPoint = MeasureBottom(MeasureAtDisplacement(-FocusedPanel.VerticalPosition) + 1)
        State.TimeSelect.EndPointLength = 0

        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub TBWrite_Click(sender As Object, e As EventArgs) Handles TBWrite.Click, mnWrite.Click
        TBSelect.Checked = False
        TBWrite.Checked = True
        TBTimeSelect.Checked = False
        mnSelect.Checked = False
        mnWrite.Checked = True
        mnTimeSelect.Checked = False

        FStatus2.Visible = False
        FStatus.Visible = True

        State.Mouse.CurrentMouseColumn = -1
        State.Mouse.CurrentMouseRow = -1
        LnDisplayLength = 0

        State.TimeSelect.StartPoint = MeasureBottom(MeasureAtDisplacement(-FocusedPanel.VerticalPosition) + 1)
        State.TimeSelect.EndPointLength = 0

        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub TBPostEffects_Click(sender As Object, e As EventArgs) Handles TBTimeSelect.Click, mnTimeSelect.Click
        TBSelect.Checked = False
        TBWrite.Checked = False
        TBTimeSelect.Checked = True
        mnSelect.Checked = False
        mnWrite.Checked = False
        mnTimeSelect.Checked = True

        FStatus.Visible = False
        FStatus2.Visible = True

        State.TimeSelect.MouseOverLine = 0
        State.Mouse.CurrentMouseColumn = -1
        State.Mouse.CurrentMouseRow = -1
        LnDisplayLength = 0
        State.TimeSelect.ValidateSelection(GetMaxVPosition())

        Dim i As Integer
        For i = 0 To UBound(Notes)
            Notes(i).Selected = False
        Next
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub CGHeight_ValueChanged(sender As Object, e As EventArgs) Handles CGHeight.ValueChanged
        Grid.HeightScale = CSng(CGHeight.Value)
        CGHeight2.Value = IIf(CGHeight.Value * 4 < CGHeight2.Maximum, CDec(CGHeight.Value * 4), CGHeight2.Maximum)
        RefreshPanelAll()
    End Sub

    Private Sub CGHeight2_Scroll(sender As Object, e As EventArgs) Handles CGHeight2.Scroll
        CGHeight.Value = CGHeight2.Value / 4
    End Sub

    Private Sub CGWidth_ValueChanged(sender As Object, e As EventArgs) Handles CGWidth.ValueChanged
        Grid.WidthScale = CSng(CGWidth.Value)
        CGWidth2.Value = IIf(CGWidth.Value * 4 < CGWidth2.Maximum, CDec(CGWidth.Value * 4), CGWidth2.Maximum)

        For Each panel In _spMain
            panel.SetWidthScale(Grid.WidthScale)
        Next

        RefreshPanelAll()
    End Sub

    Private Sub CGWidth2_Scroll(sender As Object, e As EventArgs) Handles CGWidth2.Scroll
        CGWidth.Value = CGWidth2.Value / 4
    End Sub

    Private Sub CGDivide_ValueChanged(sender As Object, e As EventArgs) Handles CGDivide.ValueChanged
        Grid.Divider = CGDivide.Value
        RefreshPanelAll()
    End Sub

    Private Sub CGSub_ValueChanged(sender As Object, e As EventArgs) Handles CGSub.ValueChanged
        Grid.Subdivider = CGSub.Value
        RefreshPanelAll()
    End Sub

    Private Sub BGSlash_Click(sender As Object, e As EventArgs) Handles BGSlash.Click
        Dim xd As Integer = Val(InputBox(Strings.Messages.PromptSlashValue, , Grid.Slash))
        If xd = 0 Then Exit Sub
        If xd > CGDivide.Maximum Then xd = CGDivide.Maximum
        If xd < CGDivide.Minimum Then xd = CGDivide.Minimum
        Grid.Slash = xd
    End Sub


    Private Sub CGSnap_CheckedChanged(sender As Object, e As EventArgs) Handles CGSnap.CheckedChanged
        Grid.IsSnapEnabled = CGSnap.Checked
        RefreshPanelAll()
    End Sub


    Private Sub TimerMiddle_Tick(sender As Object, e As EventArgs) Handles TimerMiddle.Tick
        If Not State.Mouse.MiddleButtonClicked Then
            TimerMiddle.Enabled = False
            Return
        End If

        Dim xAmt = (Cursor.Position.X - State.Mouse.MiddleButtonLocation.X) / 10 / Grid.WidthScale
        Dim yAmt = (Cursor.Position.Y - State.Mouse.MiddleButtonLocation.Y) / 5 / Grid.HeightScale

        FocusedPanel.ScrollChart(xAmt, yAmt)

        Dim xMeArgs As New MouseEventArgs(MouseButtons.Left, 0,
                                          State.Mouse.MouseMoveStatus.X,
                                          State.Mouse.MouseMoveStatus.Y,
                                          0)
        FocusedPanel.MouseMoveEvent(Me, xMeArgs)
    End Sub

    Private Sub ValidateWavListView()
        Try
            Dim xRect As Rectangle = LWAV.GetItemRectangle(LWAV.SelectedIndex)
            If xRect.Top + xRect.Height > LWAV.DisplayRectangle.Height Then SendMessage(LWAV.Handle, &H115, 1, 0)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LWAV_Click(sender As Object, e As EventArgs) Handles LWAV.Click
        If TBWrite.Checked Then FSW.Text = C10to36(LWAV.SelectedIndex + 1)

        PreviewNote("", True)
        If Not _previewOnClick Then Exit Sub
        If BmsWAV(LWAV.SelectedIndex + 1) = "" Then Exit Sub

        Dim xFileLocation As String = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName)) & "\" &
                                      BmsWAV(LWAV.SelectedIndex + 1)
        PreviewNote(xFileLocation, False)
    End Sub

    Private Sub LWAV_DoubleClick(sender As Object, e As EventArgs) Handles LWAV.DoubleClick
        Dim xDwav As New OpenFileDialog
        xDwav.DefaultExt = "wav"
        xDwav.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDwav.InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))

        If xDwav.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDwav.FileName)
        BmsWAV(LWAV.SelectedIndex + 1) = GetFileName(xDwav.FileName)
        LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": " & GetFileName(xDwav.FileName)
        If _isSaved Then SetIsSaved(False)
    End Sub

    Private Sub LWAV_KeyDown(sender As Object, e As KeyEventArgs) Handles LWAV.KeyDown
        Select Case e.KeyCode
            Case Keys.Delete
                BmsWAV(LWAV.SelectedIndex + 1) = ""
                LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": "
                If _isSaved Then SetIsSaved(False)
        End Select
    End Sub

    Private Sub LBMP_DoubleClick(sender As Object, e As EventArgs) Handles LBMP.DoubleClick
        Dim xDbmp As New OpenFileDialog
        xDbmp.DefaultExt = "bmp"
        xDbmp.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.jpeg;.gif|" &
                       Strings.FileType._movie & "|*.mpg;*.m1v;*.m2v;*.avi;*.mp4;*.m4v;*.wmv;*.webm|" &
                       Strings.FileType.BMP & "|*.bmp|" &
                       Strings.FileType.PNG & "|*.png|" &
                       Strings.FileType.JPG & "|*.jpg;*.jpeg|" &
                       Strings.FileType.GIF & "|*.gif|" &
                       Strings.FileType.MP4 & "|*.mp4;*.m4v|" &
                       Strings.FileType.AVI & "|*.avi|" &
                       Strings.FileType.MPG & "|*.mpg;*.m1v;*.m2v|" &
                       Strings.FileType.WMV & "|*.wmv|" &
                       Strings.FileType.WEBM & "|*.webm|" &
                       Strings.FileType._all & "|*.*"
        xDbmp.InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))

        If xDbmp.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDbmp.FileName)
        BmsBMP(LBMP.SelectedIndex + 1) = GetFileName(xDbmp.FileName)
        LBMP.Items.Item(LBMP.SelectedIndex) = C10to36(LBMP.SelectedIndex + 1) & ": " & GetFileName(xDbmp.FileName)
        If _isSaved Then SetIsSaved(False)
    End Sub

    Private Sub LBMP_KeyDown(sender As Object, e As KeyEventArgs) Handles LBMP.KeyDown
        Select Case e.KeyCode
            Case Keys.Delete
                BmsBMP(LBMP.SelectedIndex + 1) = ""
                LBMP.Items.Item(LBMP.SelectedIndex) = C10to36(LBMP.SelectedIndex + 1) & ": "
                If _isSaved Then SetIsSaved(False)
        End Select
    End Sub

    Private Sub TBErrorCheck_Click(sender As Object, e As EventArgs) Handles TBErrorCheck.Click, mnErrorCheck.Click
        ErrorCheck = sender.Checked
        TBErrorCheck.Checked = ErrorCheck
        mnErrorCheck.Checked = ErrorCheck
        TBErrorCheck.Image = IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN)
        mnErrorCheck.Image = IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN)
        RefreshPanelAll()
    End Sub

    Private Sub TBPreviewOnClick_Click(sender As Object, e As EventArgs) _
        Handles TBPreviewOnClick.Click, mnPreviewOnClick.Click
        PreviewNote("", True)
        _previewOnClick = sender.Checked
        TBPreviewOnClick.Checked = _previewOnClick
        mnPreviewOnClick.Checked = _previewOnClick
        TBPreviewOnClick.Image = IIf(_previewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN)
        mnPreviewOnClick.Image = IIf(_previewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN)
    End Sub

    Private Sub TBShowFileName_Click(sender As Object, e As EventArgs) _
        Handles TBShowFileName.Click, mnShowFileName.Click
        ShowFileName = sender.Checked
        TBShowFileName.Checked = ShowFileName
        mnShowFileName.Checked = ShowFileName
        TBShowFileName.Image = IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN)
        mnShowFileName.Image = IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN)
        RefreshPanelAll()
    End Sub

    Private Sub TBCut_Click(sender As Object, e As EventArgs) Handles TBCut.Click, mnCut.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        RedoRemoveNoteSelected(Notes, xUndo, xRedo)
        'Dim xRedo As String = sCmdKDs()
        'Dim xUndo As String = sCmdKs(True)

        CopyNotes(False)
        RemoveNotes(False)
        AddUndoChain(xUndo, xBaseRedo.Next)

        ValidateNotesArray()
        RefreshPanelAll()
        PoStatusRefresh()

    End Sub

    Private Sub TBCopy_Click(sender As Object, e As EventArgs) Handles TBCopy.Click, mnCopy.Click
        CopyNotes()
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub TBPaste_Click(sender As Object, e As EventArgs) Handles TBPaste.Click, mnPaste.Click
        AddNotesFromClipboard()

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        RedoAddNoteSelected(Notes, xUndo, xRedo)
        AddUndoChain(xUndo, xBaseRedo.Next)

        'AddUndo(sCmdKDs(), sCmdKs(True))

        ValidateNotesArray()
        RefreshPanelAll()
        PoStatusRefresh()

    End Sub

    'Private Function pArgPath(ByVal I As Integer)
    '    Return Mid(pArgs(I), 1, InStr(pArgs(I), vbCrLf) - 1)
    'End Function

    Private Function GetFileName(s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        Return Mid(s, IIf(fslash > bslash, fslash, bslash) + 1)
    End Function

    Private Function ExcludeFileName(s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        If (bslash Or fslash) = 0 Then Return ""
        Return Mid(s, 1, IIf(fslash > bslash, fslash, bslash) - 1)
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

        AddTempFileList(xFileName)
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

        AddTempFileList(xFileName)

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

    Private Sub AddTempFileList(s As String)
        Dim xAdd = True
        If _pTempFileNames IsNot Nothing Then
            For Each xStr1 As String In _pTempFileNames
                If xStr1 = s Then xAdd = False : Exit For
            Next
        End If

        If xAdd Then
            ReDim Preserve _pTempFileNames(UBound(_pTempFileNames) + 1)
            _pTempFileNames(UBound(_pTempFileNames)) = s
        End If
    End Sub

    Private Sub TBStatistics_Click(sender As Object, e As EventArgs) Handles TBStatistics.Click, mnStatistics.Click
        ValidateNotesArray()

        Dim data(6, 5) As Integer
        For i = 1 To UBound(Notes)
            With Notes(i)
                Dim row As Integer = -1
                Select Case .ColumnIndex
                    Case ColumnType.BPM : row = 0
                    Case ColumnType.STOPS : row = 1
                    Case ColumnType.SCROLLS : row = 2
                    Case ColumnType.A1, ColumnType.A2, ColumnType.A3, ColumnType.A4, ColumnType.A5, ColumnType.A6,
                        ColumnType.A7, ColumnType.A8 : row = 3
                    Case ColumnType.D1, ColumnType.D2, ColumnType.D3, ColumnType.D4, ColumnType.D5, ColumnType.D6,
                        ColumnType.D7, ColumnType.D8 : row = 4
                    Case Is >= ColumnType.BGM : row = 5
                    Case Else : row = 6
                End Select


StartCount:     If Not NtInput Then
                    If Not .LongNote Then data(row, 0) += 1
                    If .LongNote Then data(row, 1) += 1
                    If .Value \ 10000 = LnObj Then data(row, 2) += 1
                    If .Hidden Then data(row, 3) += 1
                    If .HasError Then data(row, 4) += 1
                    data(row, 5) += 1

                Else
                    Dim noteUnit = 1
                    If .Length = 0 Then data(row, 0) += 1
                    If .Length <> 0 Then data(row, 1) += 2 : noteUnit = 2

                    If .Value \ 10000 = LnObj Then data(row, 2) += noteUnit
                    If .Hidden Then data(row, 3) += noteUnit
                    If .HasError Then data(row, 4) += noteUnit
                    data(row, 5) += noteUnit

                End If

                If row <> 6 Then row = 6 : GoTo StartCount
            End With
        Next

        Dim dStat As New dgStatistics(data)
        dStat.ShowDialog()
    End Sub

    ''' <summary>
    '''     Remark: Pls sort and updatepairing before this process.
    ''' </summary>
    Public Sub CalculateTotalPlayableNotes()
        Dim i As Integer
        Dim xIAll = 0

        If Not NtInput Then
            For i = 1 To UBound(Notes)
                If Notes(i).ColumnIndex >= ColumnType.A1 And Notes(i).ColumnIndex <= ColumnType.A8 Then xIAll += 1
            Next

        Else
            For i = 1 To UBound(Notes)
                If Notes(i).ColumnIndex >= ColumnType.A1 And Notes(i).ColumnIndex <= ColumnType.A8 Then
                    xIAll += 1
                    If Notes(i).Length <> 0 Then xIAll += 1
                End If
            Next
        End If

        TBStatistics.Text = xIAll
    End Sub

    Public Function GetMouseVPosition(Optional snap As Boolean = True)
        Dim panHeight = FocusedPanel.Height
        Dim panDisplacement = FocusedPanel.VerticalPosition
        Dim vpos = (panHeight - panDisplacement * Grid.HeightScale - State.Mouse.MouseMoveStatus.Y - 1) / Grid.HeightScale
        If snap Then
            Return FocusedPanel.SnapToGrid(vpos)
        Else
            Return vpos
        End If
    End Function

    Public Sub PoStatusRefresh()

        If TBSelect.Checked Then
            Dim i As Integer = State.Mouse.CurrentHoveredNoteIndex
            If i < 0 Then

                UpdateMouseRowAndColumn()

                Dim xMeasure As Integer = MeasureAtDisplacement(State.Mouse.CurrentMouseRow)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = State.Mouse.CurrentMouseRow - MeasureBottom(xMeasure)
                Dim xGcd As Double = Gcd(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGcd).ToString & " / " & CInt(xMLength / xGcd).ToString & "  "
                FSP4.Text = State.Mouse.CurrentMouseRow.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
                FSC.Text = Columns.GetName(State.Mouse.CurrentMouseColumn)
                FSW.Text = ""
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = ""
                FSH.Text = ""
                FSE.Text = ""

            Else
                Dim xMeasure As Integer = MeasureAtDisplacement(Notes(i).VPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = Notes(i).VPosition - MeasureBottom(xMeasure)
                Dim xGcd As Double = Gcd(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGcd).ToString & " / " & CInt(xMLength / xGcd).ToString & "  "
                FSP4.Text = Notes(i).VPosition.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
                FSC.Text = Columns.GetName(Notes(i).ColumnIndex)
                FSW.Text = IIf(Columns.IsColumnNumeric(Notes(i).ColumnIndex),
                               Notes(i).Value / 10000,
                               C10to36(Notes(i).Value \ 10000))
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = IIf(NtInput, Strings.StatusBar.Length & " = " & Notes(i).Length,
                               IIf(Notes(i).LongNote, Strings.StatusBar.LongNote, ""))
                FSH.Text = IIf(Notes(i).Hidden, Strings.StatusBar.Hidden, "")
                FSE.Text = IIf(Notes(i).HasError, Strings.StatusBar.Err, "")

            End If

        ElseIf TBWrite.Checked Then
            If State.Mouse.CurrentMouseColumn < 0 Then Exit Sub

            Dim xMeasure As Integer = MeasureAtDisplacement(State.Mouse.CurrentMouseRow)
            Dim xMLength As Double = MeasureLength(xMeasure)
            Dim xVposMod As Double = State.Mouse.CurrentMouseRow - MeasureBottom(xMeasure)
            Dim xGcd As Double = Gcd(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

            FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
            FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
            FSP3.Text = CInt(xVposMod / xGcd).ToString & " / " & CInt(xMLength / xGcd).ToString & "  "
            FSP4.Text = State.Mouse.CurrentMouseRow.ToString() & "  "
            TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
            FSC.Text = Columns.GetName(State.Mouse.CurrentMouseColumn)
            FSW.Text = C10to36(LWAV.SelectedIndex + 1)
            FSM.Text = Add3Zeros(xMeasure)
            FST.Text = IIf(NtInput, LnDisplayLength,
                           IIf(My.Computer.Keyboard.ShiftKeyDown, Strings.StatusBar.LongNote, ""))
            FSH.Text = IIf(My.Computer.Keyboard.CtrlKeyDown, Strings.StatusBar.Hidden, "")

        ElseIf TBTimeSelect.Checked Then
            FSSS.Text = State.TimeSelect.StartPoint
            FSSL.Text = State.TimeSelect.EndPointLength
            FSSH.Text = State.TimeSelect.HalfPointLength

        End If
        FStatus.Invalidate()
    End Sub

    Public Sub UpdateMouseRowAndColumn()
        State.Mouse.CurrentMouseRow = GetMouseVPosition(Grid.IsSnapEnabled)
        State.Mouse.CurrentMouseColumn = FocusedPanel.GetColumnAtX(State.Mouse.MouseMoveStatus.X)

        If Grid.IsSnapEnabled Then
            State.Mouse.CurrentMouseRow = FocusedPanel.SnapToGrid(State.Mouse.CurrentMouseRow)
        End If
    End Sub

    Private Function GetTimeFromVPosition(vpos As Double) As Double
        Dim timingNotes = (From note In Notes
                           Where note.ColumnIndex = ColumnType.BPM Or note.ColumnIndex = ColumnType.STOPS
                           Group By column = note.ColumnIndex
                Into noteGroups = Group).ToDictionary(Function(x) x.column, Function(x) x.noteGroups)

        Dim bpmNotes = timingNotes.Item(ColumnType.BPM)

        Dim stopNotes As IEnumerable(Of Note) = Nothing

        If timingNotes.ContainsKey(ColumnType.STOPS) Then
            stopNotes = timingNotes.Item(ColumnType.STOPS)
        End If


        Dim stopContrib As Double
        Dim bpmContrib As Double

        For i = 0 To bpmNotes.Count() - 1
            ' az: sum bpm contribution first
            Dim duration = 0.0
            Dim currentNote = bpmNotes.ElementAt(i)
            Dim notevpos = Math.Max(0, currentNote.VPosition)

            If i + 1 <> bpmNotes.Count() Then
                Dim nextNote = bpmNotes.ElementAt(i + 1)
                duration = nextNote.VPosition - notevpos
            Else
                duration = vpos - notevpos
            End If

            Dim currentBps = 60 / (currentNote.Value / 10000)
            bpmContrib += currentBps * duration / 48

            If stopNotes Is Nothing Then Continue For

            Dim stops = From stp In stopNotes
                        Where stp.VPosition >= notevpos And
                          stp.VPosition < notevpos + duration

            Dim stopBeats = stops.Sum(Function(x) x.Value / 10000.0) / 48
            stopContrib += currentBps * stopBeats

        Next

        Return stopContrib + bpmContrib
    End Function


    Private Sub POBMirror_Click(sender As Object, e As EventArgs) Handles POBMirror.Click
        Dim i As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xCol = 0
        For Each note In Notes.Skip(1).Where(Function(x) x.Selected)
            Select Case Notes(i).ColumnIndex
                ' TODO: make this handle bms instead of o2jam probably
                Case ColumnType.A1 : xCol = ColumnType.A7
                Case ColumnType.A2 : xCol = ColumnType.A6
                Case ColumnType.A3 : xCol = ColumnType.A5
                Case ColumnType.A4 : xCol = ColumnType.A4
                Case ColumnType.A5 : xCol = ColumnType.A3
                Case ColumnType.A6 : xCol = ColumnType.A2
                Case ColumnType.A7 : xCol = ColumnType.A1
                Case Else : Continue For
            End Select

            RedoMoveNote(Notes(i), xCol, Notes(i).VPosition, xUndo, xRedo)
            Notes(i).ColumnIndex = xCol
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub


    Private Sub TVCM_KeyDown(sender As Object, e As KeyEventArgs) Handles TVCM.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCM.Text = Val(TVCM.Text)
            If Val(TVCM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCM.Text = 1
                TVCM.Focus()
                TVCM.SelectAll()
            Else
                BVCApply_Click(BVCApply, New EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCM_LostFocus(sender As Object, e As EventArgs) Handles TVCM.LostFocus
        TVCM.Text = Val(TVCM.Text)
        If Val(TVCM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCM.Text = 1
            TVCM.Focus()
            TVCM.SelectAll()
        End If
    End Sub

    Private Sub TVCD_KeyDown(sender As Object, e As KeyEventArgs) Handles TVCD.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCD.Text = Val(TVCD.Text)
            If Val(TVCD.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCD.Text = 1
                TVCD.Focus()
                TVCD.SelectAll()
            Else
                BVCApply_Click(BVCApply, New EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCD_LostFocus(sender As Object, e As EventArgs) Handles TVCD.LostFocus
        TVCD.Text = Val(TVCD.Text)
        If Val(TVCD.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCD.Text = 1
            TVCD.Focus()
            TVCD.SelectAll()
        End If
    End Sub

    Private Sub TVCBPM_KeyDown(sender As Object, e As KeyEventArgs) Handles TVCBPM.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCBPM.Text = Val(TVCBPM.Text)
            If Val(TVCBPM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCBPM.Text = Notes(0).Value / 10000
                TVCBPM.Focus()
                TVCBPM.SelectAll()
            Else
                BVCCalculate_Click(BVCCalculate, New EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCBPM_LostFocus(sender As Object, e As EventArgs) Handles TVCBPM.LostFocus
        TVCBPM.Text = Val(TVCBPM.Text)
        If Val(TVCBPM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCBPM.Text = Notes(0).Value / 10000
            TVCBPM.Focus()
            TVCBPM.SelectAll()
        End If
    End Sub


    Private Sub TBUndo_Click(sender As Object, e As EventArgs) Handles TBUndo.Click, mnUndo.Click
        State.Mouse.CurrentHoveredNoteIndex = -1
        'KMouseDown = -1
        ClearSelectionArray()
        If sUndo(sI).ofType = UndoRedo.opNoOperation Then Exit Sub
        PerformCommand(sUndo(sI))
        sI = sIM()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
    End Sub

    Private Sub TBRedo_Click(sender As Object, e As EventArgs) Handles TBRedo.Click, mnRedo.Click
        State.Mouse.CurrentHoveredNoteIndex = -1
        'KMouseDown = -1
        ClearSelectionArray()
        If sRedo(sIA).ofType = UndoRedo.opNoOperation Then Exit Sub
        PerformCommand(sRedo(sIA))
        sI = sIA()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
    End Sub

    Public Sub AppendNote(note As Note)
        Notes = Notes.Concat({note}).ToArray()

        If AutoincreaseWavIndex Then
            IncreaseCurrentWav()
        End If

        State.OverwriteLastUndoRedoCommand = False

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = Nothing
        RedoAddNote(note, xUndo, xRedo, AutoincreaseWavIndex)
        AddUndoChain(xUndo, xRedo)
    End Sub

    Friend Sub DeselectAllNotes()
        For j = 1 To UBound(Notes)
            Notes(j).Selected = False
        Next
    End Sub

    Private Sub TBOptions_Click(sender As Object, e As EventArgs) Handles TBVOptions.Click, mnVOptions.Click

        Dim xDiag As New OpVisual(_vo, Columns.column, LWAV.Font)
        xDiag.ShowDialog(Me)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub AddToPowav(xPath() As String)
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        If xIndices.Length < xPath.Length Then
            Dim i As Integer = xIndices.Length
            Dim currWavIndex As Integer = xIndices(UBound(xIndices)) + 1
            ReDim Preserve xIndices(UBound(xPath))

            Do While i < xIndices.Length And currWavIndex <= 1294
                Do While currWavIndex <= 1294 AndAlso BmsWAV(currWavIndex + 1) <> ""
                    currWavIndex += 1
                Loop
                If currWavIndex > 1294 Then Exit Do

                xIndices(i) = currWavIndex
                currWavIndex += 1
                i += 1
            Loop

            If currWavIndex > 1294 Then
                ReDim Preserve xPath(i - 1)
                ReDim Preserve xIndices(i - 1)
            End If
        End If

        For i = 0 To UBound(xPath)
            BmsWAV(xIndices(i) + 1) = GetFileName(xPath(i))
            LWAV.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": " & GetFileName(xPath(i))
        Next

        LWAV.SelectedIndices.Clear()
        For i = 0 To IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath))
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        If _isSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragDrop(sender As Object, e As DragEventArgs) Handles POWAV.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, _supportedAudioExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPowav(xPath)
    End Sub

    Private Sub POWAV_DragEnter(sender As Object, e As DragEventArgs) Handles POWAV.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()),
                                                     _supportedAudioExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragLeave(sender As Object, e As EventArgs) Handles POWAV.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_Resize(sender As Object, e As EventArgs) Handles POWAV.Resize
        LWAV.Height = sender.Height - 25
    End Sub

    Private Sub AddToPobmp(xPath() As String)
        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)
        If xIndices.Length = 0 Then Exit Sub

        If xIndices.Length < xPath.Length Then
            Dim i As Integer = xIndices.Length
            Dim currBmpIndex As Integer = xIndices(UBound(xIndices)) + 1
            ReDim Preserve xIndices(UBound(xPath))

            Do While i < xIndices.Length And currBmpIndex <= 1294
                Do While currBmpIndex <= 1294 AndAlso BmsBMP(currBmpIndex + 1) <> ""
                    currBmpIndex += 1
                Loop
                If currBmpIndex > 1294 Then Exit Do

                xIndices(i) = currBmpIndex
                currBmpIndex += 1
                i += 1
            Loop

            If currBmpIndex > 1294 Then
                ReDim Preserve xPath(i - 1)
                ReDim Preserve xIndices(i - 1)
            End If
        End If

        'Dim j As Integer = 0
        For i = 0 To UBound(xPath)
            'If j > UBound(xIndices) Then Exit For
            'hBMP(xIndices(j) + 1) = GetFileName(xPath(i))
            'LBMP.Items.Item(xIndices(j)) = C10to36(xIndices(j) + 1) & ": " & GetFileName(xPath(i))
            BmsBMP(xIndices(i) + 1) = GetFileName(xPath(i))
            LBMP.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": " & GetFileName(xPath(i))
            'j += 1
        Next

        LBMP.SelectedIndices.Clear()
        For i = 0 To IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath))
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        If _isSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragDrop(sender As Object, e As DragEventArgs) Handles POBMP.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, _supportedImageExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPobmp(xPath)
    End Sub

    Private Sub POBMP_DragEnter(sender As Object, e As DragEventArgs) Handles POBMP.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()),
                                                     _supportedImageExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragLeave(sender As Object, e As EventArgs) Handles POBMP.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_Resize(sender As Object, e As EventArgs) Handles POBMP.Resize
        LBMP.Height = sender.Height - 25
    End Sub

    Private Sub POBeat_Resize(sender As Object, e As EventArgs) Handles POBeat.Resize
        LBeat.Height = POBeat.Height - 25
    End Sub

    Private Sub POExpansion_Resize(sender As Object, e As EventArgs) Handles POExpansion.Resize
        TExpansion.Height = POExpansion.Height - 2
    End Sub

    Private Sub TBPOptions_Click(sender As Object, e As EventArgs) Handles TBPOptions.Click, mnPOptions.Click
        Dim xDOp As New OpPlayer(CurrentPlayer)
        xDOp.ShowDialog(Me)
    End Sub

    Private Sub THGenre_TextChanged(sender As Object, e As EventArgs) Handles _
              THGenre.TextChanged, THTitle.TextChanged,
              THArtist.TextChanged, THPlayLevel.TextChanged,
              CHRank.SelectedIndexChanged,
              TExpansion.TextChanged,
              THSubTitle.TextChanged,
              THSubArtist.TextChanged,
              THStageFile.TextChanged, THBanner.TextChanged,
              THBackBMP.TextChanged,
              CHDifficulty.SelectedIndexChanged,
              THExRank.TextChanged, THTotal.TextChanged,
              THComment.TextChanged
        If _isSaved Then SetIsSaved(False)
    End Sub

    Private Sub CHLnObj_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CHLnObj.SelectedIndexChanged
        If _isSaved Then SetIsSaved(False)
        LnObj = CHLnObj.SelectedIndex
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub ConvertBmse2Nt()
        ClearSelectionArray()
        ValidateNotesArray()

        For i2 = 0 To UBound(Notes)
            Notes(i2).Length = 0.0#
        Next

        Dim i = 1
        Dim j = 0
        Dim xUbound As Integer = UBound(Notes)

        Do While i <= xUbound
            If Not Notes(i).LongNote Then i += 1 : Continue Do

            For j = i + 1 To xUbound
                If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For

                If Notes(j).LongNote Then
                    Notes(i).Length = Notes(j).VPosition - Notes(i).VPosition
                    For j2 As Integer = j To xUbound - 1
                        Notes(j2) = Notes(j2 + 1)
                    Next
                    xUbound -= 1
                    Exit For

                ElseIf Notes(j).Value \ 10000 = LnObj Then
                    Exit For

                End If
            Next

            i += 1
        Loop

        ReDim Preserve Notes(xUbound)

        For i = 0 To xUbound
            Notes(i).LongNote = False
        Next

        ValidateNotesArray()
    End Sub

    Private Sub ConvertNt2Bmse()
        ClearSelectionArray()
        Dim newNotes(UBound(Notes)) As Note
        newNotes(0) = Notes(0)

        For i = 1 To UBound(Notes)
            newNotes(i) = New Note() With {
                .ColumnIndex = Notes(i).ColumnIndex,
                .LongNote = Notes(i).Length > 0,
                .Landmine = Notes(i).Landmine,
                .Value = Notes(i).Value,
                .VPosition = Notes(i).VPosition,
                .Selected = Notes(i).Selected,
                .Hidden = Notes(i).Hidden
            }

            If Notes(i).Length > 0 Then
                Dim note = New Note() With {
                    .ColumnIndex = Notes(i).ColumnIndex,
                    .LongNote = True,
                    .Landmine = False,
                    .Value = Notes(i).Value,
                    .VPosition = Notes(i).VPosition + Notes(i).Length,
                    .Selected = Notes(i).Selected,
                    .Hidden = Notes(i).Hidden
                }

                newNotes = newNotes.Concat({note}).ToArray()
            End If
        Next

        Notes = newNotes

        ValidateNotesArray()
    End Sub

    Private Sub TBWavIncrease_Click(sender As Object, e As EventArgs) Handles TBWavIncrease.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        TBWavIncrease.Checked = Not sender.Checked
        RedoWavIncrease(TBWavIncrease.Checked, xUndo, xRedo)
        AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub TBNTInput_Click(sender As Object, e As EventArgs) Handles TBNTInput.Click, mnNTInput.Click
        'Dim xUndo As String = "NT_" & CInt(NTInput) & "_0" & vbCrLf & "KZ" & vbCrLf & sCmdKsAll(False)
        'Dim xRedo As String = "NT_" & CInt(Not NTInput) & "_1"
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        RedoRemoveNoteAll(Notes, xUndo, xRedo)

        NtInput = sender.Checked

        TBNTInput.Checked = NtInput
        mnNTInput.Checked = NtInput
        POBLong.Enabled = Not NtInput
        POBLongShort.Enabled = Not NtInput

        State.NT.IsAdjustingNoteLength = False
        State.NT.IsAdjustingUpperEnd = False

        RedoNT(NtInput, False, xUndo, xRedo)
        If NtInput Then
            ConvertBmse2Nt()
        Else
            ConvertNt2Bmse()
        End If
        RedoAddNoteAll(Notes, xUndo, xRedo)

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
    End Sub

    Private Sub THBPM_ValueChanged(sender As Object, e As EventArgs) Handles THBPM.ValueChanged
        If Notes IsNot Nothing Then
            If Notes(0) Is Nothing Then
                Notes(0) = New Note()
            End If
            Notes(0).Value = THBPM.Value * 10000
            RefreshPanelAll()
        End If
        If _isSaved Then SetIsSaved(False)
    End Sub

    Private Sub TWPosition_ValueChanged(sender As Object, e As EventArgs) Handles TWPosition.ValueChanged
        wPosition = TWPosition.Value
        TWPosition2.Value = IIf(wPosition > TWPosition2.Maximum, TWPosition2.Maximum, wPosition)
        RefreshPanelAll()
    End Sub

    Private Sub TWLeft_ValueChanged(sender As Object, e As EventArgs) Handles TWLeft.ValueChanged
        wLeft = TWLeft.Value
        TWLeft2.Value = IIf(wLeft > TWLeft2.Maximum, TWLeft2.Maximum, wLeft)
        RefreshPanelAll()
    End Sub

    Private Sub TWWidth_ValueChanged(sender As Object, e As EventArgs) Handles TWWidth.ValueChanged
        wWidth = TWWidth.Value
        TWWidth2.Value = IIf(wWidth > TWWidth2.Maximum, TWWidth2.Maximum, wWidth)
        RefreshPanelAll()
    End Sub

    Private Sub TWPrecision_ValueChanged(sender As Object, e As EventArgs) Handles TWPrecision.ValueChanged
        wPrecision = TWPrecision.Value
        TWPrecision2.Value = IIf(wPrecision > TWPrecision2.Maximum, TWPrecision2.Maximum, wPrecision)
        RefreshPanelAll()
    End Sub

    Private Sub TWTransparency_ValueChanged(sender As Object, e As EventArgs) Handles TWTransparency.ValueChanged
        TWTransparency2.Value = TWTransparency.Value
        _vo.pBGMWav.Color = Color.FromArgb(TWTransparency.Value, _vo.pBGMWav.Color)
        RefreshPanelAll()
    End Sub

    Private Sub TWSaturation_ValueChanged(sender As Object, e As EventArgs) Handles TWSaturation.ValueChanged
        Dim xColor As Color = _vo.pBGMWav.Color
        TWSaturation2.Value = TWSaturation.Value
        _vo.pBGMWav.Color = HSL2RGB(xColor.GetHue, TWSaturation.Value, xColor.GetBrightness * 1000, xColor.A)
        RefreshPanelAll()
    End Sub

    Private Sub TWPosition2_Scroll(sender As Object, e As EventArgs) Handles TWPosition2.Scroll
        TWPosition.Value = TWPosition2.Value
    End Sub

    Private Sub TWLeft2_Scroll(sender As Object, e As EventArgs) Handles TWLeft2.Scroll
        TWLeft.Value = TWLeft2.Value
    End Sub

    Private Sub TWWidth2_Scroll(sender As Object, e As EventArgs) Handles TWWidth2.Scroll
        TWWidth.Value = TWWidth2.Value
    End Sub

    Private Sub TWPrecision2_Scroll(sender As Object, e As EventArgs) Handles TWPrecision2.Scroll
        TWPrecision.Value = TWPrecision2.Value
    End Sub

    Private Sub TWTransparency2_Scroll(sender As Object, e As EventArgs) Handles TWTransparency2.Scroll
        TWTransparency.Value = TWTransparency2.Value
    End Sub

    Private Sub TWSaturation2_Scroll(sender As Object, e As EventArgs) Handles TWSaturation2.Scroll
        TWSaturation.Value = TWSaturation2.Value
    End Sub

    Private Sub TBLangDef_Click(sender As Object, e As EventArgs) Handles TBLangDef.Click
        _dispLang = ""
        MsgBox(Strings.Messages.PreferencePostpone, MsgBoxStyle.Information)
    End Sub

    Private Sub TBLangRefresh_Click(sender As Object, e As EventArgs) Handles TBLangRefresh.Click
        For i As Integer = cmnLanguage.Items.Count - 1 To 3 Step -1
            Try
                cmnLanguage.Items.RemoveAt(i)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then _
            My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo =
                My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles(
                    "*.Lang.xml")

        For Each xStr As FileInfo In xFileNames
            LoadLocaleXML(xStr)
        Next
    End Sub


    Private Sub UpdateColumnsX()
        Columns.RecalculatePositions()

        'HSL.Maximum = Columns.GetRightBoundry()
        'HS.Maximum = Columns.GetRightBoundry()
        'HSR.Maximum = Columns.GetRightBoundry()
    End Sub

    Private Sub CHPlayer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CHPlayer.SelectedIndexChanged
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0

        Grid.IPlayer = CHPlayer.SelectedIndex
        Dim xGp2 As Boolean = Grid.IPlayer <> 0
        Columns.SetP2SideVisible(xGp2)

        UpdateNoteSelectionStatus()

        UpdateColumnsX()

        If _isInitializing Then Exit Sub
        RefreshPanelAll()
    End Sub

    Private Sub UpdateNoteSelectionStatus()
        For i = 1 To UBound(Notes)
            Notes(i).Selected = Notes(i).Selected And Columns.IsEnabled(Notes(i).ColumnIndex)
        Next
    End Sub

    Private Sub CGB_ValueChanged(sender As Object, e As EventArgs) Handles CGB.ValueChanged
        Columns.ColumnCount = ColumnType.BGM + CGB.Value - 1
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub TBGOptions_Click(sender As Object, e As EventArgs) Handles TBGOptions.Click, mnGOptions.Click
        Dim localeIndex As Integer
        Select Case EncodingToString(_textEncoding).ToUpper ' az: wow seriously? is there really no better way? 
            Case "SYSTEM ANSI" : localeIndex = 0
            Case "LITTLE ENDIAN UTF16" : localeIndex = 1
            Case "ASCII" : localeIndex = 2
            Case "BIG ENDIAN UTF16" : localeIndex = 3
            Case "LITTLE ENDIAN UTF32" : localeIndex = 4
            Case "UTF7" : localeIndex = 5
            Case "UTF8" : localeIndex = 6
            Case "SJIS" : localeIndex = 7
            Case "EUC-KR" : localeIndex = 8
            Case Else : localeIndex = 0
        End Select

        Dim _
            xDiag As _
                New OpGeneral(Grid.WheelScroll, Grid.PageUpDnScroll, MiddleButtonMoveMethod, localeIndex,
                              192.0R / _bmsGridLimit,
                              _autoSaveInterval, _beepWhileSaved, _bpMx1296, _stoPx1296,
                              AutoFocusPanelOnMouseEnter, FirstClickDisabled, _clickStopPreview)

        If xDiag.ShowDialog() = DialogResult.OK Then
            With xDiag
                Grid.WheelScroll = .zWheel
                Grid.PageUpDnScroll = .zPgUpDn
                _textEncoding = .zEncoding
                'SortingMethod = .zSort
                MiddleButtonMoveMethod = .zMiddle
                _autoSaveInterval = .zAutoSave
                _bmsGridLimit = 192.0R / .zGridPartition
                _beepWhileSaved = .cBeep.Checked
                _bpMx1296 = .cBpm1296.Checked
                _stoPx1296 = .cStop1296.Checked
                AutoFocusPanelOnMouseEnter = .cMEnterFocus.Checked
                FirstClickDisabled = .cMClickFocus.Checked
                _clickStopPreview = .cMStopPreview.Checked
            End With
            If _autoSaveInterval Then AutoSaveTimer.Interval = _autoSaveInterval
            AutoSaveTimer.Enabled = _autoSaveInterval
        End If
    End Sub

    Private Sub POBLong_Click(sender As Object, e As EventArgs) Handles POBLong.Click
        If NtInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            RedoLongNoteModify(Notes(i), Notes(i).VPosition, True, xUndo, xRedo)
            Notes(i).LongNote = True
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormal_Click(sender As Object, e As EventArgs) Handles POBShort.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not NtInput Then
            For i = 1 To UBound(Notes)
                If Not Notes(i).Selected Then Continue For

                RedoLongNoteModify(Notes(i), Notes(i).VPosition, 0, xUndo, xRedo)
                Notes(i).LongNote = False
            Next

        Else
            For i = 1 To UBound(Notes)
                If Not Notes(i).Selected Then Continue For

                RedoLongNoteModify(Notes(i), Notes(i).VPosition, 0, xUndo, xRedo)
                Notes(i).Length = 0
            Next
        End If

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormalLong_Click(sender As Object, e As EventArgs) Handles POBLongShort.Click
        If NtInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            RedoLongNoteModify(Notes(i), Notes(i).VPosition, Not Notes(i).LongNote, xUndo, xRedo)
            Notes(i).LongNote = Not Notes(i).LongNote
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBHidden_Click(sender As Object, e As EventArgs) Handles POBHidden.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            RedoHiddenNoteModify(Notes(i), True, True, xUndo, xRedo)
            Notes(i).Hidden = True
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBVisible_Click(sender As Object, e As EventArgs) Handles POBVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            RedoHiddenNoteModify(Notes(i), False, True, xUndo, xRedo)
            Notes(i).Hidden = False
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBHiddenVisible_Click(sender As Object, e As EventArgs) Handles POBHiddenVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            RedoHiddenNoteModify(Notes(i), Not Notes(i).Hidden, True, xUndo, xRedo)
            Notes(i).Hidden = Not Notes(i).Hidden
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub POBModify_Click(sender As Object, e As EventArgs) Handles POBModify.Click
        Dim xNum = False
        Dim xLbl = False
        Dim i As Integer

        For i = 1 To UBound(Notes)
            If Notes(i).Selected AndAlso Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then
                xNum = True
                Exit For
            End If
        Next
        For i = 1 To UBound(Notes)
            If Notes(i).Selected AndAlso Not Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then
                xLbl = True
                Exit For
            End If
        Next
        If Not (xNum Or xLbl) Then Exit Sub

        If xNum Then
            Dim xD1 As Double = Val(InputBox(Strings.Messages.PromptEnterNumeric, Text)) * 10000
            If Not xD1 = 0 Then
                If xD1 <= 0 Then xD1 = 1

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                For i = 1 To UBound(Notes)
                    If Not Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then Continue For
                    If Not Notes(i).Selected Then Continue For

                    RedoRelabelNote(Notes(i), xD1, xUndo, xRedo)
                    Notes(i).Value = xD1
                Next
                AddUndoChain(xUndo, xBaseRedo.Next)
            End If
        End If

        If xLbl Then
            Dim xStr As String = UCase(Trim(InputBox(Strings.Messages.PromptEnter, Me.Text)))

            If Len(xStr) = 0 Then GoTo Jump2
            If xStr = "00" Or xStr = "0" Then GoTo Jump1
            If Not Len(xStr) = 1 And Not Len(xStr) = 2 Then GoTo Jump1

            Dim xI3 As Integer = Asc(Mid(xStr, 1, 1))
            If Not ((xI3 >= 48 And xI3 <= 57) Or (xI3 >= 65 And xI3 <= 90)) Then GoTo Jump1
            If Len(xStr) = 2 Then
                Dim xI4 As Integer = Asc(Mid(xStr, 2, 1))
                If Not ((xI4 >= 48 And xI4 <= 57) Or (xI4 >= 65 And xI4 <= 90)) Then GoTo Jump1
            End If
            Dim xVal As Integer = C36to10(xStr) * 10000

            Dim xUndo As UndoRedo.LinkedURCmd = Nothing
            Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
            Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

            For i = 1 To UBound(Notes)
                If Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then Continue For
                If Not Notes(i).Selected Then Continue For

                RedoRelabelNote(Notes(i), xVal, xUndo, xRedo)
                Notes(i).Value = xVal
            Next
            AddUndoChain(xUndo, xBaseRedo.Next)
            GoTo Jump2
Jump1:
            MsgBox(Strings.Messages.InvalidLabel, MsgBoxStyle.Critical, Strings.Messages.Err)
Jump2:
        End If

        RefreshPanelAll()
    End Sub

    Private Sub TBMyO2_Click(sender As Object, e As EventArgs) Handles TBMyO2.Click, mnMyO2.Click
        Dim xDiag As New dgMyO2
        xDiag.Show()
    End Sub


    Private Sub TBFind_Click(sender As Object, e As EventArgs) Handles TBFind.Click, mnFind.Click
        Dim xDiag As New diagFind(Columns.ColumnCount, Strings.Messages.Err, Strings.Messages.InvalidLabel)
        xDiag.Show()
    End Sub


    Private Sub MInsert_Click(sender As Object, e As EventArgs) Handles MInsert.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(_menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)
        Dim xVp As Double = MeasureBottom(xMeasure)

        If NtInput Then
            Dim i = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) >= 999 Then
                    RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            Dim xdVp As Double
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVp And Notes(i).VPosition + Notes(i).Length <= MeasureBottom(999) Then
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition + xMLength, xUndo, xRedo)
                    Notes(i).VPosition += xMLength

                ElseIf Notes(i).VPosition >= xVp Then
                    xdVp = MeasureBottom(999) - 1 - Notes(i).VPosition - Notes(i).Length
                    RedoLongNoteModify(Notes(i), Notes(i).VPosition + xMLength, Notes(i).Length + xdVp, xUndo, xRedo)
                    Notes(i).VPosition += xMLength
                    Notes(i).Length += xdVp

                ElseIf Notes(i).VPosition + Notes(i).Length >= xVp Then
                    xdVp = IIf(Notes(i).VPosition + Notes(i).Length > MeasureBottom(999) - 1,
                               GetMaxVPosition() - 1 - Notes(i).VPosition - Notes(i).Length, xMLength)
                    RedoLongNoteModify(Notes(i), Notes(i).VPosition, Notes(i).Length + xdVp, xUndo, xRedo)
                    Notes(i).Length += xdVp
                End If
            Next

        Else
            Dim i = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) >= 999 Then
                    RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVp Then
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition + xMLength, xUndo, xRedo)
                    Notes(i).VPosition += xMLength
                End If
            Next
        End If

        For i = 999 To xMeasure + 1 Step -1
            MeasureLength(i) = MeasureLength(i - 1)
        Next
        UpdateMeasureBottom()

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub MRemove_Click(sender As Object, e As EventArgs) Handles MRemove.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(_menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)

        If NtInput Then
            Dim i = 1
            Do While i <= UBound(Notes)
                If _
                    MeasureAtDisplacement(Notes(i).VPosition) = xMeasure And
                    MeasureAtDisplacement(Notes(i).VPosition + Notes(i).Length) = xMeasure Then
                    RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            Dim xdVp As Double
            Dim xVP = MeasureBottom(xMeasure)
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP + xMLength Then
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition - xMLength, xUndo, xRedo)
                    Notes(i).VPosition -= xMLength

                ElseIf Notes(i).VPosition >= xVP Then
                    xdVp = xMLength + xVP - Notes(i).VPosition
                    RedoLongNoteModify(Notes(i), Notes(i).VPosition + xdVp - xMLength, Notes(i).Length - xdVp, xUndo,
                                       xRedo)
                    Notes(i).VPosition += xdVp - xMLength
                    Notes(i).Length -= xdVp

                ElseIf Notes(i).VPosition + Notes(i).Length >= xVP Then
                    xdVp = IIf(Notes(i).VPosition + Notes(i).Length >= xVP + xMLength, xMLength,
                               Notes(i).VPosition + Notes(i).Length - xVP + 1)
                    RedoLongNoteModify(Notes(i), Notes(i).VPosition, Notes(i).Length - xdVp, xUndo, xRedo)
                    Notes(i).Length -= xdVp
                End If
            Next

        Else
            Dim i = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) = xMeasure Then
                    RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            Dim xVP = MeasureBottom(xMeasure)
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP Then
                    RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition - xMLength, xUndo, xRedo)
                    Notes(i).VPosition -= xMLength
                End If
            Next
        End If

        For i = 999 To xMeasure + 1 Step -1
            MeasureLength(i - 1) = MeasureLength(i)
        Next
        UpdateMeasureBottom()

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
    End Sub

    Private Sub TBThemeDef_Click(sender As Object, e As EventArgs) Handles TBThemeDef.Click
        Dim xTempFileName As String = My.Application.Info.DirectoryPath & "\____TempFile.Theme.xml"
        My.Computer.FileSystem.WriteAllText(xTempFileName, My.Resources.O2Mania_Theme, False, Encoding.Unicode)
        LoadSettings(xTempFileName)
        File.Delete(xTempFileName)

        RefreshPanelAll()
    End Sub

    Private Sub TBThemeSave_Click(sender As Object, e As EventArgs) Handles TBThemeSave.Click
        Dim xDiag As New SaveFileDialog
        xDiag.Filter = Strings.FileType.THEME_XML & "|*.Theme.xml"
        xDiag.DefaultExt = "Theme.xml"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Data"
        If xDiag.ShowDialog = DialogResult.Cancel Then Exit Sub

        Me.SaveSettings(xDiag.FileName, True)
        If _beepWhileSaved Then Beep()
        TBThemeRefresh_Click(TBThemeRefresh, New EventArgs)
    End Sub

    Private Sub TBThemeRefresh_Click(sender As Object, e As EventArgs) Handles TBThemeRefresh.Click
        For i As Integer = cmnTheme.Items.Count - 1 To 5 Step -1
            Try
                cmnTheme.Items.RemoveAt(i)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then _
            My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo =
                My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles(
                    "*.Theme.xml")
        For Each xStr As FileInfo In xFileNames
            cmnTheme.Items.Add(xStr.Name, Nothing, AddressOf LoadTheme)
        Next
    End Sub

    Private Sub TBThemeLoadComptability_Click(sender As Object, e As EventArgs) Handles TBThemeLoadComptability.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType.TH & "|*.th"
        xDiag.DefaultExt = "th"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath
        If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\Theme") Then _
            xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Theme"
        If xDiag.ShowDialog = DialogResult.Cancel Then Exit Sub

        Me.LoadThemeComptability(xDiag.FileName)
        RefreshPanelAll()
    End Sub

    ''' <summary>
    '''     Will return Double.PositiveInfinity if canceled.
    ''' </summary>
    Private Function InputBoxDouble(prompt As String, lBound As Double, uBound As Double,
                                    Optional ByVal title As String = "", Optional ByVal defaultResponse As String = "") _
        As Double
        Dim xStr As String = InputBox(prompt, title, defaultResponse)
        If xStr = "" Then Return Double.PositiveInfinity

        InputBoxDouble = Val(xStr)
        If InputBoxDouble > uBound Then InputBoxDouble = uBound
        If InputBoxDouble < lBound Then InputBoxDouble = lBound
    End Function

    Private Sub FSSS_Click(sender As Object, e As EventArgs) Handles FSSS.Click
        Dim xMax As Double = IIf(State.TimeSelect.EndPointLength > 0,
                                 GetMaxVPosition() - State.TimeSelect.EndPointLength, GetMaxVPosition)
        Dim xMin As Double = IIf(State.TimeSelect.EndPointLength < 0, -State.TimeSelect.EndPointLength, 0)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin,
                                               xMax, , State.TimeSelect.StartPoint)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.StartPoint = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub FSSL_Click(sender As Object, e As EventArgs) Handles FSSL.Click
        Dim xMax As Double = GetMaxVPosition() - State.TimeSelect.StartPoint
        Dim xMin As Double = -State.TimeSelect.StartPoint
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin,
                                               xMax, , State.TimeSelect.EndPointLength)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.EndPointLength = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub FSSH_Click(sender As Object, e As EventArgs) Handles FSSH.Click
        Dim xMax As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.EndPointLength, 0)
        Dim xMin As Double = IIf(State.TimeSelect.EndPointLength > 0, 0, -State.TimeSelect.EndPointLength)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin,
                                               xMax, , State.TimeSelect.HalfPointLength)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.HalfPointLength = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BVCReverse_Click(sender As Object, e As EventArgs) Handles BVCReverse.Click
        State.TimeSelect.StartPoint += State.TimeSelect.EndPointLength
        State.TimeSelect.HalfPointLength -= State.TimeSelect.EndPointLength
        State.TimeSelect.EndPointLength *= -1
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub AutoSaveTimer_Tick(sender As Object, e As EventArgs) Handles AutoSaveTimer.Tick
        Dim xTime As Date = Now
        Dim xFileName As String
        With xTime
            xFileName = My.Application.Info.DirectoryPath & "\AutoSave_" &
                        .Year & "_" & .Month & "_" & .Day & "_" & .Hour & "_" & .Minute & "_" & .Second & "_" &
                        .Millisecond & ".IBMSC"
        End With
        'My.Computer.FileSystem.WriteAllText(xFileName, SaveiBMSC, False, System.Text.Encoding.Unicode)
        SaveiBmsc(xFileName)

        On Error Resume Next
        If _previousAutoSavedFileName <> "" Then File.Delete(_previousAutoSavedFileName)
        On Error GoTo 0

        _previousAutoSavedFileName = xFileName
    End Sub

    Private Sub CWAVMultiSelect_CheckedChanged(sender As Object, e As EventArgs) Handles CWAVMultiSelect.CheckedChanged
        _wavMultiSelect = CWAVMultiSelect.Checked
        LWAV.SelectionMode = IIf(_wavMultiSelect, SelectionMode.MultiExtended, SelectionMode.One)
        LBMP.SelectionMode = IIf(_wavMultiSelect, SelectionMode.MultiExtended, SelectionMode.One)
    End Sub

    Private Sub CWAVChangeLabel_CheckedChanged(sender As Object, e As EventArgs) Handles CWAVChangeLabel.CheckedChanged
        _wavChangeLabel = CWAVChangeLabel.Checked
    End Sub

    Private Sub BWAVUp_Click(sender As Object, e As EventArgs) Handles BWAVUp.Click
        If LWAV.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 0 To 1294
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsWAV(i + 1)
                BmsWAV(i + 1) = BmsWAV(i)
                BmsWAV(i) = xStr

                LWAV.Items.Item(i) = C10to36(i + 1) & ": " & BmsWAV(i + 1)
                LWAV.Items.Item(i - 1) = C10to36(i) & ": " & BmsWAV(i)

                If Not _wavChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i)
                Dim xL2 As String = C10to36(i + 1)
                For j = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        RedoRelabelNote(Notes(j), i * 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000

                    End If
                Next

1100:           xIndices(xIndex) += -1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BWAVDown_Click(sender As Object, e As EventArgs) Handles BWAVDown.Click
        If LWAV.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 1294 To 0 Step -1
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsWAV(i + 1)
                BmsWAV(i + 1) = BmsWAV(i + 2)
                BmsWAV(i + 2) = xStr

                LWAV.Items.Item(i) = C10to36(i + 1) & ": " & BmsWAV(i + 1)
                LWAV.Items.Item(i + 1) = C10to36(i + 2) & ": " & BmsWAV(i + 2)

                If Not _wavChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i + 2)
                Dim xL2 As String = C10to36(i + 1)
                For j = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 20000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 20000

                    End If
                Next

1100:           xIndices(xIndex) += 1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BWAVBrowse_Click(sender As Object, e As EventArgs) Handles BWAVBrowse.Click
        Dim xDwav As New OpenFileDialog
        xDwav.DefaultExt = "wav"
        xDwav.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDwav.InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
        xDwav.Multiselect = _wavMultiSelect

        If xDwav.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDwav.FileName)

        AddToPowav(xDwav.FileNames)
    End Sub

    Private Sub BWAVRemove_Click(sender As Object, e As EventArgs) Handles BWAVRemove.Click
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)
        For i = 0 To UBound(xIndices)
            BmsWAV(xIndices(i) + 1) = ""
            LWAV.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": "
        Next

        LWAV.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        If _isSaved Then SetIsSaved(False)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BBMPUp_Click(sender As Object, e As EventArgs) Handles BBMPUp.Click
        If LBMP.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 0 To 1294
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsBMP(i + 1)
                BmsBMP(i + 1) = BmsBMP(i)
                BmsBMP(i) = xStr

                LBMP.Items.Item(i) = C10to36(i + 1) & ": " & BmsBMP(i + 1)
                LBMP.Items.Item(i - 1) = C10to36(i) & ": " & BmsBMP(i)

                If Not _wavChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i)
                Dim xL2 As String = C10to36(i + 1)
                For j = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        RedoRelabelNote(Notes(j), i * 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000

                    End If
                Next

1100:           xIndices(xIndex) += -1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BBMPDown_Click(sender As Object, e As EventArgs) Handles BBMPDown.Click
        If LBMP.SelectedIndex = -1 Then Return

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)

        Dim xS As Integer
        For xS = 1294 To 0 Step -1
            If Array.IndexOf(xIndices, xS) = -1 Then Exit For
        Next

        Dim xStr = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsBMP(i + 1)
                BmsBMP(i + 1) = BmsBMP(i + 2)
                BmsBMP(i + 2) = xStr

                LBMP.Items.Item(i) = C10to36(i + 1) & ": " & BmsBMP(i + 1)
                LBMP.Items.Item(i + 1) = C10to36(i + 2) & ": " & BmsBMP(i + 2)

                If Not _wavChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i + 2)
                Dim xL2 As String = C10to36(i + 1)
                For j = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        RedoRelabelNote(Notes(j), i * 10000 + 20000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 20000

                    End If
                Next

1100:           xIndices(xIndex) += 1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BBMPBrowse_Click(sender As Object, e As EventArgs) Handles BBMPBrowse.Click
        Dim xDbmp As New OpenFileDialog
        xDbmp.DefaultExt = "bmp"
        xDbmp.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.jpeg;.gif|" &
                       Strings.FileType._movie & "|*.mpg;*.m1v;*.m2v;*.avi;*.mp4;*.m4v;*.wmv;*.webm|" &
                       Strings.FileType.BMP & "|*.bmp|" &
                       Strings.FileType.PNG & "|*.png|" &
                       Strings.FileType.JPG & "|*.jpg;*.jpeg|" &
                       Strings.FileType.GIF & "|*.gif|" &
                       Strings.FileType.MP4 & "|*.mp4;*.m4v|" &
                       Strings.FileType.AVI & "|*.avi|" &
                       Strings.FileType.MPG & "|*.mpg;*.m1v;*.m2v|" &
                       Strings.FileType.WMV & "|*.wmv|" &
                       Strings.FileType.WEBM & "|*.webm|" &
                       Strings.FileType._all & "|*.*"
        xDbmp.InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
        xDbmp.Multiselect = _wavMultiSelect

        If xDbmp.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDbmp.FileName)

        AddToPobmp(xDbmp.FileNames)
    End Sub

    Private Sub BBMPRemove_Click(sender As Object, e As EventArgs) Handles BBMPRemove.Click
        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)
        For i = 0 To UBound(xIndices)
            BmsBMP(xIndices(i) + 1) = ""
            LBMP.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": "
        Next

        LBMP.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        If _isSaved Then SetIsSaved(False)
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub mnMain_MouseDown(sender As Object, e As MouseEventArgs) Handles mnMain.MouseDown _
        ', TBMain.MouseDown  ', pttl.MouseDown, pIsSaved.MouseDown
        If e.Button = MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, &H112, &HF012, 0)
            If e.Clicks = 2 Then
                If Me.WindowState = FormWindowState.Maximized Then Me.WindowState = FormWindowState.Normal Else _
                    Me.WindowState = FormWindowState.Maximized
            End If
        ElseIf e.Button = MouseButtons.Right Then
            'mnSys.Show(sender, e.Location)
        End If
    End Sub

    Private Sub mnSelectAll_Click(sender As Object, e As EventArgs) Handles mnSelectAll.Click
        ' If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub
        For Each note In Notes.Skip(1)
            note.Selected = Columns.IsEnabled(note.ColumnIndex)
        Next
        If TBTimeSelect.Checked Then

            State.TimeSelect.StartPoint = 0
            State.TimeSelect.EndPointLength = MeasureBottom(MeasureAtDisplacement(GreatestVPosition)) +
                                              MeasureLength(MeasureAtDisplacement(GreatestVPosition))
        End If
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub mnDelete_Click(sender As Object, e As EventArgs) Handles mnDelete.Click
        ' If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        RedoRemoveNoteSelected(Notes, xUndo, xRedo)
        RemoveNotes(True)

        AddUndoChain(xUndo, xBaseRedo.Next)

        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub mnQuit_Click(sender As Object, e As EventArgs) Handles mnQuit.Click
        Close()
    End Sub

    Private Sub mnSMenu_Click(sender As Object, e As EventArgs) Handles mnSMenu.CheckedChanged
        mnMain.Visible = mnSMenu.Checked
    End Sub

    Private Sub mnSTB_Click(sender As Object, e As EventArgs) Handles mnSTB.CheckedChanged
        TBMain.Visible = mnSTB.Checked
    End Sub

    Private Sub mnSOP_Click(sender As Object, e As EventArgs) Handles mnSOP.CheckedChanged
        POptions.Visible = mnSOP.Checked
    End Sub

    Private Sub mnSStatus_Click(sender As Object, e As EventArgs) Handles mnSStatus.CheckedChanged
        pStatus.Visible = mnSStatus.Checked
    End Sub
    Private Sub CGShow_CheckedChanged(sender As Object, e As EventArgs) Handles CGShow.CheckedChanged
        Grid.ShowMainGrid = CGShow.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowS_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowS.CheckedChanged
        Grid.ShowSubGrid = CGShowS.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowBG_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowBG.CheckedChanged
        Grid.ShowBackground = CGShowBG.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowM_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowM.CheckedChanged
        Grid.ShowMeasureNumber = CGShowM.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowV_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowV.CheckedChanged
        Grid.ShowVerticalLines = CGShowV.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowMB_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowMB.CheckedChanged
        Grid.ShowMeasureBars = CGShowMB.Checked
        RefreshPanelAll()
    End Sub

    Private Sub CGShowC_CheckedChanged(sender As Object, e As EventArgs) Handles CGShowC.CheckedChanged
        Grid.ShowColumnCaptions = CGShowC.Checked
        RefreshPanelAll()
    End Sub

    Private Sub FixedColumnVisibilityChanged(col As ColumnType(), isVisible As Boolean)
        For Each c In col
            Columns.GetColumn(c).IsVisible = isVisible
        Next


        If _isInitializing Then Exit Sub
        For Each note In Notes.Skip(1)
            note.Selected = note.Selected And Columns.IsEnabled(note.ColumnIndex)
        Next

        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub CGBLP_CheckedChanged(sender As Object, e As EventArgs) Handles CGBLP.CheckedChanged
        Grid.ShowBgaColumn = CGBLP.Checked

        FixedColumnVisibilityChanged({
                                         ColumnType.BGA,
                                         ColumnType.LAYER,
                                         ColumnType.POOR,
                                         ColumnType.S4
                                     }, Grid.ShowBgaColumn)
    End Sub

    Private Sub CGSCROLL_CheckedChanged(sender As Object, e As EventArgs) Handles CGSCROLL.CheckedChanged
        Grid.ShowScrollColumn = CGSCROLL.Checked
        FixedColumnVisibilityChanged({ColumnType.SCROLLS}, Grid.ShowScrollColumn)
    End Sub

    Private Sub CGSTOP_CheckedChanged(sender As Object, e As EventArgs) Handles CGSTOP.CheckedChanged
        Grid.ShowStopColumn = CGSTOP.Checked
        FixedColumnVisibilityChanged({ColumnType.STOPS}, Grid.ShowStopColumn)
    End Sub

    Private Sub CGBPM_CheckedChanged(sender As Object, e As EventArgs) Handles CGBPM.CheckedChanged
        Grid.ShowBpmColumn = CGBPM.Checked
        FixedColumnVisibilityChanged({ColumnType.BPM}, Grid.ShowBpmColumn)
    End Sub

    Private Sub CGDisableVertical_CheckedChanged(sender As Object, e As EventArgs) _
        Handles CGDisableVertical.CheckedChanged
        DisableVerticalMove = CGDisableVertical.Checked
    End Sub

    Private Sub CBeatPreserve_Click(sender As Object, e As EventArgs) _
        Handles CBeatPreserve.Click, CBeatMeasure.Click, CBeatCut.Click, CBeatScale.Click
        'If Not sender.Checked Then Exit Sub
        Dim xBeatList() As RadioButton = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
        _beatChangeMode = Array.IndexOf(Of RadioButton)(xBeatList, sender)
    End Sub


    Private Sub tBeatValue_LostFocus(sender As Object, e As EventArgs) Handles tBeatValue.LostFocus
        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then tBeatValue.BackColor = Color.FromArgb(&HFFFFC0C0) Else _
                tBeatValue.BackColor = Nothing

            tBeatValue.Text = a
        End If
    End Sub


    Private Sub ApplyBeat(xRatio As Double, xDisplay As String)
        ValidateNotesArray()

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        RedoChangeMeasureLengthSelected(192 * xRatio, xUndo, xRedo)

        Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer
        LBeat.SelectedIndices.CopyTo(xIndices, 0)


        For Each i As Integer In xIndices
            Dim dLength As Double = xRatio * 192.0R - MeasureLength(i)
            Dim dRatio As Double = xRatio * 192.0R / MeasureLength(i)

            Dim xBottom As Double = 0
            For j = 0 To i - 1
                xBottom += MeasureLength(j)
            Next
            Dim xUpBefore As Double = xBottom + MeasureLength(i)
            Dim xUpAfter As Double = xUpBefore + dLength

            Select Case _beatChangeMode
                Case 1
case2:              Dim xI0 As Integer

                    If NtInput Then
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                            If Notes(xI0).VPosition + Notes(xI0).Length >= xUpBefore Then
                                RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, xUndo,
                                                   xRedo)
                                Notes(xI0).Length += dLength
                            End If
                        Next
                    Else
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                        Next
                    End If

                    For xI9 As Integer = xI0 To UBound(Notes)
                        RedoLongNoteModify(Notes(xI9), Notes(xI9).VPosition + dLength, Notes(xI9).Length, xUndo, xRedo)
                        Notes(xI9).VPosition += dLength
                    Next

                Case 2
                    If dLength < 0 Then
                        If NtInput Then
                            Dim xI0 = 1
                            Dim xU As Integer = UBound(Notes)
                            Do While xI0 <= xU
                                If Notes(xI0).VPosition < xUpAfter Then
                                    If _
                                        Notes(xI0).VPosition + Notes(xI0).Length >= xUpAfter And
                                        Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Dim nLen As Double = xUpAfter - Notes(xI0).VPosition - 1.0R
                                        RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                    End If
                                ElseIf Notes(xI0).VPosition < xUpBefore Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        RedoRemoveNote(Notes(xI0), xUndo, xRedo)
                                        RemoveNote(xI0)
                                        xI0 -= 1
                                        xU -= 1
                                    Else
                                        Dim nLen As Double = Notes(xI0).Length - xUpBefore + Notes(xI0).VPosition
                                        RedoLongNoteModify(Notes(xI0), xUpBefore, nLen, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                        Notes(xI0).VPosition = xUpBefore
                                    End If
                                End If
                                xI0 += 1
                            Loop
                        Else
                            Dim xI0 As Integer
                            Dim xI9 As Integer
                            For xI0 = 1 To UBound(Notes)
                                If Notes(xI0).VPosition >= xUpAfter Then Exit For
                            Next
                            For xI9 = xI0 To UBound(Notes)
                                If Notes(xI9).VPosition >= xUpBefore Then Exit For
                            Next

                            For xI8 As Integer = xI0 To xI9 - 1
                                RedoRemoveNote(Notes(xI8), xUndo, xRedo)
                            Next
                            For xI8 As Integer = xI9 To UBound(Notes)
                                Notes(xI8 - xI9 + xI0) = Notes(xI8)
                            Next
                            ReDim Preserve Notes(UBound(Notes) - xI9 + xI0)
                        End If
                    End If

                    GoTo case2

                Case 3
                    If NtInput Then
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition < xBottom Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength,
                                                       xUndo, xRedo)
                                    Notes(xI0).Length += dLength
                                ElseIf Notes(xI0).VPosition + Notes(xI0).Length > xBottom Then
                                    Dim nLen As Double = (Notes(xI0).Length + Notes(xI0).VPosition - xBottom) * dRatio +
                                                         xBottom - Notes(xI0).VPosition
                                    RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                End If
                            ElseIf Notes(xI0).VPosition < xUpBefore Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Dim nLen As Double = (xUpBefore - Notes(xI0).VPosition) * dRatio +
                                                         Notes(xI0).VPosition + Notes(xI0).Length - xUpBefore
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                Else
                                    Dim nLen As Double = Notes(xI0).Length * dRatio
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                End If
                            Else
                                RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition + dLength, Notes(xI0).Length, xUndo,
                                                   xRedo)
                                Notes(xI0).VPosition += dLength
                            End If
                        Next
                    Else
                        Dim xI0 As Integer
                        Dim xI9 As Integer
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xBottom Then Exit For
                        Next
                        For xI9 = xI0 To UBound(Notes)
                            If Notes(xI9).VPosition >= xUpBefore Then Exit For
                        Next

                        For xI8 As Integer = xI0 To xI9 - 1
                            Dim nVp As Double = (Notes(xI8).VPosition - xBottom) * dRatio + xBottom
                            RedoLongNoteModify(Notes(xI0), nVp, Notes(xI0).Length, xUndo, xRedo)
                            Notes(xI8).VPosition = nVp
                        Next

                        'GoTo case2

                        For xI8 As Integer = xI9 To UBound(Notes)
                            RedoLongNoteModify(Notes(xI8), Notes(xI8).VPosition + dLength, Notes(xI8).Length, xUndo,
                                               xRedo)
                            Notes(xI8).VPosition += dLength
                        Next
                    End If

            End Select

            MeasureLength(i) = xRatio * 192.0R
            LBeat.Items(i) = Add3Zeros(i) & ": " & xDisplay
        Next
        UpdateMeasureBottom()
        'xUndo &= vbCrLf & xUndo2
        'xRedo &= vbCrLf & xRedo2

        LBeat.SelectedIndices.Clear()
        For i = 0 To UBound(xIndices)
            LBeat.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()
        PoStatusRefresh()
    End Sub

    Private Sub BBeatApply_Click(sender As Object, e As EventArgs) Handles BBeatApply.Click
        Dim xxD As Integer = nBeatD.Value
        Dim xxN As Integer = nBeatN.Value
        Dim xxRatio As Double = xxN / xxD

        ApplyBeat(xxRatio, xxRatio & " ( " & xxN & " / " & xxD & " ) ")
    End Sub

    Private Sub BBeatApplyV_Click(sender As Object, e As EventArgs) Handles BBeatApplyV.Click
        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then SystemSounds.Hand.Play() : Exit Sub

            Dim xxD As Long = GetDenominator(a)

            ApplyBeat(a, a & IIf(xxD > 10000, "", " ( " & CLng(a * xxD) & " / " & xxD & " ) "))
        End If
    End Sub


    Private Sub BHStageFile_Click(sender As Object, e As EventArgs) _
        Handles BHStageFile.Click, BHBanner.Click, BHBackBMP.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpeg;*.jpg;*.gif|" &
                       Strings.FileType._all & "|*.*"
        xDiag.InitialDirectory = IIf(ExcludeFileName(_fileName) = "", _initPath, ExcludeFileName(_fileName))
        xDiag.DefaultExt = "png"

        If xDiag.ShowDialog = DialogResult.Cancel Then Exit Sub
        _initPath = ExcludeFileName(xDiag.FileName)

        If ReferenceEquals(sender, BHStageFile) Then
            THStageFile.Text = GetFileName(xDiag.FileName)
        ElseIf ReferenceEquals(sender, BHBanner) Then
            THBanner.Text = GetFileName(xDiag.FileName)
        ElseIf ReferenceEquals(sender, BHBackBMP) Then
            THBackBMP.Text = GetFileName(xDiag.FileName)
        End If
    End Sub

    Private Sub Switches_CheckedChanged(sender As Object, e As EventArgs) Handles _
                                                                              POHeaderSwitch.CheckedChanged,
                                                                              POGridSwitch.CheckedChanged,
                                                                              POWaveFormSwitch.CheckedChanged,
                                                                              POWAVSwitch.CheckedChanged,
                                                                              POBMPSwitch.CheckedChanged,
                                                                              POBeatSwitch.CheckedChanged,
                                                                              POExpansionSwitch.CheckedChanged

        Try
            Dim source = CType(sender, CheckBox)
            Dim target As Panel = Nothing

            If ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf ReferenceEquals(sender, POHeaderSwitch) Then : target = POHeaderInner
            ElseIf ReferenceEquals(sender, POGridSwitch) Then : target = POGridInner
            ElseIf ReferenceEquals(sender, POWaveFormSwitch) Then : target = POWaveFormInner
            ElseIf ReferenceEquals(sender, POWAVSwitch) Then : target = POWAVInner
            ElseIf ReferenceEquals(sender, POBMPSwitch) Then : target = POBMPInner
            ElseIf ReferenceEquals(sender, POBeatSwitch) Then : target = POBeatInner
            ElseIf ReferenceEquals(sender, POExpansionSwitch) Then : target = POExpansionInner
            End If

            If source.Checked Then
                target.Visible = True
            Else
                target.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Friend Sub SelectWavFromNote(note As Note)
        LWAV.SelectedIndices.Clear()
        LWAV.SelectedIndex = C36to10(C10to36(note.Value \ 10000)) - 1
        ValidateWavListView()
    End Sub

    Private Sub Expanders_CheckChanged(sender As Object, e As EventArgs) Handles _
                                                                             POHeaderExpander.CheckedChanged,
                                                                             POGridExpander.CheckedChanged,
                                                                             POWaveFormExpander.CheckedChanged,
                                                                             POWAVExpander.CheckedChanged,
                                                                             POBeatExpander.CheckedChanged

        Try
            Dim source = CType(sender, CheckBox)
            Dim target As Panel = Nothing
            'Dim TargetParent As Panel = Nothing

            If ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf ReferenceEquals(sender, POHeaderExpander) Then
                target = POHeaderPart2 ' : TargetParent = POHeaderInner
            ElseIf ReferenceEquals(sender, POGridExpander) Then : target = POGridPart2 ' : TargetParent = POGridInner
            ElseIf ReferenceEquals(sender, POWaveFormExpander) Then
                target = POWaveFormPart2 ' : TargetParent = POWaveFormInner
            ElseIf ReferenceEquals(sender, POWAVExpander) Then : target = POWAVPart2 ' : TargetParent = POWaveFormInner
            ElseIf ReferenceEquals(sender, POBeatExpander) Then
                target = POBeatPart2 ' : TargetParent = POWaveFormInner
            End If

            If source.Checked Then
                target.Visible = True
                'Source.Image = My.Resources.Collapse
            Else
                target.Visible = False
                'Source.Image = My.Resources.Expand
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub VerticalResizer_MouseDown(sender As Object, e As MouseEventArgs) _
        Handles POWAVResizer.MouseDown, POBMPResizer.MouseDown, POBeatResizer.MouseDown, POExpansionResizer.MouseDown
        _tempResize = e.Y
    End Sub

    Private Sub HorizontalResizer_MouseDown(sender As Object, e As MouseEventArgs) _
        Handles POptionsResizer.MouseDown
        _tempResize = e.X
    End Sub

    Private Sub POResizer_MouseMove(sender As Object, e As MouseEventArgs) _
        Handles POWAVResizer.MouseMove, POBMPResizer.MouseMove, POBeatResizer.MouseMove, POExpansionResizer.MouseMove
        If e.Button <> MouseButtons.Left Then Exit Sub
        If e.Y = _tempResize Then Exit Sub

        Try
            Dim source = CType(sender, Button)
            Dim target As Panel = source.Parent

            Dim xHeight As Integer = target.Height + e.Y - _tempResize
            If xHeight < 10 Then xHeight = 10
            target.Height = xHeight

            target.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub POptionsResizer_MouseMove(sender As Object, e As MouseEventArgs) Handles POptionsResizer.MouseMove
        If e.Button <> MouseButtons.Left Then Exit Sub
        If e.X = _tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = POptionsScroll.Width - e.X + _tempResize
            If xWidth < 25 Then xWidth = 25
            POptionsScroll.Width = xWidth

            Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub mnGotoMeasure_Click(sender As Object, e As EventArgs) Handles mnGotoMeasure.Click
        Dim s = InputBox(Strings.Messages.PromptEnterMeasure, "Enter Measure")

        Dim i As Integer
        If Int32.TryParse(s, i) Then
            If i < 0 Or i > 999 Then
                Exit Sub
            End If

            FocusedPanel.ScrollTo(-MeasureBottom(i))
        End If
    End Sub

    Public ReadOnly Property CurrentMouseoverNote As Note
        Get
            Return Notes(State.Mouse.CurrentHoveredNoteIndex)
        End Get
    End Property

    Friend Function IsTimeSelectMode() As Boolean
        Return TBTimeSelect.Checked
    End Function

    Public Sub RefreshPanelAll()
        If _isInitializing Then Exit Sub

        For Each panel In _spMain
            panel.Refresh()
        Next
    End Sub

    Public ReadOnly Property IsSelectMode As Boolean
        Get
            Return TBSelect.Checked
        End Get
    End Property

    Public ReadOnly Property IsWriteMode As Boolean
        Get
            Return TBWrite.Checked
        End Get
    End Property
End Class
