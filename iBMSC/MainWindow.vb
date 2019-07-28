Imports System.Linq
Imports iBMSC.Editor

Public Structure PlayerArguments
    Public Path As String
    Public aBegin As String
    Public aHere As String
    Public aStop As String
    Public Sub New(ByVal xPath As String, ByVal xBegin As String, ByVal xHere As String, ByVal xStop As String)
        Path = xPath
        aBegin = xBegin
        aHere = xHere
        aStop = xStop
    End Sub
End Structure

Public Class MainWindow


    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture" () As Integer
    Public MeasureLength(999) As Double
    Public MeasureBottom(999) As Double

    Public Function MeasureUpper(idx As Integer) As Double
        Return MeasureBottom(idx) + MeasureLength(idx)
    End Function


    Dim mColumn(999) As Integer  '0 = no column, 1 = 1 column, etc.
    Dim GreatestVPosition As Double    '+ 2000 = -VS.Minimum

    'Dim SortingMethod As Integer = 1
    Public MiddleButtonMoveMethod As Integer = 0
    Dim TextEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
    Dim DispLang As String = ""     'Display Language
    Dim Recent() As String = {"", "", "", "", ""}
    Public NTInput As Boolean = True
    Public ShowFileName As Boolean = False

    Dim BeepWhileSaved As Boolean = True
    Dim BPMx1296 As Boolean = False
    Dim STOPx1296 As Boolean = False

    Dim IsInitializing As Boolean = True
    Public IsFirstMouseEnterOnPanel As Boolean = True

    Dim WAVMultiSelect As Boolean = True
    Dim WAVChangeLabel As Boolean = True
    Dim BeatChangeMode As Integer = 0

    'Dim FloatTolerance As Double = 0.0001R
    Dim BMSGridLimit As Double = 1.0R



    'IO
    Dim FileName As String = "Untitled.bms"

    Dim InitPath As String = ""
    Dim IsSaved As Boolean = True

    'Variables for Drag/Drop
    Public DragDropFilename() As String = {}
    Dim SupportedFileExtension() As String = {".bms", ".bme", ".bml", ".pms", ".txt", ".sm", ".ibmsc"}
    Dim SupportedAudioExtension() As String = {".wav", ".mp3", ".ogg"}
    Dim SupportedImageExtension() As String = {".bmp", ".png", ".jpg", ".jpeg", ".gif", ".mpg", ".mpeg", ".avi", ".m1v", ".m2v", ".m4v", ".mp4", ".webm", ".wmv"}

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


    Public LNDisplayLength As Double = 0.0#

    'Variables for post effects tool


    'Variables for Full-Screen Mode
    Public IsFullscreen As Boolean = False
    Dim previousWindowState As FormWindowState = FormWindowState.Normal
    Dim previousWindowPosition As New Rectangle(0, 0, 0, 0)

    'Variables misc
    Dim menuVPosition As Double = 0.0#
    Dim tempResize As Integer = 0

    '----AutoSave Options
    Dim PreviousAutoSavedFileName As String = ""
    Dim AutoSaveInterval As Integer = 120000

    '----ErrorCheck Options
    Public ErrorCheck As Boolean = True

    '---- Grid Options
    Public Grid As Grid = New Grid

    '---- Columns
    Public Columns As ColumnList = New ColumnList

    '----Visual Options
    Dim vo As New visualSettings()

    Public Sub setVO(ByVal xvo As visualSettings)
        vo = xvo
    End Sub

    '----Preview Options


    Public pArgs() As PlayerArguments = {New PlayerArguments("<apppath>\uBMplay.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S"),
                                         New PlayerArguments("<apppath>\o2play.exe",
                                                             "-P -N0 ""<filename>""",
                                                             "-P -N<measure> ""<filename>""",
                                                             "-S")}
    Public CurrentPlayer As Integer = 0
    Dim PreviewOnClick As Boolean = True
    Dim PreviewErrorCheck As Boolean = False
    Dim ClickStopPreview As Boolean = True
    Dim pTempFileNames() As String = {}

    '----Split Panel Options
    Dim spLock() As Boolean = {False, False, False}
    Dim spDiff() As Integer = {0, 0, 0}
    Public PanelFocus As Integer = 1 '0 = Left, 1 = Middle, 2 = Right
    Public Property FocusedPanel As EditorPanel
        Get
            Return spMain(PanelFocus)
        End Get
        Set(value As EditorPanel)
            If value Is PMainL Then PanelFocus = 0
            If value Is PMain Then PanelFocus = 1
            If value Is PMainR Then PanelFocus = 2
        End Set
    End Property

    Dim CurrentHoveredPanel As Integer = 1

    ' az: guess these all have to do with focusing stuff?
    Public AutoFocusPanelOnMouseEnter As Boolean = False
    Public FirstClickDisabled As Boolean = True
    Public tempFirstMouseDown As Boolean = False

    Dim spMain() As EditorPanel = {}

    '----Find Delete Replace Options
    Dim fdriMesL As Integer
    Dim fdriMesU As Integer
    Dim fdriLblL As Integer
    Dim fdriLblU As Integer
    Dim fdriValL As Integer
    Dim fdriValU As Integer
    Dim fdriCol() As Integer


    Public Sub New()
        InitializeComponent()
        Audio.Initialize()

    End Sub

    ''' <summary>
    ''' Not a property as it is calculated every time. 
    ''' Not sure how to make it play nice with the SelectedNotes array.
    ''' </summary>
    ''' <returns></returns>
    Friend Function GetSelectedNotes() As IEnumerable(Of Note)
        Return Notes.Skip(1).Where(Function(x) x.Selected)
    End Function

    'Friend Sub RegenerateSelectedNotesArray()
    '    ReDim SelectedNotes(SelectedCount)
    '    SelectedNotes(0) = Notes(clickedNoteIndex)
    '    Notes(clickedNoteIndex).SelectedArrayIndex = 0
    '    Dim idx = 1

    '    ' Add already selected notes including this one
    '    For i = 1 To clickedNoteIndex - 1
    '        If Notes(i).Selected Then
    '            Notes(i).SelectedArrayIndex = idx
    '            SelectedNotes(idx) = Notes(i)
    '            idx += 1
    '        End If
    '    Next
    '    For i = clickedNoteIndex + 1 To UBound(Notes)
    '        If Notes(i).Selected Then
    '            Notes(i).SelectedArrayIndex = idx
    '            SelectedNotes(idx) = Notes(i)
    '            idx += 1
    '        End If
    '    Next
    'End Sub

    Private Sub DecreaseCurrentWav()
        If LWAV.SelectedIndex = -1 Then
            LWAV.SelectedIndex = 0
        Else
            Dim newIndex As Integer = LWAV.SelectedIndex - 1
            If newIndex < 0 Then newIndex = 0
            LWAV.SelectedIndices.Clear()
            LWAV.SelectedIndex = newIndex
        End If
    End Sub

    Private Sub IncreaseCurrentWav()
        If LWAV.SelectedIndex = -1 Then
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
    ''' Clears the SelectedNotes array.
    ''' To be figured out how this plays with DeselectAllNotes.
    ''' </summary>
    Public Sub ClearSelectionArray()
        ' ReDim SelectedNotes(-1)
    End Sub

    ''' <summary>
    ''' Call whenever the note order is no longer guaranteed.
    ''' Say, a modification in VPosition, insertion at the end or such.
    ''' </summary>
    Public Sub ValidateNotesArray()
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Public Sub PanelPreviewNoteIndex(NoteIndex As Integer)
        'Play wav
        If ClickStopPreview Then PreviewNote("", True)
        'My.Computer.Audio.Stop()
        If NoteIndex > 0 And PreviewOnClick AndAlso Columns.IsColumnSound(Notes(NoteIndex).ColumnIndex) Then
            Dim j As Integer = Notes(NoteIndex).Value \ 10000
            If j <= 0 Then j = 1
            If j >= 1296 Then j = 1295

            If Not BmsWAV(j) = "" Then ' AndAlso Path.GetExtension(hWAV(j)).ToLower = ".wav" Then
                Dim xFileLocation As String = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)) & "\" & BmsWAV(j)
                If Not ClickStopPreview Then PreviewNote("", True)
                PreviewNote(xFileLocation, False)
            End If
        End If
    End Sub

    Friend Sub SetToolstripVisible(v As Boolean)
        ToolStripContainer1.TopToolStripPanelVisible = v
    End Sub

    ' Why here instead of panel events?
    ' MainWindow should be taking care of the overall focus state, not the individual panels.
    Private Sub PMainInMouseEnter(ByVal sender As Object, ByVal e As EventArgs) Handles PMainIn.MouseEnter, PMainInL.MouseEnter, PMainInR.MouseEnter
        CurrentHoveredPanel = sender.Tag
        Dim xPMainIn As Panel = sender

        If AutoFocusPanelOnMouseEnter AndAlso Focused Then
            xPMainIn.Focus()
            PanelFocus = CurrentHoveredPanel
        End If

        If IsFirstMouseEnterOnPanel Then
            IsFirstMouseEnterOnPanel = False
            xPMainIn.Focus()
            PanelFocus = CurrentHoveredPanel
        End If
    End Sub

    Private Sub PMainInMouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PMainIn.MouseLeave, PMainInL.MouseLeave, PMainInR.MouseLeave
        State.Mouse.CurrentHoveredNoteIndex = -1
        ClearSelectionArray()
        State.Mouse.CurrentMouseRow = -1
        State.Mouse.CurrentMouseColumn = -1
        RefreshPanelAll()
    End Sub

    Public Function MeasureAtDisplacement(ByVal xVPos As Double) As Integer
        Dim i As Integer
        For i = 1 To 999
            If xVPos < MeasureBottom(i) Then Exit For
        Next
        Return i - 1
    End Function

    Public Function GetMaxVPosition() As Double
        Return MeasureUpper(999)
    End Function


    Public Sub CalculateGreatestVPosition()
        'If K Is Nothing Then Exit Sub
        Dim i As Integer
        GreatestVPosition = 0

        If NTInput Then
            For i = UBound(Notes) To 0 Step -1
                If Notes(i).VPosition + Notes(i).Length > GreatestVPosition Then GreatestVPosition = Notes(i).VPosition + Notes(i).Length
            Next
        Else
            For i = UBound(Notes) To 0 Step -1
                If Notes(i).VPosition > GreatestVPosition Then GreatestVPosition = Notes(i).VPosition
            Next
        End If

        Dim j As Integer = -CInt(IIf(GreatestVPosition + 2000 > GetMaxVPosition(), GetMaxVPosition, GreatestVPosition + 2000))
        MainPanelScroll.Minimum = j
        LeftPanelScroll.Minimum = j
        RightPanelScroll.Minimum = j
    End Sub

    Private Sub UpdateMeasureBottom()
        MeasureBottom(0) = 0.0#
        For i As Integer = 0 To 998
            MeasureBottom(i + 1) = MeasureBottom(i) + MeasureLength(i)
        Next
    End Sub

    Private Function PathIsValid(ByVal sPath As String) As Boolean
        Return File.Exists(sPath) Or Directory.Exists(sPath)
    End Function

    Public Function PrevCodeToReal(ByVal InitStr As String) As String
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                        IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                        ExcludeFileName(FileName)) _
                                        & "\___TempBMS.bms"
        Dim xMeasure As Integer = MeasureAtDisplacement(Math.Abs(FocusedPanel.VerticalScroll))
        Dim xS1 As String = Replace(InitStr, "<apppath>", My.Application.Info.DirectoryPath)
        Dim xS2 As String = Replace(xS1, "<measure>", xMeasure)
        Dim xS3 As String = Replace(xS2, "<filename>", xFileName)
        Return xS3
    End Function

    Private Sub SetFileName(ByVal xFileName As String)
        FileName = xFileName.Trim
        InitPath = ExcludeFileName(FileName)
        SetIsSaved(IsSaved)
    End Sub

    Private Sub SetIsSaved(ByVal isSaved As Boolean)
        'pttl.Refresh()
        'pIsSaved.Visible = Not xBool
        Dim xVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor &
                             IIf(My.Application.Info.Version.Build = 0, "", "." & My.Application.Info.Version.Build)
        Text = IIf(isSaved, "", "*") & GetFileName(FileName) & " - " & My.Application.Info.Title & " " & xVersion
        Me.IsSaved = isSaved
    End Sub

    Public Function IsLabelMatch(note As Note, note2 As Note) As Boolean
        If TBShowFileName.Checked Then
            Dim wavidx = note2.Value / 10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value / 10000) = wav Then
                Return True
            End If
        Else
            If note.Value = note2.Value Then
                Return True
            End If
        End If

        Return False
    End Function

    Public Function IsLabelMatch(note As Note, index As Integer) As Boolean
        If TBShowFileName.Checked Then
            Dim wavidx = Notes(index).Value / 10000
            Dim wav = BmsWAV(wavidx)
            If BmsWAV(note.Value / 10000) = wav Then
                Return True
            End If
        Else
            If note.Value = Notes(index).Value Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Sub PreviewNote(ByVal xFileLocation As String, ByVal bStop As Boolean)
        If bStop Then
            Audio.StopPlaying()
        End If
        Audio.Play(xFileLocation)
    End Sub

    Public Sub AddNote(note As Note,
               Optional ByVal xSelected As Boolean = False,
               Optional ByVal OverWrite As Boolean = True,
               Optional ByVal SortAndUpdatePairing As Boolean = True)

        If note.VPosition < 0 Or note.VPosition >= GetMaxVPosition() Then Exit Sub

        Dim i As Integer = 1

        If OverWrite Then
            Do While i <= UBound(Notes)
                If Notes(i).VPosition = note.VPosition And
                    Notes(i).ColumnIndex = note.ColumnIndex Then
                    RemoveNote(i)
                Else
                    i += 1
                End If
            Loop
        End If

        ReDim Preserve Notes(UBound(Notes) + 1)
        note.Selected = note.Selected And Columns.nEnabled(note.ColumnIndex)
        Notes(UBound(Notes)) = note

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Public Sub RemoveNote(ByVal I As Integer, Optional ByVal SortAndUpdatePairing As Boolean = True)
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
        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()

    End Sub



    Private Sub RemoveNotes(Optional ByVal SortAndUpdatePairing As Boolean = True)
        If UBound(Notes) = 0 Then Exit Sub

        State.Mouse.CurrentHoveredNoteIndex = -1
        Dim i As Integer = 1
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
        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub



    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If pTempFileNames IsNot Nothing Then
            For Each xStr As String In pTempFileNames
                IO.File.Delete(xStr)
            Next
        End If
        If PreviousAutoSavedFileName <> "" Then IO.File.Delete(PreviousAutoSavedFileName)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not IsSaved Then
            Dim xStr As String = Strings.Messages.SaveOnExit
            If e.CloseReason = CloseReason.WindowsShutDown Then xStr = Strings.Messages.SaveOnExit1
            If e.CloseReason = CloseReason.TaskManagerClosing Then xStr = Strings.Messages.SaveOnExit2

            Dim xResult As MsgBoxResult = MsgBox(xStr, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes Then
                If ExcludeFileName(FileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = InitPath

                    If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then e.Cancel = True : Exit Sub
                    SetFileName(xDSave.FileName)
                End If
                Dim xStrAll As String = SaveBMS()
                My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
                NewRecent(FileName)
                If BeepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then e.Cancel = True
        End If

        If Not e.Cancel Then
            SaveSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml", False)
        End If
    End Sub

    Private Function FilterFileBySupported(ByVal xFile() As String, ByVal xFilter() As String) As String()
        Dim xPath(-1) As String
        For i As Integer = 0 To UBound(xFile)
            If My.Computer.FileSystem.FileExists(xFile(i)) And Array.IndexOf(xFilter, Path.GetExtension(xFile(i))) <> -1 Then
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

    Private Sub InitializeNewBMS()
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
        For i As Integer = 0 To 999
            MeasureLength(i) = 192.0R
            MeasureBottom(i) = i * 192.0R
            LBeat.Items.Add(Add3Zeros(i) & ": 1 ( 4 / 4 )")
        Next
    End Sub

    Private Sub InitializeOpenBMS()
        CHPlayer.SelectedIndex = 0
        'THLnType.Text = ""
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedFileExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedFileExtension)
        If xPath.Length > 0 Then
            Dim xProg As New fLoadFileProgress(xPath, IsSaved)
            xProg.ShowDialog(Me)
        End If

        RefreshPanelAll()
    End Sub

    Private Sub SetFullScreen(ByVal value As Boolean)
        If value Then
            If Me.WindowState = FormWindowState.Minimized Then Exit Sub

            SuspendLayout()
            previousWindowPosition.Location = Me.Location
            previousWindowPosition.Size = Me.Size
            previousWindowState = Me.WindowState

            WindowState = FormWindowState.Normal
            FormBorderStyle = Windows.Forms.FormBorderStyle.None
            WindowState = FormWindowState.Maximized
            ToolStripContainer1.TopToolStripPanelVisible = False

            ResumeLayout()
            IsFullscreen = True
        Else
            SuspendLayout()
            FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            ToolStripContainer1.TopToolStripPanelVisible = True
            WindowState = FormWindowState.Normal

            WindowState = previousWindowState
            If Me.WindowState = FormWindowState.Normal Then
                Location = previousWindowPosition.Location
                Size = previousWindowPosition.Size
            End If

            ResumeLayout()
            IsFullscreen = False
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F11
                SetFullScreen(Not IsFullscreen)
        End Select
    End Sub

    Private Sub Form1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Friend Sub ReadFile(ByVal xPath As String)
        Select Case LCase(Path.GetExtension(xPath))
            Case ".bms", ".bme", ".bml", ".pms", ".txt"
                OpenBMS(My.Computer.FileSystem.ReadAllText(xPath, TextEncoding))
                ClearUndo()
                NewRecent(xPath)
                SetFileName(xPath)
                SetIsSaved(True)

            Case ".sm"
                If OpenSM(My.Computer.FileSystem.ReadAllText(xPath, TextEncoding)) Then Return
                InitPath = ExcludeFileName(xPath)
                ClearUndo()
                SetFileName("Untitled.bms")
                SetIsSaved(False)

            Case ".ibmsc"
                OpeniBMSC(xPath)
                InitPath = ExcludeFileName(xPath)
                NewRecent(xPath)
                SetFileName("Imported_" & GetFileName(xPath))
                SetIsSaved(False)

        End Select
    End Sub


    Public Function GCD(ByVal NumA As Double, ByVal NumB As Double) As Double
        Dim xNMax As Double = NumA
        Dim xNMin As Double = NumB
        If NumA < NumB Then
            xNMax = NumB
            xNMin = NumA
        End If
        Do While xNMin >= BMSGridLimit
            GCD = xNMax - Math.Floor(xNMax / xNMin) * xNMin
            xNMax = xNMin
            xNMin = GCD
        Loop
        GCD = xNMax
    End Function

    <DllImport("user32.dll")> Private Shared Function LoadCursorFromFile(ByVal fileName As String) As IntPtr
    End Function
    Public Shared Function ActuallyLoadCursor(ByVal path As String) As Cursor
        Return New Cursor(LoadCursorFromFile(path))
    End Function

    Private Sub Unload() Handles MyBase.Disposed
        Audio.Finalize()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'On Error Resume Next
        TopMost = True
        SuspendLayout()
        Visible = False

        SetFileName(FileName)

        InitializeNewBMS()

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

            SpL.Cursor = xRightCursor
            SpR.Cursor = xLeftCursor
        Catch ex As Exception

        End Try

        spMain = New EditorPanel() {PMainInL, PMainIn, PMainInR}
        PMainInL.Init(Me, vo, LeftPanelScroll, HSL)
        PMainIn.Init(Me, vo, MainPanelScroll, HS)
        PMainInR.Init(Me, vo, RightPanelScroll, HSR)
        PMain.SendToBack()

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

        CalculateGreatestVPosition()
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
            If LCase(Path.GetExtension(xStr(1))) = ".ibmsc" AndAlso GetFileName(xStr(1)).StartsWith("AutoSave_", True, Nothing) Then GoTo 1000
        End If

        'pIsSaved.Visible = Not IsSaved
        IsInitializing = False

        If Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length > 1 Then GoTo 1000
        Dim xFiles() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath).GetFiles("AutoSave_*.IBMSC")
        If xFiles Is Nothing OrElse xFiles.Length = 0 Then GoTo 1000

        'Me.TopMost = True
        If MsgBox(Replace(Strings.Messages.RestoreAutosavedFile, "{}", xFiles.Length), MsgBoxStyle.YesNo Or MsgBoxStyle.MsgBoxSetForeground) = MsgBoxResult.Yes Then
            For Each xF As FileInfo In xFiles
                'MsgBox(xF.FullName)
                System.Diagnostics.Process.Start(Application.ExecutablePath, """" & xF.FullName & """")
            Next
        End If

        For Each xF As FileInfo In xFiles
            ReDim Preserve pTempFileNames(UBound(pTempFileNames) + 1)
            pTempFileNames(UBound(pTempFileNames)) = xF.FullName
        Next

1000:
        IsInitializing = False
        POStatusRefresh()
        Me.ResumeLayout()

        tempResize = Me.WindowState
        Me.TopMost = False
        Me.WindowState = tempResize

        Me.Visible = True
    End Sub

    Private Sub UpdatePairing()
        Dim i As Integer, j As Integer

        If NTInput Then
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

        Dim currentMS = 0.0#
        Dim currentBPM = Notes(0).Value / 10000
        Dim currentBPMVPosition = 0.0#
        For i = 1 To UBound(Notes)
            If Notes(i).ColumnIndex = ColumnType.BPM Then
                currentMS += (Notes(i).VPosition - currentBPMVPosition) / currentBPM * 1250
                currentBPM = Notes(i).Value / 10000
                currentBPMVPosition = Notes(i).VPosition
            End If
            'K(i).TimeOffset = currentMS + (K(i).VPosition - currentBPMVPosition) / currentBPM * 1250
        Next
    End Sub

    ' az: Handle zoom in/out. Should work with any of the three splitters.
    Private Sub PMain_Scroll(sender As Object, e As MouseEventArgs) Handles PMainIn.MouseWheel, PMainInL.MouseWheel, PMainInR.MouseWheel
        If Not My.Computer.Keyboard.CtrlKeyDown Then Exit Sub
        Dim dv = Math.Round(CGHeight2.Value + e.Delta / 120)
        CGHeight2.Value = Math.Min(CGHeight2.Maximum, Math.Max(CGHeight2.Minimum, dv))
        CGHeight.Value = CGHeight2.Value / 4
    End Sub

    Public Sub ExceptionSave(ByVal Path As String)
        SaveiBMSC(Path)
    End Sub

    ''' <summary>
    ''' True if pressed cancel. False elsewise.
    ''' </summary>
    ''' <returns>True if pressed cancel. False elsewise.</returns>

    Private Function ClosingPopSave() As Boolean
        If Not IsSaved Then
            Dim xResult As MsgBoxResult = MsgBox(Strings.Messages.SaveOnExit, MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, Me.Text)

            If xResult = MsgBoxResult.Yes Then
                If ExcludeFileName(FileName) = "" Then
                    Dim xDSave As New SaveFileDialog
                    xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                                    Strings.FileType.BMS & "|*.bms|" &
                                    Strings.FileType.BME & "|*.bme|" &
                                    Strings.FileType.BML & "|*.bml|" &
                                    Strings.FileType.PMS & "|*.pms|" &
                                    Strings.FileType.TXT & "|*.txt|" &
                                    Strings.FileType._all & "|*.*"
                    xDSave.DefaultExt = "bms"
                    xDSave.InitialDirectory = InitPath

                    If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Return True
                    SetFileName(xDSave.FileName)
                End If
                Dim xStrAll As String = SaveBMS()
                My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
                NewRecent(FileName)
                If BeepWhileSaved Then Beep()
            End If

            If xResult = MsgBoxResult.Cancel Then Return True
        End If
        Return False
    End Function

    Private Sub TBNew_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles TBNew.Click, mnNew.Click

        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        ClearUndo()
        InitializeNewBMS()

        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim BmsWAV(1295)
        ReDim BmsBMP(1295)
        ReDim BmsBPM(1295)    'x10000
        ReDim BmsSTOP(1295)
        ReDim BmsSCROLL(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        With Notes(0)
            .ColumnIndex = ColumnType.BPM
            .VPosition = -1
            .Value = 1200000
        End With
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
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub



    Private Sub TBOpen_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpen.ButtonClick, mnOpen.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog With {
            .Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt",
            .DefaultExt = "bms",
            .InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        }

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDOpen.FileName)
        OpenBMS(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, TextEncoding))
        ClearUndo()
        SetFileName(xDOpen.FileName)
        NewRecent(FileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBImportIBMSC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBImportIBMSC.Click, mnImportIBMSC.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Return

        Dim xDOpen As New OpenFileDialog With {
            .Filter = Strings.FileType.IBMSC & "|*.ibmsc",
            .DefaultExt = "ibmsc",
            .InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        }

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Return
        InitPath = ExcludeFileName(xDOpen.FileName)
        SetFileName("Imported_" & GetFileName(xDOpen.FileName))
        OpeniBMSC(xDOpen.FileName)
        NewRecent(xDOpen.FileName)
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBImportSM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBImportSM.Click, mnImportSM.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog With {
            .Filter = Strings.FileType.SM & "|*.sm",
            .DefaultExt = "sm",
            .InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        }

        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        If OpenSM(My.Computer.FileSystem.ReadAllText(xDOpen.FileName, TextEncoding)) Then Exit Sub
        InitPath = ExcludeFileName(xDOpen.FileName)
        SetFileName("Untitled.bms")
        ClearUndo()
        SetIsSaved(False)
        'pIsSaved.Visible = Not IsSaved
    End Sub

    Private Sub TBSave_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSave.ButtonClick, mnSave.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1

        If ExcludeFileName(FileName) = "" Then
            Dim xDSave As New SaveFileDialog With {
                .Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                            Strings.FileType.BMS & "|*.bms|" &
                            Strings.FileType.BME & "|*.bme|" &
                            Strings.FileType.BML & "|*.bml|" &
                            Strings.FileType.PMS & "|*.pms|" &
                            Strings.FileType.TXT & "|*.txt|" &
                            Strings.FileType._all & "|*.*",
                .DefaultExt = "bms",
                .InitialDirectory = InitPath
            }

            If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
            InitPath = ExcludeFileName(xDSave.FileName)
            SetFileName(xDSave.FileName)
        End If
        Dim xStrAll As String = SaveBMS()
        My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        SetFileName(FileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
        If BeepWhileSaved Then Beep()
    End Sub

    Private Sub TBSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSaveAs.Click, mnSaveAs.Click
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
            .InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        }

        If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDSave.FileName)
        SetFileName(xDSave.FileName)
        Dim xStrAll As String = SaveBMS()
        My.Computer.FileSystem.WriteAllText(FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        SetFileName(FileName)
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved
        If BeepWhileSaved Then Beep()
    End Sub

    Private Sub TBExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBExport.Click, mnExport.Click
        'KMouseDown = -1
        ClearSelectionArray()
        State.Mouse.CurrentHoveredNoteIndex = -1

        Dim xDSave As New SaveFileDialog With {
            .Filter = Strings.FileType.IBMSC & "|*.ibmsc",
            .DefaultExt = "ibmsc",
            .InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        }
        If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        SaveiBMSC(xDSave.FileName)
        'My.Computer.FileSystem.WriteAllText(xDSave.FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        If BeepWhileSaved Then Beep()
    End Sub



    Private Sub VSGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainPanelScroll.GotFocus, LeftPanelScroll.GotFocus, RightPanelScroll.GotFocus
        PanelFocus = sender.Tag
        FocusedPanel.Focus()
    End Sub

    Private Sub PanelVerticalScrollChanged(ByVal sender As Object,
                                           ByVal e As EventArgs) Handles MainPanelScroll.ValueChanged,
        LeftPanelScroll.ValueChanged,
        RightPanelScroll.ValueChanged
        Dim iI As Integer = sender.Tag

        ' az: We got a wheel event when we're zooming in/out
        If My.Computer.Keyboard.CtrlKeyDown Then
            sender.Value = FocusedPanel.LastVerticalScroll ' Undo the scroll
            Exit Sub
        End If

        Dim currentPanel = spMain(iI)
        If iI = PanelFocus And
            Not State.Mouse.LastMouseDownLocation = New Point(-1, -1) And
            Not FocusedPanel.LastVerticalScroll = -1 Then
            State.Mouse.LastMouseDownLocation.Y += (FocusedPanel.LastVerticalScroll - sender.Value) * Grid.HeightScale
        End If


        If spLock((iI + 1) Mod 3) Then
            Dim verticalScroll As Integer = currentPanel.VerticalScroll + spDiff(iI)
            If verticalScroll > 0 Then verticalScroll = 0
            If verticalScroll < MainPanelScroll.Minimum Then verticalScroll = MainPanelScroll.Minimum
            Select Case iI
                Case 0 : MainPanelScroll.Value = verticalScroll
                Case 1 : RightPanelScroll.Value = verticalScroll
                Case 2 : LeftPanelScroll.Value = verticalScroll
            End Select
        End If

        If spLock((iI + 2) Mod 3) Then
            Dim verticalScroll As Integer = currentPanel.VerticalScroll - spDiff((iI + 2) Mod 3)
            If verticalScroll > 0 Then verticalScroll = 0
            If verticalScroll < MainPanelScroll.Minimum Then verticalScroll = MainPanelScroll.Minimum
            Select Case iI
                Case 0 : RightPanelScroll.Value = verticalScroll
                Case 1 : LeftPanelScroll.Value = verticalScroll
                Case 2 : MainPanelScroll.Value = verticalScroll
            End Select
        End If

        spDiff(iI) = spMain((iI + 1) Mod 3).VerticalScroll - currentPanel.VerticalScroll
        spDiff((iI + 2) Mod 3) = currentPanel.VerticalScroll - spMain((iI + 2) Mod 3).VerticalScroll


    End Sub

    Private Sub cVSLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cVSLockL.CheckedChanged, cVSLock.CheckedChanged, cVSLockR.CheckedChanged
        Dim iI As Integer = sender.Tag
        spLock(iI) = sender.Checked
        If Not spLock(iI) Then Return

        Dim currentPanel = spMain(iI)
        spDiff(iI) = spMain((iI + 1) Mod 3).VerticalScroll - currentPanel.VerticalScroll
        spDiff((iI + 2) Mod 3) = currentPanel.VerticalScroll - spMain((iI + 2) Mod 3).VerticalScroll

    End Sub

    Private Sub HSGotFocus(ByVal sender As Object, ByVal e As EventArgs) Handles HS.GotFocus, HSL.GotFocus, HSR.GotFocus
        PanelFocus = sender.Tag
        spMain(PanelFocus).Focus()
    End Sub

    Private Sub HSValueChanged(ByVal sender As Object, ByVal e As EventArgs) Handles HS.ValueChanged, HSL.ValueChanged, HSR.ValueChanged
        If Not State.Mouse.LastMouseDownLocation = New Point(-1, -1) Then
            State.Mouse.LastMouseDownLocation.X += (FocusedPanel.LastHorizontalScroll - sender.Value) * Grid.WidthScale
        End If
    End Sub

    Private Sub TBSelect_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles TBSelect.Click, mnSelect.Click
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
        LNDisplayLength = 0

        State.TimeSelect.StartPoint = MeasureBottom(MeasureAtDisplacement(-FocusedPanel.VerticalScroll) + 1)
        State.TimeSelect.EndPointLength = 0

        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBWrite.Click, mnWrite.Click
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
        LNDisplayLength = 0

        State.TimeSelect.StartPoint = MeasureBottom(MeasureAtDisplacement(-FocusedPanel.VerticalScroll) + 1)
        State.TimeSelect.EndPointLength = 0

        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBPostEffects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBTimeSelect.Click, mnTimeSelect.Click
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
        LNDisplayLength = 0
        State.TimeSelect.ValidateSelection(GetMaxVPosition())

        Dim i As Integer
        For i = 0 To UBound(Notes)
            Notes(i).Selected = False
        Next
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub CGHeight_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight.ValueChanged
        Grid.HeightScale = CSng(CGHeight.Value)
        CGHeight2.Value = IIf(CGHeight.Value * 4 < CGHeight2.Maximum, CDec(CGHeight.Value * 4), CGHeight2.Maximum)
        RefreshPanelAll()
    End Sub

    Private Sub CGHeight2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight2.Scroll
        CGHeight.Value = CGHeight2.Value / 4
    End Sub

    Private Sub CGWidth_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth.ValueChanged
        Grid.WidthScale = CSng(CGWidth.Value)
        CGWidth2.Value = IIf(CGWidth.Value * 4 < CGWidth2.Maximum, CDec(CGWidth.Value * 4), CGWidth2.Maximum)

        HS.LargeChange = PMainIn.Width / Grid.WidthScale
        If HS.Value > HS.Maximum - HS.LargeChange + 1 Then HS.Value = HS.Maximum - HS.LargeChange + 1
        HSL.LargeChange = PMainInL.Width / Grid.WidthScale
        If HSL.Value > HSL.Maximum - HSL.LargeChange + 1 Then HSL.Value = HSL.Maximum - HSL.LargeChange + 1
        HSR.LargeChange = PMainInR.Width / Grid.WidthScale
        If HSR.Value > HSR.Maximum - HSR.LargeChange + 1 Then HSR.Value = HSR.Maximum - HSR.LargeChange + 1

        RefreshPanelAll()
    End Sub

    Private Sub CGWidth2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth2.Scroll
        CGWidth.Value = CGWidth2.Value / 4
    End Sub

    Private Sub CGDivide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDivide.ValueChanged
        Grid.Divider = CGDivide.Value
        RefreshPanelAll()
    End Sub
    Private Sub CGSub_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSub.ValueChanged
        Grid.Subdivider = CGSub.Value
        RefreshPanelAll()
    End Sub
    Private Sub BGSlash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BGSlash.Click
        Dim xd As Integer = Val(InputBox(Strings.Messages.PromptSlashValue, , Grid.Slash))
        If xd = 0 Then Exit Sub
        If xd > CGDivide.Maximum Then xd = CGDivide.Maximum
        If xd < CGDivide.Minimum Then xd = CGDivide.Minimum
        Grid.Slash = xd
    End Sub


    Private Sub CGSnap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSnap.CheckedChanged
        Grid.IsSnapEnabled = CGSnap.Checked
        RefreshPanelAll()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim i As Integer

        With FocusedPanel.VerticalScrollBar
            i = .Value + (State.Mouse.tempY / 5) / Grid.HeightScale
            If i > 0 Then i = 0
            If i < .Minimum Then i = .Minimum
            .Value = i
        End With
        With FocusedPanel.HorizontalScrollBar
            i = .Value + (State.Mouse.tempX / 10) / Grid.WidthScale
            If i > .Maximum - .LargeChange + 1 Then i = .Maximum - .LargeChange + 1
            If i < .Minimum Then i = .Minimum
            .Value = i
        End With

        Dim xMEArgs As New MouseEventArgs(MouseButtons.Left, 0, State.Mouse.MouseMoveStatus.X, State.Mouse.MouseMoveStatus.Y, 0)
        FocusedPanel.PMainInMouseMove(Me, xMEArgs)
    End Sub

    Private Sub TimerMiddle_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMiddle.Tick
        If Not State.Mouse.MiddleButtonClicked Then
            TimerMiddle.Enabled = False
            Return
        End If

        Dim i As Integer

        With FocusedPanel.VerticalScrollBar
            i = .Value + (Cursor.Position.Y - State.Mouse.MiddleButtonLocation.Y) / 5 / Grid.HeightScale
            If i > 0 Then i = 0
            If i < .Minimum Then i = .Minimum
            .Value = i
        End With
        With FocusedPanel.HorizontalScrollBar
            i = .Value + (Cursor.Position.X - State.Mouse.MiddleButtonLocation.X) / 5 / Grid.WidthScale
            If i > .Maximum - .LargeChange + 1 Then i = .Maximum - .LargeChange + 1
            If i < .Minimum Then i = .Minimum
            .Value = i
        End With


        Dim xMEArgs As New MouseEventArgs(MouseButtons.Left, 0,
                                          State.Mouse.MouseMoveStatus.X,
                                          State.Mouse.MouseMoveStatus.Y,
                                          0)
        FocusedPanel.PMainInMouseMove(Me, xMEArgs)
    End Sub

    Private Sub ValidateWavListView()
        Try
            Dim xRect As Rectangle = LWAV.GetItemRectangle(LWAV.SelectedIndex)
            If xRect.Top + xRect.Height > LWAV.DisplayRectangle.Height Then SendMessage(LWAV.Handle, &H115, 1, 0)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LWAV_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LWAV.Click
        If TBWrite.Checked Then FSW.Text = C10to36(LWAV.SelectedIndex + 1)

        PreviewNote("", True)
        If Not PreviewOnClick Then Exit Sub
        If BmsWAV(LWAV.SelectedIndex + 1) = "" Then Exit Sub

        Dim xFileLocation As String = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)) & "\" & BmsWAV(LWAV.SelectedIndex + 1)
        PreviewNote(xFileLocation, False)
    End Sub

    Private Sub LWAV_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LWAV.DoubleClick
        Dim xDWAV As New OpenFileDialog
        xDWAV.DefaultExt = "wav"
        xDWAV.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDWAV.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)
        BmsWAV(LWAV.SelectedIndex + 1) = GetFileName(xDWAV.FileName)
        LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": " & GetFileName(xDWAV.FileName)
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub LWAV_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LWAV.KeyDown
        Select Case e.KeyCode
            Case Keys.Delete
                BmsWAV(LWAV.SelectedIndex + 1) = ""
                LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": "
                If IsSaved Then SetIsSaved(False)
        End Select
    End Sub

    Private Sub LBMP_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles LBMP.DoubleClick
        Dim xDBMP As New OpenFileDialog
        xDBMP.DefaultExt = "bmp"
        xDBMP.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.jpeg;.gif|" &
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
        xDBMP.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

        If xDBMP.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDBMP.FileName)
        BmsBMP(LBMP.SelectedIndex + 1) = GetFileName(xDBMP.FileName)
        LBMP.Items.Item(LBMP.SelectedIndex) = C10to36(LBMP.SelectedIndex + 1) & ": " & GetFileName(xDBMP.FileName)
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub LBMP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LBMP.KeyDown
        Select Case e.KeyCode
            Case Keys.Delete
                BmsBMP(LBMP.SelectedIndex + 1) = ""
                LBMP.Items.Item(LBMP.SelectedIndex) = C10to36(LBMP.SelectedIndex + 1) & ": "
                If IsSaved Then SetIsSaved(False)
        End Select
    End Sub

    Private Sub TBErrorCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBErrorCheck.Click, mnErrorCheck.Click
        ErrorCheck = sender.Checked
        TBErrorCheck.Checked = ErrorCheck
        mnErrorCheck.Checked = ErrorCheck
        TBErrorCheck.Image = IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN)
        mnErrorCheck.Image = IIf(TBErrorCheck.Checked, My.Resources.x16CheckError, My.Resources.x16CheckErrorN)
        RefreshPanelAll()
    End Sub

    Private Sub TBPreviewOnClick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPreviewOnClick.Click, mnPreviewOnClick.Click
        PreviewNote("", True)
        PreviewOnClick = sender.Checked
        TBPreviewOnClick.Checked = PreviewOnClick
        mnPreviewOnClick.Checked = PreviewOnClick
        TBPreviewOnClick.Image = IIf(PreviewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN)
        mnPreviewOnClick.Image = IIf(PreviewOnClick, My.Resources.x16PreviewOnClick, My.Resources.x16PreviewOnClickN)
    End Sub

    'Private Sub TBPreviewErrorCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    PreviewErrorCheck = TBPreviewErrorCheck.Checked
    '    TBPreviewErrorCheck.Image = IIf(PreviewErrorCheck, My.Resources.x16PreviewCheck, My.Resources.x16PreviewCheckN)
    'End Sub

    Private Sub TBShowFileName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBShowFileName.Click, mnShowFileName.Click
        ShowFileName = sender.Checked
        TBShowFileName.Checked = ShowFileName
        mnShowFileName.Checked = ShowFileName
        TBShowFileName.Image = IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN)
        mnShowFileName.Image = IIf(ShowFileName, My.Resources.x16ShowFileName, My.Resources.x16ShowFileNameN)
        RefreshPanelAll()
    End Sub

    Private Sub TBCut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCut.Click, mnCut.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Command.RedoRemoveNoteSelected(Notes, xUndo, xRedo)
        'Dim xRedo As String = sCmdKDs()
        'Dim xUndo As String = sCmdKs(True)

        CopyNotes(False)
        RemoveNotes(False)
        AddUndoChain(xUndo, xBaseRedo.Next)

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()
    End Sub

    Private Sub TBCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBCopy.Click, mnCopy.Click
        CopyNotes()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPaste.Click, mnPaste.Click
        AddNotesFromClipboard()

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        Command.RedoAddNoteSelected(Notes, xUndo, xRedo)
        AddUndoChain(xUndo, xBaseRedo.Next)

        'AddUndo(sCmdKDs(), sCmdKs(True))

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()
    End Sub

    'Private Function pArgPath(ByVal I As Integer)
    '    Return Mid(pArgs(I), 1, InStr(pArgs(I), vbCrLf) - 1)
    'End Function

    Private Function GetFileName(ByVal s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        Return Mid(s, IIf(fslash > bslash, fslash, bslash) + 1)
    End Function

    Private Function ExcludeFileName(ByVal s As String) As String
        Dim fslash As Integer = InStrRev(s, "/")
        Dim bslash As Integer = InStrRev(s, "\")
        If (bslash Or fslash) = 0 Then Return ""
        Return Mid(s, 1, IIf(fslash > bslash, fslash, bslash) - 1)
    End Function

    Private Sub PlayerMissingPrompt()
        Dim xArg As PlayerArguments = pArgs(CurrentPlayer)
        MsgBox(Strings.Messages.CannotFind.Replace("{}", PrevCodeToReal(xArg.Path)) & vbCrLf &
               Strings.Messages.PleaseRespecifyPath, MsgBoxStyle.Critical, Strings.Messages.PlayerNotFound)

        Dim xDOpen As New OpenFileDialog
        xDOpen.InitialDirectory = IIf(ExcludeFileName(PrevCodeToReal(xArg.Path)) = "",
                                      My.Application.Info.DirectoryPath,
                                      ExcludeFileName(PrevCodeToReal(xArg.Path)))
        xDOpen.FileName = PrevCodeToReal(xArg.Path)
        xDOpen.Filter = Strings.FileType.EXE & "|*.exe"
        xDOpen.DefaultExt = "exe"
        If xDOpen.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        'pArgs(CurrentPlayer) = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>") & _
        '                                           Mid(pArgs(CurrentPlayer), InStr(pArgs(CurrentPlayer), vbCrLf))
        'xStr = Split(pArgs(CurrentPlayer), vbCrLf)
        pArgs(CurrentPlayer).Path = Replace(xDOpen.FileName, My.Application.Info.DirectoryPath, "<apppath>")
        xArg = pArgs(CurrentPlayer)
    End Sub

    Private Sub TBPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPlay.Click, mnPlay.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        ' az: Treat it like we cancelled the operation
        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)
        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aHere))
    End Sub

    Private Sub TBPlayB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPlayB.Click, mnPlayB.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)

        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aBegin))
    End Sub

    Private Sub TBStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBStop.Click, mnStop.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            Exit Sub
        End If

        Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aStop))
    End Sub

    Private Sub AddTempFileList(ByVal s As String)
        Dim xAdd As Boolean = True
        If pTempFileNames IsNot Nothing Then
            For Each xStr1 As String In pTempFileNames
                If xStr1 = s Then xAdd = False : Exit For
            Next
        End If

        If xAdd Then
            ReDim Preserve pTempFileNames(UBound(pTempFileNames) + 1)
            pTempFileNames(UBound(pTempFileNames)) = s
        End If
    End Sub

    Private Sub TBStatistics_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBStatistics.Click, mnStatistics.Click
        SortByVPositionInsertion()
        UpdatePairing()

        Dim data(6, 5) As Integer
        For i As Integer = 1 To UBound(Notes)
            With Notes(i)
                Dim row As Integer = -1
                Select Case .ColumnIndex
                    Case ColumnType.BPM : row = 0
                    Case ColumnType.STOPS : row = 1
                    Case ColumnType.SCROLLS : row = 2
                    Case ColumnType.A1, ColumnType.A2, ColumnType.A3, ColumnType.A4, ColumnType.A5, ColumnType.A6, ColumnType.A7, ColumnType.A8 : row = 3
                    Case ColumnType.D1, ColumnType.D2, ColumnType.D3, ColumnType.D4, ColumnType.D5, ColumnType.D6, ColumnType.D7, ColumnType.D8 : row = 4
                    Case Is >= ColumnType.BGM : row = 5
                    Case Else : row = 6
                End Select


StartCount:     If Not NTInput Then
                    If Not .LongNote Then data(row, 0) += 1
                    If .LongNote Then data(row, 1) += 1
                    If .Value \ 10000 = LnObj Then data(row, 2) += 1
                    If .Hidden Then data(row, 3) += 1
                    If .HasError Then data(row, 4) += 1
                    data(row, 5) += 1

                Else
                    Dim noteUnit As Integer = 1
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
    ''' Remark: Pls sort and updatepairing before this process.
    ''' </summary>
    Public Sub CalculateTotalPlayableNotes()
        Dim i As Integer
        Dim xIAll As Integer = 0

        If Not NTInput Then
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
        Dim panDisplacement = FocusedPanel.VerticalScroll
        Dim vpos = (panHeight - panDisplacement * Grid.HeightScale - State.Mouse.MouseMoveStatus.Y - 1) / Grid.HeightScale
        If snap Then
            Return FocusedPanel.SnapToGrid(vpos)
        Else
            Return vpos
        End If
    End Function

    Public Sub POStatusRefresh()

        If TBSelect.Checked Then
            Dim i As Integer = State.Mouse.CurrentHoveredNoteIndex
            If i < 0 Then

                UpdateMouseRowAndColumn()

                Dim xMeasure As Integer = MeasureAtDisplacement(State.Mouse.CurrentMouseRow)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = State.Mouse.CurrentMouseRow - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = State.Mouse.CurrentMouseRow.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
                FSC.Text = Columns.nTitle(State.Mouse.CurrentMouseColumn)
                FSW.Text = ""
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = ""
                FSH.Text = ""
                FSE.Text = ""

            Else
                Dim xMeasure As Integer = MeasureAtDisplacement(Notes(i).VPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = Notes(i).VPosition - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = Notes(i).VPosition.ToString() & "  "
                TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
                FSC.Text = Columns.nTitle(Notes(i).ColumnIndex)
                FSW.Text = IIf(Columns.IsColumnNumeric(Notes(i).ColumnIndex),
                               Notes(i).Value / 10000,
                               C10to36(Notes(i).Value \ 10000))
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = IIf(NTInput, Strings.StatusBar.Length & " = " & Notes(i).Length, IIf(Notes(i).LongNote, Strings.StatusBar.LongNote, ""))
                FSH.Text = IIf(Notes(i).Hidden, Strings.StatusBar.Hidden, "")
                FSE.Text = IIf(Notes(i).HasError, Strings.StatusBar.Err, "")

            End If

        ElseIf TBWrite.Checked Then
            If State.Mouse.CurrentMouseColumn < 0 Then Exit Sub

            Dim xMeasure As Integer = MeasureAtDisplacement(State.Mouse.CurrentMouseRow)
            Dim xMLength As Double = MeasureLength(xMeasure)
            Dim xVposMod As Double = State.Mouse.CurrentMouseRow - MeasureBottom(xMeasure)
            Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

            FSP1.Text = (xVposMod * Grid.Divider / 192).ToString & " / " & (xMLength * Grid.Divider / 192).ToString & "  "
            FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
            FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
            FSP4.Text = State.Mouse.CurrentMouseRow.ToString() & "  "
            TimeStatusLabel.Text = GetTimeFromVPosition(State.Mouse.CurrentMouseRow).ToString("F4")
            FSC.Text = Columns.nTitle(State.Mouse.CurrentMouseColumn)
            FSW.Text = C10to36(LWAV.SelectedIndex + 1)
            FSM.Text = Add3Zeros(xMeasure)
            FST.Text = IIf(NTInput, LNDisplayLength, IIf(My.Computer.Keyboard.ShiftKeyDown, Strings.StatusBar.LongNote, ""))
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
        Dim timing_notes = (From note In Notes
                            Where note.ColumnIndex = ColumnType.BPM Or note.ColumnIndex = ColumnType.STOPS
                            Group By Column = note.ColumnIndex
                               Into NoteGroups = Group).ToDictionary(Function(x) x.Column, Function(x) x.NoteGroups)

        Dim bpm_notes = timing_notes.Item(ColumnType.BPM)

        Dim stop_notes As IEnumerable(Of Note) = Nothing

        If timing_notes.ContainsKey(ColumnType.STOPS) Then
            stop_notes = timing_notes.Item(ColumnType.STOPS)
        End If


        Dim stop_contrib As Double
        Dim bpm_contrib As Double

        For i = 0 To bpm_notes.Count() - 1
            ' az: sum bpm contribution first
            Dim duration = 0.0
            Dim current_note = bpm_notes.ElementAt(i)
            Dim notevpos = Math.Max(0, current_note.VPosition)

            If i + 1 <> bpm_notes.Count() Then
                Dim next_note = bpm_notes.ElementAt(i + 1)
                duration = next_note.VPosition - notevpos
            Else
                duration = vpos - notevpos
            End If

            Dim current_bps = 60 / (current_note.Value / 10000)
            bpm_contrib += current_bps * duration / 48

            If stop_notes Is Nothing Then Continue For

            Dim stops = From stp In stop_notes
                        Where stp.VPosition >= notevpos And
                            stp.VPosition < notevpos + duration

            Dim stop_beats = stops.Sum(Function(x) x.Value / 10000.0) / 48
            stop_contrib += current_bps * stop_beats

        Next

        Return stop_contrib + bpm_contrib
    End Function

    Private Sub POBStorm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBStorm.Click

    End Sub

    Private Sub POBMirror_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBMirror.Click
        Dim i As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        'xRedo &= sCmdKM(ColumnType.A1, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(ColumnType.A7) - RealColumnToEnabled(ColumnType.A1), 0, True) & vbCrLf
        'xUndo &= sCmdKM(ColumnType.A7, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(ColumnType.A1) - RealColumnToEnabled(ColumnType.A7), 0, True) & vbCrLf

        Dim xCol As Integer = 0
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

            Command.RedoMoveNote(Notes(i), xCol, Notes(i).VPosition, xUndo, xRedo)
            Notes(i).ColumnIndex = xCol
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub





    Private Sub TVCM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCM.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCM.Text = Val(TVCM.Text)
            If Val(TVCM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCM.Text = 1
                TVCM.Focus()
                TVCM.SelectAll()
            Else
                BVCApply_Click(BVCApply, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCM_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCM.LostFocus
        TVCM.Text = Val(TVCM.Text)
        If Val(TVCM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeFactorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCM.Text = 1
            TVCM.Focus()
            TVCM.SelectAll()
        End If
    End Sub

    Private Sub TVCD_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCD.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCD.Text = Val(TVCD.Text)
            If Val(TVCD.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCD.Text = 1
                TVCD.Focus()
                TVCD.SelectAll()
            Else
                BVCApply_Click(BVCApply, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCD_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCD.LostFocus
        TVCD.Text = Val(TVCD.Text)
        If Val(TVCD.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCD.Text = 1
            TVCD.Focus()
            TVCD.SelectAll()
        End If
    End Sub

    Private Sub TVCBPM_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TVCBPM.KeyDown
        If e.KeyCode = Keys.Enter Then
            TVCBPM.Text = Val(TVCBPM.Text)
            If Val(TVCBPM.Text) <= 0 Then
                MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
                TVCBPM.Text = Notes(0).Value / 10000
                TVCBPM.Focus()
                TVCBPM.SelectAll()
            Else
                BVCCalculate_Click(BVCCalculate, New System.EventArgs)
            End If
        End If
    End Sub

    Private Sub TVCBPM_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TVCBPM.LostFocus
        TVCBPM.Text = Val(TVCBPM.Text)
        If Val(TVCBPM.Text) <= 0 Then
            MsgBox(Strings.Messages.NegativeDivisorError, MsgBoxStyle.Critical, Strings.Messages.Err)
            TVCBPM.Text = Notes(0).Value / 10000
            TVCBPM.Focus()
            TVCBPM.SelectAll()
        End If
    End Sub






    Private Sub TBUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUndo.Click, mnUndo.Click
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

    Private Sub TBRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRedo.Click, mnRedo.Click
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

    'Friend Sub SelectSingleNote(clickedNote As Note)
    '    ReDim SelectedNotes(0)
    '    SelectedNotes(0) = clickedNote
    'End Sub

    Public Sub AppendNote(note As Note)
        Notes = Notes.Concat({note}).ToArray()

        ' SelectSingleNote(note)

        If AutoincreaseWavIndex Then
            IncreaseCurrentWav()
        End If

        State.uAdded = False

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = Nothing
        RedoAddNote(note, xUndo, xRedo, AutoincreaseWavIndex)
        AddUndoChain(xUndo, xRedo)
    End Sub

    Friend Sub DeselectAllNotes()
        For j As Integer = 1 To UBound(Notes)
            Notes(j).Selected = False
        Next
    End Sub

    'Undo appends before, Redo appends after.
    'After a sequence of Commands, 
    '   Undo will be the first one to execute, 
    '   Redo will be the last one to execute.
    'Remember to save the first Redo.

    'In case where undo is Nothing: Dont worry.
    'In case where redo is Nothing: 
    '   If only one redo is in a sequence, put Nothing.
    '   If several redo are in a sequence, 
    '       Create Void first. 
    '       Record its reference into a seperate copy. (xBaseRedo = xRedo)
    '       Use this xRedo as the BaseRedo.
    '       When calling AddUndo subroutine, use xBaseRedo.Next as cRedo.

    'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
    'Dim xRedo As UndoRedo.LinkedURCmd = Nothing
    '... 'Command.RedoRemoveNote(K(i), True, xUndo, xRedo)
    'AddUndo(xUndo, xRedo)

    'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
    'Dim xRedo As New UndoRedo.Void
    'Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
    '... 'Command.RedoRemoveNote(K(i), True, xUndo, xRedo)
    'AddUndo(xUndo, xBaseRedo.Next)


    Private Sub TBOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBVOptions.Click, mnVOptions.Click

        Dim xDiag As New OpVisual(vo, Columns.column, LWAV.Font)
        xDiag.ShowDialog(Me)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub AddToPOWAV(ByVal xPath() As String)
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

        For i As Integer = 0 To UBound(xPath)
            BmsWAV(xIndices(i) + 1) = GetFileName(xPath(i))
            LWAV.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": " & GetFileName(xPath(i))
        Next

        LWAV.SelectedIndices.Clear()
        For i As Integer = 0 To IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath))
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedAudioExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPOWAV(xPath)
    End Sub

    Private Sub POWAV_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedAudioExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.Resize
        LWAV.Height = sender.Height - 25
    End Sub

    Private Sub AddToPOBMP(ByVal xPath() As String)
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
        For i As Integer = 0 To UBound(xPath)
            'If j > UBound(xIndices) Then Exit For
            'hBMP(xIndices(j) + 1) = GetFileName(xPath(i))
            'LBMP.Items.Item(xIndices(j)) = C10to36(xIndices(j) + 1) & ": " & GetFileName(xPath(i))
            BmsBMP(xIndices(i) + 1) = GetFileName(xPath(i))
            LBMP.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": " & GetFileName(xPath(i))
            'j += 1
        Next

        LBMP.SelectedIndices.Clear()
        For i As Integer = 0 To IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath))
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POBMP.DragDrop
        ReDim DragDropFilename(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedImageExtension)
        Array.Sort(xPath)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPOBMP(xPath)
    End Sub

    Private Sub POBMP_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POBMP.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DragDropFilename = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedImageExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBMP.DragLeave
        ReDim DragDropFilename(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POBMP_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBMP.Resize
        LBMP.Height = sender.Height - 25
    End Sub
    Private Sub POBeat_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POBeat.Resize
        LBeat.Height = POBeat.Height - 25
    End Sub
    Private Sub POExpansion_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POExpansion.Resize
        TExpansion.Height = POExpansion.Height - 2
    End Sub

    Private Sub mn_DropDownClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.ForeColor = Color.White
    End Sub
    Private Sub mn_DropDownOpened(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.ForeColor = Color.Black
    End Sub
    Private Sub mn_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.Pressed Then Return
        sender.ForeColor = Color.Black
    End Sub
    Private Sub mn_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.Pressed Then Return
        sender.ForeColor = Color.White
    End Sub

    Private Sub TBPOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPOptions.Click, mnPOptions.Click
        Dim xDOp As New OpPlayer(CurrentPlayer)
        xDOp.ShowDialog(Me)
    End Sub

    Private Sub THGenre_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    THGenre.TextChanged, THTitle.TextChanged, THArtist.TextChanged, THPlayLevel.TextChanged, CHRank.SelectedIndexChanged, TExpansion.TextChanged,
    THSubTitle.TextChanged, THSubArtist.TextChanged, THStageFile.TextChanged, THBanner.TextChanged, THBackBMP.TextChanged,
    CHDifficulty.SelectedIndexChanged, THExRank.TextChanged, THTotal.TextChanged, THComment.TextChanged
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub CHLnObj_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHLnObj.SelectedIndexChanged
        If IsSaved Then SetIsSaved(False)
        LnObj = CHLnObj.SelectedIndex
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub ConvertBMSE2NT()
        ClearSelectionArray()
        SortByVPositionInsertion()

        For i2 As Integer = 0 To UBound(Notes)
            Notes(i2).Length = 0.0#
        Next

        Dim i As Integer = 1
        Dim j As Integer = 0
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

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub ConvertNT2BMSE()
        ClearSelectionArray()
        Dim xK(0) As Note
        xK(0) = Notes(0)

        For i As Integer = 1 To UBound(Notes)
            ReDim Preserve xK(UBound(xK) + 1)
            With xK(UBound(xK))
                .ColumnIndex = Notes(i).ColumnIndex
                .LongNote = Notes(i).Length > 0
                .Landmine = Notes(i).Landmine
                .Value = Notes(i).Value
                .VPosition = Notes(i).VPosition
                .Selected = Notes(i).Selected
                .Hidden = Notes(i).Hidden
            End With

            If Notes(i).Length > 0 Then
                ReDim Preserve xK(UBound(xK) + 1)
                With xK(UBound(xK))
                    .ColumnIndex = Notes(i).ColumnIndex
                    .LongNote = True
                    .Landmine = False
                    .Value = Notes(i).Value
                    .VPosition = Notes(i).VPosition + Notes(i).Length
                    .Selected = Notes(i).Selected
                    .Hidden = Notes(i).Hidden
                End With
            End If
        Next

        Notes = xK

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub TBWavIncrease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBWavIncrease.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        TBWavIncrease.Checked = Not sender.Checked
        Command.RedoWavIncrease(TBWavIncrease.Checked, xUndo, xRedo)
        AddUndoChain(xUndo, xBaseRedo.Next)
    End Sub

    Private Sub TBNTInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBNTInput.Click, mnNTInput.Click
        'Dim xUndo As String = "NT_" & CInt(NTInput) & "_0" & vbCrLf & "KZ" & vbCrLf & sCmdKsAll(False)
        'Dim xRedo As String = "NT_" & CInt(Not NTInput) & "_1"
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Command.RedoRemoveNoteAll(Notes, xUndo, xRedo)

        NTInput = sender.Checked

        TBNTInput.Checked = NTInput
        mnNTInput.Checked = NTInput
        POBLong.Enabled = Not NTInput
        POBLongShort.Enabled = Not NTInput

        State.NT.IsAdjustingNoteLength = False
        State.NT.IsAdjustingUpperEnd = False

        Command.RedoNT(NTInput, False, xUndo, xRedo)
        If NTInput Then
            ConvertBMSE2NT()
        Else
            ConvertNT2BMSE()
        End If
        Command.RedoAddNoteAll(Notes, xUndo, xRedo)

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
    End Sub

    Private Sub THBPM_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles THBPM.ValueChanged
        If Notes IsNot Nothing Then Notes(0).Value = THBPM.Value * 10000 : RefreshPanelAll()
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub TWPosition_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPosition.ValueChanged
        wPosition = TWPosition.Value
        TWPosition2.Value = IIf(wPosition > TWPosition2.Maximum, TWPosition2.Maximum, wPosition)
        RefreshPanelAll()
    End Sub

    Private Sub TWLeft_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWLeft.ValueChanged
        wLeft = TWLeft.Value
        TWLeft2.Value = IIf(wLeft > TWLeft2.Maximum, TWLeft2.Maximum, wLeft)
        RefreshPanelAll()
    End Sub

    Private Sub TWWidth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWWidth.ValueChanged
        wWidth = TWWidth.Value
        TWWidth2.Value = IIf(wWidth > TWWidth2.Maximum, TWWidth2.Maximum, wWidth)
        RefreshPanelAll()
    End Sub

    Private Sub TWPrecision_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPrecision.ValueChanged
        wPrecision = TWPrecision.Value
        TWPrecision2.Value = IIf(wPrecision > TWPrecision2.Maximum, TWPrecision2.Maximum, wPrecision)
        RefreshPanelAll()
    End Sub

    Private Sub TWTransparency_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWTransparency.ValueChanged
        TWTransparency2.Value = TWTransparency.Value
        vo.pBGMWav.Color = Color.FromArgb(TWTransparency.Value, vo.pBGMWav.Color)
        RefreshPanelAll()
    End Sub

    Private Sub TWSaturation_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWSaturation.ValueChanged
        Dim xColor As Color = vo.pBGMWav.Color
        TWSaturation2.Value = TWSaturation.Value
        vo.pBGMWav.Color = HSL2RGB(xColor.GetHue, TWSaturation.Value, xColor.GetBrightness * 1000, xColor.A)
        RefreshPanelAll()
    End Sub

    Private Sub TWPosition2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPosition2.Scroll
        TWPosition.Value = TWPosition2.Value
    End Sub

    Private Sub TWLeft2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWLeft2.Scroll
        TWLeft.Value = TWLeft2.Value
    End Sub

    Private Sub TWWidth2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWWidth2.Scroll
        TWWidth.Value = TWWidth2.Value
    End Sub

    Private Sub TWPrecision2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWPrecision2.Scroll
        TWPrecision.Value = TWPrecision2.Value
    End Sub

    Private Sub TWTransparency2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWTransparency2.Scroll
        TWTransparency.Value = TWTransparency2.Value
    End Sub

    Private Sub TWSaturation2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TWSaturation2.Scroll
        TWSaturation.Value = TWSaturation2.Value
    End Sub

    Private Sub TBLangDef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBLangDef.Click
        DispLang = ""
        MsgBox(Strings.Messages.PreferencePostpone, MsgBoxStyle.Information)
    End Sub

    Private Sub TBLangRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBLangRefresh.Click
        For i As Integer = cmnLanguage.Items.Count - 1 To 3 Step -1
            Try
                cmnLanguage.Items.RemoveAt(i)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles("*.Lang.xml")

        For Each xStr As FileInfo In xFileNames
            LoadLocaleXML(xStr)
        Next
    End Sub


    Private Sub UpdateColumnsX()
        Columns.RecalculatePositions()

        HSL.Maximum = Columns.GetRightBoundry()
        HS.Maximum = Columns.GetRightBoundry()
        HSR.Maximum = Columns.GetRightBoundry()
    End Sub

    Private Sub CHPlayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHPlayer.SelectedIndexChanged
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0

        Grid.IPlayer = CHPlayer.SelectedIndex
        Dim xGP2 As Boolean = Grid.IPlayer <> 0
        Columns.SetP2SideVisible(xGP2)

        UpdateNoteSelectionStatus()

        UpdateColumnsX()

        If IsInitializing Then Exit Sub
        RefreshPanelAll()
    End Sub

    Private Sub UpdateNoteSelectionStatus()
        For i As Integer = 1 To UBound(Notes)
            Notes(i).Selected = Notes(i).Selected And Columns.nEnabled(Notes(i).ColumnIndex)
        Next
    End Sub

    Private Sub CGB_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGB.ValueChanged
        Columns.ColumnCount = ColumnType.BGM + CGB.Value - 1
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub TBGOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBGOptions.Click, mnGOptions.Click
        Dim localeIndex As Integer
        Select Case EncodingToString(TextEncoding).ToUpper ' az: wow seriously? is there really no better way? 
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

        Dim xDiag As New OpGeneral(Grid.WheelScroll, Grid.PageUpDnScroll, MiddleButtonMoveMethod, localeIndex, 192.0R / BMSGridLimit,
            AutoSaveInterval, BeepWhileSaved, BPMx1296, STOPx1296,
            AutoFocusPanelOnMouseEnter, FirstClickDisabled, ClickStopPreview)

        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With xDiag
                Grid.WheelScroll = .zWheel
                Grid.PageUpDnScroll = .zPgUpDn
                TextEncoding = .zEncoding
                'SortingMethod = .zSort
                MiddleButtonMoveMethod = .zMiddle
                AutoSaveInterval = .zAutoSave
                BMSGridLimit = 192.0R / .zGridPartition
                BeepWhileSaved = .cBeep.Checked
                BPMx1296 = .cBpm1296.Checked
                STOPx1296 = .cStop1296.Checked
                AutoFocusPanelOnMouseEnter = .cMEnterFocus.Checked
                FirstClickDisabled = .cMClickFocus.Checked
                ClickStopPreview = .cMStopPreview.Checked
            End With
            If AutoSaveInterval Then AutoSaveTimer.Interval = AutoSaveInterval
            AutoSaveTimer.Enabled = AutoSaveInterval
        End If
    End Sub

    Private Sub POBLong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLong.Click
        If NTInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i As Integer = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, True, xUndo, xRedo)
            Notes(i).LongNote = True
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBShort.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not NTInput Then
            For i As Integer = 1 To UBound(Notes)
                If Not Notes(i).Selected Then Continue For

                Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, 0, xUndo, xRedo)
                Notes(i).LongNote = False
            Next

        Else
            For i As Integer = 1 To UBound(Notes)
                If Not Notes(i).Selected Then Continue For

                Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, 0, xUndo, xRedo)
                Notes(i).Length = 0
            Next
        End If

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormalLong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLongShort.Click
        If NTInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i As Integer = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, Not Notes(i).LongNote, xUndo, xRedo)
            Notes(i).LongNote = Not Notes(i).LongNote
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHidden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHidden.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i As Integer = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            Command.RedoHiddenNoteModify(Notes(i), True, True, xUndo, xRedo)
            Notes(i).Hidden = True
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i As Integer = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            Command.RedoHiddenNoteModify(Notes(i), False, True, xUndo, xRedo)
            Notes(i).Hidden = False
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHiddenVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHiddenVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For i As Integer = 1 To UBound(Notes)
            If Not Notes(i).Selected Then Continue For

            Command.RedoHiddenNoteModify(Notes(i), Not Notes(i).Hidden, True, xUndo, xRedo)
            Notes(i).Hidden = Not Notes(i).Hidden
        Next
        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBModify_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles POBModify.Click
        Dim xNum As Boolean = False
        Dim xLbl As Boolean = False
        Dim i As Integer

        For i = 1 To UBound(Notes)
            If Notes(i).Selected AndAlso Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then xNum = True : Exit For
        Next
        For i = 1 To UBound(Notes)
            If Notes(i).Selected AndAlso Not Columns.IsColumnNumeric(Notes(i).ColumnIndex) Then xLbl = True : Exit For
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

                    Command.RedoRelabelNote(Notes(i), xD1, xUndo, xRedo)
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

                Command.RedoRelabelNote(Notes(i), xVal, xUndo, xRedo)
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

    Private Sub TBMyO2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBMyO2.Click, mnMyO2.Click
        Dim xDiag As New dgMyO2
        xDiag.Show()
    End Sub


    Private Sub TBFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBFind.Click, mnFind.Click
        Dim xDiag As New diagFind(Columns.ColumnCount, Strings.Messages.Err, Strings.Messages.InvalidLabel)
        xDiag.Show()
    End Sub

    Private Function fdrCheck(ByVal xNote As Note) As Boolean
        Return xNote.VPosition >= MeasureBottom(fdriMesL) And xNote.VPosition < MeasureBottom(fdriMesU) + MeasureLength(fdriMesU) AndAlso
               IIf(Columns.IsColumnNumeric(xNote.ColumnIndex),
                   xNote.Value >= fdriValL And xNote.Value <= fdriValU,
                   xNote.Value >= fdriLblL And xNote.Value <= fdriLblU) AndAlso
               Array.IndexOf(fdriCol, xNote.ColumnIndex) <> -1
    End Function

    Private Function fdrRangeS(ByVal xbLim1 As Boolean, ByVal xbLim2 As Boolean, ByVal xVal As Boolean) As Boolean
        Return (Not xbLim1 And xbLim2 And xVal) Or (xbLim1 And Not xbLim2 And Not xVal) Or (xbLim1 And xbLim2)
    End Function

    Public Sub fdrSelect(ByVal iRange As Integer,
                         ByVal xMesL As Integer, ByVal xMesU As Integer,
                         ByVal xLblL As String, ByVal xLblU As String,
                         ByVal xValL As Integer, ByVal xValU As Integer,
                         ByVal iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xSel(UBound(Notes)) As Boolean
        For i As Integer = 1 To UBound(Notes)
            xSel(i) = Notes(i).Selected
        Next

        'Main process
        For i As Integer = 1 To UBound(Notes)
            Dim bbba As Boolean = xbSel And xSel(i)
            Dim bbbb As Boolean = xbUnsel And Not xSel(i)
            Dim bbbc As Boolean = Columns.nEnabled(Notes(i).ColumnIndex)
            Dim bbbd As Boolean = fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote))
            Dim bbbe As Boolean = fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden)
            Dim bbbf As Boolean = fdrCheck(Notes(i))

            If ((xbSel And xSel(i)) Or (xbUnsel And Not xSel(i))) AndAlso
                    Columns.nEnabled(Notes(i).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                Notes(i).Selected = fdrCheck(Notes(i))
            End If
        Next

        RefreshPanelAll()
        Beep()
    End Sub

    Public Sub fdrUnselect(ByVal iRange As Integer,
                           ByVal xMesL As Integer, ByVal xMesU As Integer,
                           ByVal xLblL As String, ByVal xLblU As String,
                           ByVal xValL As Integer, ByVal xValU As Integer,
                           ByVal iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xSel(UBound(Notes)) As Boolean
        For i As Integer = 1 To UBound(Notes)
            xSel(i) = Notes(i).Selected
        Next

        'Main process
        For i As Integer = 1 To UBound(Notes)
            If ((xbSel And xSel(i)) Or (xbUnsel And Not xSel(i))) AndAlso
                        Columns.nEnabled(Notes(i).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                Notes(i).Selected = Not fdrCheck(Notes(i))
            End If
        Next

        RefreshPanelAll()
        Beep()
    End Sub

    Public Sub fdrDelete(ByVal iRange As Integer,
                         ByVal xMesL As Integer, ByVal xMesU As Integer,
                         ByVal xLblL As String, ByVal xLblU As String,
                         ByVal xValL As Integer, ByVal xValU As Integer,
                         ByVal iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'Main process
        Dim i As Integer = 1
        Do While i <= UBound(Notes)
            If ((xbSel And Notes(i).Selected) Or (xbUnsel And Not Notes(i).Selected)) AndAlso
                        fdrCheck(Notes(i)) AndAlso
                        Columns.nEnabled(Notes(i).ColumnIndex) AndAlso
                        fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And
                        fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                RedoRemoveNote(Notes(i), xUndo, xRedo)
                RemoveNote(i, False)
            Else
                i += 1
            End If
        Loop

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        CalculateTotalPlayableNotes()
        Beep()
    End Sub

    Public Sub fdrReplaceL(ByVal iRange As Integer,
                           ByVal xMesL As Integer, ByVal xMesU As Integer,
                           ByVal xLblL As String, ByVal xLblU As String,
                           ByVal xValL As Integer, ByVal xValU As Integer,
                           ByVal iCol() As Integer, ByVal xReplaceLbl As String)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xxLbl As Integer = C36to10(xReplaceLbl) * 10000

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'Main process
        For Each note In Notes
            If ((xbSel And note.Selected) Or (xbUnsel And Not note.Selected)) AndAlso
                    fdrCheck(note) AndAlso
                    Columns.nEnabled(note.ColumnIndex) And
                    Not Columns.IsColumnNumeric(note.ColumnIndex) AndAlso
                    fdrRangeS(xbShort, xbLong, IIf(NTInput, note.Length, note.LongNote)) And
                    fdrRangeS(xbVisible, xbHidden, note.Hidden) Then
                RedoRelabelNote(note, xxLbl, xUndo, xRedo)
                note.Value = xxLbl
            End If
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        Beep()
    End Sub

    Public Sub fdrReplaceV(ByVal iRange As Integer,
                           ByVal xMesL As Integer, ByVal xMesU As Integer,
                           ByVal xLblL As String, ByVal xLblU As String,
                           ByVal xValL As Integer, ByVal xValU As Integer,
                           ByVal iCol() As Integer, ByVal xReplaceVal As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL) * 10000
        fdriLblU = C36to10(xLblU) * 10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'Main process
        For Each note In Notes.Skip(1)
            If ((xbSel And note.Selected) Or (xbUnsel And Not note.Selected)) AndAlso
                    fdrCheck(note) AndAlso
                    Columns.nEnabled(note.ColumnIndex) And
                    Columns.IsColumnNumeric(note.ColumnIndex) AndAlso
                    fdrRangeS(xbShort, xbLong, IIf(NTInput, note.Length, note.LongNote)) And fdrRangeS(xbVisible, xbHidden, note.Hidden) Then
                RedoRelabelNote(note, xReplaceVal, xUndo, xRedo)
                note.Value = xReplaceVal
            End If
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        Beep()
    End Sub

    Private Sub MInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MInsert.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)
        Dim xVP As Double = MeasureBottom(xMeasure)

        If NTInput Then
            Dim i As Integer = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) >= 999 Then
                    Command.RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            Dim xdVP As Double
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP And Notes(i).VPosition + Notes(i).Length <= MeasureBottom(999) Then
                    Command.RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition + xMLength, xUndo, xRedo)
                    Notes(i).VPosition += xMLength

                ElseIf Notes(i).VPosition >= xVP Then
                    xdVP = MeasureBottom(999) - 1 - Notes(i).VPosition - Notes(i).Length
                    Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition + xMLength, Notes(i).Length + xdVP, xUndo, xRedo)
                    Notes(i).VPosition += xMLength
                    Notes(i).Length += xdVP

                ElseIf Notes(i).VPosition + Notes(i).Length >= xVP Then
                    xdVP = IIf(Notes(i).VPosition + Notes(i).Length > MeasureBottom(999) - 1, GetMaxVPosition() - 1 - Notes(i).VPosition - Notes(i).Length, xMLength)
                    Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, Notes(i).Length + xdVP, xUndo, xRedo)
                    Notes(i).Length += xdVP
                End If
            Next

        Else
            Dim i As Integer = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) >= 999 Then
                    Command.RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP Then
                    Command.RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition + xMLength, xUndo, xRedo)
                    Notes(i).VPosition += xMLength
                End If
            Next
        End If

        For i As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(i) = MeasureLength(i - 1)
        Next
        UpdateMeasureBottom()

        AddUndoChain(xUndo, xBaseRedo.Next)
        UpdatePairing()
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
    End Sub

    Private Sub MRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MRemove.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Dim xMeasure As Integer = MeasureAtDisplacement(menuVPosition)
        Dim xMLength As Double = MeasureLength(xMeasure)
        Dim xVP As Double = MeasureBottom(xMeasure)

        If NTInput Then
            Dim i As Integer = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) = xMeasure And MeasureAtDisplacement(Notes(i).VPosition + Notes(i).Length) = xMeasure Then
                    Command.RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            Dim xdVP As Double
            xVP = MeasureBottom(xMeasure)
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP + xMLength Then
                    Command.RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition - xMLength, xUndo, xRedo)
                    Notes(i).VPosition -= xMLength

                ElseIf Notes(i).VPosition >= xVP Then
                    xdVP = xMLength + xVP - Notes(i).VPosition
                    Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition + xdVP - xMLength, Notes(i).Length - xdVP, xUndo, xRedo)
                    Notes(i).VPosition += xdVP - xMLength
                    Notes(i).Length -= xdVP

                ElseIf Notes(i).VPosition + Notes(i).Length >= xVP Then
                    xdVP = IIf(Notes(i).VPosition + Notes(i).Length >= xVP + xMLength, xMLength, Notes(i).VPosition + Notes(i).Length - xVP + 1)
                    Command.RedoLongNoteModify(Notes(i), Notes(i).VPosition, Notes(i).Length - xdVP, xUndo, xRedo)
                    Notes(i).Length -= xdVP
                End If
            Next

        Else
            Dim i As Integer = 1
            Do While i <= UBound(Notes)
                If MeasureAtDisplacement(Notes(i).VPosition) = xMeasure Then
                    Command.RedoRemoveNote(Notes(i), xUndo, xRedo)
                    RemoveNote(i, False)
                Else
                    i += 1
                End If
            Loop

            xVP = MeasureBottom(xMeasure)
            For i = 1 To UBound(Notes)
                If Notes(i).VPosition >= xVP Then
                    Command.RedoMoveNote(Notes(i), Notes(i).ColumnIndex, Notes(i).VPosition - xMLength, xUndo, xRedo)
                    Notes(i).VPosition -= xMLength
                End If
            Next
        End If

        For i As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(i - 1) = MeasureLength(i)
        Next
        UpdateMeasureBottom()

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
    End Sub

    Private Sub TBThemeDef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeDef.Click
        Dim xTempFileName As String = My.Application.Info.DirectoryPath & "\____TempFile.Theme.xml"
        My.Computer.FileSystem.WriteAllText(xTempFileName, My.Resources.O2Mania_Theme, False, System.Text.Encoding.Unicode)
        LoadSettings(xTempFileName)
        File.Delete(xTempFileName)

        RefreshPanelAll()
    End Sub

    Private Sub TBThemeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeSave.Click
        Dim xDiag As New SaveFileDialog
        xDiag.Filter = Strings.FileType.THEME_XML & "|*.Theme.xml"
        xDiag.DefaultExt = "Theme.xml"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Data"
        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        Me.SaveSettings(xDiag.FileName, True)
        If BeepWhileSaved Then Beep()
        TBThemeRefresh_Click(TBThemeRefresh, New System.EventArgs)
    End Sub

    Private Sub TBThemeRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeRefresh.Click
        For i As Integer = cmnTheme.Items.Count - 1 To 5 Step -1
            Try
                cmnTheme.Items.RemoveAt(i)
            Catch ex As Exception
            End Try
        Next

        If Not Directory.Exists(My.Application.Info.DirectoryPath & "\Data") Then My.Computer.FileSystem.CreateDirectory(My.Application.Info.DirectoryPath & "\Data")
        Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(My.Application.Info.DirectoryPath & "\Data").GetFiles("*.Theme.xml")
        For Each xStr As FileInfo In xFileNames
            cmnTheme.Items.Add(xStr.Name, Nothing, AddressOf LoadTheme)
        Next
    End Sub

    Private Sub TBThemeLoadComptability_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBThemeLoadComptability.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType.TH & "|*.th"
        xDiag.DefaultExt = "th"
        xDiag.InitialDirectory = My.Application.Info.DirectoryPath
        If My.Computer.FileSystem.DirectoryExists(My.Application.Info.DirectoryPath & "\Theme") Then xDiag.InitialDirectory = My.Application.Info.DirectoryPath & "\Theme"
        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        Me.LoadThemeComptability(xDiag.FileName)
        RefreshPanelAll()
    End Sub

    ''' <summary>
    ''' Will return Double.PositiveInfinity if canceled.
    ''' </summary>
    Private Function InputBoxDouble(ByVal Prompt As String, ByVal LBound As Double, ByVal UBound As Double, Optional ByVal Title As String = "", Optional ByVal DefaultResponse As String = "") As Double
        Dim xStr As String = InputBox(Prompt, Title, DefaultResponse)
        If xStr = "" Then Return Double.PositiveInfinity

        InputBoxDouble = Val(xStr)
        If InputBoxDouble > UBound Then InputBoxDouble = UBound
        If InputBoxDouble < LBound Then InputBoxDouble = LBound
    End Function

    Private Sub FSSS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSS.Click
        Dim xMax As Double = IIf(State.TimeSelect.EndPointLength > 0, GetMaxVPosition() - State.TimeSelect.EndPointLength, GetMaxVPosition)
        Dim xMin As Double = IIf(State.TimeSelect.EndPointLength < 0, -State.TimeSelect.EndPointLength, 0)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , State.TimeSelect.StartPoint)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.StartPoint = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSL.Click
        Dim xMax As Double = GetMaxVPosition() - State.TimeSelect.StartPoint
        Dim xMin As Double = -State.TimeSelect.StartPoint
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , State.TimeSelect.EndPointLength)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.EndPointLength = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSH.Click
        Dim xMax As Double = IIf(State.TimeSelect.EndPointLength > 0, State.TimeSelect.EndPointLength, 0)
        Dim xMin As Double = IIf(State.TimeSelect.EndPointLength > 0, 0, -State.TimeSelect.EndPointLength)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , State.TimeSelect.HalfPointLength)
        If xDouble = Double.PositiveInfinity Then Return

        State.TimeSelect.HalfPointLength = xDouble
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BVCReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCReverse.Click
        State.TimeSelect.StartPoint += State.TimeSelect.EndPointLength
        State.TimeSelect.HalfPointLength -= State.TimeSelect.EndPointLength
        State.TimeSelect.EndPointLength *= -1
        State.TimeSelect.ValidateSelection(GetMaxVPosition())
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub AutoSaveTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoSaveTimer.Tick
        Dim xTime As Date = Now
        Dim xFileName As String
        With xTime
            xFileName = My.Application.Info.DirectoryPath & "\AutoSave_" &
                              .Year & "_" & .Month & "_" & .Day & "_" & .Hour & "_" & .Minute & "_" & .Second & "_" & .Millisecond & ".IBMSC"
        End With
        'My.Computer.FileSystem.WriteAllText(xFileName, SaveiBMSC, False, System.Text.Encoding.Unicode)
        SaveiBMSC(xFileName)

        On Error Resume Next
        If PreviousAutoSavedFileName <> "" Then IO.File.Delete(PreviousAutoSavedFileName)
        On Error GoTo 0

        PreviousAutoSavedFileName = xFileName
    End Sub

    Private Sub CWAVMultiSelect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CWAVMultiSelect.CheckedChanged
        WAVMultiSelect = CWAVMultiSelect.Checked
        LWAV.SelectionMode = IIf(WAVMultiSelect, SelectionMode.MultiExtended, SelectionMode.One)
        LBMP.SelectionMode = IIf(WAVMultiSelect, SelectionMode.MultiExtended, SelectionMode.One)
    End Sub

    Private Sub CWAVChangeLabel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CWAVChangeLabel.CheckedChanged
        WAVChangeLabel = CWAVChangeLabel.Checked
    End Sub

    Private Sub BWAVUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVUp.Click
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

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsWAV(i + 1)
                BmsWAV(i + 1) = BmsWAV(i)
                BmsWAV(i) = xStr

                LWAV.Items.Item(i) = C10to36(i + 1) & ": " & BmsWAV(i + 1)
                LWAV.Items.Item(i - 1) = C10to36(i) & ": " & BmsWAV(i)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i)
                Dim xL2 As String = C10to36(i + 1)
                For j As Integer = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000

                    End If
                Next

1100:           xIndices(xIndex) += -1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BWAVDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVDown.Click
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

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsWAV(i + 1)
                BmsWAV(i + 1) = BmsWAV(i + 2)
                BmsWAV(i + 2) = xStr

                LWAV.Items.Item(i) = C10to36(i + 1) & ": " & BmsWAV(i + 1)
                LWAV.Items.Item(i + 1) = C10to36(i + 2) & ": " & BmsWAV(i + 2)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i + 2)
                Dim xL2 As String = C10to36(i + 1)
                For j As Integer = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 20000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 20000

                    End If
                Next

1100:           xIndices(xIndex) += 1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BWAVBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVBrowse.Click
        Dim xDWAV As New OpenFileDialog
        xDWAV.DefaultExt = "wav"
        xDWAV.Filter = Strings.FileType._wave & "|*.wav;*.ogg;*.mp3|" &
                       Strings.FileType.WAV & "|*.wav|" &
                       Strings.FileType.OGG & "|*.ogg|" &
                       Strings.FileType.MP3 & "|*.mp3|" &
                       Strings.FileType._all & "|*.*"
        xDWAV.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        xDWAV.Multiselect = WAVMultiSelect

        If xDWAV.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDWAV.FileName)

        AddToPOWAV(xDWAV.FileNames)
    End Sub

    Private Sub BWAVRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BWAVRemove.Click
        Dim xIndices(LWAV.SelectedIndices.Count - 1) As Integer
        LWAV.SelectedIndices.CopyTo(xIndices, 0)
        For i As Integer = 0 To UBound(xIndices)
            BmsWAV(xIndices(i) + 1) = ""
            LWAV.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": "
        Next

        LWAV.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(i))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBMPUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPUp.Click
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

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsBMP(i + 1)
                BmsBMP(i + 1) = BmsBMP(i)
                BmsBMP(i) = xStr

                LBMP.Items.Item(i) = C10to36(i + 1) & ": " & BmsBMP(i + 1)
                LBMP.Items.Item(i - 1) = C10to36(i) & ": " & BmsBMP(i)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i)
                Dim xL2 As String = C10to36(i + 1)
                For j As Integer = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000

                    End If
                Next

1100:           xIndices(xIndex) += -1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBMPDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPDown.Click
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

        Dim xStr As String = ""
        Dim xIndex As Integer = -1
        For i As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, i)
            If xIndex <> -1 Then
                xStr = BmsBMP(i + 1)
                BmsBMP(i + 1) = BmsBMP(i + 2)
                BmsBMP(i + 2) = xStr

                LBMP.Items.Item(i) = C10to36(i + 1) & ": " & BmsBMP(i + 1)
                LBMP.Items.Item(i + 1) = C10to36(i + 2) & ": " & BmsBMP(i + 2)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(i + 2)
                Dim xL2 As String = C10to36(i + 1)
                For j As Integer = 1 To UBound(Notes)
                    If Columns.IsColumnNumeric(Notes(j).ColumnIndex) Then Continue For

                    If C10to36(Notes(j).Value \ 10000) = xL1 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 10000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 10000

                    ElseIf C10to36(Notes(j).Value \ 10000) = xL2 Then
                        Command.RedoRelabelNote(Notes(j), i * 10000 + 20000, xUndo, xRedo)
                        Notes(j).Value = i * 10000 + 20000

                    End If
                Next

1100:           xIndices(xIndex) += 1
            End If
        Next

        LBMP.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBMPBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPBrowse.Click
        Dim xDBMP As New OpenFileDialog
        xDBMP.DefaultExt = "bmp"
        xDBMP.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.jpeg;.gif|" &
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
        xDBMP.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        xDBMP.Multiselect = WAVMultiSelect

        If xDBMP.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDBMP.FileName)

        AddToPOBMP(xDBMP.FileNames)
    End Sub

    Private Sub BBMPRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBMPRemove.Click
        Dim xIndices(LBMP.SelectedIndices.Count - 1) As Integer
        LBMP.SelectedIndices.CopyTo(xIndices, 0)
        For i As Integer = 0 To UBound(xIndices)
            BmsBMP(xIndices(i) + 1) = ""
            LBMP.Items.Item(xIndices(i)) = C10to36(xIndices(i) + 1) & ": "
        Next

        LBMP.SelectedIndices.Clear()
        For i As Integer = 0 To UBound(xIndices)
            LBMP.SelectedIndices.Add(xIndices(i))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnMain_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mnMain.MouseDown ', TBMain.MouseDown  ', pttl.MouseDown, pIsSaved.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()
            SendMessage(Me.Handle, &H112, &HF012, 0)
            If e.Clicks = 2 Then
                If Me.WindowState = FormWindowState.Maximized Then Me.WindowState = FormWindowState.Normal Else Me.WindowState = FormWindowState.Maximized
            End If
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            'mnSys.Show(sender, e.Location)
        End If
    End Sub

    Private Sub mnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSelectAll.Click
        If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub
        For Each note In Notes.Skip(1)
            note.Selected = Columns.nEnabled(note.ColumnIndex)
        Next
        If TBTimeSelect.Checked Then
            CalculateGreatestVPosition()
            State.TimeSelect.StartPoint = 0
            State.TimeSelect.EndPointLength = MeasureBottom(MeasureAtDisplacement(GreatestVPosition)) + MeasureLength(MeasureAtDisplacement(GreatestVPosition))
        End If
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnDelete.Click
        If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        RedoRemoveNoteSelected(Notes, xUndo, xRedo)
        RemoveNotes(True)

        AddUndoChain(xUndo, xBaseRedo.Next)
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("http://www.cs.mcgill.ca/~ryang6/iBMSC/")
    End Sub

    Private Sub mnUpdateC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Process.Start("http://bbs.rohome.net/thread-1074065-1-1.html")
    End Sub

    Private Sub mnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnQuit.Click
        Close()
    End Sub

    Private Sub mnSMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSMenu.CheckedChanged
        mnMain.Visible = mnSMenu.Checked
    End Sub
    Private Sub mnSTB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSTB.CheckedChanged
        TBMain.Visible = mnSTB.Checked
    End Sub
    Private Sub mnSOP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSOP.CheckedChanged
        POptions.Visible = mnSOP.Checked
    End Sub
    Private Sub mnSStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSStatus.CheckedChanged
        pStatus.Visible = mnSStatus.Checked
    End Sub
    Private Sub mnSLSplitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSLSplitter.CheckedChanged
        SpL.Visible = mnSLSplitter.Checked
    End Sub
    Private Sub mnSRSplitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnSRSplitter.CheckedChanged
        SpR.Visible = mnSRSplitter.Checked
    End Sub
    Private Sub CGShow_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShow.CheckedChanged
        Grid.ShowMainGrid = CGShow.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowS.CheckedChanged
        Grid.ShowSubGrid = CGShowS.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowBG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowBG.CheckedChanged
        Grid.ShowBackground = CGShowBG.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowM.CheckedChanged
        Grid.ShowMeasureNumber = CGShowM.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowV.CheckedChanged
        Grid.ShowVerticalLines = CGShowV.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowMB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowMB.CheckedChanged
        Grid.ShowMeasureBars = CGShowMB.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowC.CheckedChanged
        Grid.ShowColumnCaptions = CGShowC.Checked
        RefreshPanelAll()
    End Sub

    Private Sub FixedColumnVisibilityChanged(col As ColumnType(), isVisible As Boolean)
        For Each c In col
            Columns.GetColumn(c).IsVisible = isVisible
        Next


        If IsInitializing Then Exit Sub
        For Each note In Notes.Skip(1)
            note.Selected = note.Selected And Columns.nEnabled(note.ColumnIndex)
        Next

        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub CGBLP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBLP.CheckedChanged
        Grid.ShowBgaColumn = CGBLP.Checked

        FixedColumnVisibilityChanged({
            ColumnType.BGA,
            ColumnType.LAYER,
            ColumnType.POOR,
            ColumnType.S4
        }, Grid.ShowBgaColumn)

    End Sub

    Private Sub CGSCROLL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSCROLL.CheckedChanged
        Grid.ShowScrollColumn = CGSCROLL.Checked
        FixedColumnVisibilityChanged({ColumnType.SCROLLS}, Grid.ShowScrollColumn)
    End Sub

    Private Sub CGSTOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSTOP.CheckedChanged
        Grid.ShowStopColumn = CGSTOP.Checked
        FixedColumnVisibilityChanged({ColumnType.STOPS}, Grid.ShowStopColumn)
    End Sub

    Private Sub CGBPM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBPM.CheckedChanged
        Grid.ShowBpmColumn = CGBPM.Checked
        FixedColumnVisibilityChanged({ColumnType.BPM}, Grid.ShowBpmColumn)
    End Sub

    Private Sub CGDisableVertical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDisableVertical.CheckedChanged
        DisableVerticalMove = CGDisableVertical.Checked
    End Sub

    Private Sub CBeatPreserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBeatPreserve.Click, CBeatMeasure.Click, CBeatCut.Click, CBeatScale.Click
        'If Not sender.Checked Then Exit Sub
        Dim xBeatList() As RadioButton = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
        BeatChangeMode = Array.IndexOf(Of RadioButton)(xBeatList, sender)
    End Sub


    Private Sub tBeatValue_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tBeatValue.LostFocus
        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then tBeatValue.BackColor = Color.FromArgb(&HFFFFC0C0) Else tBeatValue.BackColor = Nothing

            tBeatValue.Text = a
        End If
    End Sub



    Private Sub ApplyBeat(ByVal xRatio As Double, ByVal xDisplay As String)
        SortByVPositionInsertion()

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Me.RedoChangeMeasureLengthSelected(192 * xRatio, xUndo, xRedo)

        Dim xIndices(LBeat.SelectedIndices.Count - 1) As Integer
        LBeat.SelectedIndices.CopyTo(xIndices, 0)


        For Each i As Integer In xIndices
            Dim dLength As Double = xRatio * 192.0R - MeasureLength(i)
            Dim dRatio As Double = xRatio * 192.0R / MeasureLength(i)

            Dim xBottom As Double = 0
            For j As Integer = 0 To i - 1
                xBottom += MeasureLength(j)
            Next
            Dim xUpBefore As Double = xBottom + MeasureLength(i)
            Dim xUpAfter As Double = xUpBefore + dLength

            Select Case BeatChangeMode
                Case 1
case2:              Dim xI0 As Integer

                    If NTInput Then
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                            If Notes(xI0).VPosition + Notes(xI0).Length >= xUpBefore Then
                                Command.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, xUndo, xRedo)
                                Notes(xI0).Length += dLength
                            End If
                        Next
                    Else
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                        Next
                    End If

                    For xI9 As Integer = xI0 To UBound(Notes)
                        Command.RedoLongNoteModify(Notes(xI9), Notes(xI9).VPosition + dLength, Notes(xI9).Length, xUndo, xRedo)
                        Notes(xI9).VPosition += dLength
                    Next

                Case 2
                    If dLength < 0 Then
                        If NTInput Then
                            Dim xI0 As Integer = 1
                            Dim xU As Integer = UBound(Notes)
                            Do While xI0 <= xU
                                If Notes(xI0).VPosition < xUpAfter Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length >= xUpAfter And Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Dim nLen As Double = xUpAfter - Notes(xI0).VPosition - 1.0R
                                        Command.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                    End If
                                ElseIf Notes(xI0).VPosition < xUpBefore Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Command.RedoRemoveNote(Notes(xI0), xUndo, xRedo)
                                        RemoveNote(xI0)
                                        xI0 -= 1
                                        xU -= 1
                                    Else
                                        Dim nLen As Double = Notes(xI0).Length - xUpBefore + Notes(xI0).VPosition
                                        Command.RedoLongNoteModify(Notes(xI0), xUpBefore, nLen, xUndo, xRedo)
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
                                Command.RedoRemoveNote(Notes(xI8), xUndo, xRedo)
                            Next
                            For xI8 As Integer = xI9 To UBound(Notes)
                                Notes(xI8 - xI9 + xI0) = Notes(xI8)
                            Next
                            ReDim Preserve Notes(UBound(Notes) - xI9 + xI0)
                        End If
                    End If

                    GoTo case2

                Case 3
                    If NTInput Then
                        For xI0 As Integer = 1 To UBound(Notes)
                            If Notes(xI0).VPosition < xBottom Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Command.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, xUndo, xRedo)
                                    Notes(xI0).Length += dLength
                                ElseIf Notes(xI0).VPosition + Notes(xI0).Length > xBottom Then
                                    Dim nLen As Double = (Notes(xI0).Length + Notes(xI0).VPosition - xBottom) * dRatio + xBottom - Notes(xI0).VPosition
                                    Command.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                End If
                            ElseIf Notes(xI0).VPosition < xUpBefore Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Dim nLen As Double = (xUpBefore - Notes(xI0).VPosition) * dRatio + Notes(xI0).VPosition + Notes(xI0).Length - xUpBefore
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Command.RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                Else
                                    Dim nLen As Double = Notes(xI0).Length * dRatio
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Command.RedoLongNoteModify(Notes(xI0), nVPos, nLen, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                End If
                            Else
                                Command.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition + dLength, Notes(xI0).Length, xUndo, xRedo)
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
                            Dim nVP As Double = (Notes(xI8).VPosition - xBottom) * dRatio + xBottom
                            Command.RedoLongNoteModify(Notes(xI0), nVP, Notes(xI0).Length, xUndo, xRedo)
                            Notes(xI8).VPosition = nVP
                        Next

                        'GoTo case2

                        For xI8 As Integer = xI9 To UBound(Notes)
                            Command.RedoLongNoteModify(Notes(xI8), Notes(xI8).VPosition + dLength, Notes(xI8).Length, xUndo, xRedo)
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
        For i As Integer = 0 To UBound(xIndices)
            LBeat.SelectedIndices.Add(xIndices(i))
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BBeatApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBeatApply.Click
        Dim xxD As Integer = nBeatD.Value
        Dim xxN As Integer = nBeatN.Value
        Dim xxRatio As Double = xxN / xxD

        ApplyBeat(xxRatio, xxRatio & " ( " & xxN & " / " & xxD & " ) ")
    End Sub

    Private Sub BBeatApplyV_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BBeatApplyV.Click
        Dim a As Double
        If Double.TryParse(tBeatValue.Text, a) Then
            If a <= 0.0# Or a >= 1000.0# Then System.Media.SystemSounds.Hand.Play() : Exit Sub

            Dim xxD As Long = GetDenominator(a)

            ApplyBeat(a, a & IIf(xxD > 10000, "", " ( " & CLng(a * xxD) & " / " & xxD & " ) "))
        End If
    End Sub


    Private Sub BHStageFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BHStageFile.Click, BHBanner.Click, BHBackBMP.Click
        Dim xDiag As New OpenFileDialog
        xDiag.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpeg;*.jpg;*.gif|" &
                       Strings.FileType._all & "|*.*"
        xDiag.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        xDiag.DefaultExt = "png"

        If xDiag.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub
        InitPath = ExcludeFileName(xDiag.FileName)

        If [Object].ReferenceEquals(sender, BHStageFile) Then
            THStageFile.Text = GetFileName(xDiag.FileName)
        ElseIf [Object].ReferenceEquals(sender, BHBanner) Then
            THBanner.Text = GetFileName(xDiag.FileName)
        ElseIf [Object].ReferenceEquals(sender, BHBackBMP) Then
            THBackBMP.Text = GetFileName(xDiag.FileName)
        End If
    End Sub

    Private Sub Switches_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    POHeaderSwitch.CheckedChanged,
    POGridSwitch.CheckedChanged,
    POWaveFormSwitch.CheckedChanged,
    POWAVSwitch.CheckedChanged,
    POBMPSwitch.CheckedChanged,
    POBeatSwitch.CheckedChanged,
    POExpansionSwitch.CheckedChanged

        Try
            Dim Source As CheckBox = CType(sender, CheckBox)
            Dim Target As Panel = Nothing

            If ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf ReferenceEquals(sender, POHeaderSwitch) Then : Target = POHeaderInner
            ElseIf ReferenceEquals(sender, POGridSwitch) Then : Target = POGridInner
            ElseIf ReferenceEquals(sender, POWaveFormSwitch) Then : Target = POWaveFormInner
            ElseIf ReferenceEquals(sender, POWAVSwitch) Then : Target = POWAVInner
            ElseIf ReferenceEquals(sender, POBMPSwitch) Then : Target = POBMPInner
            ElseIf ReferenceEquals(sender, POBeatSwitch) Then : Target = POBeatInner
            ElseIf ReferenceEquals(sender, POExpansionSwitch) Then : Target = POExpansionInner
            End If

            If Source.Checked Then
                Target.Visible = True
            Else
                Target.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Friend Sub SelectWavFromNote(note As Note)
        LWAV.SelectedIndices.Clear()
        LWAV.SelectedIndex = C36to10(C10to36(note.Value \ 10000)) - 1
        ValidateWavListView()
    End Sub

    Private Sub Expanders_CheckChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _
    POHeaderExpander.CheckedChanged,
    POGridExpander.CheckedChanged,
    POWaveFormExpander.CheckedChanged,
    POWAVExpander.CheckedChanged,
    POBeatExpander.CheckedChanged

        Try
            Dim Source As CheckBox = CType(sender, CheckBox)
            Dim Target As Panel = Nothing
            'Dim TargetParent As Panel = Nothing

            If Object.ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf Object.ReferenceEquals(sender, POHeaderExpander) Then : Target = POHeaderPart2 ' : TargetParent = POHeaderInner
            ElseIf Object.ReferenceEquals(sender, POGridExpander) Then : Target = POGridPart2 ' : TargetParent = POGridInner
            ElseIf Object.ReferenceEquals(sender, POWaveFormExpander) Then : Target = POWaveFormPart2 ' : TargetParent = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POWAVExpander) Then : Target = POWAVPart2 ' : TargetParent = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POBeatExpander) Then : Target = POBeatPart2 ' : TargetParent = POWaveFormInner
            End If

            If Source.Checked Then
                Target.Visible = True
                'Source.Image = My.Resources.Collapse
            Else
                Target.Visible = False
                'Source.Image = My.Resources.Expand
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub VerticalResizer_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseDown, POBMPResizer.MouseDown, POBeatResizer.MouseDown, POExpansionResizer.MouseDown
        tempResize = e.Y
    End Sub

    Private Sub HorizontalResizer_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POptionsResizer.MouseDown, SpL.MouseDown, SpR.MouseDown
        tempResize = e.X
    End Sub

    Private Sub POResizer_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseMove, POBMPResizer.MouseMove, POBeatResizer.MouseMove, POExpansionResizer.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.Y = tempResize Then Exit Sub

        Try
            Dim Source As Button = CType(sender, Button)
            Dim Target As Panel = Source.Parent

            Dim xHeight As Integer = Target.Height + e.Y - tempResize
            If xHeight < 10 Then xHeight = 10
            Target.Height = xHeight

            Target.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub POptionsResizer_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POptionsResizer.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = POptionsScroll.Width - e.X + tempResize
            If xWidth < 25 Then xWidth = 25
            POptionsScroll.Width = xWidth

            Me.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SpR_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SpR.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = PMainR.Width - e.X + tempResize
            If xWidth < 0 Then xWidth = 0
            PMainR.Width = xWidth

            Me.ToolStripContainer1.Refresh()
            Application.DoEvents()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SpL_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SpL.MouseMove
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.X = tempResize Then Exit Sub

        Try
            Dim xWidth As Integer = PMainL.Width + e.X - tempResize
            If xWidth < 0 Then xWidth = 0
            PMainL.Width = xWidth

            Me.ToolStripContainer1.Refresh()
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

            FocusedPanel.VerticalScroll = -MeasureBottom(i)
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
        If IsInitializing Then Exit Sub
        PMainInL.Refresh()
        PMainIn.Refresh()
        PMainInR.Refresh()
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

    Private Sub PMain_Paint(sender As Object, e As PaintEventArgs)
        ' do nothing, on purpose.
    End Sub
End Class
