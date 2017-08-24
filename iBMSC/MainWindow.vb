Imports iBMSC.Editor


Public Class MainWindow


    'Public Structure MARGINS
    '    Public Left As Integer
    '    Public Right As Integer
    '    Public Top As Integer
    '    Public Bottom As Integer
    'End Structure

    '<System.Runtime.InteropServices.DllImport("dwmapi.dll")> _
    'Public Shared Function DwmIsCompositionEnabled(ByRef en As Integer) As Integer
    'End Function
    '<System.Runtime.InteropServices.DllImport("dwmapi.dll")> _
    'Public Shared Function DwmExtendFrameIntoClientArea(ByVal hwnd As IntPtr, ByRef margin As MARGINS) As Integer
    'End Function
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Public Declare Function ReleaseCapture Lib "user32.dll" Alias "ReleaseCapture" () As Integer

    'Private Declare Auto Function GetWindowLong Lib "user32" (ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    'Private Declare Auto Function SetWindowLong Lib "user32" (ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    'Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    '<DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    'Private Shared Function SetWindowText(ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean
    'End Function

    'Private Const GWL_STYLE As Integer = -16
    'Private Const WS_CAPTION As Integer = &HC00000
    'Private Const SWP_NOSIZE As Integer = &H1
    'Private Const SWP_NOMOVE As Integer = &H2
    'Private Const SWP_NOZORDER As Integer = &H4
    'Private Const SWP_NOACTIVATE As Integer = &H10
    'Private Const SWP_FRAMECHANGED As Integer = &H20
    'Private Const SWP_REFRESH As Integer = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_FRAMECHANGED

    Public Const niMeasure As Integer = 0
    Public Const niBPM As Integer = 1
    Public Const niSTOP As Integer = 2
    Public Const niS1 As Integer = 3

    Public Const niA1 As Integer = 4
    Public Const niA2 As Integer = 5
    Public Const niA3 As Integer = 6
    Public Const niA4 As Integer = 7
    Public Const niA5 As Integer = 8
    Public Const niA6 As Integer = 9
    Public Const niA7 As Integer = 10
    Public Const niA8 As Integer = 11
    Public Const niS2 As Integer = 12

    Public Const niD1 As Integer = 13
    Public Const niD2 As Integer = 14
    Public Const niD3 As Integer = 15
    Public Const niD4 As Integer = 16
    Public Const niD5 As Integer = 17
    Public Const niD6 As Integer = 18
    Public Const niD7 As Integer = 19
    Public Const niD8 As Integer = 20
    Public Const niS3 As Integer = 21

    Public Const niBGA As Integer = 22
    Public Const niLAYER As Integer = 23
    Public Const niPOOR As Integer = 24
    Public Const niS4 As Integer = 25
    Public Const niB As Integer = 26

    Public Const idflBPM As Integer = 5
    Public Const idLong As Integer = 40
    Public Const idHidden As Integer = 20

    Public column() As Column = {New Column(0, 50, "Measure", False, True, True, 0, 0, &HFF00FFFF, 0, &HFF00FFFF, 0),
                              New Column(50, 60, "BPM", True, True, True, 3, 0, &HFFFF0000, 0, &HFFFF0000, 0),
                              New Column(110, 50, "STOP", True, True, True, 9, 0, &HFFFF0000, 0, &HFFFF0000, 0),
                              New Column(110, 5, "", False, False, True, 0, 0, 0, 0, 0, 0),
                              New Column(115, 42, "A1", True, False, True, 16, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(157, 30, "A2", True, False, True, 11, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(187, 42, "A3", True, False, True, 12, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(229, 45, "A4", True, False, True, 13, &HFFFFC862, &HFF000000, &HFFF7C66A, &HFF000000, &H16F38B0C),
                              New Column(274, 42, "A5", True, False, True, 14, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(316, 30, "A6", True, False, True, 15, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(346, 42, "A7", True, False, True, 18, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(388, 40, "A8", True, False, True, 19, &HFF808080, &HFF000000, &HFF909090, &HFF000000, 0),
                              New Column(388, 5, "", False, False, True, 0, 0, 0, 0, 0, 0),
                              New Column(393, 42, "D1", True, False, False, 21, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(393, 30, "D2", True, False, False, 22, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(393, 42, "D3", True, False, False, 23, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(393, 45, "D4", True, False, False, 24, &HFFFFC862, &HFF000000, &HFFF7C66A, &HFF000000, &H16F38B0C),
                              New Column(393, 42, "D5", True, False, False, 25, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(393, 30, "D6", True, False, False, 28, &HFF62B0FF, &HFF000000, &HFF6AB0F7, &HFF000000, &H140033FF),
                              New Column(393, 42, "D7", True, False, False, 29, &HFFB0B0B0, &HFF000000, &HFFC0C0C0, &HFF000000, &H14FFFFFF),
                              New Column(393, 40, "D8", True, False, False, 26, &HFF808080, &HFF000000, &HFF909090, &HFF000000, 0),
                              New Column(393, 5, "", False, False, False, 0, 0, 0, 0, 0, 0),
                              New Column(393, 40, "BGA", True, False, False, 4, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(393, 40, "LAYER", True, False, False, 7, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(393, 40, "POOR", True, False, False, 6, &HFF8CD78A, &HFF000000, &HFF90D38E, &HFF000000, 0),
                              New Column(393, 5, "", False, False, False, 0, 0, 0, 0, 0, 0),
                              New Column(393, 40, "B", True, False, True, 1, &HFFE18080, &HFF000000, &HFFDC8585, &HFF000000, 0)}

    Dim MeasureLength(999) As Double
    Dim MeasureBottom(999) As Double

    Public Function MeasureUpper(idx As Integer) As Double
        Return MeasureBottom(idx) + MeasureLength(idx)
    End Function


    Dim Notes() As Note = {New Note(niBPM, -1, 1200000, 0, False)}
    Dim mColumn(999) As Integer  '0 = no column, 1 = 1 column, etc.
    Dim GreatestVPosition As Double    '+ 2000 = -VS.Minimum

    Dim VSValue As Integer = 0 'Store value before ValueChange event
    Dim HSValue As Integer = 0 'Store value before ValueChange event

    'Dim SortingMethod As Integer = 1
    Dim MiddleButtonMoveMethod As Integer = 0
    Dim TextEncoding As System.Text.Encoding = System.Text.Encoding.UTF8
    Dim DispLang As String = ""     'Display Language
    Dim Recent() As String = {"", "", "", "", ""}
    Dim NTInput As Boolean = True
    Dim ShowFileName As Boolean = False

    Dim BeepWhileSaved As Boolean = True
    Dim BPMx1296 As Boolean = False
    Dim STOPx1296 As Boolean = False

    Dim IsInitializing As Boolean = True
    Dim FirstMouseEnter As Boolean = True

    Dim WAVMultiSelect As Boolean = True
    Dim WAVChangeLabel As Boolean = True
    Dim BeatChangeMode As Integer = 0

    'Dim FloatTolerance As Double = 0.0001R
    Dim BMSGridLimit As Double = 1.0R

    Dim LnObj As Integer = 0    '0 for none, 1-1295 for 01-ZZ

    'IO
    Dim FileName As String = "Untitled.bms"
    'Dim TitlePath As New Drawing2D.GraphicsPath
    Dim InitPath As String = ""
    Dim IsSaved As Boolean = True

    'Variables for Drag/Drop
    Dim DDFileName() As String = {}
    Dim SupportedFileExtension() As String = {".bms", ".bme", ".bml", ".pms", ".txt", ".sm", ".ibmsc"}
    Dim SupportedAudioExtension() As String = {".wav", ".mp3", ".ogg"}

    'Variables for theme
    'Dim SaveTheme As Boolean = True

    'Variables for undo/redo
    Dim sUndo(99) As UndoRedo.LinkedURCmd
    Dim sRedo(99) As UndoRedo.LinkedURCmd
    Dim sI As Integer = 0

    'Variables for select tool
    Dim DisableVerticalMove As Boolean = False
    Dim KMouseOver As Integer = -1   'Mouse is on which note (for drawing green outline)
    Dim LastMouseDownLocation As PointF = New Point(-1, -1)          'Mouse is clicked on which point (location for display) (for selection box)
    Dim pMouseMove As PointF = New Point(-1, -1)          'Mouse is moved to which point   (location for display) (for selection box)
    'Dim KMouseDown As Integer = -1   'Mouse is clicked on which note (for moving)
    Dim deltaVPosition As Double = 0   'difference between mouse and VPosition of K
    Dim bAdjustLength As Boolean     'If adjusting note length instead of moving it
    Dim bAdjustUpper As Boolean      'true = Adjusting upper end, false = adjusting lower end
    Dim bAdjustSingle As Boolean     'true if there is only one note to be adjusted
    Dim tempY As Integer
    Dim tempV As Integer
    Dim tempX As Integer
    Dim tempH As Integer
    Dim MiddleButtonLocation As New Point(0, 0)
    Dim MiddleButtonClicked As Boolean = False
    Dim MouseMoveStatus As Point = New Point(0, 0)  'mouse is moved to which point (For Status Panel)
    'Dim uCol As Integer         'temp variables for undo, original enabled columnindex
    'Dim uVPos As Double         'temp variables for undo, original vposition
    'Dim uPairWithI As Double    'temp variables for undo, original note length
    Dim uAdded As Boolean       'temp variables for undo, if undo command is added
    'Dim uNote As Note           'temp variables for undo, original note
    Dim SelectedNotes(-1) As Note        'temp notes for undo
    Dim ctrlPressed As Boolean = False          'Indicates if the CTRL key is pressed while mousedown
    Dim ctrlForDuplicate As Boolean = False     'Indicates if duplicate notes of select/unselect note

    'Variables for write tool
    Dim ShouldDrawTempNote As Boolean = False
    Dim SelectedColumn As Integer = -1
    Dim TempVPosition As Double = -1.0#
    Dim TempLength As Double = 0.0#

    'Variables for post effects tool
    Dim vSelStart As Double = 192.0#
    Dim vSelLength As Double = 0.0#
    Dim vSelHalf As Double = 0.0#
    Dim vSelMouseOverLine As Integer = 0  '0 = nothing, 1 = start, 2 = half, 3 = end
    Dim vSelAdjust As Boolean = False
    Dim vSelK() As Note = {}
    Dim vSelPStart As Double = 192.0#
    Dim vSelPLength As Double = 0.0#
    Dim vSelPHalf As Double = 0.0#

    'Variables for Full-Screen Mode
    Dim isFullScreen As Boolean = False
    Dim previousWindowState As FormWindowState = FormWindowState.Normal
    Dim previousWindowPosition As New Rectangle(0, 0, 0, 0)

    'Variables misc
    Dim menuVPosition As Double = 0.0#
    Dim tempResize As Integer = 0

    '----AutoSave Options
    Dim PreviousAutoSavedFileName As String = ""
    Dim AutoSaveInterval As Integer = 120000

    '----ErrorCheck Options
    Dim ErrorCheck As Boolean = True

    '----Header Options
    Dim hWAV(1295) As String
    Dim hBPM(1295) As Long   'x10000
    Dim hSTOP(1295) As Long

    '----Grid Options
    Dim gSnap As Boolean = True
    Dim gShowGrid As Boolean = True 'Grid
    Dim gShowSubGrid As Boolean = True 'Sub
    Dim gShowBG As Boolean = True 'BG Color
    Dim gShowMeasureNumber As Boolean = True 'Measure Label
    Dim gShowVerticalLine As Boolean = True 'Vertical
    Dim gShowMeasureBar As Boolean = True 'Measure Barline
    Dim gShowC As Boolean = True 'Column Caption
    Dim gDivide As Integer = 16
    Dim gSub As Integer = 4
    Dim gSlash As Integer = 192
    Dim gxHeight As Single = 1.0!
    Dim gxWidth As Single = 1.0!
    Dim gWheel As Integer = 96
    Dim gPgUpDn As Integer = 384

    Dim gDisplayBGAColumn As Boolean = False
    Dim gSTOP As Boolean = False
    Dim gBPM As Boolean = True
    'Dim gA8 As Boolean = False
    Dim iPlayer As Integer = 0
    Dim gColumns As Integer = 46

    '----Visual Options
    Dim vo As New visualSettings()

    Public Sub setVO(ByVal xvo As visualSettings)
        vo = xvo
    End Sub

    '----Preview Options
    Structure PlayerArguments
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
    Dim spWidth() As Single = {0, 100, 0}
    Dim spH() As Integer = {0, 0, 0}
    Dim spV() As Integer = {0, 0, 0}
    Dim spLock() As Boolean = {False, False, False}
    Dim spDiff() As Integer = {0, 0, 0}
    Dim spFocus As Integer = 1 '0 = Left, 1 = Middle, 2 = Right
    Dim spMouseOver As Integer = 1

    Dim AutoFocusMouseEnter As Boolean = False
    Dim FirstClickDisabled As Boolean = True
    Dim tempFirstMouseDown As Boolean = False

    Dim spMain() As Panel = {}

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

    Private Function GetBMSChannelBy(note As Note) As String
        Dim iCol = note.ColumnIndex
        Dim xVal = note.Value
        Dim xLong = note.LongNote
        Dim xHidden = note.Hidden
        Dim xI As Integer = GetColumn(iCol).Identifier
        If iCol = niBPM AndAlso (xVal / 10000 <> xVal \ 10000 Or xVal >= 2560000 Or xVal < 0) Then xI += idflBPM
        If (iCol >= niA1 And iCol <= niA8) Or (iCol >= niD1 And iCol <= niD8) Then
            If xLong Then xI += idLong
            If xHidden Then xI += idHidden
        End If
        Return Add2Zeros(xI)
    End Function

    Private Function nLeft(ByVal iCol As Integer) As Integer
        If iCol < niB Then Return column(iCol).Left Else Return column(niB).Left + (iCol - niB) * column(niB).Width
    End Function
    Private Function getColumnWidth(ByVal iCol As Integer) As Integer
        If Not GetColumn(iCol).isVisible Then Return 0
        If iCol < niB Then Return column(iCol).Width Else Return column(niB).Width
    End Function
    Private Function nTitle(ByVal iCol As Integer) As String
        If iCol < niB Then Return column(iCol).Title Else Return column(niB).Title & (iCol - niB + 1).ToString
    End Function
    Private Function nEnabled(ByVal iCol As Integer) As Boolean
        'If iCol < niB Then Return col(iCol).Enabled And col(iCol).Visible Else Return col(niB).Enabled And col(niB).Visible
        If iCol < niB Then Return column(iCol).isEnabledAfterAll Else Return column(niB).isEnabledAfterAll
    End Function
    Private Function isColumnNumeric(ByVal iCol As Integer) As Boolean
        If iCol < niB Then Return column(iCol).isNumeric Else Return column(niB).isNumeric
    End Function



    Private Function GetColumn(ByVal iCol As Integer) As Column
        If iCol < niB Then Return column(iCol) Else Return column(niB)
    End Function

    Private Function BMSEChannelToColumnIndex(ByVal I As String)
        Dim Ivalue = Val(I)
        If Ivalue > 100 Then
            Return niB + Ivalue - 101
        ElseIf Ivalue < 100 And Ivalue > 0 Then
            Return BMSChannelToColumn(Mid(I, 2, 2))
        End If
        Return niB ' ??? how did a negative number get here?
    End Function

    Private Function BMSChannelToColumn(ByVal I As String) As Integer
        Select Case I
            Case "01" : Return niB
            Case "03", "08" : Return niBPM
            Case "09" : Return niSTOP
            Case "04" : Return niBGA
            Case "07" : Return niLAYER
            Case "06" : Return niPOOR

            Case "16", "36", "56", "76" : Return niA1
            Case "11", "31", "51", "71" : Return niA2
            Case "12", "32", "52", "72" : Return niA3
            Case "13", "33", "53", "73" : Return niA4
            Case "14", "34", "54", "74" : Return niA5
            Case "15", "35", "55", "75" : Return niA6
            Case "18", "38", "58", "78" : Return niA7
            Case "19", "39", "59", "79" : Return niA8

            Case "21", "41", "61", "81" : Return niD1
            Case "22", "42", "62", "82" : Return niD2
            Case "23", "43", "63", "83" : Return niD3
            Case "24", "44", "64", "84" : Return niD4
            Case "25", "45", "65", "85" : Return niD5
            Case "28", "48", "68", "88" : Return niD6
            Case "29", "49", "69", "89" : Return niD7
            Case "26", "46", "66", "86" : Return niD8

            Case Else : Return 0
        End Select
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="xHPosition">Original horizontal position.</param>
    ''' <param name="xHSVal">HS.Value</param>


    Private Function HorizontalPositiontoDisplay(ByVal xHPosition As Integer, ByVal xHSVal As Long) As Integer
        Return CInt(xHPosition * gxWidth - xHSVal * gxWidth)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="xVPosition">Original vertical position.</param>
    ''' <param name="xVSVal">VS.Value</param>
    ''' <param name="xTHeight">Height of the panel. (DisplayRectangle, but not ClipRectangle)</param>


    Private Function VerticalPositiontoDisplay(ByVal xVPosition As Double, ByVal xVSVal As Long, ByVal xTHeight As Integer) As Integer
        Return xTHeight - CInt((xVPosition + xVSVal) * gxHeight) - 1
    End Function

    Public Function MeasureAtDisplacement(ByVal xVPos As Double) As Integer
        'Return Math.Floor((xVPos + FloatTolerance) / 192)
        'Return Math.Floor(xVPos / 192)
        Dim xI1 As Integer
        For xI1 = 1 To 999
            If xVPos < MeasureBottom(xI1) Then Exit For
        Next
        Return xI1 - 1
    End Function

    Private Function VPosition1000() As Double
        Return MeasureUpper(999)
    End Function

    Private Function SnapToGrid(ByVal xVPos As Double) As Double
        Dim xOffset As Double = MeasureBottom(MeasureAtDisplacement(xVPos))
        Dim xRatio As Double = 192.0R / gDivide
        Return Math.Floor((xVPos - xOffset) / xRatio) * xRatio + xOffset
    End Function

    Private Sub CalculateGreatestVPosition()
        'If K Is Nothing Then Exit Sub
        Dim xI1 As Integer
        GreatestVPosition = 0

        If NTInput Then
            For xI1 = UBound(Notes) To 0 Step -1
                If Notes(xI1).VPosition + Notes(xI1).Length > GreatestVPosition Then GreatestVPosition = Notes(xI1).VPosition + Notes(xI1).Length
            Next
        Else
            For xI1 = UBound(Notes) To 0 Step -1
                If Notes(xI1).VPosition > GreatestVPosition Then GreatestVPosition = Notes(xI1).VPosition
            Next
        End If

        Dim xI2 As Integer = -CInt(IIf(GreatestVPosition + 2000 > VPosition1000(), VPosition1000, GreatestVPosition + 2000))
        VS.Minimum = xI2
        VSL.Minimum = xI2
        VSR.Minimum = xI2
    End Sub


    Private Sub SortByVPositionInsertion() 'Insertion Sort
        If UBound(Notes) <= 0 Then Exit Sub
        Dim xNote As Note
        Dim xI1 As Integer
        Dim xI2 As Integer
        For xI1 = 2 To UBound(Notes)
            xNote = Notes(xI1)
            For xI2 = xI1 - 1 To 1 Step -1
                If Notes(xI2).VPosition > xNote.VPosition Then
                    Notes(xI2 + 1) = Notes(xI2)
                    '                    If KMouseDown = xI2 Then KMouseDown += 1
                    If xI2 = 1 Then
                        Notes(xI2) = xNote
                        '                       If KMouseDown = xI1 Then KMouseDown = xI2
                        Exit For
                    End If
                Else
                    Notes(xI2 + 1) = xNote
                    '                    If KMouseDown = xI1 Then KMouseDown = xI2 + 1
                    Exit For
                End If
            Next
        Next

    End Sub

    Private Sub SortByVPositionQuick(ByVal xMin As Integer, ByVal xMax As Integer) 'Quick Sort
        Dim xNote As Note
        Dim iHi As Integer
        Dim iLo As Integer
        Dim xI1 As Integer

        ' If min >= max, the list contains 0 or 1 items so it is sorted.
        If xMin >= xMax Then Exit Sub

        ' Pick the dividing value.
        xI1 = CInt((xMax - xMin) / 2) + xMin
        xNote = Notes(xI1)

        ' Swap it to the front.
        Notes(xI1) = Notes(xMin)

        iLo = xMin
        iHi = xMax
        Do
            ' Look down from hi for a value < med_value.
            Do While Notes(iHi).VPosition >= xNote.VPosition
                iHi = iHi - 1
                If iHi <= iLo Then Exit Do
            Loop
            If iHi <= iLo Then
                Notes(iLo) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iLo) = Notes(iHi)

            ' Look up from lo for a value >= med_value.
            iLo = iLo + 1
            Do While Notes(iLo).VPosition < xNote.VPosition
                iLo = iLo + 1
                If iLo >= iHi Then Exit Do
            Loop
            If iLo >= iHi Then
                iLo = iHi
                Notes(iHi) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            Notes(iHi) = Notes(iLo)
        Loop

        ' Sort the two sublists.
        SortByVPositionQuick(xMin, iLo - 1)
        SortByVPositionQuick(iLo + 1, xMax)
    End Sub

    Private Sub SortByVPositionQuick3(ByVal xMin As Integer, ByVal xMax As Integer)
        Dim xxMin As Integer
        Dim xxMax As Integer
        Dim xxMid As Integer
        Dim xNote As Note
        Dim xNoteMid As Note
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        'If xMax = 0 Then
        '    xMin = LBound(K1)
        '    xMax = UBound(K1)
        'End If
        xxMin = xMin
        xxMax = xMax
        xxMid = xMax - xMin + 1
        xI1 = CInt(Int(xxMid * Rnd())) + xMin
        xI2 = CInt(Int(xxMid * Rnd())) + xMin
        xI3 = CInt(Int(xxMid * Rnd())) + xMin
        If Notes(xI1).VPosition <= Notes(xI2).VPosition And Notes(xI2).VPosition <= Notes(xI3).VPosition Then
            xxMid = xI2
        Else
            If Notes(xI2).VPosition <= Notes(xI1).VPosition And Notes(xI1).VPosition <= Notes(xI3).VPosition Then
                xxMid = xI1
            Else
                xxMid = xI3
            End If
        End If
        xNoteMid = Notes(xxMid)
        Do
            Do While Notes(xxMin).VPosition < xNoteMid.VPosition And xxMin < xMax
                xxMin = xxMin + 1
            Loop
            Do While xNoteMid.VPosition < Notes(xxMax).VPosition And xxMax > xMin
                xxMax = xxMax - 1
            Loop
            If xxMin <= xxMax Then
                xNote = Notes(xxMin)
                Notes(xxMin) = Notes(xxMax)
                Notes(xxMax) = xNote
                xxMin = xxMin + 1
                xxMax = xxMax - 1
            End If
        Loop Until xxMin > xxMax
        If xxMax - xMin < xMax - xxMin Then
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
        Else
            If xxMin < xMax Then SortByVPositionQuick3(xxMin, xMax)
            If xMin < xxMax Then SortByVPositionQuick3(xMin, xxMax)
        End If
    End Sub


    Private Sub UpdateMeasureBottom()
        MeasureBottom(0) = 0.0#
        For xI1 As Integer = 0 To 998
            MeasureBottom(xI1 + 1) = MeasureBottom(xI1) + MeasureLength(xI1)
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
        Dim xMeasure As Integer = MeasureAtDisplacement(Math.Abs(spV(spFocus)))
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

    Private Sub PreviewNote(ByVal xFileLocation As String, ByVal bStop As Boolean)
        If bStop Then
            Audio.StopPlaying()
        End If
        Audio.Play(xFileLocation)
    End Sub

    Private Sub AddNote(ByVal xVPosition As Double,
                        ByVal xColumnIndex As Integer,
                        ByVal xValue As Long,
                        ByVal xLongNote As Boolean,
                        ByVal xHidden As Boolean,
               Optional ByVal xSelected As Boolean = False,
               Optional ByVal OverWrite As Boolean = True,
               Optional ByVal SortAndUpdatePairing As Boolean = True)

        If xVPosition < 0 Or xVPosition >= VPosition1000() Then Exit Sub

        Dim xI1 As Integer = 1

        If OverWrite Then
            Do While xI1 <= UBound(Notes)
                If Notes(xI1).VPosition = xVPosition And Notes(xI1).ColumnIndex = xColumnIndex Then RemoveNote(xI1) Else xI1 += 1
            Loop
        End If

        ReDim Preserve Notes(UBound(Notes) + 1)
        With Notes(UBound(Notes))
            .VPosition = xVPosition
            .ColumnIndex = xColumnIndex
            .Value = xValue
            .LongNote = xLongNote
            .Hidden = xHidden
            .Selected = xSelected And nEnabled(xColumnIndex)
        End With

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub RemoveNote(ByVal I As Integer, Optional ByVal SortAndUpdatePairing As Boolean = True)
        KMouseOver = -1
        Dim xI2 As Integer
        For xI2 = I + 1 To UBound(Notes)
            Notes(xI2 - 1) = Notes(xI2)
        Next
        ReDim Preserve Notes(UBound(Notes) - 1)
        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        'CalculateTotalNotes()
    End Sub

    Private Sub AddNotesFromClipboard(Optional ByVal xSelected As Boolean = True, Optional ByVal SortAndUpdatePairing As Boolean = True)
        Dim xStrLine() As String = Split(Clipboard.GetText, vbCrLf)

        Dim xI1 As Integer
        For xI1 = 0 To UBound(Notes)
            Notes(xI1).Selected = False
        Next

        Dim xVS As Long = spV(spFocus)
        Dim xTempVP As Double
        Dim xKbu() As Note = Notes

        If xStrLine(0) = "iBMSC Clipboard Data" Then
            If NTInput Then ReDim Preserve Notes(0)

            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = Val(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 4 And xTempVP >= 0 And xTempVP < VPosition1000() Then
                    ReDim Preserve Notes(UBound(Notes) + 1)
                    With Notes(UBound(Notes))
                        .ColumnIndex = Val(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = Val(xStrSub(2))
                        .LongNote = CBool(Val(xStrSub(3)))
                        .Hidden = CBool(Val(xStrSub(4)))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If NTInput Then
                ConvertBMSE2NT()

                For xI1 = 1 To UBound(Notes)
                    Notes(xI1 - 1) = Notes(xI1)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For xI1 = xIStart To UBound(Notes)
                    Notes(xI1) = xKn(xI1 - xIStart)
                Next
            End If

        ElseIf xStrLine(0) = "iBMSC Clipboard Data xNT" Then
            If Not NTInput Then ReDim Preserve Notes(0)

            'paste
            Dim xStrSub() As String
            For xI1 = 1 To UBound(xStrLine)
                If xStrLine(xI1).Trim = "" Then Continue For
                xStrSub = Split(xStrLine(xI1), " ")
                xTempVP = Val(xStrSub(1)) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)
                If UBound(xStrSub) = 4 And xTempVP >= 0 And xTempVP < VPosition1000() Then
                    ReDim Preserve Notes(UBound(Notes) + 1)
                    With Notes(UBound(Notes))
                        .ColumnIndex = Val(xStrSub(0))
                        .VPosition = xTempVP
                        .Value = Val(xStrSub(2))
                        .Length = Val(xStrSub(3))
                        .Hidden = CBool(Val(xStrSub(4)))
                        .Selected = xSelected
                    End With
                End If
            Next

            'convert
            If Not NTInput Then
                ConvertNT2BMSE()

                For xI1 = 1 To UBound(Notes)
                    Notes(xI1 - 1) = Notes(xI1)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For xI1 = xIStart To UBound(Notes)
                    Notes(xI1) = xKn(xI1 - xIStart)
                Next
            End If

        ElseIf xStrLine(0) = "BMSE ClipBoard Object Data Format" Then
            If NTInput Then ReDim Preserve Notes(0)

            'paste
            For xI1 = 1 To UBound(xStrLine)
                ' zdr: holy crap this is obtuse
                Dim posStr = Mid(xStrLine(xI1), 5, 7)
                Dim vPos = Val(posStr) + MeasureBottom(MeasureAtDisplacement(-xVS) + 1)

                Dim bmsIdent = Mid(xStrLine(xI1), 1, 3)
                Dim lineCol = BMSEChannelToColumnIndex(bmsIdent)

                Dim Value = Val(Mid(xStrLine(xI1), 12)) * 10000

                Dim attribute = Mid(xStrLine(xI1), 4, 1)

                If Len(xStrLine(xI1)) > 11 And ' Valid length
                    lineCol > 0 And  ' Valid column 
                    vPos >= 0 And vPos < VPosition1000() Then ' In range
                    ReDim Preserve Notes(UBound(Notes) + 1)

                    With Notes(UBound(Notes))
                        .ColumnIndex = lineCol
                        .VPosition = vPos
                        .Value = Value
                        .LongNote = attribute = "2"
                        .Hidden = attribute = "1"
                        .Selected = xSelected And nEnabled(.ColumnIndex)
                    End With
                End If
            Next

            'convert
            If NTInput Then
                ConvertBMSE2NT()

                For xI1 = 1 To UBound(Notes)
                    Notes(xI1 - 1) = Notes(xI1)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)

                Dim xKn() As Note = Notes
                Notes = xKbu

                Dim xIStart As Integer = Notes.Length
                ReDim Preserve Notes(UBound(Notes) + xKn.Length)

                For xI1 = xIStart To UBound(Notes)
                    Notes(xI1) = xKn(xI1 - xIStart)
                Next
            End If
        End If

        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub CopyNotes(Optional ByVal Unselect As Boolean = True)
        Dim xStrAll As String = "iBMSC Clipboard Data" & IIf(NTInput, " xNT", "")
        Dim xI1 As Integer
        Dim MinMeasure As Double = 999

        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected And MeasureAtDisplacement(Notes(xI1).VPosition) < MinMeasure Then MinMeasure = MeasureAtDisplacement(Notes(xI1).VPosition)
        Next
        MinMeasure = MeasureBottom(MinMeasure)

        If Not NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Selected Then
                    xStrAll &= vbCrLf & Notes(xI1).ColumnIndex.ToString & " " &
                                       (Notes(xI1).VPosition - MinMeasure).ToString & " " &
                                        Notes(xI1).Value.ToString & " " &
                                   CInt(Notes(xI1).LongNote).ToString & " " &
                                   CInt(Notes(xI1).Hidden).ToString
                    Notes(xI1).Selected = Not Unselect
                End If
            Next

        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Selected Then
                    xStrAll &= vbCrLf & Notes(xI1).ColumnIndex.ToString & " " &
                                       (Notes(xI1).VPosition - MinMeasure).ToString & " " &
                                        Notes(xI1).Value.ToString & " " &
                                        Notes(xI1).Length.ToString & " " &
                                   CInt(Notes(xI1).Hidden).ToString
                    Notes(xI1).Selected = Not Unselect
                End If
            Next
        End If

        Clipboard.SetText(xStrAll)
    End Sub

    Private Sub RemoveNotes(Optional ByVal SortAndUpdatePairing As Boolean = True)
        If UBound(Notes) = 0 Then Exit Sub

        KMouseOver = -1
        Dim xI1 As Integer = 1
        Dim xI2 As Integer
        Do
            If Notes(xI1).Selected Then
                For xI2 = xI1 + 1 To UBound(Notes)
                    Notes(xI2 - 1) = Notes(xI2)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
                xI1 = 0
            End If
            xI1 += 1
        Loop While xI1 < UBound(Notes) + 1
        If SortAndUpdatePairing Then SortByVPositionInsertion() : UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Function EnabledColumnToReal(ByVal cEnabled As Integer) As Integer
        Dim xI1 As Integer = 0
        Do
            If xI1 >= gColumns Then Exit Do
            If Not nEnabled(xI1) Then cEnabled += 1
            If xI1 >= cEnabled Then Exit Do
            xI1 += 1
        Loop
        Return cEnabled
    End Function

    Private Function RealColumnToEnabled(ByVal cReal As Integer) As Integer
        Dim xI1 As Integer
        For xI1 = 0 To cReal - 1
            If Not nEnabled(xI1) Then cReal -= 1
        Next
        Return cReal
    End Function

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
            'If SaveTheme Then
            '    My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Skin.cff", SaveSkinCFF, False, System.Text.Encoding.Unicode)
            'Else
            '    My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Skin.cff", "", False, System.Text.Encoding.Unicode)
            'End If
            '
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\PlayerArgs.cff", SavePlayerCFF, False, System.Text.Encoding.Unicode)
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\Config.cff", SaveCFF, False, System.Text.Encoding.Unicode)
            'My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath & "\PreConfig.cff", "", False, System.Text.Encoding.Unicode)
            Me.SaveSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml", False)
        End If
    End Sub

    Private Function FilterFileBySupported(ByVal xFile() As String, ByVal xFilter() As String) As String()
        Dim xPath(-1) As String
        For xI1 As Integer = 0 To UBound(xFile)
            If My.Computer.FileSystem.FileExists(xFile(xI1)) And Array.IndexOf(xFilter, Path.GetExtension(xFile(xI1))) <> -1 Then
                ReDim Preserve xPath(UBound(xPath) + 1)
                xPath(UBound(xPath)) = xFile(xI1)
            End If

            If My.Computer.FileSystem.DirectoryExists(xFile(xI1)) Then
                Dim xFileNames() As FileInfo = My.Computer.FileSystem.GetDirectoryInfo(xFile(xI1)).GetFiles()
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
        'ReDim K(0)
        'With K(0)
        ' .ColumnIndex = niBPM
        ' .VPosition = -1
        ' .LongNote = False
        ' .Selected = False
        ' .Value = 1200000
        'End With

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
        For xI1 As Integer = 0 To 999
            MeasureLength(xI1) = 192.0R
            MeasureBottom(xI1) = xI1 * 192.0R
            LBeat.Items.Add(Add3Zeros(xI1) & ": 1 ( 4 / 4 )")
        Next
    End Sub

    Private Sub InitializeOpenBMS()
        CHPlayer.SelectedIndex = 0
        'THLnType.Text = ""
    End Sub

    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DDFileName = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedFileExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DragLeave
        ReDim DDFileName(-1)
        RefreshPanelAll()
    End Sub

    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragDrop
        ReDim DDFileName(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedFileExtension)
        If xPath.Length > 0 Then
            Dim xProg As New fLoadFileProgress(xPath, IsSaved)
            xProg.ShowDialog(Me)
        End If

        RefreshPanelAll()
    End Sub

    Private Sub setFullScreen(ByVal value As Boolean)
        If value Then
            If Me.WindowState = FormWindowState.Minimized Then Exit Sub

            Me.SuspendLayout()
            previousWindowPosition.Location = Me.Location
            previousWindowPosition.Size = Me.Size
            previousWindowState = Me.WindowState

            Me.WindowState = FormWindowState.Normal
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
            Me.WindowState = FormWindowState.Maximized
            ToolStripContainer1.TopToolStripPanelVisible = False

            Me.ResumeLayout()
            isFullScreen = True
        Else
            Me.SuspendLayout()
            Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
            ToolStripContainer1.TopToolStripPanelVisible = True
            Me.WindowState = FormWindowState.Normal

            Me.WindowState = previousWindowState
            If Me.WindowState = FormWindowState.Normal Then
                Me.Location = previousWindowPosition.Location
                Me.Size = previousWindowPosition.Size
            End If

            Me.ResumeLayout()
            isFullScreen = False
        End If
    End Sub

    Private Sub Form1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F4
                If Not My.Computer.Keyboard.AltKeyDown Then MsgBox(Strings.Messages.EraserObsolete, MsgBoxStyle.Information, "Eraser tool")
            Case Keys.F11
                setFullScreen(Not isFullScreen)
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

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'On Error Resume Next
        Me.TopMost = True
        Me.SuspendLayout()
        Me.Visible = False

        'POBMP.Dispose()
        'POBGA.Dispose()

        'Me.MaximizedBounds = Screen.GetWorkingArea(Me)
        'Me.Visible = False
        SetFileName(FileName)
        'Me.ShowCaption = False
        'SetWindowText(Me.Handle.ToInt32, FileName)

        InitializeNewBMS()
        'nBeatD.SelectedIndex = 4

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
            POBeatResizer.Cursor = xDownCursor
            POExpansionResizer.Cursor = xDownCursor

            POptionsResizer.Cursor = xLeftCursor

            SpL.Cursor = xRightCursor
            SpR.Cursor = xLeftCursor
        Catch ex As Exception

        End Try

        spMain = New Panel() {PMainInL, PMainIn, PMainInR}

        Dim xI1 As Integer

        sUndo(0) = New UndoRedo.NoOperation
        sUndo(1) = New UndoRedo.NoOperation
        sRedo(0) = New UndoRedo.NoOperation
        sRedo(1) = New UndoRedo.NoOperation
        sI = 0

        LWAV.Items.Clear()
        For xI1 = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ":")
        Next
        LWAV.SelectedIndex = 0
        CHPlayer.SelectedIndex = 0

        CalculateGreatestVPosition()
        TBLangRefresh_Click(TBLangRefresh, Nothing)
        TBThemeRefresh_Click(TBThemeRefresh, Nothing)

        POHeaderPart2.Visible = False
        POGridPart2.Visible = False
        POWaveFormPart2.Visible = False

        If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml") Then
            LoadSettings(My.Application.Info.DirectoryPath & "\iBMSC.Settings.xml")
            'Else
            '---- Settings for first-time start-up ---------------------------------------------------------------------------
            'Me.LoadLocale(My.Application.Info.DirectoryPath & "\Data\chs.Lang.xml")
            '-----------------------------------------------------------------------------------------------------------------
        End If
        'On Error GoTo 0
        SetIsSaved(True)

        Dim xStr() As String = Environment.GetCommandLineArgs
        'Dim xStr() As String = {Application.ExecutablePath, "C:\Users\User\Desktop\yang run xuan\SoFtwArES\Games\O2Mania\music\SHOOT!\shoot! -NM-.bms"}

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
                Notes(i).PairWithI = 0
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

                    If Notes(i).Value \ 10000 = LnObj AndAlso Not isColumnNumeric(Notes(i).ColumnIndex) Then
                        For j = i - 1 To 1 Step -1
                            If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                            If Notes(j).Hidden Then Continue For

                            If Notes(j).Length <> 0 OrElse Notes(j).Value \ 10000 = LnObj Then
                                Notes(i).HasError = True
                            Else
                                Notes(i).PairWithI = j
                                Notes(j).PairWithI = i
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
                Notes(i).PairWithI = 0
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
                        ElseIf Notes(j).LongNote And Notes(j).PairWithI = i Then
                            Notes(i).PairWithI = j
                            GoTo EndSearch
                        Else
                            Exit For
                        End If
                    Next

                    For j = i + 1 To UBound(Notes)
                        If Notes(j).ColumnIndex <> Notes(i).ColumnIndex Then Continue For
                        Notes(i).PairWithI = j
                        Notes(j).PairWithI = i
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
                    Not isColumnNumeric(Notes(i).ColumnIndex) Then
                    'LnObj: Match anything below.
                    '           If matching a LongNote not matching back, then error on below.
                    '           If overlapping a note, then error.
                    '           If mathcing a LnObj below, then error on below.
                    '       If nothing below, then error.
                    For j = i - 1 To 1 Step -1
                        If Notes(i).ColumnIndex <> Notes(j).ColumnIndex Then Continue For
                        If Notes(j).PairWithI <> 0 And Notes(j).PairWithI <> i Then
                            Notes(j).HasError = True
                        End If
                        Notes(i).PairWithI = j
                        Notes(j).PairWithI = i
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
            If Notes(i).ColumnIndex = niBPM Then
                currentMS += (Notes(i).VPosition - currentBPMVPosition) / currentBPM * 1250
                currentBPM = Notes(i).Value / 10000
                currentBPMVPosition = Notes(i).VPosition
            End If
            'K(i).TimeOffset = currentMS + (K(i).VPosition - currentBPMVPosition) / currentBPM * 1250
        Next
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
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        ClearUndo()
        InitializeNewBMS()

        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With
        THBPM.Value = 120

        LWAV.Items.Clear()
        Dim xI1 As Integer
        For xI1 = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ": " & hWAV(xI1))
        Next
        LWAV.SelectedIndex = 0

        SetFileName("Untitled.bms")
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        CalculateTotalPlayableNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub TBNewC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles TBNewC.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        ClearUndo()

        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        THGenre.Text = ""
        THTitle.Text = ""
        THArtist.Text = ""
        THPlayLevel.Text = ""

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With
        THBPM.Value = 120

        SetFileName("Untitled.bms")
        SetIsSaved(True)
        'pIsSaved.Visible = Not IsSaved

        If MsgBox("Please copy your code to clipboard and click OK.", MsgBoxStyle.OkCancel, "Create from code") = MsgBoxResult.Cancel Then Exit Sub
        OpenBMS(Clipboard.GetText)
    End Sub

    Private Sub TBOpen_ButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBOpen.ButtonClick, mnOpen.Click
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt"
        xDOpen.DefaultExt = "bms"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

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
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Return

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType.IBMSC & "|*.ibmsc"
        xDOpen.DefaultExt = "ibmsc"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

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
        ReDim SelectedNotes(-1)
        KMouseOver = -1
        If ClosingPopSave() Then Exit Sub

        Dim xDOpen As New OpenFileDialog
        xDOpen.Filter = Strings.FileType.SM & "|*.sm"
        xDOpen.DefaultExt = "sm"
        xDOpen.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

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
        ReDim SelectedNotes(-1)
        KMouseOver = -1

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
        ReDim SelectedNotes(-1)
        KMouseOver = -1

        Dim xDSave As New SaveFileDialog
        xDSave.Filter = Strings.FileType._bms & "|*.bms;*.bme;*.bml;*.pms;*.txt|" &
                        Strings.FileType.BMS & "|*.bms|" &
                        Strings.FileType.BME & "|*.bme|" &
                        Strings.FileType.BML & "|*.bml|" &
                        Strings.FileType.PMS & "|*.pms|" &
                        Strings.FileType.TXT & "|*.txt|" &
                        Strings.FileType._all & "|*.*"
        xDSave.DefaultExt = "bms"
        xDSave.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))

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
        ReDim SelectedNotes(-1)
        KMouseOver = -1

        Dim xDSave As New SaveFileDialog
        xDSave.Filter = Strings.FileType.IBMSC & "|*.ibmsc"
        xDSave.DefaultExt = "ibmsc"
        xDSave.InitialDirectory = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName))
        If xDSave.ShowDialog = Windows.Forms.DialogResult.Cancel Then Exit Sub

        SaveiBMSC(xDSave.FileName)
        'My.Computer.FileSystem.WriteAllText(xDSave.FileName, xStrAll, False, TextEncoding)
        NewRecent(FileName)
        If BeepWhileSaved Then Beep()
    End Sub



    Private Sub VSGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles VS.GotFocus, VSL.GotFocus, VSR.GotFocus
        spFocus = sender.Tag
        spMain(spFocus).Focus()
    End Sub

    Private Sub VSValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles VS.ValueChanged, VSL.ValueChanged, VSR.ValueChanged
        Dim iI As Integer = sender.Tag

        ' az: We got a wheel event when we're zooming in/out
        If My.Computer.Keyboard.CtrlKeyDown Then
            sender.Value = VSValue ' Undo the scroll
            Exit Sub
        End If

        If iI = spFocus And Not LastMouseDownLocation = New Point(-1, -1) And Not VSValue = -1 Then LastMouseDownLocation.Y += (VSValue - sender.Value) * gxHeight
        spV(iI) = sender.Value

        If spLock((iI + 1) Mod 3) Then
            Dim xVS As Integer = spV(iI) + spDiff(iI)
            If xVS > 0 Then xVS = 0
            If xVS < VS.Minimum Then xVS = VS.Minimum
            Select Case iI
                Case 0 : VS.Value = xVS
                Case 1 : VSR.Value = xVS
                Case 2 : VSL.Value = xVS
            End Select
        End If

        If spLock((iI + 2) Mod 3) Then
            Dim xVS As Integer = spV(iI) - spDiff((iI + 2) Mod 3)
            If xVS > 0 Then xVS = 0
            If xVS < VS.Minimum Then xVS = VS.Minimum
            Select Case iI
                Case 0 : VSR.Value = xVS
                Case 1 : VSL.Value = xVS
                Case 2 : VS.Value = xVS
            End Select
        End If

        spDiff(iI) = spV((iI + 1) Mod 3) - spV(iI)
        spDiff((iI + 2) Mod 3) = spV(iI) - spV((iI + 2) Mod 3)

        VSValue = sender.Value
        RefreshPanel(iI, spMain(iI).DisplayRectangle)
    End Sub

    Private Sub cVSLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cVSLockL.CheckedChanged, cVSLock.CheckedChanged, cVSLockR.CheckedChanged
        Dim iI As Integer = sender.Tag
        spLock(iI) = sender.Checked
        If Not spLock(iI) Then Return

        spDiff(iI) = spV((iI + 1) Mod 3) - spV(iI)
        spDiff((iI + 2) Mod 3) = spV(iI) - spV((iI + 2) Mod 3)

        'POHeaderB.Text = spDiff(0) & "_" & spDiff(1) & "_" & spDiff(2)
    End Sub

    Private Sub HSGotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles HS.GotFocus, HSL.GotFocus, HSR.GotFocus
        spFocus = sender.Tag
        spMain(spFocus).Focus()
    End Sub

    Private Sub HSValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles HS.ValueChanged, HSL.ValueChanged, HSR.ValueChanged
        Dim iI As Integer = sender.Tag
        If Not LastMouseDownLocation = New Point(-1, -1) And Not HSValue = -1 Then LastMouseDownLocation.X += (HSValue - sender.Value) * gxWidth
        spH(iI) = sender.Value
        HSValue = sender.Value
        RefreshPanel(iI, spMain(iI).DisplayRectangle)
    End Sub

    Private Sub TBSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBSelect.Click, mnSelect.Click
        TBSelect.Checked = True
        TBWrite.Checked = False
        TBTimeSelect.Checked = False
        mnSelect.Checked = True
        mnWrite.Checked = False
        mnTimeSelect.Checked = False

        FStatus2.Visible = False
        FStatus.Visible = True

        ShouldDrawTempNote = False
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0

        vSelStart = MeasureBottom(MeasureAtDisplacement(-spV(spFocus)) + 1)
        vSelLength = 0

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

        ShouldDrawTempNote = True
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0

        vSelStart = MeasureBottom(MeasureAtDisplacement(-spV(spFocus)) + 1)
        vSelLength = 0

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

        vSelMouseOverLine = 0
        ShouldDrawTempNote = False
        SelectedColumn = -1
        TempVPosition = -1
        TempLength = 0
        ValidateSelection()

        Dim xI1 As Integer
        For xI1 = 0 To UBound(Notes)
            Notes(xI1).Selected = False
        Next
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub CGHeight_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight.ValueChanged
        gxHeight = CSng(CGHeight.Value)
        CGHeight2.Value = IIf(CGHeight.Value * 4 < CGHeight2.Maximum, CDec(CGHeight.Value * 4), CGHeight2.Maximum)
        RefreshPanelAll()
    End Sub

    Private Sub CGHeight2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGHeight2.Scroll
        CGHeight.Value = CGHeight2.Value / 4
    End Sub

    Private Sub CGWidth_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth.ValueChanged
        gxWidth = CSng(CGWidth.Value)
        CGWidth2.Value = IIf(CGWidth.Value * 4 < CGWidth2.Maximum, CDec(CGWidth.Value * 4), CGWidth2.Maximum)

        HS.LargeChange = PMainIn.Width / gxWidth
        If HS.Value > HS.Maximum - HS.LargeChange + 1 Then HS.Value = HS.Maximum - HS.LargeChange + 1
        HSL.LargeChange = PMainInL.Width / gxWidth
        If HSL.Value > HSL.Maximum - HSL.LargeChange + 1 Then HSL.Value = HSL.Maximum - HSL.LargeChange + 1
        HSR.LargeChange = PMainInR.Width / gxWidth
        If HSR.Value > HSR.Maximum - HSR.LargeChange + 1 Then HSR.Value = HSR.Maximum - HSR.LargeChange + 1

        RefreshPanelAll()
    End Sub

    Private Sub CGWidth2_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles CGWidth2.Scroll
        CGWidth.Value = CGWidth2.Value / 4
    End Sub

    Private Sub CGDivide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDivide.ValueChanged
        gDivide = CGDivide.Value
        RefreshPanelAll()
    End Sub
    Private Sub CGSub_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSub.ValueChanged
        gSub = CGSub.Value
        RefreshPanelAll()
    End Sub
    Private Sub BGSlash_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BGSlash.Click
        Dim xd As Integer = Val(InputBox(Strings.Messages.PromptSlashValue, , gSlash))
        If xd = 0 Then Exit Sub
        If xd > CGDivide.Maximum Then xd = CGDivide.Maximum
        If xd < CGDivide.Minimum Then xd = CGDivide.Minimum
        gSlash = xd
    End Sub


    Private Sub CGSnap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSnap.CheckedChanged
        gSnap = CGSnap.Checked
        RefreshPanelAll()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim xI1 As Integer

        Select Case spFocus
            Case 0
                With VSL
                    xI1 = .Value + (tempY / 5) / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSL
                    xI1 = .Value + (tempX / 10) / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 1
                With VS
                    xI1 = .Value + (tempY / 5) / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HS
                    xI1 = .Value + (tempX / 10) / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 2
                With VSR
                    xI1 = .Value + (tempY / 5) / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSR
                    xI1 = .Value + (tempX / 10) / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
        End Select

        Dim xMEArgs As New System.Windows.Forms.MouseEventArgs(Windows.Forms.MouseButtons.Left, 0, MouseMoveStatus.X, MouseMoveStatus.Y, 0)
        PMainInMouseMove(spMain(spFocus), xMEArgs)

    End Sub

    Private Sub TimerMiddle_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMiddle.Tick
        If Not MiddleButtonClicked Then TimerMiddle.Enabled = False : Return

        Dim xI1 As Integer

        Select Case spFocus
            Case 0
                With VSL
                    xI1 = .Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSL
                    xI1 = .Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 1
                With VS
                    xI1 = .Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HS
                    xI1 = .Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With

            Case 2
                With VSR
                    xI1 = .Value + (Cursor.Position.Y - MiddleButtonLocation.Y) / 5 / gxHeight
                    If xI1 > 0 Then xI1 = 0
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
                With HSR
                    xI1 = .Value + (Cursor.Position.X - MiddleButtonLocation.X) / 5 / gxWidth
                    If xI1 > .Maximum - .LargeChange + 1 Then xI1 = .Maximum - .LargeChange + 1
                    If xI1 < .Minimum Then xI1 = .Minimum
                    .Value = xI1
                End With
        End Select

        Dim xMEArgs As New System.Windows.Forms.MouseEventArgs(Windows.Forms.MouseButtons.Left, 0, MouseMoveStatus.X, MouseMoveStatus.Y, 0)
        PMainInMouseMove(spMain(spFocus), xMEArgs)
    End Sub

    Private Sub validate_LWAV_view()
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
        If hWAV(LWAV.SelectedIndex + 1) = "" Then Exit Sub

        Dim xFileLocation As String = IIf(ExcludeFileName(FileName) = "", InitPath, ExcludeFileName(FileName)) & "\" & hWAV(LWAV.SelectedIndex + 1)
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
        hWAV(LWAV.SelectedIndex + 1) = GetFileName(xDWAV.FileName)
        LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": " & GetFileName(xDWAV.FileName)
        If IsSaved Then SetIsSaved(False)
    End Sub

    Private Sub LWAV_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LWAV.KeyDown
        Select Case e.KeyCode
            Case Keys.Delete
                hWAV(LWAV.SelectedIndex + 1) = ""
                LWAV.Items.Item(LWAV.SelectedIndex) = C10to36(LWAV.SelectedIndex + 1) & ": "
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
        For xI1 As Integer = 1 To UBound(Notes)
            Me.RedoRemoveNoteSelected(True, xUndo, xRedo)
        Next
        'Dim xRedo As String = sCmdKDs()
        'Dim xUndo As String = sCmdKs(True)

        CopyNotes(False)
        RemoveNotes(False)
        AddUndo(xUndo, xBaseRedo.Next)

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
        Me.RedoAddNoteSelected(True, xUndo, xRedo)
        AddUndo(xUndo, xBaseRedo.Next)

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
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)
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
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)
        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aHere))
    End Sub

    Private Sub TBPlayB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBPlayB.Click, mnPlayB.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        Dim xStrAll As String = SaveBMS()
        Dim xFileName As String = IIf(Not PathIsValid(FileName),
                                      IIf(InitPath = "", My.Application.Info.DirectoryPath, InitPath),
                                      ExcludeFileName(FileName)) & "\___TempBMS.bms"
        My.Computer.FileSystem.WriteAllText(xFileName, xStrAll, False, TextEncoding)

        AddTempFileList(xFileName)
        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aBegin))
    End Sub

    Private Sub TBStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBStop.Click, mnStop.Click
        'Dim xStr() As String = Split(pArgs(CurrentPlayer), vbCrLf)
        Dim xArg As MainWindow.PlayerArguments = pArgs(CurrentPlayer)

        If Not File.Exists(PrevCodeToReal(xArg.Path)) Then
            PlayerMissingPrompt()
            xArg = pArgs(CurrentPlayer)
        End If

        System.Diagnostics.Process.Start(PrevCodeToReal(xArg.Path), PrevCodeToReal(xArg.aStop))
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

        Dim data(5, 5) As Integer
        For i As Integer = 1 To UBound(Notes)
            With Notes(i)
                Dim row As Integer = -1
                Select Case .ColumnIndex
                    Case niBPM : row = 0
                    Case niSTOP : row = 1
                    Case niA1, niA2, niA3, niA4, niA5, niA6, niA7, niA8 : row = 2
                    Case niD1, niD2, niD3, niD4, niD5, niD6, niD7, niD8 : row = 3
                    Case Is >= niB : row = 4
                    Case Else : row = 5
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

                If row <> 5 Then row = 5 : GoTo StartCount
            End With
        Next

        Dim dStat As New dgStatistics(data)
        dStat.ShowDialog()
    End Sub

    ''' <summary>
    ''' Remark: Pls sort and updatepairing before this process.
    ''' </summary>

    Private Sub CalculateTotalPlayableNotes()
        Dim xI1 As Integer
        Dim xIAll As Integer = 0

        If Not NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).ColumnIndex >= niA1 And Notes(xI1).ColumnIndex <= niA8 Then xIAll += 1
            Next

        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).ColumnIndex >= niA1 And Notes(xI1).ColumnIndex <= niA8 Then
                    xIAll += 1
                    If Notes(xI1).Length <> 0 Then xIAll += 1
                End If
            Next
        End If

        TBStatistics.Text = xIAll
    End Sub

    Private Sub POStatusRefresh()

        If TBSelect.Checked Then
            Dim xI1 As Integer = KMouseOver
            If xI1 < 0 Then

                TempVPosition = (spMain(spFocus).Height - spV(spFocus) * gxHeight - MouseMoveStatus.Y - 1) / gxHeight 'VPosition of the mouse
                If gSnap Then TempVPosition = SnapToGrid(TempVPosition)

                xI1 = 0
                Dim mLeft As Integer = MouseMoveStatus.X / gxWidth + spH(spFocus) 'horizontal position of the mouse
                If mLeft >= 0 Then
                    Do
                        If mLeft < nLeft(xI1 + 1) Or xI1 >= gColumns Then SelectedColumn = xI1 : Exit Do 'get the column where mouse is 
                        xI1 += 1
                    Loop
                End If
                SelectedColumn = EnabledColumnToReal(RealColumnToEnabled(SelectedColumn))  'get the enabled column where mouse is 

                Dim xMeasure As Integer = MeasureAtDisplacement(TempVPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = TempVPosition - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = TempVPosition.ToString() & "  "
                FSC.Text = nTitle(SelectedColumn)
                FSW.Text = ""
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = ""
                FSH.Text = ""
                FSE.Text = ""

            Else
                Dim xMeasure As Integer = MeasureAtDisplacement(Notes(xI1).VPosition)
                Dim xMLength As Double = MeasureLength(xMeasure)
                Dim xVposMod As Double = Notes(xI1).VPosition - MeasureBottom(xMeasure)
                Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

                FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
                FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
                FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
                FSP4.Text = Notes(xI1).VPosition.ToString() & "  "
                FSC.Text = nTitle(Notes(xI1).ColumnIndex)
                FSW.Text = IIf(isColumnNumeric(Notes(xI1).ColumnIndex),
                               Notes(xI1).Value / 10000,
                               C10to36(Notes(xI1).Value \ 10000))
                FSM.Text = Add3Zeros(xMeasure)
                FST.Text = IIf(NTInput, Strings.StatusBar.Length & " = " & Notes(xI1).Length, IIf(Notes(xI1).LongNote, Strings.StatusBar.LongNote, ""))
                FSH.Text = IIf(Notes(xI1).Hidden, Strings.StatusBar.Hidden, "")
                FSE.Text = IIf(Notes(xI1).HasError, Strings.StatusBar.Err, "")

            End If

        ElseIf TBWrite.Checked Then
            If SelectedColumn < 0 Then Exit Sub

            Dim xMeasure As Integer = MeasureAtDisplacement(TempVPosition)
            Dim xMLength As Double = MeasureLength(xMeasure)
            Dim xVposMod As Double = TempVPosition - MeasureBottom(xMeasure)
            Dim xGCD As Double = GCD(IIf(xVposMod = 0, xMLength, xVposMod), xMLength)

            FSP1.Text = (xVposMod * gDivide / 192).ToString & " / " & (xMLength * gDivide / 192).ToString & "  "
            FSP2.Text = xVposMod.ToString & " / " & xMLength & "  "
            FSP3.Text = CInt(xVposMod / xGCD).ToString & " / " & CInt(xMLength / xGCD).ToString & "  "
            FSP4.Text = TempVPosition.ToString() & "  "
            FSC.Text = nTitle(SelectedColumn)
            FSW.Text = C10to36(LWAV.SelectedIndex + 1)
            FSM.Text = Add3Zeros(xMeasure)
            FST.Text = IIf(NTInput, TempLength, IIf(My.Computer.Keyboard.ShiftKeyDown, Strings.StatusBar.LongNote, ""))
            FSH.Text = IIf(My.Computer.Keyboard.CtrlKeyDown, Strings.StatusBar.Hidden, "")

        ElseIf TBTimeSelect.Checked Then
            FSSS.Text = vSelStart
            FSSL.Text = vSelLength
            FSSH.Text = vSelHalf

        End If

    End Sub

    Private Sub POBStorm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBStorm.Click

    End Sub

    Private Sub POBMirror_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBMirror.Click
        Dim xI1 As Integer
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
        'xRedo &= sCmdKM(niA1, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(niA7) - RealColumnToEnabled(niA1), 0, True) & vbCrLf
        'xUndo &= sCmdKM(niA7, .VPosition, .Value, IIf(NTInput, .Length, .LongNote), .Hidden, RealColumnToEnabled(niA1) - RealColumnToEnabled(niA7), 0, True) & vbCrLf

        Dim xCol As Integer = 0
        For xI1 = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Select Case Notes(xI1).ColumnIndex
                Case niA1 : xCol = niA7
                Case niA2 : xCol = niA6
                Case niA3 : xCol = niA5
                Case niA4 : xCol = niA4
                Case niA5 : xCol = niA3
                Case niA6 : xCol = niA2
                Case niA7 : xCol = niA1
                Case Else : Continue For
            End Select

            Me.RedoMoveNote(Notes(xI1), xCol, Notes(xI1).VPosition, True, xUndo, xRedo)
            Notes(xI1).ColumnIndex = xCol
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        UpdatePairing()
        RefreshPanelAll()
    End Sub



    Private Sub BVCApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCApply.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        SortByVPositionInsertion()
        BPMChangeTop(Val(TVCM.Text) / Val(TVCD.Text))

        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        POStatusRefresh()
        CalculateGreatestVPosition()

        Beep()
        TVCM.Focus()
        'Select Case spFocus
        '    Case 0 : PMainInL.Focus()
        '    Case 1 : PMainIn.Focus()
        '    Case 2 : PMainInR.Focus()
        'End Select
    End Sub

    Private Sub BVCCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCCalculate.Click
        If Not TBTimeSelect.Checked Then Exit Sub

        SortByVPositionInsertion()
        BPMChangeByValue(Val(TVCBPM.Text) * 10000)

        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
        POStatusRefresh()

        Beep()
        TVCBPM.Focus()
        'Select Case spFocus
        '    Case 0 : PMainInL.Focus()
        '    Case 1 : PMainIn.Focus()
        '    Case 2 : PMainInR.Focus()
        'End Select
    End Sub

    Private Sub ValidateSelection()
        If vSelStart < 0 Then vSelLength += vSelStart : vSelHalf += vSelStart : vSelStart = 0
        If vSelStart > VPosition1000() - 1 Then vSelLength += vSelStart - VPosition1000() + 1 : vSelHalf += vSelStart - VPosition1000() + 1 : vSelStart = VPosition1000() - 1
        If vSelStart + vSelLength < 0 Then vSelLength = -vSelStart
        If vSelStart + vSelLength > VPosition1000() - 1 Then vSelLength = VPosition1000() - 1 - vSelStart

        If Math.Sign(vSelHalf) <> Math.Sign(vSelLength) Then vSelHalf = 0
        If Math.Abs(vSelHalf) > Math.Abs(vSelLength) Then vSelHalf = vSelLength
    End Sub

    Private Sub BPMChangeTop(ByVal xRatio As Double, Optional ByVal bAddUndo As Boolean = True, Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then GoTo EndofSub
        If xRatio = 1 Or xRatio <= 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If xVLower < 0 Then xVLower = 0
        If xVUpper >= VPosition1000() Then xVUpper = VPosition1000() - 1

        Dim xBPM As Integer = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueU As Integer = xBPM

        'Save undo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        'Start
        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
            Next
            xValueL = xBPM
            xI2 = xI1

            'Within Selection
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVUpper Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio <= 655359999, Notes(xI1).Value * xRatio, 655359999)
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio + xVLower
            Next
            xValueU = xBPM
            xI2 = xI1

            'Above Selection
            For xI1 = xI2 To UBound(Notes)
                Notes(xI1).VPosition += (xRatio - 1) * (xVUpper - xVLower)
            Next

            'Add BPMs
            AddNote(xVLower, niBPM, IIf(xValueL * xRatio <= 655359999, xValueL * xRatio, 655359999), False, False, , , False)
            AddNote(xVUpper + (xRatio - 1) * (xVUpper - xVLower), niBPM, xValueU, False, False, , , False)

        Else
            Dim xAddBPML As Boolean = True
            Dim xAddBPMU As Boolean = True

            For xI1 = 1 To UBound(Notes)
                'Modify notes
                If Notes(xI1).VPosition <= xVLower Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueL = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVLower Then xAddBPML = False : Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio <= 655359999, Notes(xI1).Value * xRatio, 655359999)
                    End If

                    'If longnote then adjust length
                    If Notes(xI1).VPosition + Notes(xI1).Length > xVLower Then
                        Notes(xI1).Length += (IIf(xVUpper < Notes(xI1).VPosition + Notes(xI1).Length, xVUpper, Notes(xI1).VPosition + Notes(xI1).Length) - xVLower) * (xRatio - 1)
                    End If

                ElseIf Notes(xI1).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVUpper Then xAddBPMU = False Else Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio <= 655359999, Notes(xI1).Value * xRatio, 655359999)
                    End If

                    'Adjust Length
                    Notes(xI1).Length += (IIf(xVUpper < Notes(xI1).Length + Notes(xI1).VPosition, xVUpper, Notes(xI1).Length + Notes(xI1).VPosition) - Notes(xI1).VPosition) * (xRatio - 1)

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio + xVLower

                Else
                    Notes(xI1).VPosition += (xVUpper - xVLower) * (xRatio - 1)
                End If
            Next

            'Add BPMs
            If xAddBPML Then AddNote(xVLower, niBPM, IIf(xValueL * xRatio <= 655359999, xValueL * xRatio, 655359999), False, False, , , False)
            If xAddBPMU Then AddNote((xVUpper - xVLower) * xRatio + xVLower, niBPM, xValueU, False, False, , , False)
        End If

        'Check BPM Overflow
        For xI3 = 1 To UBound(Notes)
            If Notes(xI3).ColumnIndex = niBPM AndAlso Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
        Next

        Me.RedoAddNoteAll(False, xUndo, xRedo)

        'Restore selection
        Dim pSelStart As Double = vSelStart
        Dim pSelLength As Double = vSelLength
        Dim pSelHalf As Double = vSelHalf
        If vSelLength < 0 Then vSelStart += (xRatio - 1) * (xVUpper - xVLower)
        vSelLength = vSelLength * xRatio
        vSelHalf = vSelHalf * xRatio
        ValidateSelection()
        Me.RedoChangeTimeSelection(pSelStart, pSelLength, pSelHalf, vSelStart, vSelLength, vSelHalf, True, xUndo, xRedo)

        'Save redo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xRedo = "KZ" & vbCrLf & _
        '                      sCmdKs(False) & vbCrLf & _
        '                      "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

EndofSub:
        If bAddUndo Then AddUndo(xUndo, xBaseRedo.Next, bOverWriteUndo)
    End Sub

    Private Sub BPMChangeHalf(ByVal dVPosition As Double, Optional ByVal bAddUndo As Boolean = True, Optional ByVal bOverWriteUndo As Boolean = False)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then GoTo EndofSub
        If dVPosition = 0 Then GoTo EndofSub

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVHalf As Double = vSelStart + vSelHalf
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= VPosition1000() Then xVUpper = VPosition1000() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM As Integer = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xValueL As Integer = xBPM
        Dim xValueM As Integer = xBPM
        Dim xValueU As Integer = xBPM

        Dim xRatio1 As Double = (xVHalf - xVLower + dVPosition) / (xVHalf - xVLower)
        Dim xRatio2 As Double = (xVUpper - xVHalf - dVPosition) / (xVUpper - xVHalf)

        'Save undo
        'For xI3 = 1 To UBound(K)
        'K(xI3).Selected = True
        'Next
        'xUndo = "KZ" & vbCrLf & _
        '        sCmdKs(False) & vbCrLf & _
        '        "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
            Next
            xValueL = xBPM
            xI2 = xI1

            'Below Half
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVHalf Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio1 <= 655359999, Notes(xI1).Value * xRatio1, 655359999)
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio1 + xVLower
            Next
            xValueM = xBPM
            xI2 = xI1

            'Above Half
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition > xVUpper Then Exit For
                If Notes(xI1).ColumnIndex = niBPM Then
                    xBPM = Notes(xI1).Value
                    Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                End If
                Notes(xI1).VPosition = (Notes(xI1).VPosition - xVHalf) * xRatio2 + xVHalf + dVPosition
            Next
            xValueU = xBPM
            xI2 = xI1

            'Above Selection
            'For xI1 = xI2 To UBound(K)
            '    K(xI1).VPosition += (xRatio - 1) * (xVUpper - xVLower)
            'Next

            'Add BPMs
            AddNote(xVLower, niBPM, IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999), False, False, , , False)
            AddNote(xVHalf + dVPosition, niBPM, IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999), False, False, , , False)
            AddNote(xVUpper, niBPM, xValueU, False, False, , , False)

        Else
            Dim xAddBPML As Boolean = True
            Dim xAddBPMM As Boolean = True
            Dim xAddBPMU As Boolean = True

            'Modify notes
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition <= xVLower Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueL = Notes(xI1).Value
                        xValueM = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVLower Then
                            xAddBPML = False
                            Notes(xI1).Value = IIf(xVHalf <> xVLower AndAlso Notes(xI1).Value * xRatio1 <= 655359999, Notes(xI1).Value * xRatio1, 655359999)
                        End If
                    End If

                    'If longnote then adjust length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                    ElseIf xEnd > xVHalf Then
                        Notes(xI1).Length = (xEnd - xVHalf) * xRatio2 + xVHalf + dVPosition - Notes(xI1).VPosition
                    ElseIf xEnd > xVLower Then
                        Notes(xI1).Length = (xEnd - xVLower) * xRatio1 + xVLower - Notes(xI1).VPosition
                    End If

                ElseIf Notes(xI1).VPosition <= xVHalf Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueM = Notes(xI1).Value
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVHalf Then
                            xAddBPMM = False
                            Notes(xI1).Value = IIf(xVHalf <> xVUpper AndAlso Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                        Else
                            Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio1 <= 655359999, Notes(xI1).Value * xRatio1, 655359999)
                        End If
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                        Notes(xI1).Length = xEnd - xVLower - (Notes(xI1).VPosition - xVLower) * xRatio1
                    ElseIf xEnd > xVHalf Then
                        Notes(xI1).Length = (xVHalf - Notes(xI1).VPosition) * xRatio1 + (xEnd - xVHalf) * xRatio2
                    Else
                        Notes(xI1).Length *= xRatio1
                    End If

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVLower) * xRatio1 + xVLower

                ElseIf Notes(xI1).VPosition <= xVUpper Then
                    'check BPM
                    If Notes(xI1).ColumnIndex = niBPM Then
                        xValueU = Notes(xI1).Value
                        If Notes(xI1).VPosition = xVUpper Then xAddBPMU = False Else Notes(xI1).Value = IIf(Notes(xI1).Value * xRatio2 <= 655359999, Notes(xI1).Value * xRatio2, 655359999)
                    End If

                    'Adjust Length
                    Dim xEnd As Double = Notes(xI1).VPosition + Notes(xI1).Length
                    If xEnd > xVUpper Then
                        Notes(xI1).Length = (xVUpper - Notes(xI1).VPosition) * xRatio2 + xEnd - xVUpper
                    Else
                        Notes(xI1).Length *= xRatio2
                    End If

                    'Adjust VPosition
                    Notes(xI1).VPosition = (Notes(xI1).VPosition - xVHalf) * xRatio2 + xVHalf + dVPosition

                    'Else
                    '    K(xI1).VPosition += (xVUpper - xVLower) * (xRatio - 1)
                End If
            Next

            'Add BPMs
            If xAddBPML Then AddNote(xVLower, niBPM, IIf(xVHalf <> xVLower AndAlso xValueL * xRatio1 <= 655359999, xValueL * xRatio1, 655359999), False, False, , , False)
            If xAddBPMM Then AddNote(xVHalf + dVPosition, niBPM, IIf(xVHalf <> xVUpper AndAlso xValueM * xRatio2 <= 655359999, xValueM * xRatio2, 655359999), False, False, , , False)
            If xAddBPMU Then AddNote(xVUpper, niBPM, xValueU, False, False, , , False)
        End If

        'Check BPM Overflow
        For xI3 = 1 To UBound(Notes)
            If Notes(xI3).ColumnIndex = niBPM Then
                If Notes(xI3).Value > 655359999 Then Notes(xI3).Value = 655359999
                If Notes(xI3).Value < 1 Then Notes(xI3).Value = 1
            End If
        Next

        'Restore selection
        'If vSelLength < 0 Then vSelStart += (xRatio - 1) * (xVUpper - xVLower)
        'vSelLength = vSelLength * xRatio
        Dim pSelHalf As Double = vSelHalf
        vSelHalf += dVPosition
        ValidateSelection()
        Me.RedoChangeTimeSelection(vSelStart, vSelLength, pSelHalf, vSelStart, vSelStart, vSelHalf, True, xUndo, xRedo)

        Me.RedoAddNoteAll(False, xUndo, xRedo)


        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

EndofSub:
        If bAddUndo Then AddUndo(xUndo, xBaseRedo.Next, bOverWriteUndo)
    End Sub

    Private Sub BPMChangeByValue(ByVal xValue As Integer)
        'Dim xUndo As String = vbCrLf
        'Dim xRedo As String = vbCrLf
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If vSelLength = 0 Then Return

        Dim xVLower As Double = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        Dim xVHalf As Double = vSelStart + vSelHalf
        Dim xVUpper As Double = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If xVHalf = xVUpper Then xVHalf = xVLower
        'If dVPosition + xVHalf <= xVLower Or dVPosition + xVHalf >= xVUpper Then GoTo EndofSub

        If xVLower < 0 Then xVLower = 0
        If xVUpper >= VPosition1000() Then xVUpper = VPosition1000() - 1
        If xVHalf > xVUpper Then xVHalf = xVUpper
        If xVHalf < xVLower Then xVHalf = xVLower

        Dim xBPM As Integer = Notes(0).Value
        Dim xI1 As Integer
        Dim xI2 As Integer
        Dim xI3 As Integer

        Dim xConstBPM As Double = 0

        'Below Selection
        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).VPosition > xVLower Then Exit For
            If Notes(xI1).ColumnIndex = niBPM Then xBPM = Notes(xI1).Value
        Next
        xI2 = xI1
        Dim xVPos() As Double = {xVLower}
        Dim xVal() As Integer = {xBPM}

        'Within Selection
        Dim xU As Integer = 0
        For xI1 = xI2 To UBound(Notes)
            If Notes(xI1).VPosition > xVUpper Then Exit For

            If Notes(xI1).ColumnIndex = niBPM Then
                xU = UBound(xVPos) + 1
                ReDim Preserve xVPos(xU)
                ReDim Preserve xVal(xU)
                xVPos(xU) = Notes(xI1).VPosition
                xVal(xU) = Notes(xI1).Value
            End If
        Next
        ReDim Preserve xVPos(xU + 1)
        xVPos(xU + 1) = xVUpper

        'Calculate Constant BPM
        For xI1 = 0 To xU
            xConstBPM += (xVPos(xI1 + 1) - xVPos(xI1)) / xVal(xI1)
        Next
        xConstBPM = (xVUpper - xVLower) / xConstBPM

        'Compare BPM        '(xVHalf - xVLower) / xValue + (xVUpper - xVHalf) / xResult = (xVUpper - xVLower) / xConstBPM
        If (xVUpper - xVLower) / xConstBPM < (xVHalf - xVLower) / xValue Then _
            MsgBox("Please enter a value that is at least " & ((xVHalf - xVLower) * xConstBPM / (xVUpper - xVLower) / 10000) & ".", MsgBoxStyle.Critical, Strings.Messages.Err) : Return
        Dim xResult As Integer
        Dim xTempDivider As Double = xConstBPM * xVHalf - xConstBPM * xVLower - xValue * xVUpper + xValue * xVLower
        Dim xTemp001 As Double = (xVHalf - xVUpper) * xValue * xConstBPM / xTempDivider

        xResult = (xVHalf - xVUpper) * xValue * xConstBPM / xTempDivider

        'Save undo
        'For xI3 = 1 To UBound(K)
        ' K(xI3).Selected = True
        ' Next
        ' xUndo = "KZ" & vbCrLf & _
        '         sCmdKs(False) & vbCrLf & _
        '         "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        'Adjust note
        If Not NTInput Then
            'Below Selection
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition > xVLower Then Exit For
            Next
            xI2 = xI1
            If xI2 > UBound(Notes) Then GoTo EndOfAdjustment

            'Within Selection
            Dim xTempTime As Double
            Dim xTempVPos As Double
            For xI1 = xI2 To UBound(Notes)
                If Notes(xI1).VPosition >= xVUpper Then Exit For
                xTempTime = 0

                xTempVPos = Notes(xI1).VPosition
                For xI3 = 0 To xU
                    If xTempVPos < xVPos(xI3 + 1) Then Exit For
                    xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                Next
                xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                    Notes(xI1).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                Else
                    Notes(xI1).VPosition = xTempTime * xValue + xVLower
                End If
            Next

        Else
            Dim xTempTime As Double
            Dim xTempVPos As Double
            Dim xTempEnd As Double

            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).Length Then xTempEnd = Notes(xI1).VPosition + Notes(xI1).Length

                If Notes(xI1).VPosition > xVLower And Notes(xI1).VPosition < xVUpper Then
                    xTempTime = 0

                    xTempVPos = Notes(xI1).VPosition
                    For xI3 = 0 To xU
                        If xTempVPos < xVPos(xI3 + 1) Then Exit For
                        xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                    Next
                    xTempTime += (xTempVPos - xVPos(xI3)) / xVal(xI3)

                    If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                        Notes(xI1).VPosition = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf
                    Else
                        Notes(xI1).VPosition = xTempTime * xValue + xVLower
                    End If
                End If

                If Notes(xI1).Length Then
                    If xTempEnd > xVLower And xTempEnd < xVUpper Then
                        xTempTime = 0

                        For xI3 = 0 To xU
                            If xTempEnd < xVPos(xI3 + 1) Then Exit For
                            xTempTime += (xVPos(xI3 + 1) - xVPos(xI3)) / xVal(xI3)
                        Next
                        xTempTime += (xTempEnd - xVPos(xI3)) / xVal(xI3)

                        If xTempTime - (xVHalf - xVLower) / xValue > 0 Then
                            Notes(xI1).Length = (xTempTime - (xVHalf - xVLower) / xValue) * xResult + xVHalf - Notes(xI1).VPosition
                        Else
                            Notes(xI1).Length = xTempTime * xValue + xVLower - Notes(xI1).VPosition
                        End If

                    Else
                        Notes(xI1).Length = xTempEnd - Notes(xI1).VPosition
                    End If
                End If

            Next
        End If

EndOfAdjustment:

        'Delete BPMs
        xI1 = 1
        Do While xI1 <= UBound(Notes)
            If Notes(xI1).VPosition > xVUpper Then Exit Do
            If Notes(xI1).VPosition >= xVLower And Notes(xI1).ColumnIndex = niBPM Then
                For xI3 = xI1 + 1 To UBound(Notes)
                    Notes(xI3 - 1) = Notes(xI3)
                Next
                ReDim Preserve Notes(UBound(Notes) - 1)
            Else
                xI1 += 1
            End If
        Loop

        'Add BPMs
        ReDim Preserve Notes(UBound(Notes) + 2)
        With Notes(UBound(Notes) - 1)
            .ColumnIndex = niBPM
            .VPosition = xVHalf
            .Value = xResult
        End With
        With Notes(UBound(Notes))
            .ColumnIndex = niBPM
            .VPosition = xVUpper
            .Value = xVal(xU)
        End With
        If xVLower <> xVHalf Then
            ReDim Preserve Notes(UBound(Notes) + 1)
            With Notes(UBound(Notes))
                .ColumnIndex = niBPM
                .VPosition = xVLower
                .Value = xValue
            End With
        End If

        'Save redo
        'For xI3 = 1 To UBound(K)
        '    K(xI3).Selected = True
        'Next
        'xRedo = "KZ" & vbCrLf & _
        '                      sCmdKs(False) & vbCrLf & _
        '                      "SA_" & vSelStart & "_" & vSelLength & "_" & vSelHalf & "_1"
        Me.RedoAddNoteAll(False, xUndo, xRedo)

        'Restore note selection
        xVLower = IIf(vSelLength > 0, vSelStart, vSelStart + vSelLength)
        xVUpper = IIf(vSelLength < 0, vSelStart, vSelStart + vSelLength)
        If Not NTInput Then
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition >= xVLower And Notes(xI3).VPosition < xVUpper And nEnabled(Notes(xI3).ColumnIndex)
            Next
        Else
            For xI3 = 1 To UBound(Notes)
                Notes(xI3).Selected = Notes(xI3).VPosition < xVUpper And Notes(xI3).VPosition + Notes(xI3).Length >= xVLower And nEnabled(Notes(xI3).ColumnIndex)
            Next
        End If

        'EndofSub:
        AddUndo(xUndo, xBaseRedo.Next)
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

    Private Function FindNoteIndex(ByVal nColumnIndex As Integer, ByVal nVposition As Double, ByVal nValue As Long, ByVal nLongNote As Double, ByVal nHidden As Boolean) As Integer
        Dim xI1 As Integer
        If NTInput Then
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).equalsNT(nColumnIndex, nVposition, nValue, nLongNote, nHidden) Then Return xI1
            Next
        Else
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).equalsBMSE(nColumnIndex, nVposition, nValue, nLongNote, nHidden) Then Return xI1
            Next
        End If
        Return xI1
    End Function




    Private Function sIA() As Integer
        Return IIf(sI > 98, 0, sI + 1)
    End Function

    Private Function sIM() As Integer
        Return IIf(sI < 1, 99, sI - 1)
    End Function



    Private Sub TBUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBUndo.Click, mnUndo.Click
        KMouseOver = -1
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        If sUndo(sI).ofType = UndoRedo.opNoOperation Then Exit Sub
        Operate(sUndo(sI))
        sI = sIM()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
    End Sub

    Private Sub TBRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBRedo.Click, mnRedo.Click
        KMouseOver = -1
        'KMouseDown = -1
        ReDim SelectedNotes(-1)
        If sRedo(sIA).ofType = UndoRedo.opNoOperation Then Exit Sub
        Operate(sRedo(sIA))
        sI = sIA()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
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
    '... 'Me.RedoRemoveNote(K(xI1), True, xUndo, xRedo)
    'AddUndo(xUndo, xRedo)

    'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
    'Dim xRedo As New UndoRedo.Void
    'Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo
    '... 'Me.RedoRemoveNote(K(xI1), True, xUndo, xRedo)
    'AddUndo(xUndo, xBaseRedo.Next)



    Private Sub TBAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAbout.Click, mnAbout1.Click
        Dim Aboutboxx1 As New AboutBox1()
        'If My.Computer.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\About.png") Then
        Aboutboxx1.bBitmap = My.Resources.About0
        'Aboutboxx1.SelectBitmap()
        Aboutboxx1.ClientSize = New Size(1000, 500)
        Aboutboxx1.ClickToCopy.Visible = True
        Aboutboxx1.ShowDialog(Me)
        'Else
        '    MsgBox(locale.Messages.cannotfind & " ""About.png""", MsgBoxStyle.Critical, locale.Messages.err)
        'End If
    End Sub

    Private Sub TBOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBVOptions.Click, mnVOptions.Click

        Dim xDiag As New OpVisual(vo, column, LWAV.Font)
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
                Do While currWavIndex <= 1294 AndAlso hWAV(currWavIndex + 1) <> ""
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

        'Dim xI2 As Integer = 0
        For xI1 As Integer = 0 To UBound(xPath)
            'If xI2 > UBound(xIndices) Then Exit For
            'hWAV(xIndices(xI2) + 1) = GetFileName(xPath(xI1))
            'LWAV.Items.Item(xIndices(xI2)) = C10to36(xIndices(xI2) + 1) & ": " & GetFileName(xPath(xI1))
            hWAV(xIndices(xI1) + 1) = GetFileName(xPath(xI1))
            LWAV.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": " & GetFileName(xPath(xI1))
            'xI2 += 1
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To IIf(UBound(xIndices) < UBound(xPath), UBound(xIndices), UBound(xPath))
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        If IsSaved Then SetIsSaved(False)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragDrop
        ReDim DDFileName(-1)
        If Not e.Data.GetDataPresent(DataFormats.FileDrop) Then Return

        Dim xOrigPath() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        Dim xPath() As String = FilterFileBySupported(xOrigPath, SupportedAudioExtension)
        If xPath.Length = 0 Then
            RefreshPanelAll()
            Exit Sub
        End If

        AddToPOWAV(xPath)
    End Sub

    Private Sub POWAV_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles POWAV.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            DDFileName = FilterFileBySupported(CType(e.Data.GetData(DataFormats.FileDrop), String()), SupportedAudioExtension)
        Else
            e.Effect = DragDropEffects.None
        End If
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.DragLeave
        ReDim DDFileName(-1)
        RefreshPanelAll()
    End Sub

    Private Sub POWAV_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles POWAV.Resize
        LWAV.Height = sender.Height - 25
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
        ReDim SelectedNotes(-1)
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
        ReDim SelectedNotes(-1)
        Dim xK(0) As Note
        xK(0) = Notes(0)

        For xI1 As Integer = 1 To UBound(Notes)
            ReDim Preserve xK(UBound(xK) + 1)
            With xK(UBound(xK))
                .ColumnIndex = Notes(xI1).ColumnIndex
                .LongNote = Notes(xI1).Length > 0
                .Value = Notes(xI1).Value
                .VPosition = Notes(xI1).VPosition
                .Selected = Notes(xI1).Selected
                .Hidden = Notes(xI1).Hidden
            End With

            If Notes(xI1).Length > 0 Then
                ReDim Preserve xK(UBound(xK) + 1)
                With xK(UBound(xK))
                    .ColumnIndex = Notes(xI1).ColumnIndex
                    .LongNote = True
                    .Value = Notes(xI1).Value
                    .VPosition = Notes(xI1).VPosition + Notes(xI1).Length
                    .Selected = Notes(xI1).Selected
                    .Hidden = Notes(xI1).Hidden
                End With
            End If
        Next

        Notes = xK

        SortByVPositionInsertion()
        UpdatePairing()
        CalculateTotalPlayableNotes()
    End Sub

    Private Sub TBNTInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBNTInput.Click, mnNTInput.Click
        'Dim xUndo As String = "NT_" & CInt(NTInput) & "_0" & vbCrLf & "KZ" & vbCrLf & sCmdKsAll(False)
        'Dim xRedo As String = "NT_" & CInt(Not NTInput) & "_1"
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Me.RedoRemoveNoteAll(False, xUndo, xRedo)

        NTInput = sender.Checked

        TBNTInput.Checked = NTInput
        mnNTInput.Checked = NTInput
        POBLong.Enabled = Not NTInput
        POBLongShort.Enabled = Not NTInput

        bAdjustLength = False
        bAdjustUpper = False

        Me.RedoNT(NTInput, False, xUndo, xRedo)
        If NTInput Then
            ConvertBMSE2NT()
        Else
            ConvertNT2BMSE()
        End If
        Me.RedoAddNoteAll(False, xUndo, xRedo)

        AddUndo(xUndo, xBaseRedo.Next)
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
        For xI1 As Integer = cmnLanguage.Items.Count - 1 To 3 Step -1
            Try
                cmnLanguage.Items.RemoveAt(xI1)
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
        column(0).Left = 0
        'If col(0).Width = 0 Then col(0).Visible = False

        For xI1 As Integer = 1 To UBound(column)
            column(xI1).Left = column(xI1 - 1).Left + IIf(column(xI1 - 1).isVisible, column(xI1 - 1).Width, 0)
            'If col(xI1).Width = 0 Then col(xI1).Visible = False
        Next
        HSL.Maximum = nLeft(gColumns) + column(niB).Width
        HS.Maximum = nLeft(gColumns) + column(niB).Width
        HSR.Maximum = nLeft(gColumns) + column(niB).Width
    End Sub

    Private Sub CHPlayer_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHPlayer.SelectedIndexChanged
        If CHPlayer.SelectedIndex = -1 Then CHPlayer.SelectedIndex = 0

        iPlayer = CHPlayer.SelectedIndex
        Dim xGP2 As Boolean = iPlayer <> 0
        column(niD1).isVisible = xGP2
        column(niD2).isVisible = xGP2
        column(niD3).isVisible = xGP2
        column(niD4).isVisible = xGP2
        column(niD5).isVisible = xGP2
        column(niD6).isVisible = xGP2
        column(niD7).isVisible = xGP2
        column(niD8).isVisible = xGP2
        column(niS3).isVisible = xGP2

        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()

        If IsInitializing Then Exit Sub
        RefreshPanelAll()
    End Sub

    Private Sub CGB_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGB.ValueChanged
        gColumns = niB + CGB.Value - 1
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub TBGOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBGOptions.Click, mnGOptions.Click
        Dim xTE As Integer
        Select Case UCase(EncodingToString(TextEncoding)) ' az: wow seriously? is there really no better way? 
            Case "ANSI" : xTE = 0
            Case "UNICODE" : xTE = 1
            Case "ASCII" : xTE = 2
            Case "BIGENDIAN" : xTE = 3
            Case "UTF32" : xTE = 4
            Case "UTF7" : xTE = 5
            Case "UTF8" : xTE = 6
        End Select

        Dim xDiag As New OpGeneral(gWheel, gPgUpDn, MiddleButtonMoveMethod, xTE, 192.0R / BMSGridLimit,
            AutoSaveInterval, BeepWhileSaved, BPMx1296, STOPx1296,
            AutoFocusMouseEnter, FirstClickDisabled, ClickStopPreview)

        If xDiag.ShowDialog() = Windows.Forms.DialogResult.OK Then
            With xDiag
                gWheel = .zWheel
                gPgUpDn = .zPgUpDn
                TextEncoding = .zEncoding
                'SortingMethod = .zSort
                MiddleButtonMoveMethod = .zMiddle
                AutoSaveInterval = .zAutoSave
                BMSGridLimit = 192.0R / .zGridPartition
                BeepWhileSaved = .cBeep.Checked
                BPMx1296 = .cBpm1296.Checked
                STOPx1296 = .cStop1296.Checked
                AutoFocusMouseEnter = .cMEnterFocus.Checked
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

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, True, True, xUndo, xRedo)
            Notes(xI1).LongNote = True
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBShort.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        If Not NTInput Then
            For xI1 As Integer = 1 To UBound(Notes)
                If Not Notes(xI1).Selected Then Continue For

                Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, 0, True, xUndo, xRedo)
                Notes(xI1).LongNote = False
            Next

        Else
            For xI1 As Integer = 1 To UBound(Notes)
                If Not Notes(xI1).Selected Then Continue For

                Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, 0, True, xUndo, xRedo)
                Notes(xI1).Length = 0
            Next
        End If

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBNormalLong_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBLongShort.Click
        If NTInput Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Not Notes(xI1).LongNote, True, xUndo, xRedo)
            Notes(xI1).LongNote = Not Notes(xI1).LongNote
        Next

        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHidden_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHidden.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), True, True, xUndo, xRedo)
            Notes(xI1).Hidden = True
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), False, True, xUndo, xRedo)
            Notes(xI1).Hidden = False
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBHiddenVisible_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles POBHiddenVisible.Click
        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        For xI1 As Integer = 1 To UBound(Notes)
            If Not Notes(xI1).Selected Then Continue For

            Me.RedoHiddenNoteModify(Notes(xI1), Not Notes(xI1).Hidden, True, xUndo, xRedo)
            Notes(xI1).Hidden = Not Notes(xI1).Hidden
        Next
        AddUndo(xUndo, xBaseRedo.Next)
        SortByVPositionInsertion()
        UpdatePairing()
        RefreshPanelAll()
    End Sub

    Private Sub POBModify_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles POBModify.Click
        Dim xNum As Boolean = False
        Dim xLbl As Boolean = False
        Dim xI1 As Integer

        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected AndAlso isColumnNumeric(Notes(xI1).ColumnIndex) Then xNum = True : Exit For
        Next
        For xI1 = 1 To UBound(Notes)
            If Notes(xI1).Selected AndAlso Not isColumnNumeric(Notes(xI1).ColumnIndex) Then xLbl = True : Exit For
        Next
        If Not (xNum Or xLbl) Then Exit Sub

        If xNum Then
            Dim xD1 As Double = Val(InputBox(Strings.Messages.PromptEnterNumeric, Text)) * 10000
            If Not xD1 = 0 Then
                If xD1 <= 0 Then xD1 = 1

                Dim xUndo As UndoRedo.LinkedURCmd = Nothing
                Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
                Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

                For xI1 = 1 To UBound(Notes)
                    If Not isColumnNumeric(Notes(xI1).ColumnIndex) Then Continue For
                    If Not Notes(xI1).Selected Then Continue For

                    Me.RedoRelabelNote(Notes(xI1), xD1, True, xUndo, xRedo)
                    Notes(xI1).Value = xD1
                Next
                AddUndo(xUndo, xBaseRedo.Next)
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

            For xI1 = 1 To UBound(Notes)
                If isColumnNumeric(Notes(xI1).ColumnIndex) Then Continue For
                If Not Notes(xI1).Selected Then Continue For

                Me.RedoRelabelNote(Notes(xI1), xVal, True, xUndo, xRedo)
                Notes(xI1).Value = xVal
            Next
            AddUndo(xUndo, xBaseRedo.Next)
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
        Dim xDiag As New diagFind(gColumns, Strings.Messages.Err, Strings.Messages.InvalidLabel)
        xDiag.Show()
    End Sub

    Private Function fdrCheck(ByVal xNote As Note) As Boolean
        Return xNote.VPosition >= MeasureBottom(fdriMesL) And xNote.VPosition < MeasureBottom(fdriMesU) + MeasureLength(fdriMesU) AndAlso
               IIf(isColumnNumeric(xNote.ColumnIndex),
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
        For xI1 As Integer = 1 To UBound(Notes)
            xSel(xI1) = Notes(xI1).Selected
        Next

        'Main process
        For xI1 As Integer = 1 To UBound(Notes)
            Dim bbba As Boolean = xbSel And xSel(xI1)
            Dim bbbb As Boolean = xbUnsel And Not xSel(xI1)
            Dim bbbc As Boolean = nEnabled(Notes(xI1).ColumnIndex)
            Dim bbbd As Boolean = fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote))
            Dim bbbe As Boolean = fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden)
            Dim bbbf As Boolean = fdrCheck(Notes(xI1))

            If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                    nEnabled(Notes(xI1).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden) Then
                Notes(xI1).Selected = fdrCheck(Notes(xI1))
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
        For xI1 As Integer = 1 To UBound(Notes)
            xSel(xI1) = Notes(xI1).Selected
        Next

        'Main process
        For xI1 As Integer = 1 To UBound(Notes)
            If ((xbSel And xSel(xI1)) Or (xbUnsel And Not xSel(xI1))) AndAlso
                        nEnabled(Notes(xI1).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden) Then
                Notes(xI1).Selected = Not fdrCheck(Notes(xI1))
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
        Dim xI1 As Integer = 1
        Do While xI1 <= UBound(Notes)
            If ((xbSel And Notes(xI1).Selected) Or (xbUnsel And Not Notes(xI1).Selected)) AndAlso
                        fdrCheck(Notes(xI1)) AndAlso nEnabled(Notes(xI1).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden) Then
                'xUndo &= sCmdK(K(xI1).ColumnIndex, K(xI1).VPosition, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, True) & vbCrLf
                'xRedo &= sCmdKD(K(xI1).ColumnIndex, K(xI1).VPosition, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden) & vbCrLf
                Me.RedoRemoveNote(Notes(xI1), True, xUndo, xRedo)
                RemoveNote(xI1, False)
            Else
                xI1 += 1
            End If
        Loop

        AddUndo(xUndo, xBaseRedo.Next)
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
        For xI1 As Integer = 1 To UBound(Notes)
            If ((xbSel And Notes(xI1).Selected) Or (xbUnsel And Not Notes(xI1).Selected)) AndAlso
                    fdrCheck(Notes(xI1)) AndAlso nEnabled(Notes(xI1).ColumnIndex) And Not isColumnNumeric(Notes(xI1).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden) Then
                'xUndo &= sCmdKC(K(xI1).ColumnIndex, K(xI1).VPosition, xxLbl, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, 0, 0, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, True) & vbCrLf
                'xRedo &= sCmdKC(K(xI1).ColumnIndex, K(xI1).VPosition, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, 0, 0, xxLbl, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, True) & vbCrLf
                Me.RedoRelabelNote(Notes(xI1), xxLbl, True, xUndo, xRedo)
                Notes(xI1).Value = xxLbl
            End If
        Next

        AddUndo(xUndo, xBaseRedo.Next)
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
        For xI1 As Integer = 1 To UBound(Notes)
            If ((xbSel And Notes(xI1).Selected) Or (xbUnsel And Not Notes(xI1).Selected)) AndAlso
                    fdrCheck(Notes(xI1)) AndAlso nEnabled(Notes(xI1).ColumnIndex) And isColumnNumeric(Notes(xI1).ColumnIndex) AndAlso fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(xI1).Length, Notes(xI1).LongNote)) And fdrRangeS(xbVisible, xbHidden, Notes(xI1).Hidden) Then
                'xUndo &= sCmdKC(K(xI1).ColumnIndex, K(xI1).VPosition, xReplaceVal, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, 0, 0, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, True) & vbCrLf
                'xRedo &= sCmdKC(K(xI1).ColumnIndex, K(xI1).VPosition, K(xI1).Value, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, 0, 0, xReplaceVal, IIf(NTInput, K(xI1).Length, K(xI1).LongNote), K(xI1).Hidden, True) & vbCrLf
                Me.RedoRelabelNote(Notes(xI1), xReplaceVal, True, xUndo, xRedo)
                Notes(xI1).Value = xReplaceVal
            End If
        Next

        AddUndo(xUndo, xBaseRedo.Next)
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
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) >= 999 Then
                    Me.RedoRemoveNote(Notes(xI1), True, xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            Dim xdVP As Double
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP And Notes(xI1).VPosition + Notes(xI1).Length <= MeasureBottom(999) Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition + xMLength, True, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength

                ElseIf Notes(xI1).VPosition >= xVP Then
                    xdVP = MeasureBottom(999) - 1 - Notes(xI1).VPosition - Notes(xI1).Length
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition + xMLength, Notes(xI1).Length + xdVP, True, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength
                    Notes(xI1).Length += xdVP

                ElseIf Notes(xI1).VPosition + Notes(xI1).Length >= xVP Then
                    xdVP = IIf(Notes(xI1).VPosition + Notes(xI1).Length > MeasureBottom(999) - 1, VPosition1000() - 1 - Notes(xI1).VPosition - Notes(xI1).Length, xMLength)
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Notes(xI1).Length + xdVP, True, xUndo, xRedo)
                    Notes(xI1).Length += xdVP
                End If
            Next

        Else
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) >= 999 Then
                    Me.RedoRemoveNote(Notes(xI1), True, xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition + xMLength, True, xUndo, xRedo)
                    Notes(xI1).VPosition += xMLength
                End If
            Next
        End If

        For xI1 As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(xI1) = MeasureLength(xI1 - 1)
        Next
        UpdateMeasureBottom()

        AddUndo(xUndo, xBaseRedo.Next)
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
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) = xMeasure And MeasureAtDisplacement(Notes(xI1).VPosition + Notes(xI1).Length) = xMeasure Then
                    Me.RedoRemoveNote(Notes(xI1), True, xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            Dim xdVP As Double
            xVP = MeasureBottom(xMeasure)
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP + xMLength Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition - xMLength, True, xUndo, xRedo)
                    Notes(xI1).VPosition -= xMLength

                ElseIf Notes(xI1).VPosition >= xVP Then
                    xdVP = xMLength + xVP - Notes(xI1).VPosition
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition + xdVP - xMLength, Notes(xI1).Length - xdVP, True, xUndo, xRedo)
                    Notes(xI1).VPosition += xdVP - xMLength
                    Notes(xI1).Length -= xdVP

                ElseIf Notes(xI1).VPosition + Notes(xI1).Length >= xVP Then
                    xdVP = IIf(Notes(xI1).VPosition + Notes(xI1).Length >= xVP + xMLength, xMLength, Notes(xI1).VPosition + Notes(xI1).Length - xVP + 1)
                    Me.RedoLongNoteModify(Notes(xI1), Notes(xI1).VPosition, Notes(xI1).Length - xdVP, True, xUndo, xRedo)
                    Notes(xI1).Length -= xdVP
                End If
            Next

        Else
            Dim xI1 As Integer = 1
            Do While xI1 <= UBound(Notes)
                If MeasureAtDisplacement(Notes(xI1).VPosition) = xMeasure Then
                    Me.RedoRemoveNote(Notes(xI1), True, xUndo, xRedo)
                    RemoveNote(xI1, False)
                Else
                    xI1 += 1
                End If
            Loop

            xVP = MeasureBottom(xMeasure)
            For xI1 = 1 To UBound(Notes)
                If Notes(xI1).VPosition >= xVP Then
                    Me.RedoMoveNote(Notes(xI1), Notes(xI1).ColumnIndex, Notes(xI1).VPosition - xMLength, True, xUndo, xRedo)
                    Notes(xI1).VPosition -= xMLength
                End If
            Next
        End If

        For xI1 As Integer = 999 To xMeasure + 1 Step -1
            MeasureLength(xI1 - 1) = MeasureLength(xI1)
        Next
        UpdateMeasureBottom()

        AddUndo(xUndo, xBaseRedo.Next)
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
        System.IO.File.Delete(xTempFileName)

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
        For xI1 As Integer = cmnTheme.Items.Count - 1 To 5 Step -1
            Try
                cmnTheme.Items.RemoveAt(xI1)
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
        Dim xMax As Double = IIf(vSelLength > 0, VPosition1000() - vSelLength, VPosition1000)
        Dim xMin As Double = IIf(vSelLength < 0, -vSelLength, 0)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelStart)
        If xDouble = Double.PositiveInfinity Then Return

        vSelStart = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSL.Click
        Dim xMax As Double = VPosition1000() - vSelStart
        Dim xMin As Double = -vSelStart
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelLength)
        If xDouble = Double.PositiveInfinity Then Return

        vSelLength = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub FSSH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FSSH.Click
        Dim xMax As Double = IIf(vSelLength > 0, vSelLength, 0)
        Dim xMin As Double = IIf(vSelLength > 0, 0, -vSelLength)
        Dim xDouble As Double = InputBoxDouble("Please enter a number between " & xMin & " and " & xMax & ".", xMin, xMax, , vSelHalf)
        If xDouble = Double.PositiveInfinity Then Return

        vSelHalf = xDouble
        ValidateSelection()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub BVCReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BVCReverse.Click
        vSelStart += vSelLength
        vSelHalf -= vSelLength
        vSelLength *= -1
        ValidateSelection()
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
        For xI1 As Integer = xS To 1294
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hWAV(xI1 + 1)
                hWAV(xI1 + 1) = hWAV(xI1)
                hWAV(xI1) = xStr

                LWAV.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hWAV(xI1 + 1)
                LWAV.Items.Item(xI1 - 1) = C10to36(xI1) & ": " & hWAV(xI1)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(xI1)
                Dim xL2 As String = C10to36(xI1 + 1)
                For xI2 As Integer = 1 To UBound(Notes)
                    If isColumnNumeric(Notes(xI2).ColumnIndex) Then Continue For

                    If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                        Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, True, xUndo, xRedo)
                        Notes(xI2).Value = xI1 * 10000 + 10000

                    ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                        Me.RedoRelabelNote(Notes(xI2), xI1 * 10000, True, xUndo, xRedo)
                        Notes(xI2).Value = xI1 * 10000

                    End If
                Next

1100:           xIndices(xIndex) += -1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
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
        For xI1 As Integer = xS To 0 Step -1
            xIndex = Array.IndexOf(xIndices, xI1)
            If xIndex <> -1 Then
                xStr = hWAV(xI1 + 1)
                hWAV(xI1 + 1) = hWAV(xI1 + 2)
                hWAV(xI1 + 2) = xStr

                LWAV.Items.Item(xI1) = C10to36(xI1 + 1) & ": " & hWAV(xI1 + 1)
                LWAV.Items.Item(xI1 + 1) = C10to36(xI1 + 2) & ": " & hWAV(xI1 + 2)

                If Not WAVChangeLabel Then GoTo 1100

                Dim xL1 As String = C10to36(xI1 + 2)
                Dim xL2 As String = C10to36(xI1 + 1)
                For xI2 As Integer = 1 To UBound(Notes)
                    If isColumnNumeric(Notes(xI2).ColumnIndex) Then Continue For

                    If C10to36(Notes(xI2).Value \ 10000) = xL1 Then
                        Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 10000, True, xUndo, xRedo)
                        Notes(xI2).Value = xI1 * 10000 + 10000

                    ElseIf C10to36(Notes(xI2).Value \ 10000) = xL2 Then
                        Me.RedoRelabelNote(Notes(xI2), xI1 * 10000 + 20000, True, xUndo, xRedo)
                        Notes(xI2).Value = xI1 * 10000 + 20000

                    End If
                Next

1100:           xIndices(xIndex) += 1
            End If
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
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
        For xI1 As Integer = 0 To UBound(xIndices)
            hWAV(xIndices(xI1) + 1) = ""
            LWAV.Items.Item(xIndices(xI1)) = C10to36(xIndices(xI1) + 1) & ": "
        Next

        LWAV.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LWAV.SelectedIndices.Add(xIndices(xI1))
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
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = nEnabled(Notes(xI1).ColumnIndex)
        Next
        If TBTimeSelect.Checked Then
            CalculateGreatestVPosition()
            vSelStart = 0
            vSelLength = MeasureBottom(MeasureAtDisplacement(GreatestVPosition)) + MeasureLength(MeasureAtDisplacement(GreatestVPosition))
        End If
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnDelete.Click
        If Not (PMainIn.Focused OrElse PMainInL.Focused Or PMainInR.Focused) Then Exit Sub

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        Me.RedoRemoveNoteSelected(True, xUndo, xRedo)
        RemoveNotes(True)

        AddUndo(xUndo, xBaseRedo.Next)
        CalculateGreatestVPosition()
        CalculateTotalPlayableNotes()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub mnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnUpdate.Click
        Process.Start("http://www.cs.mcgill.ca/~ryang6/iBMSC/")
    End Sub

    Private Sub mnUpdateC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnUpdateC.Click
        Process.Start("http://bbs.rohome.net/thread-1074065-1-1.html")
    End Sub

    Private Sub mnQuit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnQuit.Click
        Close()
    End Sub


    Private Sub EnableDWM()
        mnMain.BackColor = Color.Black
        'TBMain.BackColor = Color.FromArgb(64, 64, 64)

        For Each xmn As ToolStripMenuItem In mnMain.Items
            xmn.ForeColor = Color.White
            AddHandler xmn.DropDownClosed, AddressOf mn_DropDownClosed
            AddHandler xmn.DropDownOpened, AddressOf mn_DropDownOpened
            AddHandler xmn.MouseEnter, AddressOf mn_MouseEnter
            AddHandler xmn.MouseLeave, AddressOf mn_MouseLeave
        Next
    End Sub

    Private Sub DisableDWM()
        mnMain.BackColor = SystemColors.Control
        'TBMain.BackColor = SystemColors.Control

        For Each xmn As ToolStripMenuItem In mnMain.Items
            xmn.ForeColor = SystemColors.ControlText
            RemoveHandler xmn.DropDownClosed, AddressOf mn_DropDownClosed
            RemoveHandler xmn.DropDownOpened, AddressOf mn_DropDownOpened
            RemoveHandler xmn.MouseEnter, AddressOf mn_MouseEnter
            RemoveHandler xmn.MouseLeave, AddressOf mn_MouseLeave
        Next
    End Sub

    Private Sub ttlIcon_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ttlIcon.MouseDown
        ttlIcon.Image = My.Resources.icon2_16
        'mnSys.Show(ttlIcon, 0, ttlIcon.Height)
    End Sub
    Private Sub ttlIcon_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ttlIcon.MouseEnter
        ttlIcon.Image = My.Resources.icon2_16_highlight
    End Sub
    Private Sub ttlIcon_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ttlIcon.MouseLeave
        ttlIcon.Image = My.Resources.icon2_16
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
        gShowGrid = CGShow.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowS.CheckedChanged
        gShowSubGrid = CGShowS.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowBG_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowBG.CheckedChanged
        gShowBG = CGShowBG.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowM.CheckedChanged
        gShowMeasureNumber = CGShowM.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowV_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowV.CheckedChanged
        gShowVerticalLine = CGShowV.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowMB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowMB.CheckedChanged
        gShowMeasureBar = CGShowMB.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGShowC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGShowC.CheckedChanged
        gShowC = CGShowC.Checked
        RefreshPanelAll()
    End Sub
    Private Sub CGBLP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBLP.CheckedChanged
        gDisplayBGAColumn = CGBLP.Checked

        column(niBGA).isVisible = gDisplayBGAColumn
        column(niLAYER).isVisible = gDisplayBGAColumn
        column(niPOOR).isVisible = gDisplayBGAColumn
        column(niS4).isVisible = gDisplayBGAColumn

        If IsInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub
    Private Sub CGSTOP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGSTOP.CheckedChanged
        gSTOP = CGSTOP.Checked

        column(niSTOP).isVisible = gSTOP

        If IsInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub
    Private Sub CGBPM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGBPM.CheckedChanged
        'Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        'Dim xRedo As UndoRedo.LinkedURCmd = Nothing
        'Me.RedoChangeVisibleColumns(gBLP, gSTOP, iPlayer, gBLP, CGSTOP.Checked, iPlayer, xUndo, xRedo)
        gBPM = CGBPM.Checked

        column(niBPM).isVisible = gBPM

        If IsInitializing Then Exit Sub
        For xI1 As Integer = 1 To UBound(Notes)
            Notes(xI1).Selected = Notes(xI1).Selected And nEnabled(Notes(xI1).ColumnIndex)
        Next
        'AddUndo(xUndo, xRedo)
        UpdateColumnsX()
        RefreshPanelAll()
    End Sub

    Private Sub CGDisableVertical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CGDisableVertical.CheckedChanged
        DisableVerticalMove = CGDisableVertical.Checked
    End Sub

    Private Sub CBeatPreserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBeatPreserve.Click, CBeatMeasure.Click, CBeatCut.Click, CBeatScale.Click
        'If Not sender.Checked Then Exit Sub
        Dim xBeatList() As RadioButton = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
        BeatChangeMode = Array.IndexOf(Of RadioButton)(xBeatList, sender)
        'For xI1 As Integer = 0 To mnBeat.Items.Count - 1
        'If xI1 <> BeatChangeMode Then CType(mnBeat.Items(xI1), ToolStripMenuItem).Checked = False
        'Next
        'sender.Checked = True
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


        For Each xI1 As Integer In xIndices
            Dim dLength As Double = xRatio * 192.0R - MeasureLength(xI1)
            Dim dRatio As Double = xRatio * 192.0R / MeasureLength(xI1)

            Dim xBottom As Double = 0
            For xI2 As Integer = 0 To xI1 - 1
                xBottom += MeasureLength(xI2)
            Next
            Dim xUpBefore As Double = xBottom + MeasureLength(xI1)
            Dim xUpAfter As Double = xUpBefore + dLength

            Select Case BeatChangeMode
                Case 1
case2:              Dim xI0 As Integer

                    If NTInput Then
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                            If Notes(xI0).VPosition + Notes(xI0).Length >= xUpBefore Then
                                Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, False, xUndo, xRedo)
                                Notes(xI0).Length += dLength
                            End If
                        Next
                    Else
                        For xI0 = 1 To UBound(Notes)
                            If Notes(xI0).VPosition >= xUpBefore Then Exit For
                        Next
                    End If

                    For xI9 As Integer = xI0 To UBound(Notes)
                        Me.RedoLongNoteModify(Notes(xI9), Notes(xI9).VPosition + dLength, Notes(xI9).Length, False, xUndo, xRedo)
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
                                        Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, False, xUndo, xRedo)
                                        Notes(xI0).Length = nLen
                                    End If
                                ElseIf Notes(xI0).VPosition < xUpBefore Then
                                    If Notes(xI0).VPosition + Notes(xI0).Length < xUpBefore Then
                                        Me.RedoRemoveNote(Notes(xI0), False, xUndo, xRedo)
                                        RemoveNote(xI0)
                                        xI0 -= 1
                                        xU -= 1
                                    Else
                                        Dim nLen As Double = Notes(xI0).Length - xUpBefore + Notes(xI0).VPosition
                                        Me.RedoLongNoteModify(Notes(xI0), xUpBefore, nLen, False, xUndo, xRedo)
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
                                Me.RedoRemoveNote(Notes(xI8), False, xUndo, xRedo)
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
                                    Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, Notes(xI0).Length + dLength, False, xUndo, xRedo)
                                    Notes(xI0).Length += dLength
                                ElseIf Notes(xI0).VPosition + Notes(xI0).Length > xBottom Then
                                    Dim nLen As Double = (Notes(xI0).Length + Notes(xI0).VPosition - xBottom) * dRatio + xBottom - Notes(xI0).VPosition
                                    Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition, nLen, False, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                End If
                            ElseIf Notes(xI0).VPosition < xUpBefore Then
                                If Notes(xI0).VPosition + Notes(xI0).Length > xUpBefore Then
                                    Dim nLen As Double = (xUpBefore - Notes(xI0).VPosition) * dRatio + Notes(xI0).VPosition + Notes(xI0).Length - xUpBefore
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Me.RedoLongNoteModify(Notes(xI0), nVPos, nLen, False, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                Else
                                    Dim nLen As Double = Notes(xI0).Length * dRatio
                                    Dim nVPos As Double = (Notes(xI0).VPosition - xBottom) * dRatio + xBottom
                                    Me.RedoLongNoteModify(Notes(xI0), nVPos, nLen, False, xUndo, xRedo)
                                    Notes(xI0).Length = nLen
                                    Notes(xI0).VPosition = nVPos
                                End If
                            Else
                                Me.RedoLongNoteModify(Notes(xI0), Notes(xI0).VPosition + dLength, Notes(xI0).Length, False, xUndo, xRedo)
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
                            Me.RedoLongNoteModify(Notes(xI0), nVP, Notes(xI0).Length, False, xUndo, xRedo)
                            Notes(xI8).VPosition = nVP
                        Next

                        'GoTo case2

                        For xI8 As Integer = xI9 To UBound(Notes)
                            Me.RedoLongNoteModify(Notes(xI8), Notes(xI8).VPosition + dLength, Notes(xI8).Length, False, xUndo, xRedo)
                            Notes(xI8).VPosition += dLength
                        Next
                    End If

            End Select

            MeasureLength(xI1) = xRatio * 192.0R
            LBeat.Items(xI1) = Add3Zeros(xI1) & ": " & xDisplay
        Next
        UpdateMeasureBottom()
        'xUndo &= vbCrLf & xUndo2
        'xRedo &= vbCrLf & xRedo2

        LBeat.SelectedIndices.Clear()
        For xI1 As Integer = 0 To UBound(xIndices)
            LBeat.SelectedIndices.Add(xIndices(xI1))
        Next

        AddUndo(xUndo, xBaseRedo.Next)
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
        xDiag.Filter = Strings.FileType._image & "|*.bmp;*.png;*.jpg;*.gif|" &
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
    POBeatSwitch.CheckedChanged,
    POExpansionSwitch.CheckedChanged

        Try
            Dim Source As CheckBox = CType(sender, CheckBox)
            Dim Target As Panel = Nothing

            If Object.ReferenceEquals(sender, Nothing) Then : Exit Sub
            ElseIf Object.ReferenceEquals(sender, POHeaderSwitch) Then : Target = POHeaderInner
            ElseIf Object.ReferenceEquals(sender, POGridSwitch) Then : Target = POGridInner
            ElseIf Object.ReferenceEquals(sender, POWaveFormSwitch) Then : Target = POWaveFormInner
            ElseIf Object.ReferenceEquals(sender, POWAVSwitch) Then : Target = POWAVInner
            ElseIf Object.ReferenceEquals(sender, POBeatSwitch) Then : Target = POBeatInner
            ElseIf Object.ReferenceEquals(sender, POExpansionSwitch) Then : Target = POExpansionInner
            End If

            If Source.Checked Then
                Target.Visible = True
            Else
                Target.Visible = False
            End If

        Catch ex As Exception

        End Try
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

    Private Sub VerticalResizer_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseDown, POBeatResizer.MouseDown, POExpansionResizer.MouseDown
        tempResize = e.Y
    End Sub

    Private Sub HorizontalResizer_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POptionsResizer.MouseDown, SpL.MouseDown, SpR.MouseDown
        tempResize = e.X
    End Sub

    Private Sub POResizer_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles POWAVResizer.MouseMove, POBeatResizer.MouseMove, POExpansionResizer.MouseMove
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

    Private Sub LWAV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LWAV.SelectedIndexChanged

    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click

    End Sub

    Private Sub tBeatValue_TextChanged(sender As Object, e As EventArgs) Handles tBeatValue.TextChanged

    End Sub

    Private Sub mnSys_Click(sender As Object, e As EventArgs) Handles mnSys.Click

    End Sub
End Class
