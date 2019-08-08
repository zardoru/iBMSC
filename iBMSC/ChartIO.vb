Imports System.Linq
Imports System.Text
Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub OpenBms(xStrAll As String)
        State.Mouse.CurrentHoveredNoteIndex = - 1

        'Line feed validation: will remove some empty lines
        xStrAll = Replace(Replace(Replace(xStrAll, vbLf, vbCr), vbCr & vbCr, vbCr), vbCr, vbCrLf)

        Dim xStrLine() As String = Split(xStrAll, vbCrLf, , CompareMethod.Text)
        Dim i As Integer
        Dim sLine As String
        Dim xExpansion = ""
        ReDim Notes(0)
        ReDim _mColumn(999)
        ReDim BmsWAV(1295)
        ReDim BmsBPM(1295)    'x10000
        ReDim BmsSTOP(1295)
        ReDim BmsSCROLL(1295)
        InitializeNewBMS()
        InitializeOpenBMS()

        Notes(0) = New Note
        With Notes(0)
            .ColumnIndex = ColumnType.BPM
            .VPosition = - 1
            .Value = 1200000
        End With

        'random, setRandom      0
        'endRandom              0
        'if             +1
        'else           0
        'endif          -1
        'switch, setSwitch      +1
        'case, skip, def        0
        'endSw                  -1
        Dim xStack = 0

        For Each sLine In xStrLine
            Dim sLineTrim As String = sLine.Trim
            If xStack > 0 Then GoTo Expansion

            If sLineTrim.StartsWith("#") And Mid(sLineTrim, 5, 3) = "02:" Then
                Dim xIndex As Integer = Val(Mid(sLineTrim, 2, 3))
                Dim xRatio As Double = Val(Mid(sLineTrim, 8))
                Dim xxD As Long = GetDenominator(xRatio)
                MeasureLength(xIndex) = xRatio*192.0R
                LBeat.Items(xIndex) = Add3Zeros(xIndex) & ": " & xRatio &
                                      IIf(xxD > 10000, "", " ( " & CLng(xRatio*xxD) & " / " & xxD & " ) ")

            ElseIf sLineTrim.StartsWith("#WAV", StringComparison.CurrentCultureIgnoreCase) Then
                BmsWAV(C36to10(Mid(sLineTrim, Len("#WAV") + 1, 2))) = Mid(sLineTrim, Len("#WAV") + 4)

            ElseIf sLineTrim.StartsWith("#BMP", StringComparison.CurrentCultureIgnoreCase) Then
                BmsBMP(C36to10(Mid(sLineTrim, Len("#BMP") + 1, 2))) = Mid(sLineTrim, Len("#BMP") + 4)

            ElseIf _
                sLineTrim.StartsWith("#BPM", StringComparison.CurrentCultureIgnoreCase) And
                Not Mid(sLineTrim, Len("#BPM") + 1, 1).Trim = "" Then 'If BPM##
                ' zdr: No limits on BPM editing.. they don't make much sense.
                BmsBPM(C36to10(Mid(sLineTrim, Len("#BPM") + 1, 2))) = Val(Mid(sLineTrim, Len("#BPM") + 4))*10000

                'No limits on STOPs either.
            ElseIf sLineTrim.StartsWith("#STOP", StringComparison.CurrentCultureIgnoreCase) Then
                BmsSTOP(C36to10(Mid(sLineTrim, Len("#STOP") + 1, 2))) = Val(Mid(sLineTrim, Len("#STOP") + 4))*10000

            ElseIf sLineTrim.StartsWith("#SCROLL", StringComparison.CurrentCultureIgnoreCase) Then
                BmsSCROLL(C36to10(Mid(sLineTrim, Len("#SCROLL") + 1, 2))) = Val(Mid(sLineTrim, Len("#SCROLL") + 4))*
                                                                            10000


            ElseIf sLineTrim.StartsWith("#TITLE", StringComparison.CurrentCultureIgnoreCase) Then
                THTitle.Text = Mid(sLineTrim, Len("#TITLE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#ARTIST", StringComparison.CurrentCultureIgnoreCase) Then
                THArtist.Text = Mid(sLineTrim, Len("#ARTIST") + 1).Trim

            ElseIf sLineTrim.StartsWith("#GENRE", StringComparison.CurrentCultureIgnoreCase) Then
                THGenre.Text = Mid(sLineTrim, Len("#GENRE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#BPM", StringComparison.CurrentCultureIgnoreCase) Then 'If BPM ####
                Notes(0).Value = Val(Mid(sLineTrim, Len("#BPM") + 1).Trim)*10000
                THBPM.Value = Notes(0).Value/10000

            ElseIf sLineTrim.StartsWith("#PLAYER", StringComparison.CurrentCultureIgnoreCase) Then
                Dim xInt As Integer = Val(Mid(sLineTrim, Len("#PLAYER") + 1).Trim)
                If xInt >= 1 And xInt <= 4 Then _
                    CHPlayer.SelectedIndex = xInt - 1

            ElseIf sLineTrim.StartsWith("#RANK", StringComparison.CurrentCultureIgnoreCase) Then
                Dim xInt As Integer = Val(Mid(sLineTrim, Len("#RANK") + 1).Trim)
                If xInt >= 0 And xInt <= 4 Then _
                    CHRank.SelectedIndex = xInt

            ElseIf sLineTrim.StartsWith("#PLAYLEVEL", StringComparison.CurrentCultureIgnoreCase) Then
                THPlayLevel.Text = Mid(sLineTrim, Len("#PLAYLEVEL") + 1).Trim


            ElseIf sLineTrim.StartsWith("#SUBTITLE", StringComparison.CurrentCultureIgnoreCase) Then
                THSubTitle.Text = Mid(sLineTrim, Len("#SUBTITLE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#SUBARTIST", StringComparison.CurrentCultureIgnoreCase) Then
                THSubArtist.Text = Mid(sLineTrim, Len("#SUBARTIST") + 1).Trim

            ElseIf sLineTrim.StartsWith("#STAGEFILE", StringComparison.CurrentCultureIgnoreCase) Then
                THStageFile.Text = Mid(sLineTrim, Len("#STAGEFILE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#BANNER", StringComparison.CurrentCultureIgnoreCase) Then
                THBanner.Text = Mid(sLineTrim, Len("#BANNER") + 1).Trim

            ElseIf sLineTrim.StartsWith("#BACKBMP", StringComparison.CurrentCultureIgnoreCase) Then
                THBackBMP.Text = Mid(sLineTrim, Len("#BACKBMP") + 1).Trim

            ElseIf sLineTrim.StartsWith("#DIFFICULTY", StringComparison.CurrentCultureIgnoreCase) Then
                Try
                    CHDifficulty.SelectedIndex = Integer.Parse(Mid(sLineTrim, Len("#DIFFICULTY") + 1).Trim)
                Catch ex As Exception
                End Try

            ElseIf sLineTrim.StartsWith("#DEFEXRANK", StringComparison.CurrentCultureIgnoreCase) Then
                THExRank.Text = Mid(sLineTrim, Len("#DEFEXRANK") + 1).Trim

            ElseIf sLineTrim.StartsWith("#TOTAL", StringComparison.CurrentCultureIgnoreCase) Then
                Dim xStr As String = Mid(sLineTrim, Len("#TOTAL") + 1).Trim
                'If xStr.EndsWith("%") Then xStr = Mid(xStr, 1, Len(xStr) - 1)
                THTotal.Text = xStr

            ElseIf sLineTrim.StartsWith("#COMMENT", StringComparison.CurrentCultureIgnoreCase) Then
                Dim xStr As String = Mid(sLineTrim, Len("#COMMENT") + 1).Trim
                If xStr.StartsWith("""") Then xStr = Mid(xStr, 2)
                If xStr.EndsWith("""") Then xStr = Mid(xStr, 1, Len(xStr) - 1)
                THComment.Text = xStr

            ElseIf sLineTrim.StartsWith("#LNTYPE", StringComparison.CurrentCultureIgnoreCase) Then
                'THLnType.Text = Mid(sLineTrim, Len("#LNTYPE") + 1).Trim
                If Val(Mid(sLineTrim, Len("#LNTYPE") + 1).Trim) = 1 Then CHLnObj.SelectedIndex = 0

            ElseIf sLineTrim.StartsWith("#LNOBJ", StringComparison.CurrentCultureIgnoreCase) Then
                Dim xValue As Integer = C36to10(Mid(sLineTrim, Len("#LNOBJ") + 1).Trim)
                CHLnObj.SelectedIndex = xValue

                'TODO: LNOBJ value validation

                'ElseIf sLineTrim.StartsWith("#LNTYPE", StringComparison.CurrentCultureIgnoreCase) Then
                '    CAdLNTYPE.Checked = True
                '    If Mid(sLineTrim, 9) = "" Or Mid(sLineTrim, 9) = "1" Or Mid(sLineTrim, 9) = "01" Then CAdLNTYPEb.Text = "1"
                '    CAdLNTYPEb.Text = Mid(sLineTrim, 9)

            ElseIf sLineTrim.StartsWith("#") And Mid(sLineTrim, 7, 1) = ":" Then 'If the line contains Ks
                Dim channel As String = Mid(sLineTrim, 5, 2)
                If Columns.BMSChannelToColumn(channel) = 0 Then GoTo AddExpansion

            Else
                Expansion: If sLineTrim.StartsWith("#IF", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1
                    GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#ENDIF", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack -= 1
                    GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#SWITCH", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1
                    GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#SETSWITCH", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1
                    GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#ENDSW", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack -= 1
                    GoTo AddExpansion

                ElseIf sLineTrim.StartsWith("#") Then
                    AddExpansion: xExpansion &= sLine & vbCrLf
                End If

            End If
        Next

        UpdateMeasureBottom()

        xStack = 0
        For Each sLine In xStrLine
            Dim sLineTrim As String = sLine.Trim
            If xStack > 0 Then Continue For

            If Not (sLineTrim.StartsWith("#") And Mid(sLineTrim, 7, 1) = ":") Then Continue For 'If the line contains Ks

            ' >> Measure =           Mid(sLine, 2, 3)
            ' >> Column Identifier = Mid(sLine, 5, 2)
            ' >> K =                 Mid(sLine, i, 2)
            Dim xMeasure As Integer = Val(Mid(sLineTrim, 2, 3))
            Dim channel As String = Mid(sLineTrim, 5, 2)
            If Columns.BMSChannelToColumn(Channel) = 0 Then Continue For

            If Channel = "01" Then _mColumn(xMeasure) += 1 'If the identifier is 01 then add a B column in that measure
            For i = 8 To Len(sLineTrim) - 1 Step 2 'For all Ks within that line ( - 1 can be ommitted )
                If Mid(sLineTrim, i, 2) = "00" Then Continue For 'If the K is not 00

                Notes = Notes.Concat({New Note}).ToArray()

                With Notes(UBound(Notes))
                    .ColumnIndex = Columns.BMSChannelToColumn(Channel) +
                                   IIf(Channel = "01", 1, 0)*(_mColumn(xMeasure) - 1)
                    .LongNote = IsChannelLongNote(Channel)
                    .Hidden = IsChannelHidden(Channel)
                    .Landmine = IsChannelLandmine(Channel)
                    .Selected = False
                    .VPosition = MeasureBottom(xMeasure) + MeasureLength(xMeasure)*(i/2 - 4)/((Len(sLineTrim) - 7)/2)
                    .Value = C36to10(Mid(sLineTrim, i, 2))*10000

                    If Channel = "03" Then .Value = Convert.ToInt32(Mid(sLineTrim, i, 2), 16)*10000
                    If Channel = "08" Then .Value = BmsBPM(C36to10(Mid(sLineTrim, i, 2)))
                    If Channel = "09" Then .Value = BmsSTOP(C36to10(Mid(sLineTrim, i, 2)))
                    If Channel = "SC" Then .Value = BmsSCROLL(C36to10(Mid(sLineTrim, i, 2)))
                End With

            Next
        Next

        If NTInput Then ConvertBMSE2NT()

        LWAV.Visible = False
        LWAV.Items.Clear()
        LBMP.Visible = False
        LBMP.Items.Clear()
        For i = 1 To 1295
            LWAV.Items.Add(C10to36(i) & ": " & BmsWAV(i))
            LBMP.Items.Add(C10to36(i) & ": " & BmsBMP(i))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True
        LBMP.SelectedIndex = 0
        LBMP.Visible = True

        TExpansion.Text = xExpansion

        ValidateNotesArray()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    ReadOnly _bmsChannelList() As String = {"01", "03", "04", "06", "07", "08", "09",
                                           "11", "12", "13", "14", "15", "16", "18", "19",
                                           "21", "22", "23", "24", "25", "26", "28", "29",
                                           "31", "32", "33", "34", "35", "36", "38", "39",
                                           "41", "42", "43", "44", "45", "46", "48", "49",
                                           "51", "52", "53", "54", "55", "56", "58", "59",
                                           "61", "62", "63", "64", "65", "66", "68", "69",
                                           "D1", "D2", "D3", "D4", "D5", "D6", "D8", "D9",
                                           "E1", "E2", "E3", "E4", "E5", "E6", "E8", "E9",
                                           "SC"}
    ' 71 through 89 are reserved
    '"71", "72", "73", "74", "75", "76", "78", "79",
    '"81", "82", "83", "84", "85", "86", "88", "89",


    Private Function SaveBms() As String
        ValidateNotesArray()

        Dim hasOverlapping = False

        ' We regenerate these when traversing the bms event list.
        ReDim BmsBPM(0)
        ReDim BmsSTOP(0)
        ReDim BmsSCROLL(0)

        Dim restoreNtInput As Boolean = NtInput
        Dim xKBackUp() As Note = Notes
        If restoreNtInput Then
            NtInput = False
            ConvertNt2Bmse()
        End If

        Dim xprevNotes(-1) As Note  'Notes too close to the next measure

        Dim measureString = CreateBmsMeasureData(hasOverlapping, xprevNotes)

        ' Warn about 255 limit if neccesary.
        If hasOverlapping Then MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                                      Strings.Messages.NoteOverlapError & vbCrLf &
                                      Strings.Messages.SavedFileWillContainErrors, MsgBoxStyle.Exclamation)
        If UBound(BmsBPM) > IIf(_bpMx1296, 1295, 255) Then
            MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                   Strings.Messages.BPMOverflowError & UBound(BmsBPM) &
                   " > " & IIf(_bpMx1296, 1295, 255) & vbCrLf &
                   Strings.Messages.SavedFileWillContainErrors,
                   MsgBoxStyle.Exclamation)
        End If
        If UBound(BmsSTOP) > IIf(_stoPx1296, 1295, 255) Then
            MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                   Strings.Messages.STOPOverflowError & UBound(BmsSTOP) &
                   " > " & IIf(_stoPx1296, 1295, 255) & vbCrLf &
                   Strings.Messages.SavedFileWillContainErrors,
                   MsgBoxStyle.Exclamation)
        End If
        If UBound(BmsSCROLL) > 1295 Then
            MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                   Strings.Messages.SCROLLOverflowError & UBound(BmsSCROLL) & " > " & 1295 &
                   vbCrLf &
                   Strings.Messages.SavedFileWillContainErrors, MsgBoxStyle.Exclamation)
        End If

        ' Add expansion text
        Dim expansionText = String.Format("\r\n*---------------------- EXPANSION FIELD\r\n{0}\r\n", TExpansion.Text)
        If TExpansion.Text = "" Then expansionText = ""

        ' Output main data field.
        Dim dataText = String.Format("\r\n*---------------------- MAIN DATA FIELD\r\n{0}\r\n", measureString)

        If restoreNtInput Then
            Notes = xKBackUp
            NtInput = True
        End If

        ' Generate headers now, since we have the unique BPM/STOP/etc declarations.
        Dim headerText As String = GenerateHeaderMeta()
        headerText &= GenerateHeaderIndexedData()

        Dim outputBms As String = Join({headerText, expansionText, dataText}, vbCrLf)
        Return outputBms
    End Function

    Private Function CreateBmsMeasureData(ByRef hasOverlapping As Boolean, ByRef xprevNotes() As Note) As String
        Dim measureString(MeasureAtDisplacement(GreatestVPosition) + 1) As String
        Dim measureIndex As Integer

        For measureIndex = 0 To MeasureAtDisplacement(GreatestVPosition) + 1 'For i in each measure
            Dim consistentDecimalStr = WriteDecimalWithDot(MeasureLength(measureIndex) / 192.0R)

            ' Handle fractional measure
            If MeasureLength(measureIndex) <> 192.0R Then _
                measureString(measureIndex) &= String.Format("#{0}02:{1}\r\n", Add3Zeros(measureIndex), consistentDecimalStr)

            ' Get note count in current measure
            Dim lowerLimit As Integer = Nothing
            Dim upperLimit As Integer = Nothing
            GetMeasureLimits(measureIndex, lowerLimit, upperLimit)

            If upperLimit - lowerLimit = 0 Then
                Continue For 'If there is no K in the current measure then end this loop
            End If

            ' Get notes from this measure
            Dim xUPrevText As Integer = UBound(xprevNotes)
            Dim notesInMeasure(upperLimit - lowerLimit + xUPrevText) As Note

            ' Copy notes from previous array
            For i = 0 To xUPrevText
                notesInMeasure(i) = xprevNotes(i)
            Next

            ' Copy notes in current measure
            For i = lowerLimit To upperLimit - 1
                notesInMeasure(i - lowerLimit + xprevNotes.Length) = Notes(i)
            Next

            ' Find greatest column.
            ' Since background tracks have the highest column values
            ' this - Columns.BGM will yield the number of B columns.
            Dim greatestColumn = (From tn In notesInMeasure Select tn.ColumnIndex).Concat({0}).Max()

            ReDim xprevNotes(-1)
            measureString(measureIndex) &= GenerateBackgroundTracks(measureIndex,
                                                                    hasOverlapping,
                                                                    notesInMeasure,
                                                                    greatestColumn,
                                                                    xprevNotes)
            measureString(measureIndex) &= GenerateKeyTracks(measureIndex, hasOverlapping, notesInMeasure, xprevNotes)
        Next

        Return Join(measureString, vbCrLf)
    End Function

    Private Function GenerateHeaderMeta() As String
        Dim xStrHeader As String = vbCrLf & "*---------------------- HEADER FIELD" & vbCrLf & vbCrLf
        xStrHeader &= "#PLAYER " & (CHPlayer.SelectedIndex + 1) & vbCrLf
        xStrHeader &= "#GENRE " & THGenre.Text & vbCrLf
        xStrHeader &= "#TITLE " & THTitle.Text & vbCrLf
        xStrHeader &= "#ARTIST " & THArtist.Text & vbCrLf
        xStrHeader &= "#BPM " & WriteDecimalWithDot(Notes(0).Value / 10000) & vbCrLf
        xStrHeader &= "#PLAYLEVEL " & THPlayLevel.Text & vbCrLf
        xStrHeader &= "#RANK " & CHRank.SelectedIndex & vbCrLf
        xStrHeader &= vbCrLf
        If THSubTitle.Text <> "" Then xStrHeader &= "#SUBTITLE " & THSubTitle.Text & vbCrLf
        If THSubArtist.Text <> "" Then xStrHeader &= "#SUBARTIST " & THSubArtist.Text & vbCrLf
        If THStageFile.Text <> "" Then xStrHeader &= "#STAGEFILE " & THStageFile.Text & vbCrLf
        If THBanner.Text <> "" Then xStrHeader &= "#BANNER " & THBanner.Text & vbCrLf
        If THBackBMP.Text <> "" Then xStrHeader &= "#BACKBMP " & THBackBMP.Text & vbCrLf
        xStrHeader &= vbCrLf
        If CHDifficulty.SelectedIndex Then xStrHeader &= "#DIFFICULTY " & CHDifficulty.SelectedIndex & vbCrLf
        If THExRank.Text <> "" Then xStrHeader &= "#DEFEXRANK " & THExRank.Text & vbCrLf
        If THTotal.Text <> "" Then xStrHeader &= "#TOTAL " & THTotal.Text & vbCrLf
        If THComment.Text <> "" Then xStrHeader &= "#COMMENT """ & THComment.Text & """" & vbCrLf
        'If THLnType.Text <> "" Then xStrHeader &= "#LNTYPE " & THLnType.Text & vbCrLf
        If CHLnObj.SelectedIndex > 0 Then xStrHeader &= "#LNOBJ " & C10to36(CHLnObj.SelectedIndex) & vbCrLf _
            Else xStrHeader &= "#LNTYPE 1" & vbCrLf
        xStrHeader &= vbCrLf
        Return xStrHeader
    End Function

    Private Function GenerateHeaderIndexedData() As String
        Dim xStrHeader = ""

        For i = 1 To UBound(BmsWAV)
            If Not BmsWAV(i) = "" Then xStrHeader &= "#WAV" & C10to36(i) &
                                                     " " & BmsWAV(i) & vbCrLf
        Next
        For i = 1 To UBound(BmsBMP)
            If Not BmsBMP(i) = "" Then xStrHeader &= "#BMP" & C10to36(i) &
                                                     " " & BmsBMP(i) & vbCrLf
        Next
        For i = 1 To UBound(BmsBPM)
            xStrHeader &= "#BPM" &
                          IIf(_bpMx1296, C10to36(i), Mid("0" & Hex(i), Len(Hex(i)))) &
                          " " & WriteDecimalWithDot(BmsBPM(i) / 10000) & vbCrLf
        Next
        For i = 1 To UBound(BmsSTOP)
            xStrHeader &= "#STOP" &
                          IIf(_stoPx1296, C10to36(i), Mid("0" & Hex(i), Len(Hex(i)))) &
                          " " & WriteDecimalWithDot(BmsSTOP(i) / 10000) & vbCrLf
        Next
        For i = 1 To UBound(BmsSCROLL)
            xStrHeader &= "#SCROLL" &
                          C10to36(i) & " " & WriteDecimalWithDot(BmsSCROLL(i) / 10000) & vbCrLf
        Next

        Return xStrHeader
    End Function

    Private Sub GetMeasureLimits(measureIndex As Integer, ByRef lowerLimit As Integer, ByRef upperLimit As Integer)
        Dim noteCount = UBound(Notes)
        lowerLimit = 0

        For i = 1 To noteCount 'Collect Ks in the same measure
            If MeasureAtDisplacement(Notes(i).VPosition) >= measureIndex Then
                lowerLimit = i
                Exit For
            End If 'Lower limit found
        Next

        upperLimit = 0

        For i = lowerLimit To noteCount
            If MeasureAtDisplacement(Notes(i).VPosition) > measureIndex Then
                upperLimit = i
                Exit For 'Upper limit found
            End If
        Next

        If upperLimit < lowerLimit Then upperLimit = noteCount + 1
    End Sub

    Private Function GenerateKeyTracks(measureIndex As Integer, ByRef hasOverlapping As Boolean,
                                       notesInMeasure() As Note, ByRef xprevNotes() As Note) As String
        Dim currentBmsChannel As String
        Dim ret = ""

        For Each currentBmsChannel In _bmsChannelList 'Start rendering other notes
            Dim relativeMeasurePos(-1) 'Ks in the same column
            Dim noteStrings(-1)      'Ks in the same column

            ' Background tracks take care of this.
            If currentBmsChannel = "01" Then Continue For


            For NoteIndex = 0 To UBound(notesInMeasure) 'Find Ks in the same column (xI4 is TK index)

                Dim currentNote As Note = notesInMeasure(NoteIndex)
                If Columns.GetBMSChannelBy(currentNote) = currentBmsChannel Then

                    ReDim Preserve relativeMeasurePos(UBound(relativeMeasurePos) + 1)
                    ReDim Preserve noteStrings(UBound(noteStrings) + 1)
                    relativeMeasurePos(UBound(relativeMeasurePos)) = currentNote.VPosition -
                                                                     MeasureBottom(
                                                                         MeasureAtDisplacement(currentNote.VPosition))
                    If relativeMeasurePos(UBound(relativeMeasurePos)) < 0 Then _
                        relativeMeasurePos(UBound(relativeMeasurePos)) = 0

                    If currentBmsChannel = "03" Then 'If integer bpm
                        noteStrings(UBound(noteStrings)) = Mid("0" & Hex(currentNote.Value \ 10000),
                                                               Len(Hex(currentNote.Value \ 10000)))
                    ElseIf currentBmsChannel = "08" Then 'If bpm requires declaration
                        Dim bpmIndex
                        For bpmIndex = 1 To UBound(BmsBPM) ' find BPM value in existing array
                            If currentNote.Value = BmsBPM(bpmIndex) Then Exit For
                        Next
                        If bpmIndex > UBound(BmsBPM) Then ' Didn't find it, add it
                            ReDim Preserve BmsBPM(UBound(BmsBPM) + 1)
                            BmsBPM(UBound(BmsBPM)) = currentNote.Value
                        End If
                        noteStrings(UBound(noteStrings)) = IIf(_bpMx1296, C10to36(bpmIndex),
                                                               Mid("0" & Hex(bpmIndex), Len(Hex(bpmIndex))))
                    ElseIf currentBmsChannel = "09" Then 'If STOP
                        Dim stopIndex
                        For stopIndex = 1 To UBound(BmsSTOP) ' find STOP value in existing array
                            If currentNote.Value = BmsSTOP(stopIndex) Then Exit For
                        Next

                        If stopIndex > UBound(BmsSTOP) Then ' Didn't find it, add it
                            ReDim Preserve BmsSTOP(UBound(BmsSTOP) + 1)
                            BmsSTOP(UBound(BmsSTOP)) = currentNote.Value
                        End If
                        noteStrings(UBound(noteStrings)) = IIf(_stoPx1296, C10to36(stopIndex),
                                                               Mid("0" & Hex(stopIndex), Len(Hex(stopIndex))))
                    ElseIf currentBmsChannel = "SC" Then 'If SCROLL
                        Dim scrollIndex
                        For scrollIndex = 1 To UBound(BmsSCROLL) ' find SCROLL value in existing array
                            If currentNote.Value = BmsSCROLL(scrollIndex) Then Exit For
                        Next

                        If scrollIndex > UBound(BmsSCROLL) Then ' Didn't find it, add it
                            ReDim Preserve BmsSCROLL(UBound(BmsSCROLL) + 1)
                            BmsSCROLL(UBound(BmsSCROLL)) = currentNote.Value
                        End If
                        noteStrings(UBound(noteStrings)) = C10to36(scrollIndex)
                    Else
                        noteStrings(UBound(noteStrings)) = C10to36(currentNote.Value \ 10000)
                    End If
                End If
            Next

            If relativeMeasurePos.Length = 0 Then Continue For

            Dim xGcd As Double = MeasureLength(measureIndex)
            For i = 0 To UBound(relativeMeasurePos) 'find greatest common divisor
                If relativeMeasurePos(i) > 0 Then xGcd = Gcd(xGcd, relativeMeasurePos(i))
            Next

            Dim xStrKey() As String
            ReDim xStrKey(CInt(MeasureLength(measureIndex) / xGcd) - 1)
            For i = 0 To UBound(xStrKey) 'assign 00 to all keys
                xStrKey(i) = "00"
            Next

            For i = 0 To UBound(relativeMeasurePos) 'assign K texts
                If CInt(relativeMeasurePos(i) / xGcd) > UBound(xStrKey) Then
                    ReDim Preserve xprevNotes(UBound(xprevNotes) + 1)
                    With xprevNotes(UBound(xprevNotes))
                        .ColumnIndex = Columns.BMSChannelToColumn(_bmsChannelList(currentBmsChannel))
                        .LongNote = IsChannelLongNote(_bmsChannelList(currentBmsChannel))
                        .Hidden = IsChannelHidden(_bmsChannelList(currentBmsChannel))
                        .VPosition = MeasureBottom(measureIndex)
                        .Value = C36to10(noteStrings(i))
                    End With
                    If _bmsChannelList(currentBmsChannel) = "08" Then _
                        xprevNotes(UBound(xprevNotes)).Value = IIf(_bpMx1296, BmsBPM(C36to10(noteStrings(i))),
                                                                   BmsBPM(Convert.ToInt32(noteStrings(i), 16)))
                    If _bmsChannelList(currentBmsChannel) = "09" Then _
                        xprevNotes(UBound(xprevNotes)).Value = IIf(_stoPx1296, BmsSTOP(C36to10(noteStrings(i))),
                                                                   BmsSTOP(Convert.ToInt32(noteStrings(i), 16)))
                    If _bmsChannelList(CurrentBMSChannel) = "SC" Then _
                        xprevNotes(UBound(xprevNotes)).Value = BmsSCROLL(C36to10(NoteStrings(i)))
                    Continue For
                End If
                If xStrKey(CInt(relativeMeasurePos(i)/xGCD)) <> "00" Then
                    hasOverlapping = True
                End If

                xStrKey(CInt(relativeMeasurePos(i)/xGCD)) = NoteStrings(i)
            Next

            Ret &= "#" & Add3Zeros(MeasureIndex) & CurrentBMSChannel & ":" & Join(xStrKey, "") & vbCrLf
        Next

        Return Ret
    End Function

    Private Function GenerateBackgroundTracks(measureIndex As Integer, ByRef hasOverlapping As Boolean,
                                              notesInMeasure() As Note, greatestColumn As Integer,
                                              ByRef xprevNotes() As Note) As String
        Dim relativeNotePositions() As Double 'Ks in the same column
        Dim ret = ""

        For ColIndex = ColumnType.BGM To greatestColumn 'Start rendering B notes (xI3 is columnindex)
            ReDim relativeNotePositions(-1) 'Ks in the same column
            Dim columnNoteStrings(-1) As String     'Ks in the same column

            For I = 0 To UBound(notesInMeasure) 'Find Ks in the same column (xI4 is TK index)
                If notesInMeasure(I).ColumnIndex = ColIndex Then

                    ReDim Preserve relativeNotePositions(UBound(relativeNotePositions) + 1)
                    ReDim Preserve columnNoteStrings(UBound(columnNoteStrings) + 1)

                    relativeNotePositions(UBound(relativeNotePositions)) = notesInMeasure(I).VPosition -
                                                                           MeasureBottom(
                                                                               MeasureAtDisplacement(
                                                                                   notesInMeasure(I).VPosition))
                    If relativeNotePositions(UBound(relativeNotePositions)) < 0 Then _
                        relativeNotePositions(UBound(relativeNotePositions)) = 0

                    columnNoteStrings(UBound(columnNoteStrings)) = C10to36(notesInMeasure(I).Value \ 10000)
                End If
            Next

            Dim xGcd As Double = MeasureLength(measureIndex)
            For i = 0 To UBound(relativeNotePositions) 'find greatest common divisor
                If relativeNotePositions(i) > 0 Then xGcd = Gcd(xGcd, relativeNotePositions(i))
            Next

            Dim notesString(CInt(MeasureLength(measureIndex) / xGcd) - 1) As String
            For i = 0 To UBound(notesString) 'assign 00 to all keys
                notesString(i) = "00"
            Next

            For i = 0 To UBound(relativeNotePositions) 'assign K texts
                If CInt(relativeNotePositions(i) / xGcd) > UBound(notesString) Then

                    ReDim Preserve xprevNotes(UBound(xprevNotes) + 1)

                    With xprevNotes(UBound(xprevNotes))
                        .ColumnIndex = ColIndex
                        .VPosition = MeasureBottom(measureIndex)
                        .Value = C36to10(columnNoteStrings(i))
                    End With

                    Continue For
                End If
                If notesString(relativeNotePositions(i) / xGcd) <> "00" Then hasOverlapping = True
                notesString(relativeNotePositions(i) / xGcd) = columnNoteStrings(i)
            Next

            ret &= String.Format("#{0}01:{1}\r\n", Add3Zeros(measureIndex), Join(notesString, ""))
        Next

        Return ret
    End Function

    Private Function OpenSm(xStrAll As String) As Boolean
        State.Mouse.CurrentHoveredNoteIndex = - 1

        Dim xStrLine() As String = Split(xStrAll, vbCrLf)
        'Remove comments starting with "//"
        For i = 0 To UBound(xStrLine)
            If xStrLine(i).Contains("//") Then xStrLine(i) = Mid(xStrLine(i), 1, InStr(xStrLine(i), "//") - 1)
        Next

        xStrAll = Join(xStrLine, "")
        xStrLine = Split(xStrAll, ";")

        Dim iDiff = 0
        Dim iCurrentDiff = 0
        Dim xTempSplit() As String = Split(xStrAll, "#NOTES:")
        Dim xTempStr() As String = {}
        If xTempSplit.Length > 2 Then
            ReDim Preserve xTempStr(UBound(xTempSplit) - 1)
            For i = 1 To UBound(xTempSplit)
                xTempSplit(i) = Mid(xTempSplit(i), InStr(xTempSplit(i), ":") + 1)
                xTempSplit(i) = Mid(xTempSplit(i), InStr(xTempSplit(i), ":") + 1).Trim
                xTempStr(i - 1) = Mid(xTempSplit(i), 1, InStr(xTempSplit(i), ":") - 1)
                xTempSplit(i) = Mid(xTempSplit(i), InStr(xTempSplit(i), ":") + 1).Trim
                xTempStr(i - 1) &= " : " & Mid(xTempSplit(i), 1, InStr(xTempSplit(i), ":") - 1)
            Next

            Dim xDiag As New dgImportSM(xTempStr)
            If xDiag.ShowDialog() = DialogResult.Cancel Then Return True
            iDiff = xDiag.iResult
        End If

        Dim sL As String
        ReDim Notes(0)
        ReDim _mColumn(999)
        ReDim BmsWAV(1295)
        ReDim BmsBMP(1295)
        ReDim BmsBPM(1295)    'x10000
        ReDim BmsSTOP(1295)
        ReDim BmsSCROLL(1295)
        Me.InitializeNewBMS()

        With Notes(0)
            .ColumnIndex = ColumnType.BPM
            .VPosition = - 1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With

        For Each sL In xStrLine
            If UCase(sL).StartsWith("#TITLE:") Then
                THTitle.Text = Mid(sL, Len("#TITLE:") + 1)

            ElseIf UCase(sL).StartsWith("#SUBTITLE:") Then
                If Not UCase(sL).EndsWith("#SUBTITLE:") Then THTitle.Text &= " " & Mid(sL, Len("#SUBTITLE:") + 1)

            ElseIf UCase(sL).StartsWith("#ARTIST:") Then
                THArtist.Text = Mid(sL, Len("#ARTIST:") + 1)

            ElseIf UCase(sL).StartsWith("#GENRE:") Then
                THGenre.Text = Mid(sL, Len("#GENRE:") + 1)

            ElseIf UCase(sL).StartsWith("#BPMS:") Then
                Dim xLine As String = Mid(sL, Len("#BPMS:") + 1)
                Dim xItem() As String = Split(xLine, ",")

                Dim xVal1 As Double
                Dim xVal2 As Double

                For i = 0 To UBound(xItem)
                    xVal1 = Mid(xItem(i), 1, InStr(xItem(i), "=") - 1)
                    xVal2 = Mid(xItem(i), InStr(xItem(i), "=") + 1)

                    If xVal1 <> 0 Then
                        ReDim Preserve Notes(Notes.Length)
                        With Notes(UBound(Notes))
                            .ColumnIndex = ColumnType.BPM
                            '.LongNote = False
                            '.Hidden = False
                            '.Selected = False
                            .VPosition = xVal1*48
                            .Value = xVal2*10000
                        End With
                    Else
                        Notes(0).Value = xVal2*10000
                    End If
                Next

            ElseIf UCase(sL).StartsWith("#NOTES:") Then
                If iCurrentDiff <> iDiff Then iCurrentDiff += 1 : GoTo Jump1

                iCurrentDiff += 1
                Dim xLine As String = Mid(sL, Len("#NOTES:") + 1)
                Dim xItem() As String = Split(xLine, ":")
                For i = 0 To UBound(xItem)
                    xItem(i) = xItem(i).Trim
                Next

                If xItem.Length <> 6 Then GoTo Jump1

                THPlayLevel.Text = xItem(3)

                Dim xM() As String = Split(xItem(5), ",")
                For i = 0 To UBound(xM)
                    xM(i) = xM(i).Trim
                Next

                For i = 0 To UBound(xM)
                    For j = 0 To Len(xM(i)) - 1 Step 4
                        If xM(i)(j) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = ColumnType.A1
                                .LongNote = xM(i)(j) = "2" Or xM(i)(j) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192\(Len(xM(i))\4))*j\4 + i*192
                                .Value = 10000
                            End With
                        End If
                        If xM(i)(j + 1) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = ColumnType.A2
                                .LongNote = xM(i)(j + 1) = "2" Or xM(i)(j + 1) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192\(Len(xM(i))\4))*j\4 + i*192
                                .Value = 10000
                            End With
                        End If
                        If xM(i)(j + 2) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = ColumnType.A3
                                .LongNote = xM(i)(j + 2) = "2" Or xM(i)(j + 2) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192\(Len(xM(i))\4))*j\4 + i*192
                                .Value = 10000
                            End With
                        End If
                        If xM(i)(j + 3) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = ColumnType.A4
                                .LongNote = xM(i)(j + 3) = "2" Or xM(i)(j + 3) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192\(Len(xM(i))\4))*j\4 + i*192
                                .Value = 10000
                            End With
                        End If
                    Next
                Next
                Jump1:
            End If
        Next

        If NTInput Then ConvertBMSE2NT()

        LWAV.Visible = False
        LWAV.Items.Clear()
        LBMP.Visible = False
        LBMP.Items.Clear()
        For i = 1 To 1295
            LWAV.Items.Add(C10to36(i) & ": " & BmsWAV(i))
            LBMP.Items.Add(C10to36(i) & ": " & BmsBMP(i))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True
        LBMP.SelectedIndex = 0
        LBMP.Visible = True

        THBPM.Value = Notes(0).Value/10000
        ValidateNotesArray()

        RefreshPanelAll()
        POStatusRefresh()
        Return False
    End Function

    ''' <summary>Do not clear Undo.</summary>
    Private Sub OpeniBmsc(path As String)
        State.Mouse.CurrentHoveredNoteIndex = - 1

        Dim br As New BinaryReader(New FileStream(Path, FileMode.Open, FileAccess.Read), Encoding.Unicode)

        If br.ReadInt32 <> &H534D4269 Then GoTo EndOfSub
        If br.ReadByte <> CByte(&H43) Then GoTo EndOfSub
        Dim xMajor As Integer = br.ReadByte
        Dim xMinor As Integer = br.ReadByte
        Dim xBuild As Integer = br.ReadByte

        ClearUndo()
        ReDim Notes(0)
        ReDim _mColumn(999)
        ReDim BmsWAV(1295)
        ReDim BmsBMP(1295)
        Me.InitializeNewBMS()
        Me.InitializeOpenBMS()

        With Notes(0)
            .ColumnIndex = ColumnType.BPM
            .VPosition = - 1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With

        Do Until br.BaseStream.Position >= br.BaseStream.Length
            Dim blockId As Integer = br.ReadInt32()

            Select Case BlockID

                Case &H66657250 'Preferences
                    Dim xPref As Integer = br.ReadInt32

                    NTInput = xPref And &H1
                    TBNTInput.Checked = NTInput
                    mnNTInput.Checked = NTInput
                    POBLong.Enabled = Not NTInput
                    POBLongShort.Enabled = Not NTInput

                    ErrorCheck = xPref And &H2
                    TBErrorCheck.Checked = ErrorCheck
                    TBErrorCheck_Click(TBErrorCheck, New EventArgs)

                    _previewOnClick = xPref And &H4
                    TBPreviewOnClick.Checked = _previewOnClick
                    TBPreviewOnClick_Click(TBPreviewOnClick, New EventArgs)

                    ShowFileName = xPref And &H8
                    TBShowFileName.Checked = ShowFileName
                    TBShowFileName_Click(TBShowFileName, New EventArgs)

                    mnSMenu.Checked = xPref And &H100
                    mnSTB.Checked = xPref And &H200
                    mnSOP.Checked = xPref And &H400
                    mnSStatus.Checked = xPref And &H800
                    mnSLSplitter.Checked = xPref And &H1000
                    mnSRSplitter.Checked = xPref And &H2000

                    CGShow.Checked = xPref And &H4000
                    CGShowS.Checked = xPref And &H8000
                    CGShowBG.Checked = xPref And &H10000
                    CGShowM.Checked = xPref And &H20000
                    CGShowMB.Checked = xPref And &H40000
                    CGShowV.Checked = xPref And &H80000
                    CGShowC.Checked = xPref And &H100000
                    CGBLP.Checked = xPref And &H200000
                    CGSTOP.Checked = xPref And &H400000
                    CGSCROLL.Checked = xPref And &H20000000
                    CGBPM.Checked = xPref And &H800000

                    CGSnap.Checked = xPref And &H1000000
                    CGDisableVertical.Checked = xPref And &H2000000
                    cVSLockL.Checked = xPref And &H4000000
                    cVSLock.Checked = xPref And &H8000000
                    cVSLockR.Checked = xPref And &H10000000

                    CGDivide.Value = br.ReadInt32
                    CGSub.Value = br.ReadInt32
                    Grid.Slash = br.ReadInt32
                    CGHeight.Value = br.ReadSingle
                    CGWidth.Value = br.ReadSingle
                    CGB.Value = br.ReadInt32

                Case &H64616548 'Header
                    THTitle.Text = br.ReadString
                    THArtist.Text = br.ReadString
                    THGenre.Text = br.ReadString
                    Notes(0).Value = br.ReadInt64
                    Dim xPlayerRank As Integer = br.ReadByte
                    THPlayLevel.Text = br.ReadString

                    CHPlayer.SelectedIndex = xPlayerRank And &HF
                    CHRank.SelectedIndex = xPlayerRank >> 4

                    THSubTitle.Text = br.ReadString
                    THSubArtist.Text = br.ReadString
                    'THMaker.Text = br.ReadString
                    THStageFile.Text = br.ReadString
                    THBanner.Text = br.ReadString
                    THBackBMP.Text = br.ReadString
                    'THMidiFile.Text = br.ReadString
                    CHDifficulty.SelectedIndex = br.ReadByte
                    THExRank.Text = br.ReadString
                    THTotal.Text = br.ReadString
                    'THVolWAV.Text = br.ReadString
                    THComment.Text = br.ReadString
                    'THLnType.Text = br.ReadString
                    CHLnObj.SelectedIndex = br.ReadInt16

                Case &H564157 'WAV List
                    Dim xWavOptions As Integer = br.ReadByte
                    _wavMultiSelect = xWAVOptions And &H1
                    CWAVMultiSelect.Checked = _wavMultiSelect
                    CWAVMultiSelect_CheckedChanged(CWAVMultiSelect, New EventArgs)
                    _wavChangeLabel = xWAVOptions And &H2
                    CWAVChangeLabel.Checked = _wavChangeLabel
                    CWAVChangeLabel_CheckedChanged(CWAVChangeLabel, New EventArgs)

                    Dim xWavCount As Integer = br.ReadInt32
                    For xxi = 1 To xWAVCount
                        Dim xI As Integer = br.ReadInt16
                        BmsWAV(xI) = br.ReadString
                    Next

                Case &H504D42 'BMP List

                    Dim xBmpCount As Integer = br.ReadInt32
                    For xxi = 1 To xBMPCount
                        Dim xI As Integer = br.ReadInt16
                        BmsBMP(xI) = br.ReadString
                    Next

                Case &H74616542 'Beat
                    nBeatN.Value = br.ReadInt16
                    nBeatD.Value = br.ReadInt16
                    'nBeatD.SelectedIndex = br.ReadByte

                    Dim xBeatChangeMode As Integer = br.ReadByte
                    Dim xBeatChangeList As RadioButton() = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
                    xBeatChangeList(xBeatChangeMode).Checked = True
                    CBeatPreserve_Click(xBeatChangeList(xBeatChangeMode), New EventArgs)

                    Dim xBeatCount As Integer = br.ReadInt32
                    For xxi = 1 To xBeatCount
                        Dim xIndex As Integer = br.ReadInt16
                        MeasureLength(xIndex) = br.ReadDouble
                        Dim xRatio As Double = MeasureLength(xIndex)/192.0R
                        Dim xxD As Long = GetDenominator(xRatio)
                        LBeat.Items(xIndex) = Add3Zeros(xIndex) & ": " & xRatio &
                                              IIf(xxD > 10000, "", " ( " & CLng(xRatio*xxD) & " / " & xxD & " ) ")
                    Next

                Case &H6E707845 'Expansion Code
                    TExpansion.Text = br.ReadString

                Case &H65746F4E 'Note
                    Dim xNoteUbound As Integer = br.ReadInt32
                    ReDim Preserve Notes(xNoteUbound)
                    For i = 1 To UBound(Notes)
                        Notes(i).FromBinReader(br)
                    Next

                Case &H6F646E55 'Undo / Redo Commands
                    Dim urCount As Integer = br.ReadInt32   'Should be 100
                    sI = br.ReadInt32

                    For xI = 0 To 99
                        Dim xUndoCount As Integer = br.ReadInt32
                        Dim xBaseUndo As New UndoRedo.Void
                        Dim xIteratorUndo As UndoRedo.LinkedURCmd = xBaseUndo

                        For xxj = 1 To xUndoCount
                            Dim xByteLen As Integer = br.ReadInt32
                            Dim xByte() As Byte = br.ReadBytes(xByteLen)
                            xIteratorUndo.Next = UndoRedo.fromBytes(xByte)
                            xIteratorUndo = xIteratorUndo.Next
                        Next

                        sUndo(xI) = xBaseUndo.Next

                        Dim xRedoCount As Integer = br.ReadInt32
                        Dim xBaseRedo As New UndoRedo.Void
                        Dim xIteratorRedo As UndoRedo.LinkedURCmd = xBaseRedo
                        For xxj = 1 To xRedoCount
                            Dim xByteLen As Integer = br.ReadInt32
                            Dim xByte() As Byte = br.ReadBytes(xByteLen)
                            xIteratorRedo.Next = UndoRedo.fromBytes(xByte)
                            xIteratorRedo = xIteratorRedo.Next
                        Next
                        sRedo(xI) = xBaseRedo.Next
                    Next

            End Select
        Loop

        EndOfSub:
        br.Close()

        TBUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        TBRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation
        mnUndo.Enabled = sUndo(sI).ofType <> UndoRedo.opNoOperation
        mnRedo.Enabled = sRedo(sIA).ofType <> UndoRedo.opNoOperation

        LBMP.Visible = False
        LBMP.Items.Clear()
        LWAV.Visible = False
        LWAV.Items.Clear()
        For i = 1 To 1295
            LWAV.Items.Add(C10to36(i) & ": " & BmsWAV(i))
            LBMP.Items.Add(C10to36(i) & ": " & BmsBMP(i))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True
        LBMP.SelectedIndex = 0
        LBMP.Visible = True

        THBPM.Value = Notes(0).Value/10000
        ValidateNotesArray()
        UpdateMeasureBottom()
        CalculateTotalPlayableNotes()
        
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub SaveiBmsc(path As String)

        ValidateNotesArray()

        Try

            Dim bw As New BinaryWriter(New FileStream(Path, FileMode.Create), Encoding.Unicode)

            'bw.Write("iBMSC".ToCharArray)
            bw.Write(&H534D4269)
            bw.Write(CByte(&H43))
            bw.Write(CByte(My.Application.Info.Version.Major))
            bw.Write(CByte(My.Application.Info.Version.Minor))
            bw.Write(CByte(My.Application.Info.Version.Build))

            'Preferences
            'bw.Write("Pref".ToCharArray)
            bw.Write(&H66657250)
            Dim xPref = 0
            If NTInput Then xPref = xPref Or &H1
            If ErrorCheck Then xPref = xPref Or &H2
            If _previewOnClick Then xPref = xPref Or &H4
            If ShowFileName Then xPref = xPref Or &H8
            If mnSMenu.Checked Then xPref = xPref Or &H100
            If mnSTB.Checked Then xPref = xPref Or &H200
            If mnSOP.Checked Then xPref = xPref Or &H400
            If mnSStatus.Checked Then xPref = xPref Or &H800
            If mnSLSplitter.Checked Then xPref = xPref Or &H1000
            If mnSRSplitter.Checked Then xPref = xPref Or &H2000
            If Grid.ShowMainGrid Then xPref = xPref Or &H4000
            If Grid.ShowSubGrid Then xPref = xPref Or &H8000
            If Grid.ShowBackground Then xPref = xPref Or &H10000
            If Grid.ShowMeasureNumber Then xPref = xPref Or &H20000
            If Grid.ShowMeasureBars Then xPref = xPref Or &H40000
            If Grid.ShowVerticalLines Then xPref = xPref Or &H80000
            If Grid.ShowColumnCaptions Then xPref = xPref Or &H100000
            If Grid.ShowColumnCaptions Then xPref = xPref Or &H200000
            If Grid.ShowStopColumn Then xPref = xPref Or &H400000
            If Grid.ShowBpmColumn Then xPref = xPref Or &H800000
            If Grid.ShowScrollColumn Then xPref = xPref Or &H20000000
            If Grid.IsSnapEnabled Then xPref = xPref Or &H1000000
            If DisableVerticalMove Then xPref = xPref Or &H2000000
            If _spLock(0) Then xPref = xPref Or &H4000000
            If _spLock(1) Then xPref = xPref Or &H8000000
            If _spLock(2) Then xPref = xPref Or &H10000000
            bw.Write(xPref)
            bw.Write(BitConverter.GetBytes(Grid.ShowMainGrid))
            bw.Write(BitConverter.GetBytes(Grid.ShowSubGrid))
            bw.Write(BitConverter.GetBytes(Grid.Slash))
            bw.Write(BitConverter.GetBytes(Grid.HeightScale))
            bw.Write(BitConverter.GetBytes(Grid.WidthScale))
            bw.Write(BitConverter.GetBytes(Columns.ColumnCount))

            'Header
            'bw.Write("Head".ToCharArray)
            bw.Write(&H64616548)
            bw.Write(THTitle.Text)
            bw.Write(THArtist.Text)
            bw.Write(THGenre.Text)
            bw.Write(Notes(0).Value)
            Dim xPlayer As Integer = CHPlayer.SelectedIndex
            Dim xRank As Integer = CHRank.SelectedIndex << 4
            bw.Write(CByte(xPlayer Or xRank))
            bw.Write(THPlayLevel.Text)

            bw.Write(THSubTitle.Text)
            bw.Write(THSubArtist.Text)
            'bw.Write(THMaker.Text)
            bw.Write(THStageFile.Text)
            bw.Write(THBanner.Text)
            bw.Write(THBackBMP.Text)
            'bw.Write(THMidiFile.Text)
            bw.Write(CByte(CHDifficulty.SelectedIndex))
            bw.Write(THExRank.Text)
            bw.Write(THTotal.Text)
            'bw.Write(THVolWAV.Text)
            bw.Write(THComment.Text)
            'bw.Write(THLnType.Text)
            bw.Write(CShort(CHLnObj.SelectedIndex))

            'Wav List
            'bw.Write(("WAV" & vbNullChar).ToCharArray)
            bw.Write(&H564157)

            Dim xWavOptions = 0
            If _wavMultiSelect Then xWAVOptions = xWAVOptions Or &H1
            If _wavChangeLabel Then xWAVOptions = xWAVOptions Or &H2
            bw.Write(CByte(xWAVOptions))

            Dim xWavCount = 0
            For i = 1 To UBound(BmsWAV)
                If BmsWAV(i) <> "" Then xWAVCount += 1
            Next
            bw.Write(xWAVCount)

            For i = 1 To UBound(BmsWAV)
                If BmsWAV(i) = "" Then Continue For
                bw.Write(CShort(i))
                bw.Write(BmsWAV(i))
            Next

            'Bmp List
            'bw.Write(("BMP" & vbNullChar).ToCharArray)
            bw.Write(&H504D42)

            Dim xBmpCount = 0
            For i = 1 To UBound(BmsBMP)
                If BmsBMP(i) <> "" Then xBMPCount += 1
            Next
            bw.Write(xBMPCount)

            For i = 1 To UBound(BmsBMP)
                If BmsBMP(i) = "" Then Continue For
                bw.Write(CShort(i))
                bw.Write(BmsBMP(i))
            Next

            'Beat
            'bw.Write("Beat".ToCharArray)
            bw.Write(&H74616542)
            'Dim xNumerator As Short = nBeatN.Value
            'Dim xDenominator As Short = nBeatD.Value
            'Dim xBeatChangeMode As Byte = BeatChangeMode
            bw.Write(CShort(nBeatN.Value))
            bw.Write(CShort(nBeatD.Value))
            bw.Write(CByte(_beatChangeMode))

            Dim xBeatCount = 0
            For i = 0 To UBound(MeasureLength)
                If MeasureLength(i) <> 192.0R Then xBeatCount += 1
            Next
            bw.Write(xBeatCount)

            For i = 0 To UBound(MeasureLength)
                If MeasureLength(i) = 192.0R Then Continue For
                bw.Write(CShort(i))
                bw.Write(MeasureLength(i))
            Next

            'Expansion Code
            'bw.Write("Expn".ToCharArray)
            bw.Write(&H6E707845)
            bw.Write(TExpansion.Text)

            'Note
            'bw.Write("Note".ToCharArray)
            bw.Write(&H65746F4E)
            bw.Write(UBound(Notes))
            For i = 1 To UBound(Notes)
                Notes(i).WriteBinWriter(bw)
            Next

            'Undo / Redo Commands
            'bw.Write("Undo".ToCharArray)
            bw.Write(&H6F646E55)
            bw.Write(100)
            bw.Write(sI)

            For i = 0 To 99
                'UndoCommandsCount
                Dim countUndo = 0
                Dim pUndo As UndoRedo.LinkedURCmd = sUndo(i)
                While pUndo IsNot Nothing
                    countUndo += 1
                    pUndo = pUndo.Next
                End While
                bw.Write(countUndo)

                'UndoCommands
                pUndo = sUndo(i)
                For xxi = 1 To countUndo
                    Dim bUndo() As Byte = pUndo.toBytes
                    bw.Write(bUndo.Length)  'Length
                    bw.Write(bUndo)         'Command
                    pUndo = pUndo.Next
                Next

                'RedoCommandsCount
                Dim countRedo = 0
                Dim pRedo As UndoRedo.LinkedURCmd = sRedo(i)
                While pRedo IsNot Nothing
                    countRedo += 1
                    pRedo = pRedo.Next
                End While
                bw.Write(countRedo)

                'RedoCommands
                pRedo = sRedo(i)
                For xxi = 1 To countRedo
                    Dim bRedo() As Byte = pRedo.toBytes
                    bw.Write(bRedo.Length)
                    bw.Write(bRedo)
                    pRedo = pRedo.Next
                Next
            Next

            bw.Close()

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
    End Sub
End Class
