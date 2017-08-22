Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub OpenBMS(ByVal xStrAll As String)
        KMouseOver = -1

        'Line feed validation: will remove some empty lines
        xStrAll = Replace(Replace(Replace(xStrAll, vbLf, vbCr), vbCr & vbCr, vbCr), vbCr, vbCrLf)

        Dim xStrLine() As String = Split(xStrAll, vbCrLf, , CompareMethod.Text)
        Dim xI1 As Integer
        Dim sLine As String
        Dim xExpansion As String = ""
        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        Me.InitializeNewBMS()
        Me.InitializeOpenBMS()

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
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
        Dim xStack As Integer = 0

        For Each sLine In xStrLine
            Dim sLineTrim As String = sLine.Trim
            If xStack > 0 Then GoTo Expansion

            If sLineTrim.StartsWith("#") And Mid(sLineTrim, 5, 3) = "02:" Then
                Dim xIndex As Integer = Val(Mid(sLineTrim, 2, 3))
                Dim xRatio As Double = Val(Mid(sLineTrim, 8))
                Dim xxD As Long = GetDenominator(xRatio)
                MeasureLength(xIndex) = xRatio * 192.0R
                LBeat.Items(xIndex) = Add3Zeros(xIndex) & ": " & xRatio & IIf(xxD > 10000, "", " ( " & CLng(xRatio * xxD) & " / " & xxD & " ) ")

            ElseIf sLineTrim.StartsWith("#WAV", StringComparison.CurrentCultureIgnoreCase) Then
                hWAV(C36to10(Mid(sLineTrim, Len("#WAV") + 1, 2))) = Mid(sLineTrim, Len("#WAV") + 4)

            ElseIf sLineTrim.StartsWith("#BPM", StringComparison.CurrentCultureIgnoreCase) And Not Mid(sLineTrim, Len("#BPM") + 1, 1).Trim = "" Then  'If BPM##
                ' zdr: No limits on BPM editing.. they don't make much sense.
                hBPM(C36to10(Mid(sLineTrim, Len("#BPM") + 1, 2))) = Val(Mid(sLineTrim, Len("#BPM") + 4)) * 10000

                'No limits on STOPs either.
            ElseIf sLineTrim.StartsWith("#STOP", StringComparison.CurrentCultureIgnoreCase) Then
                hSTOP(C36to10(Mid(sLineTrim, Len("#STOP") + 1, 2))) = Val(Mid(sLineTrim, Len("#STOP") + 4)) * 10000


            ElseIf sLineTrim.StartsWith("#TITLE", StringComparison.CurrentCultureIgnoreCase) Then
                THTitle.Text = Mid(sLineTrim, Len("#TITLE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#ARTIST", StringComparison.CurrentCultureIgnoreCase) Then
                THArtist.Text = Mid(sLineTrim, Len("#ARTIST") + 1).Trim

            ElseIf sLineTrim.StartsWith("#GENRE", StringComparison.CurrentCultureIgnoreCase) Then
                THGenre.Text = Mid(sLineTrim, Len("#GENRE") + 1).Trim

            ElseIf sLineTrim.StartsWith("#BPM", StringComparison.CurrentCultureIgnoreCase) Then  'If BPM ####
                Notes(0).Value = Val(Mid(sLineTrim, Len("#BPM") + 1).Trim) * 10000
                THBPM.Value = Notes(0).Value / 10000

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

            ElseIf sLineTrim.StartsWith("#EXRANK", StringComparison.CurrentCultureIgnoreCase) Then
                THExRank.Text = Mid(sLineTrim, Len("#EXRANK") + 1).Trim

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

            ElseIf sLineTrim.StartsWith("#") And Mid(sLineTrim, 7, 1) = ":" Then   'If the line contains Ks
                Dim xIdentifier As String = Mid(sLineTrim, 5, 2)
                If IdentifiertoColumnIndex(xIdentifier) = 0 Then GoTo AddExpansion

            Else
Expansion:      If sLineTrim.StartsWith("#IF", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1 : GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#ENDIF", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack -= 1 : GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#SWITCH", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1 : GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#SETSWITCH", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack += 1 : GoTo AddExpansion
                ElseIf sLineTrim.StartsWith("#ENDSW", StringComparison.CurrentCultureIgnoreCase) Then
                    xStack -= 1 : GoTo AddExpansion

                ElseIf sLineTrim.StartsWith("#") Then
AddExpansion:       xExpansion &= sLine & vbCrLf
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
            ' >> K =                 Mid(sLine, xI1, 2)
            Dim xMeasure As Integer = Val(Mid(sLineTrim, 2, 3))
            Dim xIdentifier As String = Mid(sLineTrim, 5, 2)
            If IdentifiertoColumnIndex(xIdentifier) = 0 Then Continue For

            If xIdentifier = "01" Then mColumn(xMeasure) += 1 'If the identifier is 01 then add a B column in that measure
            For xI1 = 8 To Len(sLineTrim) - 1 Step 2   'For all Ks within that line ( - 1 can be ommitted )
                If Mid(sLineTrim, xI1, 2) = "00" Then Continue For 'If the K is not 00

                ReDim Preserve Notes(Notes.Length)
                With Notes(UBound(Notes))
                    .ColumnIndex = IdentifiertoColumnIndex(xIdentifier) +
                                        IIf(xIdentifier = "01", 1, 0) * (mColumn(xMeasure) - 1)
                    .LongNote = IdentifiertoLongNote(xIdentifier)
                    .Hidden = IdentifiertoHidden(xIdentifier)
                    .Selected = False
                    .VPosition = MeasureBottom(xMeasure) + MeasureLength(xMeasure) * (xI1 / 2 - 4) / ((Len(sLineTrim) - 7) / 2)
                    .Value = C36to10(Mid(sLineTrim, xI1, 2)) * 10000
                    If xIdentifier = "03" Then .Value = Convert.ToInt32(Mid(sLineTrim, xI1, 2), 16) * 10000
                    If xIdentifier = "08" Then .Value = hBPM(C36to10(Mid(sLineTrim, xI1, 2)))
                    If xIdentifier = "09" Then .Value = hSTOP(C36to10(Mid(sLineTrim, xI1, 2)))
                End With
            Next
        Next

        If NTInput Then ConvertBMSE2NT()

        LWAV.Visible = False
        LWAV.Items.Clear()
        For xI1 = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ": " & hWAV(xI1))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True

        TExpansion.Text = xExpansion

        SortByVPositionQuick(0, UBound(Notes))
        UpdatePairing()
        CalculateTotalNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Function SaveBMS() As String
        CalculateGreatestVPosition()
        SortByVPositionInsertion()
        UpdatePairing()
        Dim MeasureIndex As Integer
        Dim xI2 As Integer
        Dim CurrentBMSChannel As Integer
        Dim xI4 As Integer
        Dim hasOverlapping As Boolean = False
        'Dim xStrAll As String = ""   'for all 
        Dim xStrMeasure(MeasureAtDisplacement(GreatestVPosition) + 1) As String
        Dim BMSChannelList() As String = {"01", "03", "04", "06", "07", "08", "09",
                                       "11", "12", "13", "14", "15", "16", "18", "19",
                                       "21", "22", "23", "24", "25", "26", "28", "29",
                                       "31", "32", "33", "34", "35", "36", "38", "39",
                                       "41", "42", "43", "44", "45", "46", "48", "49",
                                       "51", "52", "53", "54", "55", "56", "58", "59",
                                       "61", "62", "63", "64", "65", "66", "68", "69",
                                       "71", "72", "73", "74", "75", "76", "78", "79",
                                       "81", "82", "83", "84", "85", "86", "88", "89"}
        ReDim hBPM(0)
        ReDim hSTOP(0)

        Dim xNTInput As Boolean = NTInput
        Dim xKBackUp() As Note = Notes
        If xNTInput Then
            NTInput = False
            ConvertNT2BMSE()
        End If

        'Dim xNumPlayer As String = "1"
        'For xI1 = 1 To UBound(K)
        ' If K(xI1).ColumnIndex >= niD1 And K(xI1).ColumnIndex <= niD8 Then xNumPlayer = "2" : Exit For
        'Next

        Dim xStrHeader As String = vbCrLf & "*---------------------- HEADER FIELD" & vbCrLf & vbCrLf
        xStrHeader &= "#PLAYER " & (CHPlayer.SelectedIndex + 1) & vbCrLf
        xStrHeader &= "#GENRE " & THGenre.Text & vbCrLf
        xStrHeader &= "#TITLE " & THTitle.Text & vbCrLf
        xStrHeader &= "#ARTIST " & THArtist.Text & vbCrLf
        xStrHeader &= "#BPM " & Decimalify(Notes(0).Value / 10000) & vbCrLf
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
        If THExRank.Text <> "" Then xStrHeader &= "#EXRANK " & THExRank.Text & vbCrLf
        If THTotal.Text <> "" Then xStrHeader &= "#TOTAL " & THTotal.Text & vbCrLf
        If THComment.Text <> "" Then xStrHeader &= "#COMMENT """ & THComment.Text & """" & vbCrLf
        'If THLnType.Text <> "" Then xStrHeader &= "#LNTYPE " & THLnType.Text & vbCrLf
        If CHLnObj.SelectedIndex > 0 Then xStrHeader &= "#LNOBJ " & C10to36(CHLnObj.SelectedIndex) & vbCrLf _
                                     Else xStrHeader &= "#LNTYPE 1" & vbCrLf
        xStrHeader &= vbCrLf

        Dim NotesInMeasure() As Note   'Temp K for storing Ks in the same measure
        Dim tempNote As Note     'Temp K
        Dim GreatestColumn As Integer = 0  'greatest column for B column

        Dim xprevNotes(-1) As Note  'Notes too close to the next measure

        For MeasureIndex = 0 To MeasureAtDisplacement(GreatestVPosition) + 1  'For xI1 in each measure
            xStrMeasure(MeasureIndex) = vbCrLf

            Dim consistentDecimalStr = Decimalify(MeasureLength(MeasureIndex) / 192.0R)

            If MeasureLength(MeasureIndex) <> 192.0R Then xStrMeasure(MeasureIndex) &= "#" & Add3Zeros(MeasureIndex) & "02:" & consistentDecimalStr & vbCrLf

            Dim NoteCount = UBound(Notes)
            Dim LowerLimit
            For LowerLimit = 1 To NoteCount  'Collect Ks in the same measure
                If MeasureAtDisplacement(Notes(LowerLimit).VPosition) >= MeasureIndex Then Exit For 'Lower limit found
            Next

            For NoteIndex = LowerLimit To NoteCount
                If MeasureAtDisplacement(Notes(NoteIndex).VPosition) > MeasureIndex Then Exit For 'Upper limit found
            Next
            If NoteCount - LowerLimit = 0 Then Continue For 'If there is no K in the current measure then end this loop

            'Start collecting Ks
            Dim xUPrevText As Integer = UBound(xprevNotes)
            ReDim NotesInMeasure(CurrentBMSChannel - LowerLimit + xUPrevText)
            For xI4 = 0 To xUPrevText
                NotesInMeasure(xI4) = xprevNotes(xI4)
            Next
            For xI4 = LowerLimit To CurrentBMSChannel - 1
                NotesInMeasure(xI4 - LowerLimit + xprevNotes.Length) = Notes(xI4)
            Next
            ReDim xprevNotes(-1)

            GreatestColumn = 0
            For Each tempNote In NotesInMeasure  'Find greatest column
                GreatestColumn = IIf(tempNote.ColumnIndex > GreatestColumn, tempNote.ColumnIndex, GreatestColumn)
            Next

            WriteBGM(MeasureIndex, hasOverlapping, xStrMeasure, NotesInMeasure, GreatestColumn, xprevNotes)

            For CurrentBMSChannel = 1 To UBound(BMSChannelList) 'Start rendering other notes
                Dim xVPosition(-1) 'Ks in the same column
                Dim xText(-1)      'Ks in the same column

                For xI4 = 0 To UBound(NotesInMeasure) 'Find Ks in the same column (xI4 is TK index)
                    If nIdentifier(NotesInMeasure(xI4).ColumnIndex, NotesInMeasure(xI4).Value, NotesInMeasure(xI4).LongNote, NotesInMeasure(xI4).Hidden) = BMSChannelList(CurrentBMSChannel) Then
                        ReDim Preserve xVPosition(UBound(xVPosition) + 1)
                        ReDim Preserve xText(UBound(xText) + 1)
                        xVPosition(UBound(xVPosition)) = NotesInMeasure(xI4).VPosition - MeasureBottom(MeasureAtDisplacement(NotesInMeasure(xI4).VPosition))
                        If xVPosition(UBound(xVPosition)) < 0 Then xVPosition(UBound(xVPosition)) = 0

                        If BMSChannelList(CurrentBMSChannel) = "03" Then 'If integer bpm
                            xText(UBound(xText)) = Mid("0" & Hex(NotesInMeasure(xI4).Value \ 10000), Len(Hex(NotesInMeasure(xI4).Value \ 10000)))
                        ElseIf BMSChannelList(CurrentBMSChannel) = "08" Then 'If bpm requires declaration
                            Dim BpmIndex
                            For BpmIndex = 1 To UBound(hBPM) ' find BPM value in existing array
                                If NotesInMeasure(xI4).Value = hBPM(BpmIndex) Then Exit For
                            Next
                            If BpmIndex > UBound(hBPM) Then ' Didn't find it, add it
                                ReDim Preserve hBPM(UBound(hBPM) + 1)
                                hBPM(UBound(hBPM)) = NotesInMeasure(xI4).Value
                            End If
                            xText(UBound(xText)) = IIf(BPMx1296, C10to36(BpmIndex), Mid("0" & Hex(BpmIndex), Len(Hex(BpmIndex))))
                        ElseIf BMSChannelList(CurrentBMSChannel) = "09" Then 'If STOP
                            Dim StopIndex
                            For StopIndex = 1 To UBound(hSTOP) ' find STOP value in existing array
                                If NotesInMeasure(xI4).Value = hSTOP(StopIndex) Then Exit For
                            Next
                            If StopIndex > UBound(hSTOP) Then ' Didn't find it, add it
                                ReDim Preserve hSTOP(UBound(hSTOP) + 1)
                                hSTOP(UBound(hSTOP)) = NotesInMeasure(xI4).Value
                            End If
                            xText(UBound(xText)) = IIf(STOPx1296,
                                                       C10to36(StopIndex),
                                                       Mid("0" & Hex(StopIndex), Len(Hex(StopIndex))))
                        Else
                            xText(UBound(xText)) = C10to36(NotesInMeasure(xI4).Value \ 10000)
                        End If
                    End If
                Next

                If xVPosition.Length = 0 Then Continue For

                Dim xGCD As Double = MeasureLength(MeasureIndex)
                For xI2 = 0 To UBound(xVPosition)        'find greatest common divisor
                    If xVPosition(xI2) > 0 Then xGCD = GCD(xGCD, xVPosition(xI2))
                Next

                Dim xStrKey() As String
                ReDim xStrKey(CInt(MeasureLength(MeasureIndex) / xGCD) - 1)
                For xI2 = 0 To UBound(xStrKey)           'assign 00 to all keys
                    xStrKey(xI2) = "00"
                Next

                For xI2 = 0 To UBound(xVPosition)        'assign K texts
                    If CInt(xVPosition(xI2) / xGCD) > UBound(xStrKey) Then
                        ReDim Preserve xprevNotes(UBound(xprevNotes) + 1)
                        With xprevNotes(UBound(xprevNotes))
                            .ColumnIndex = IdentifiertoColumnIndex(BMSChannelList(CurrentBMSChannel))
                            .LongNote = IdentifiertoLongNote(BMSChannelList(CurrentBMSChannel))
                            .Hidden = IdentifiertoHidden(BMSChannelList(CurrentBMSChannel))
                            .VPosition = MeasureBottom(MeasureIndex)
                            .Value = C36to10(xText(xI2))
                        End With
                        If BMSChannelList(CurrentBMSChannel) = "08" Then _
                            xprevNotes(UBound(xprevNotes)).Value = IIf(BPMx1296, hBPM(C36to10(xText(xI2))), hBPM(Convert.ToInt32(xText(xI2), 16)))
                        If BMSChannelList(CurrentBMSChannel) = "09" Then _
                            xprevNotes(UBound(xprevNotes)).Value = IIf(STOPx1296, hSTOP(C36to10(xText(xI2))), hSTOP(Convert.ToInt32(xText(xI2), 16)))
                        Continue For
                    End If
                    If xStrKey(CInt(xVPosition(xI2) / xGCD)) <> "00" Then hasOverlapping = True
                    xStrKey(CInt(xVPosition(xI2) / xGCD)) = xText(xI2)
                Next

                xStrMeasure(MeasureIndex) &= "#" & Add3Zeros(MeasureIndex) & BMSChannelList(CurrentBMSChannel) & ":" & Join(xStrKey, "") & vbCrLf
            Next

        Next

        For MeasureIndex = 1 To UBound(hWAV)
            If Not hWAV(MeasureIndex) = "" Then xStrHeader &= "#WAV" & C10to36(MeasureIndex) &
                                                    " " & hWAV(MeasureIndex) & vbCrLf
        Next
        For MeasureIndex = 1 To UBound(hBPM)
            xStrHeader &= "#BPM" &
                IIf(BPMx1296, C10to36(MeasureIndex), Mid("0" & Hex(MeasureIndex), Len(Hex(MeasureIndex)))) &
                " " & Decimalify(hBPM(MeasureIndex) / 10000) & vbCrLf
        Next
        For MeasureIndex = 1 To UBound(hSTOP)
            xStrHeader &= "#STOP" &
                IIf(STOPx1296, C10to36(MeasureIndex), Mid("0" & Hex(MeasureIndex), Len(Hex(MeasureIndex)))) &
                " " & Decimalify(hSTOP(MeasureIndex) / 10000) & vbCrLf
        Next

        If hasOverlapping Then MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                                                          Strings.Messages.NoteOverlapError & vbCrLf &
                                                Strings.Messages.SavedFileWillContainErrors, MsgBoxStyle.Exclamation)
        If UBound(hBPM) > IIf(BPMx1296, 1295, 255) Then MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                                                          Strings.Messages.BPMOverflowError & UBound(hBPM) & " > " & IIf(BPMx1296, 1295, 255) & vbCrLf &
                                                Strings.Messages.SavedFileWillContainErrors, MsgBoxStyle.Exclamation)
        If UBound(hSTOP) > IIf(STOPx1296, 1295, 255) Then MsgBox(Strings.Messages.SaveWarning & vbCrLf &
                                                           Strings.Messages.STOPOverflowError & UBound(hSTOP) & " > " & IIf(STOPx1296, 1295, 255) & vbCrLf &
                                                  Strings.Messages.SavedFileWillContainErrors, MsgBoxStyle.Exclamation)

        Dim xStrExp As String = vbCrLf & "*---------------------- EXPANSION FIELD" & vbCrLf & TExpansion.Text & vbCrLf & vbCrLf
        If TExpansion.Text = "" Then xStrExp = ""

        Dim xStrMain As String = "*---------------------- MAIN DATA FIELD" & vbCrLf & vbCrLf & Join(xStrMeasure, "") & vbCrLf

        If xNTInput Then
            Notes = xKBackUp
            NTInput = True
            'SortByVPositionInsertion()
            'UpdatePairing()
        End If

        Dim xStrAll As String = xStrHeader & vbCrLf & xStrExp & vbCrLf & xStrMain
        Return xStrAll
    End Function

    Private Sub WriteBGM(MeasureIndex As Integer,
                         ByRef hasOverlapping As Boolean, xStrMeasure() As String,
                         TempK() As Note, GreatestColumn As Integer, ByRef xprevNotes() As Note)
        Dim xI2 = 0
        Dim xI4 = 0
        Dim xVPosition() As Double 'Ks in the same column
        Dim xText() As String    'Ks in the same column

        For ColumnIndex = niB To GreatestColumn 'Start rendering B notes (xI3 is columnindex)
            ReDim xVPosition(-1) 'Ks in the same column
            ReDim xText(-1)      'Ks in the same column

            For xI4 = 0 To UBound(TempK) 'Find Ks in the same column (xI4 is TK index)
                If TempK(xI4).ColumnIndex = ColumnIndex Then
                    ReDim Preserve xVPosition(UBound(xVPosition) + 1)
                    ReDim Preserve xText(UBound(xText) + 1)
                    xVPosition(UBound(xVPosition)) = TempK(xI4).VPosition - MeasureBottom(MeasureAtDisplacement(TempK(xI4).VPosition))
                    If xVPosition(UBound(xVPosition)) < 0 Then xVPosition(UBound(xVPosition)) = 0
                    xText(UBound(xText)) = C10to36(TempK(xI4).Value \ 10000)
                End If
            Next

            Dim xGCD As Double = MeasureLength(MeasureIndex)
            For xI2 = 0 To UBound(xVPosition)        'find greatest common divisor
                If xVPosition(xI2) > 0 Then xGCD = GCD(xGCD, xVPosition(xI2))
            Next

            Dim xStrKey(CInt(MeasureLength(MeasureIndex) / xGCD) - 1) As String
            For xI2 = 0 To UBound(xStrKey)           'assign 00 to all keys
                xStrKey(xI2) = "00"
            Next

            For xI2 = 0 To UBound(xVPosition)        'assign K texts
                If CInt(xVPosition(xI2) / xGCD) > UBound(xStrKey) Then
                    ReDim Preserve xprevNotes(UBound(xprevNotes) + 1)
                    With xprevNotes(UBound(xprevNotes))
                        .ColumnIndex = ColumnIndex
                        .VPosition = MeasureBottom(MeasureIndex)
                        .Value = C36to10(xText(xI2))
                    End With
                    Continue For
                End If
                If xStrKey(CInt(xVPosition(xI2) / xGCD)) <> "00" Then hasOverlapping = True
                xStrKey(CInt(xVPosition(xI2) / xGCD)) = xText(xI2)
            Next

            xStrMeasure(MeasureIndex) &= "#" & Add3Zeros(MeasureIndex) & "01:" & Join(xStrKey, "") & vbCrLf
        Next
    End Sub

    Private Function Decimalify(v As Double) As String
        Static nfi As New System.Globalization.NumberFormatInfo()
        nfi.NumberDecimalSeparator = "."
        Return v.ToString(nfi)
    End Function

    Private Function OpenSM(ByVal xStrAll As String) As Boolean
        KMouseOver = -1

        Dim xStrLine() As String = Split(xStrAll, vbCrLf)
        'Remove comments starting with "//"
        For xI1 As Integer = 0 To UBound(xStrLine)
            If xStrLine(xI1).Contains("//") Then xStrLine(xI1) = Mid(xStrLine(xI1), 1, InStr(xStrLine(xI1), "//") - 1)
        Next

        xStrAll = Join(xStrLine, "")
        xStrLine = Split(xStrAll, ";")

        Dim iDiff As Integer = 0
        Dim iCurrentDiff As Integer = 0
        Dim xTempSplit() As String = Split(xStrAll, "#NOTES:")
        Dim xTempStr() As String = {}
        If xTempSplit.Length > 2 Then
            ReDim Preserve xTempStr(UBound(xTempSplit) - 1)
            For xI1 As Integer = 1 To UBound(xTempSplit)
                xTempSplit(xI1) = Mid(xTempSplit(xI1), InStr(xTempSplit(xI1), ":") + 1)
                xTempSplit(xI1) = Mid(xTempSplit(xI1), InStr(xTempSplit(xI1), ":") + 1).Trim
                xTempStr(xI1 - 1) = Mid(xTempSplit(xI1), 1, InStr(xTempSplit(xI1), ":") - 1)
                xTempSplit(xI1) = Mid(xTempSplit(xI1), InStr(xTempSplit(xI1), ":") + 1).Trim
                xTempStr(xI1 - 1) &= " : " & Mid(xTempSplit(xI1), 1, InStr(xTempSplit(xI1), ":") - 1)
            Next

            Dim xDiag As New dgImportSM(xTempStr)
            If xDiag.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Return True
            iDiff = xDiag.iResult
        End If

        Dim sL As String
        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        ReDim hBPM(1295)    'x10000
        ReDim hSTOP(1295)
        Me.InitializeNewBMS()

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
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

                For xI1 As Integer = 0 To UBound(xItem)
                    xVal1 = Mid(xItem(xI1), 1, InStr(xItem(xI1), "=") - 1)
                    xVal2 = Mid(xItem(xI1), InStr(xItem(xI1), "=") + 1)

                    If xVal1 <> 0 Then
                        ReDim Preserve Notes(Notes.Length)
                        With Notes(UBound(Notes))
                            .ColumnIndex = niBPM
                            '.LongNote = False
                            '.Hidden = False
                            '.Selected = False
                            .VPosition = xVal1 * 48
                            .Value = xVal2 * 10000
                        End With
                    Else
                        Notes(0).Value = xVal2 * 10000
                    End If
                Next

            ElseIf UCase(sL).StartsWith("#NOTES:") Then
                If iCurrentDiff <> iDiff Then iCurrentDiff += 1 : GoTo Jump1

                iCurrentDiff += 1
                Dim xLine As String = Mid(sL, Len("#NOTES:") + 1)
                Dim xItem() As String = Split(xLine, ":")
                For xI1 As Integer = 0 To UBound(xItem)
                    xItem(xI1) = xItem(xI1).Trim
                Next

                If xItem.Length <> 6 Then GoTo Jump1

                THPlayLevel.Text = xItem(3)

                Dim xM() As String = Split(xItem(5), ",")
                For xI1 As Integer = 0 To UBound(xM)
                    xM(xI1) = xM(xI1).Trim
                Next

                For xI1 As Integer = 0 To UBound(xM)
                    For xI2 As Integer = 0 To Len(xM(xI1)) - 1 Step 4
                        If xM(xI1)(xI2) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = niA1
                                .LongNote = xM(xI1)(xI2) = "2" Or xM(xI1)(xI2) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192 \ (Len(xM(xI1)) \ 4)) * xI2 \ 4 + xI1 * 192
                                .Value = 10000
                            End With
                        End If
                        If xM(xI1)(xI2 + 1) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = niA2
                                .LongNote = xM(xI1)(xI2 + 1) = "2" Or xM(xI1)(xI2 + 1) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192 \ (Len(xM(xI1)) \ 4)) * xI2 \ 4 + xI1 * 192
                                .Value = 10000
                            End With
                        End If
                        If xM(xI1)(xI2 + 2) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = niA3
                                .LongNote = xM(xI1)(xI2 + 2) = "2" Or xM(xI1)(xI2 + 2) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192 \ (Len(xM(xI1)) \ 4)) * xI2 \ 4 + xI1 * 192
                                .Value = 10000
                            End With
                        End If
                        If xM(xI1)(xI2 + 3) <> "0" Then
                            ReDim Preserve Notes(Notes.Length)
                            With Notes(UBound(Notes))
                                .ColumnIndex = niA4
                                .LongNote = xM(xI1)(xI2 + 3) = "2" Or xM(xI1)(xI2 + 3) = "3"
                                '.Hidden = False
                                '.Selected = False
                                .VPosition = (192 \ (Len(xM(xI1)) \ 4)) * xI2 \ 4 + xI1 * 192
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
        For xI1 As Integer = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ": " & hWAV(xI1))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True

        THBPM.Value = Notes(0).Value / 10000
        SortByVPositionQuick(0, UBound(Notes))
        UpdatePairing()
        CalculateTotalNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
        Return False
    End Function

    ''' <summary>Do not clear Undo.</summary>
    Private Sub OpeniBMSC(ByVal Path As String)
        KMouseOver = -1

        Dim br As New BinaryReader(New FileStream(Path, FileMode.Open, FileAccess.Read), System.Text.Encoding.Unicode)

        If br.ReadInt32 <> &H534D4269 Then GoTo EndOfSub
        If br.ReadByte <> CByte(&H43) Then GoTo EndOfSub
        Dim xMajor As Integer = br.ReadByte
        Dim xMinor As Integer = br.ReadByte
        Dim xBuild As Integer = br.ReadByte

        ClearUndo()
        ReDim Notes(0)
        ReDim mColumn(999)
        ReDim hWAV(1295)
        Me.InitializeNewBMS()
        Me.InitializeOpenBMS()

        With Notes(0)
            .ColumnIndex = niBPM
            .VPosition = -1
            '.LongNote = False
            '.Selected = False
            .Value = 1200000
        End With

        Do Until br.BaseStream.Position >= br.BaseStream.Length
            Dim BlockID As Integer = br.ReadInt32()

            Select Case BlockID

                Case &H66657250     'Preferences
                    Dim xPref As Integer = br.ReadInt32

                    NTInput = xPref And &H1
                    TBNTInput.Checked = NTInput
                    mnNTInput.Checked = NTInput
                    POBLong.Enabled = Not NTInput
                    POBLongShort.Enabled = Not NTInput

                    ErrorCheck = xPref And &H2
                    TBErrorCheck.Checked = ErrorCheck
                    TBErrorCheck_Click(TBErrorCheck, New System.EventArgs)

                    PreviewOnClick = xPref And &H4
                    TBPreviewOnClick.Checked = PreviewOnClick
                    TBPreviewOnClick_Click(TBPreviewOnClick, New System.EventArgs)

                    ShowFileName = xPref And &H8
                    TBShowFileName.Checked = ShowFileName
                    TBShowFileName_Click(TBShowFileName, New System.EventArgs)

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
                    CGBPM.Checked = xPref And &H800000

                    CGSnap.Checked = xPref And &H1000000
                    CGDisableVertical.Checked = xPref And &H2000000
                    cVSLockL.Checked = xPref And &H4000000
                    cVSLock.Checked = xPref And &H8000000
                    cVSLockR.Checked = xPref And &H10000000

                    CGDivide.Value = br.ReadInt32
                    CGSub.Value = br.ReadInt32
                    gSlash = br.ReadInt32
                    CGHeight.Value = br.ReadSingle
                    CGWidth.Value = br.ReadSingle
                    CGB.Value = br.ReadInt32

                Case &H64616548     'Header
                    THTitle.Text = br.ReadString
                    THArtist.Text = br.ReadString
                    THGenre.Text = br.ReadString
                    Notes(0).Value = br.ReadInt32
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

                Case &H564157       'WAV List
                    Dim xWAVOptions As Integer = br.ReadByte
                    WAVMultiSelect = xWAVOptions And &H1
                    CWAVMultiSelect.Checked = WAVMultiSelect
                    CWAVMultiSelect_CheckedChanged(CWAVMultiSelect, New EventArgs)
                    WAVChangeLabel = xWAVOptions And &H2
                    CWAVChangeLabel.Checked = WAVChangeLabel
                    CWAVChangeLabel_CheckedChanged(CWAVChangeLabel, New EventArgs)

                    Dim xWAVCount As Integer = br.ReadInt32
                    For xxi As Integer = 1 To xWAVCount
                        Dim xI As Integer = br.ReadInt16
                        hWAV(xI) = br.ReadString
                    Next

                Case &H74616542     'Beat
                    nBeatN.Value = br.ReadInt16
                    nBeatD.Value = br.ReadInt16
                    'nBeatD.SelectedIndex = br.ReadByte

                    Dim xBeatChangeMode As Integer = br.ReadByte
                    Dim xBeatChangeList As RadioButton() = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
                    xBeatChangeList(xBeatChangeMode).Checked = True
                    CBeatPreserve_Click(xBeatChangeList(xBeatChangeMode), New System.EventArgs)

                    Dim xBeatCount As Integer = br.ReadInt32
                    For xxi As Integer = 1 To xBeatCount
                        Dim xIndex As Integer = br.ReadInt16
                        MeasureLength(xIndex) = br.ReadDouble
                        Dim xRatio As Double = MeasureLength(xIndex) / 192.0R
                        Dim xxD As Long = GetDenominator(xRatio)
                        LBeat.Items(xIndex) = Add3Zeros(xIndex) & ": " & xRatio & IIf(xxD > 10000, "", " ( " & CLng(xRatio * xxD) & " / " & xxD & " ) ")
                    Next

                Case &H6E707845     'Expansion Code
                    TExpansion.Text = br.ReadString

                Case &H65746F4E     'Note
                    Dim xNoteUbound As Integer = br.ReadInt32
                    ReDim Preserve Notes(xNoteUbound)
                    For i As Integer = 1 To UBound(Notes)
                        Notes(i).VPosition = br.ReadDouble
                        Notes(i).ColumnIndex = br.ReadInt32
                        Notes(i).Value = br.ReadInt32
                        Dim xFormat As Integer = br.ReadByte
                        Notes(i).Length = br.ReadDouble

                        Notes(i).LongNote = xFormat And &H1
                        Notes(i).Hidden = xFormat And &H2
                        Notes(i).Selected = xFormat And &H4
                    Next

                Case &H6F646E55     'Undo / Redo Commands
                    Dim URCount As Integer = br.ReadInt32   'Should be 100
                    sI = br.ReadInt32

                    For xI As Integer = 0 To 99
                        Dim xUndoCount As Integer = br.ReadInt32
                        Dim xBaseUndo As New UndoRedo.Void
                        Dim xIteratorUndo As UndoRedo.LinkedURCmd = xBaseUndo
                        For xxj As Integer = 1 To xUndoCount
                            Dim xByteLen As Integer = br.ReadInt32
                            Dim xByte() As Byte = br.ReadBytes(xByteLen)
                            xIteratorUndo.Next = UndoRedo.fromBytes(xByte)
                            xIteratorUndo = xIteratorUndo.Next
                        Next
                        sUndo(xI) = xBaseUndo.Next

                        Dim xRedoCount As Integer = br.ReadInt32
                        Dim xBaseRedo As New UndoRedo.Void
                        Dim xIteratorRedo As UndoRedo.LinkedURCmd = xBaseRedo
                        For xxj As Integer = 1 To xRedoCount
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

        LWAV.Visible = False
        LWAV.Items.Clear()
        For xI1 As Integer = 1 To 1295
            LWAV.Items.Add(C10to36(xI1) & ": " & hWAV(xI1))
        Next
        LWAV.SelectedIndex = 0
        LWAV.Visible = True

        THBPM.Value = Notes(0).Value / 10000
        SortByVPositionQuick(0, UBound(Notes))
        UpdatePairing()
        UpdateMeasureBottom()
        CalculateTotalNotes()
        CalculateGreatestVPosition()
        RefreshPanelAll()
        POStatusRefresh()
    End Sub

    Private Sub SaveiBMSC(ByVal Path As String)
        CalculateGreatestVPosition()
        SortByVPositionInsertion()
        UpdatePairing()

        Try

            Dim bw As New BinaryWriter(New IO.FileStream(Path, FileMode.Create), System.Text.Encoding.Unicode)

            'bw.Write("iBMSC".ToCharArray)
            bw.Write(&H534D4269)
            bw.Write(CByte(&H43))
            bw.Write(CByte(My.Application.Info.Version.Major))
            bw.Write(CByte(My.Application.Info.Version.Minor))
            bw.Write(CByte(My.Application.Info.Version.Build))

            'Preferences
            'bw.Write("Pref".ToCharArray)
            bw.Write(&H66657250)
            Dim xPref As Integer = 0
            If NTInput Then xPref = xPref Or &H1
            If ErrorCheck Then xPref = xPref Or &H2
            If PreviewOnClick Then xPref = xPref Or &H4
            If ShowFileName Then xPref = xPref Or &H8
            If mnSMenu.Checked Then xPref = xPref Or &H100
            If mnSTB.Checked Then xPref = xPref Or &H200
            If mnSOP.Checked Then xPref = xPref Or &H400
            If mnSStatus.Checked Then xPref = xPref Or &H800
            If mnSLSplitter.Checked Then xPref = xPref Or &H1000
            If mnSRSplitter.Checked Then xPref = xPref Or &H2000
            If gShowGrid Then xPref = xPref Or &H4000
            If gShowSubGrid Then xPref = xPref Or &H8000
            If gShowBG Then xPref = xPref Or &H10000
            If gShowMeasureNumber Then xPref = xPref Or &H20000
            If gShowMeasureBar Then xPref = xPref Or &H40000
            If gShowVerticalLine Then xPref = xPref Or &H80000
            If gShowC Then xPref = xPref Or &H100000
            If gBLP Then xPref = xPref Or &H200000
            If gSTOP Then xPref = xPref Or &H400000
            If gBPM Then xPref = xPref Or &H800000
            If gSnap Then xPref = xPref Or &H1000000
            If DisableVerticalMove Then xPref = xPref Or &H2000000
            If spLock(0) Then xPref = xPref Or &H4000000
            If spLock(1) Then xPref = xPref Or &H8000000
            If spLock(2) Then xPref = xPref Or &H10000000
            bw.Write(xPref)
            bw.Write(BitConverter.GetBytes(gDivide))
            bw.Write(BitConverter.GetBytes(gSub))
            bw.Write(BitConverter.GetBytes(gSlash))
            bw.Write(BitConverter.GetBytes(gxHeight))
            bw.Write(BitConverter.GetBytes(gxWidth))
            bw.Write(BitConverter.GetBytes(gColumns))

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

            Dim xWAVOptions As Integer = 0
            If WAVMultiSelect Then xWAVOptions = xWAVOptions Or &H1
            If WAVChangeLabel Then xWAVOptions = xWAVOptions Or &H2
            bw.Write(CByte(xWAVOptions))

            Dim xWAVCount As Integer = 0
            For i As Integer = 1 To UBound(hWAV)
                If hWAV(i) <> "" Then xWAVCount += 1
            Next
            bw.Write(xWAVCount)

            For i As Integer = 1 To UBound(hWAV)
                If hWAV(i) = "" Then Continue For
                bw.Write(CShort(i))
                bw.Write(hWAV(i))
            Next

            'Beat
            'bw.Write("Beat".ToCharArray)
            bw.Write(&H74616542)
            'Dim xNumerator As Short = nBeatN.Value
            'Dim xDenominator As Short = nBeatD.Value
            'Dim xBeatChangeMode As Byte = BeatChangeMode
            bw.Write(CShort(nBeatN.Value))
            bw.Write(CShort(nBeatD.Value))
            bw.Write(CByte(BeatChangeMode))

            Dim xBeatCount As Integer = 0
            For i As Integer = 0 To UBound(MeasureLength)
                If MeasureLength(i) <> 192.0R Then xBeatCount += 1
            Next
            bw.Write(xBeatCount)

            For i As Integer = 0 To UBound(MeasureLength)
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
            For i As Integer = 1 To UBound(Notes)
                Dim xFormat As Integer = 0
                If Notes(i).LongNote Then xFormat = xFormat Or &H1
                If Notes(i).Hidden Then xFormat = xFormat Or &H2
                If Notes(i).Selected Then xFormat = xFormat Or &H4

                bw.Write(Notes(i).VPosition)
                bw.Write(Notes(i).ColumnIndex)
                bw.Write(Notes(i).Value)
                bw.Write(CByte(xFormat))
                bw.Write(Notes(i).Length)
            Next

            'Undo / Redo Commands
            'bw.Write("Undo".ToCharArray)
            bw.Write(&H6F646E55)
            bw.Write(100)
            bw.Write(sI)

            For i As Integer = 0 To 99
                'UndoCommandsCount
                Dim countUndo As Integer = 0
                Dim pUndo As UndoRedo.LinkedURCmd = sUndo(i)
                While pUndo IsNot Nothing
                    countUndo += 1
                    pUndo = pUndo.Next
                End While
                bw.Write(countUndo)

                'UndoCommands
                pUndo = sUndo(i)
                For xxi As Integer = 1 To countUndo
                    Dim bUndo() As Byte = pUndo.toBytes
                    bw.Write(bUndo.Length)  'Length
                    bw.Write(bUndo)         'Command
                    pUndo = pUndo.Next
                Next

                'RedoCommandsCount
                Dim countRedo As Integer = 0
                Dim pRedo As UndoRedo.LinkedURCmd = sRedo(i)
                While pRedo IsNot Nothing
                    countRedo += 1
                    pRedo = pRedo.Next
                End While
                bw.Write(countRedo)

                'RedoCommands
                pRedo = sRedo(i)
                For xxi As Integer = 1 To countRedo
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
