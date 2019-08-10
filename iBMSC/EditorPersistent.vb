Imports System.Text
Imports iBMSC.Editor

Partial Public Class MainWindow
    Private Sub XMLWriteColumn(w As XmlTextWriter, I As Integer)
        w.WriteStartElement("Column")
        w.WriteAttributeString("Index", I)
        With Columns.column(I)
            'w.WriteAttributeString("Left", .Left)
            w.WriteAttributeString("Width", .Width)
            w.WriteAttributeString("Title", .Title)
            'w.WriteAttributeString("Text", .Text)
            'w.WriteAttributeString("Enabled", .Enabled)
            'w.WriteAttributeString("isNumeric", .isNumeric)
            'w.WriteAttributeString("Visible", .Visible)
            'w.WriteAttributeString("Identifier", .Identifier)
            w.WriteAttributeString("NoteColor", .cNote)
            w.WriteAttributeString("TextColor", .cText.ToArgb)
            w.WriteAttributeString("LongNoteColor", .cLNote)
            w.WriteAttributeString("LongTextColor", .cLText.ToArgb)
            w.WriteAttributeString("BG", .cBG.ToArgb)
        End With
        w.WriteEndElement()
    End Sub

    Private Sub XMLWriteFont(w As XmlTextWriter, local As String, f As Font)
        w.WriteStartElement(local)
        w.WriteAttributeString("Name", f.FontFamily.Name)
        w.WriteAttributeString("Size", f.SizeInPoints)
        w.WriteAttributeString("Style", f.Style)
        w.WriteEndElement()
    End Sub

    Private Sub XMLWritePlayerArguments(w As XmlTextWriter, I As Integer)
        w.WriteStartElement("Player")
        w.WriteAttributeString("Index", I)
        w.WriteAttributeString("Path", pArgs(I).Path)
        w.WriteAttributeString("FromBeginning", pArgs(I).aBegin)
        w.WriteAttributeString("FromHere", pArgs(I).aHere)
        w.WriteAttributeString("Stop", pArgs(I).aStop)
        w.WriteEndElement()
    End Sub

    Private Sub SaveSettings(Path As String, ThemeOnly As Boolean)
        Dim w As New XmlTextWriter(Path, Encoding.Unicode)
        With w
            .WriteStartDocument()
            .Formatting = Formatting.Indented
            .Indentation = 4

            .WriteStartElement("iBMSC")
            .WriteAttributeString("Major", My.Application.Info.Version.Major)
            .WriteAttributeString("Minor", My.Application.Info.Version.Minor)
            .WriteAttributeString("Build", My.Application.Info.Version.Build)

            If ThemeOnly Then GoTo 5000

            .WriteStartElement("Form")
            .WriteAttributeString("WindowState", IIf(IsFullscreen, _previousWindowState, Me.WindowState))
            .WriteAttributeString("Width", IIf(IsFullscreen, _previousWindowPosition.Width, Me.Width))
            .WriteAttributeString("Height", IIf(IsFullscreen, _previousWindowPosition.Height, Me.Height))
            .WriteAttributeString("Top", IIf(IsFullscreen, _previousWindowPosition.Top, Me.Top))
            .WriteAttributeString("Left", IIf(IsFullscreen, _previousWindowPosition.Left, Me.Left))
            .WriteEndElement()

            .WriteStartElement("Recent")
            .WriteAttributeString("Recent0", _recent(0))
            .WriteAttributeString("Recent1", _recent(1))
            .WriteAttributeString("Recent2", _recent(2))
            .WriteAttributeString("Recent3", _recent(3))
            .WriteAttributeString("Recent4", _recent(4))
            .WriteEndElement()

            .WriteStartElement("Edit")
            .WriteAttributeString("NTInput", NTInput)
            .WriteAttributeString("Language", _dispLang)
            '.WriteAttributeString("SortingMethod", SortingMethod)
            .WriteAttributeString("ErrorCheck", ErrorCheck)
            .WriteAttributeString("AutoFocusMouseEnter", AutoFocusPanelOnMouseEnter)
            .WriteAttributeString("FirstClickDisabled", FirstClickDisabled)
            .WriteAttributeString("ShowFileName", ShowFileName)
            .WriteAttributeString("MiddleButtonMoveMethod", MiddleButtonMoveMethod)
            .WriteAttributeString("AutoSaveInterval", _autoSaveInterval)
            .WriteAttributeString("PreviewOnClick", _previewOnClick)
            '.WriteAttributeString("PreviewErrorCheck", PreviewErrorCheck)
            .WriteAttributeString("ClickStopPreview", _clickStopPreview)
            .WriteEndElement()

            .WriteStartElement("Save")
            .WriteAttributeString("TextEncoding", EncodingToString(_textEncoding))
            .WriteAttributeString("BMSGridLimit", _bmsGridLimit)
            .WriteAttributeString("BeepWhileSaved", _beepWhileSaved)
            .WriteAttributeString("BPMx1296", _bpMx1296)
            .WriteAttributeString("STOPx1296", _stoPx1296)
            .WriteEndElement()

            .WriteStartElement("WAV")
            .WriteAttributeString("WAVMultiSelect", _wavMultiSelect)
            .WriteAttributeString("WAVChangeLabel", _wavChangeLabel)
            .WriteAttributeString("BeatChangeMode", _beatChangeMode)
            .WriteEndElement()

            .WriteStartElement("ShowHide")
            .WriteAttributeString("showMenu", mnSMenu.Checked)
            .WriteAttributeString("showTB", mnSTB.Checked)
            .WriteAttributeString("showOpPanel", mnSOP.Checked)
            .WriteAttributeString("showStatus", mnSStatus.Checked)
            .WriteAttributeString("showLSplit", mnSLSplitter.Checked)
            .WriteAttributeString("showRSplit", mnSRSplitter.Checked)
            .WriteEndElement()

            .WriteStartElement("Grid")
            .WriteAttributeString("Snap", Grid.IsSnapEnabled)
            .WriteAttributeString("WheelScroll", Grid.WheelScroll)
            .WriteAttributeString("PageUpDnScroll", Grid.PageUpDnScroll)
            .WriteAttributeString("gShow", Grid.ShowMainGrid)
            .WriteAttributeString("gShowS", Grid.ShowSubGrid)
            .WriteAttributeString("gShowBG", Grid.ShowBackground)
            .WriteAttributeString("gShowM", Grid.ShowMeasureNumber)
            .WriteAttributeString("gShowV", Grid.ShowVerticalLines)
            .WriteAttributeString("gShowMB", Grid.ShowMeasureBars)
            .WriteAttributeString("ShowColumnCaptions", Grid.ShowColumnCaptions)
            .WriteAttributeString("ShowBpmColumn", Grid.ShowBpmColumn)
            .WriteAttributeString("ShowStopColumn", Grid.ShowStopColumn)
            .WriteAttributeString("ShowScrollColumn", Grid.ShowScrollColumn)
            .WriteAttributeString("gBLP", Grid.ShowBgaColumn)
            .WriteAttributeString("gP2", CHPlayer.SelectedIndex)
            .WriteAttributeString("gCol", CGB.Value)
            .WriteAttributeString("Divider", Grid.Divider)
            .WriteAttributeString("Subdivider", Grid.Subdivider)
            .WriteAttributeString("Slash", Grid.Slash)
            .WriteAttributeString("gxHeight", Grid.HeightScale)
            .WriteAttributeString("WidthScale", Grid.WidthScale)
            .WriteEndElement()

            .WriteStartElement("WaveForm")
            .WriteAttributeString("wLock", wLock)
            .WriteAttributeString("wPosition", wPosition)
            .WriteAttributeString("wLeft", wLeft)
            .WriteAttributeString("wWidth", wWidth)
            .WriteAttributeString("wPrecision", wPrecision)
            .WriteEndElement()

            .WriteStartElement("Player")
            .WriteAttributeString("Count", pArgs.Length)
            .WriteAttributeString("CurrentPlayer", CurrentPlayer)
            For i = 0 To UBound(pArgs)
                XMLWritePlayerArguments(w, i) : Next
            .WriteEndElement()

            5000: .WriteStartElement("Columns")
            '.WriteAttributeString("Count", col.Length)
            For i = 0 To UBound(Columns.column)
                XMLWriteColumn(w, i) : Next
            .WriteEndElement()

            .WriteStartElement("VisualSettings")
            XMLWriteValue(w, "ColumnTitle", _theme.ColumnTitle.Color.ToArgb)
            XMLWriteFont(w, "ColumnTitleFont", _theme.ColumnTitleFont)
            XMLWriteValue(w, "Bg", _theme.Bg.Color.ToArgb)
            XMLWriteValue(w, "Grid", _theme.pGrid.Color.ToArgb)
            XMLWriteValue(w, "Sub", _theme.pSub.Color.ToArgb)
            XMLWriteValue(w, "VLine", _theme.pVLine.Color.ToArgb)
            XMLWriteValue(w, "MLine", _theme.pMLine.Color.ToArgb)
            XMLWriteValue(w, "BGMWav", _theme.pBGMWav.Color.ToArgb)
            XMLWriteValue(w, "SelBox", _theme.SelBox.Color.ToArgb)
            XMLWriteValue(w, "TSCursor", _theme.PECursor.Color.ToArgb)
            XMLWriteValue(w, "TSHalf", _theme.PEHalf.Color.ToArgb)
            XMLWriteValue(w, "TSDeltaMouseOver", _theme.PEDeltaMouseOver)
            XMLWriteValue(w, "TSMouseOver", _theme.PEMouseOver.Color.ToArgb)
            XMLWriteValue(w, "TSSel", _theme.PESel.Color.ToArgb)
            XMLWriteValue(w, "TSBPM", _theme.PEBPM.Color.ToArgb)
            XMLWriteFont(w, "TSBPMFont", _theme.PEBPMFont)
            XMLWriteValue(w, "MiddleDeltaRelease", _theme.MiddleDeltaRelease)
            XMLWriteValue(w, "kHeight", _theme.NoteHeight)
            XMLWriteFont(w, "kFont", _theme.kFont)
            XMLWriteFont(w, "kMFont", _theme.kMFont)
            XMLWriteValue(w, "kLabelVShift", _theme.kLabelVShift)
            XMLWriteValue(w, "kLabelHShift", _theme.kLabelHShift)
            XMLWriteValue(w, "kLabelHShiftL", _theme.kLabelHShiftL)
            XMLWriteValue(w, "kMouseOver", _theme.kMouseOver.Color.ToArgb)
            XMLWriteValue(w, "kMouseOverE", _theme.kMouseOverE.Color.ToArgb)
            XMLWriteValue(w, "kSelected", _theme.kSelected.Color.ToArgb)
            XMLWriteValue(w, "kOpacity", _theme.kOpacity)
            .WriteEndElement()

            .WriteEndElement()
            .WriteEndDocument()
            .Close()
        End With
    End Sub

    Private Sub XMLLoadElementValue(n As XmlElement, ByRef v As Integer)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Private Sub XMLLoadElementValue(n As XmlElement, ByRef v As Single)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Private Sub XMLLoadElementValue(n As XmlElement, ByRef v As Color)
        If n Is Nothing Then Exit Sub
        XMLLoadAttribute(n.GetAttribute("Value"), v)
    End Sub

    Private Sub XMLLoadElementValue(n As XmlElement, ByRef v As Font)
        If n Is Nothing Then Exit Sub

        Dim xName As String = Me.Font.FontFamily.Name
        Dim xSize As Integer = Me.Font.Size
        Dim xStyle As Integer = Me.Font.Style
        XMLLoadAttribute(n.GetAttribute("Name"), xName)
        XMLLoadAttribute(n.GetAttribute("Size"), xSize)
        XMLLoadAttribute(n.GetAttribute("Style"), xStyle)
        v = New Font(xName, xSize, CType(xStyle, FontStyle))
    End Sub

    Private Sub XMLLoadPlayer(n As XmlElement)
        Dim i As Integer = - 1
        XMLLoadAttribute(n.GetAttribute("Index"), i)
        If i < 0 Or i > UBound(pArgs) Then Exit Sub

        XMLLoadAttribute(n.GetAttribute("Path"), pArgs(i).Path)
        XMLLoadAttribute(n.GetAttribute("FromBeginning"), pArgs(i).aBegin)
        XMLLoadAttribute(n.GetAttribute("FromHere"), pArgs(i).aHere)
        XMLLoadAttribute(n.GetAttribute("Stop"), pArgs(i).aStop)
    End Sub

    Private Sub XMLLoadColumn(n As XmlElement)
        Dim i As Integer = - 1
        XMLLoadAttribute(n.GetAttribute("Index"), i)
        If i < 0 Or i > UBound(Columns.column) Then Exit Sub

        With Columns.column(i)
            'XMLLoadAttribute(n.GetAttribute("Left"), .Left)
            XMLLoadAttribute(n.GetAttribute("Width"), .Width)
            XMLLoadAttribute(n.GetAttribute("Title"), .Title)
            'XMLLoadAttribute(n.GetAttribute("Text"), .Text)
            Dim Display As Boolean
            Dim attr = n.GetAttribute("Display")
            XMLLoadAttribute(attr, Display)
            .IsVisible = IIf(String.IsNullOrEmpty(attr), .IsVisible, Display)

            'XMLLoadAttribute(n.GetAttribute("isNumeric"), .isNumeric)
            'XMLLoadAttribute(n.GetAttribute("Visible"), .Visible)
            'XMLLoadAttribute(n.GetAttribute("Identifier"), .Identifier)
            XMLLoadAttribute(n.GetAttribute("NoteColor"), .cNote)
            .setNoteColor(.cNote)
            XMLLoadAttribute(n.GetAttribute("TextColor"), .cText)
            XMLLoadAttribute(n.GetAttribute("LongNoteColor"), .cLNote)
            .setLNoteColor(.cLNote)
            XMLLoadAttribute(n.GetAttribute("LongTextColor"), .cLText)
            XMLLoadAttribute(n.GetAttribute("BG"), .cBG)
        End With
    End Sub

    Private Sub LoadSettings(Path As String)
        If Not My.Computer.FileSystem.FileExists(Path) Then Return

        'Dim xTempFileName As String = ""
        'Do
        'Try
        'xTempFileName = Me.RandomFileName(".xml")
        'File.Copy(Path, xTempFileName)
        'Catch
        'Continue Do
        'End Try
        'Exit Do
        'Loop
        Dim Doc As New XmlDocument
        Dim FileStream As New FileStream(Path, FileMode.Open, FileAccess.Read)
        Doc.Load(FileStream)

        Dim Root As XmlElement = Doc.Item("iBMSC")
        If Root Is Nothing Then GoTo EndOfSub

        'version
        Dim Major As Integer = My.Application.Info.Version.Major
        Dim Minor As Integer = My.Application.Info.Version.Minor
        Dim Build As Integer = My.Application.Info.Version.Build
        Try
            Dim xMajor As Integer = Val(Root.Attributes("Major").Value)
            Dim xMinor As Integer = Val(Root.Attributes("Minor").Value)
            Dim xBuild As Integer = Val(Root.Attributes("Build").Value)
            Major = xMajor
            Minor = xMinor
            Build = xBuild
        Catch ex As Exception
        End Try

        'form
        Dim eForm As XmlElement = Root.Item("Form")
        If eForm IsNot Nothing Then
            With eForm
                Select Case Val(.GetAttribute("WindowState"))
                    Case FormWindowState.Normal
                        Me.WindowState = FormWindowState.Normal
                        XMLLoadAttribute(.GetAttribute("Width"), Me.Width)
                        XMLLoadAttribute(.GetAttribute("Height"), Me.Height)
                        XMLLoadAttribute(.GetAttribute("Top"), Me.Top)
                        XMLLoadAttribute(.GetAttribute("Left"), Me.Left)
                    Case FormWindowState.Maximized
                        Me.WindowState = FormWindowState.Maximized
                End Select
            End With
        End If

        'recent
        Dim eRecent As XmlElement = Root.Item("Recent")
        If eRecent IsNot Nothing Then
            With eRecent
                XMLLoadAttribute(.GetAttribute("Recent0"), _recent(0))
                SetRecent(0, _recent(0))
                XMLLoadAttribute(.GetAttribute("Recent1"), _recent(1))
                SetRecent(1, _recent(1))
                XMLLoadAttribute(.GetAttribute("Recent2"), _recent(2))
                SetRecent(2, _recent(2))
                XMLLoadAttribute(.GetAttribute("Recent3"), _recent(3))
                SetRecent(3, _recent(3))
                XMLLoadAttribute(.GetAttribute("Recent4"), _recent(4))
                SetRecent(4, _recent(4))
            End With
        End If

        'edit
        Dim eEdit As XmlElement = Root.Item("Edit")
        If eEdit IsNot Nothing Then
            With eEdit
                XMLLoadAttribute(.GetAttribute("NTInput"), NTInput)
                TBNTInput.Checked = NTInput
                mnNTInput.Checked = NTInput
                POBLong.Enabled = Not NTInput
                POBLongShort.Enabled = Not NTInput

                LoadLocale(My.Application.Info.DirectoryPath & "\" & .GetAttribute("Language"))

                'XMLLoadAttribute(.GetAttribute("SortingMethod"), SortingMethod)

                XMLLoadAttribute(.GetAttribute("ErrorCheck"), ErrorCheck)
                TBErrorCheck.Checked = ErrorCheck
                TBErrorCheck_Click(TBErrorCheck, New EventArgs)

                XMLLoadAttribute(.GetAttribute("ShowFileName"), ShowFileName)
                TBShowFileName.Checked = ShowFileName
                TBShowFileName_Click(TBShowFileName, New EventArgs)

                XMLLoadAttribute(.GetAttribute("MiddleButtonMoveMethod"), MiddleButtonMoveMethod)
                XMLLoadAttribute(.GetAttribute("AutoFocusMouseEnter"), AutoFocusPanelOnMouseEnter)
                XMLLoadAttribute(.GetAttribute("FirstClickDisabled"), FirstClickDisabled)

                XMLLoadAttribute(.GetAttribute("AutoSaveInterval"), _autoSaveInterval)
                If _autoSaveInterval Then AutoSaveTimer.Interval = _autoSaveInterval Else AutoSaveTimer.Enabled = False

                XMLLoadAttribute(.GetAttribute("PreviewOnClick"), _previewOnClick)
                TBPreviewOnClick.Checked = _previewOnClick
                TBPreviewOnClick_Click(TBPreviewOnClick, New EventArgs)

                XMLLoadAttribute(.GetAttribute("ClickStopPreview"), _clickStopPreview)
            End With
        End If

        'save
        Dim eSave As XmlElement = Root.Item("Save")
        If eSave IsNot Nothing Then
            With eSave
                Select Case UCase(.GetAttribute("TextEncoding"))
                    Case "SYSTEM ANSI" : _textEncoding = Encoding.Default
                    Case "LITTLE ENDIAN UTF16" : _textEncoding = Encoding.Unicode
                    Case "ASCII" : _textEncoding = Encoding.ASCII
                    Case "BIG ENDIAN UTF16" : _textEncoding = Encoding.BigEndianUnicode
                    Case "LITTLE ENDIAN UTF32" : _textEncoding = Encoding.UTF32
                    Case "UTF7" : _textEncoding = Encoding.UTF7
                    Case "UTF8" : _textEncoding = Encoding.UTF8
                    Case "SJIS" : _textEncoding = Encoding.GetEncoding(932)
                    Case "EUC-KR" : _textEncoding = Encoding.GetEncoding(51949)
                        ' leave current encoding
                        ' Case Else 
                End Select

                XMLLoadAttribute(.GetAttribute("BMSGridLimit"), _bmsGridLimit)
                XMLLoadAttribute(.GetAttribute("BeepWhileSaved"), _beepWhileSaved)
                XMLLoadAttribute(.GetAttribute("BPMx1296"), _bpMx1296)
                XMLLoadAttribute(.GetAttribute("STOPx1296"), _stoPx1296)
            End With
        End If

        'WAV
        Dim eWAV As XmlElement = Root.Item("WAV")
        If eWAV IsNot Nothing Then
            With eWAV
                XMLLoadAttribute(.GetAttribute("WAVMultiSelect"), _wavMultiSelect)
                CWAVMultiSelect.Checked = _wavMultiSelect
                CWAVMultiSelect_CheckedChanged(CWAVMultiSelect, New EventArgs)

                XMLLoadAttribute(.GetAttribute("WAVChangeLabel"), _wavChangeLabel)
                CWAVChangeLabel.Checked = _wavChangeLabel
                CWAVChangeLabel_CheckedChanged(CWAVChangeLabel, New EventArgs)

                Dim xInt = CInt(.GetAttribute("BeatChangeMode"))
                Dim xBeatOpList As RadioButton() = {CBeatPreserve, CBeatMeasure, CBeatCut, CBeatScale}
                If xInt >= 0 And xInt < xBeatOpList.Length Then
                    xBeatOpList(xInt).Checked = True
                    CBeatPreserve_Click(xBeatOpList(xInt), New EventArgs)
                End If
            End With
        End If

        'ShowHide
        Dim eShowHide As XmlElement = Root.Item("ShowHide")
        If eShowHide IsNot Nothing Then
            With eShowHide
                XMLLoadAttribute(.GetAttribute("showMenu"), mnSMenu.Checked)
                XMLLoadAttribute(.GetAttribute("showTB"), mnSTB.Checked)
                XMLLoadAttribute(.GetAttribute("showOpPanel"), mnSOP.Checked)
                XMLLoadAttribute(.GetAttribute("showStatus"), mnSStatus.Checked)
                XMLLoadAttribute(.GetAttribute("showLSplit"), mnSLSplitter.Checked)
                XMLLoadAttribute(.GetAttribute("showRSplit"), mnSRSplitter.Checked)
            End With
        End If

        'Grid
        Dim eGrid As XmlElement = Root.Item("Grid")
        If eGrid IsNot Nothing Then
            With eGrid
                XMLLoadAttribute(.GetAttribute("Snap"), CGSnap.Checked)
                XMLLoadAttribute(.GetAttribute("WheelScroll"), Grid.WheelScroll)
                XMLLoadAttribute(.GetAttribute("PageUpDnScroll"), Grid.PageUpDnScroll)
                XMLLoadAttribute(.GetAttribute("gShow"), CGShow.Checked)
                XMLLoadAttribute(.GetAttribute("gShowS"), CGShowS.Checked)
                XMLLoadAttribute(.GetAttribute("gShowBG"), CGShowBG.Checked)
                XMLLoadAttribute(.GetAttribute("gShowM"), CGShowM.Checked)
                XMLLoadAttribute(.GetAttribute("gShowV"), CGShowV.Checked)
                XMLLoadAttribute(.GetAttribute("gShowMB"), CGShowMB.Checked)
                XMLLoadAttribute(.GetAttribute("ShowColumnCaptions"), CGShowC.Checked)
                XMLLoadAttribute(.GetAttribute("ShowBpmColumn"), CGBPM.Checked)
                XMLLoadAttribute(.GetAttribute("ShowStopColumn"), CGSTOP.Checked)
                XMLLoadAttribute(.GetAttribute("ShowScrollColumn"), CGSCROLL.Checked)
                XMLLoadAttribute(.GetAttribute("gBLP"), CGBLP.Checked)
                XMLLoadAttribute(.GetAttribute("gP2"), CHPlayer.SelectedIndex)
                XMLLoadAttribute(.GetAttribute("gCol"), CGB.Value)
                XMLLoadAttribute(.GetAttribute("gxHeight"), CGHeight.Value)
                XMLLoadAttribute(.GetAttribute("WidthScale"), CGWidth.Value)
                XMLLoadAttribute(.GetAttribute("Slash"), Grid.Slash)

                Dim xgDivide = CInt(.GetAttribute("Divider"))
                If xgDivide >= CGDivide.Minimum And xgDivide <= CGDivide.Maximum Then CGDivide.Value = xgDivide

                Dim xgSub = CInt(.GetAttribute("Subdivider"))
                If xgSub >= CGSub.Minimum And xgSub <= CGSub.Maximum Then CGSub.Value = xgSub
            End With
        End If

        'WaveForm
        Dim eWaveForm As XmlElement = Root.Item("WaveForm")
        If eWaveForm IsNot Nothing Then
            With eWaveForm
                XMLLoadAttribute(.GetAttribute("wLock"), BWLock.Checked)
                XMLLoadAttribute(.GetAttribute("wPosition"), TWPosition.Value)
                XMLLoadAttribute(.GetAttribute("wLeft"), TWLeft.Value)
                XMLLoadAttribute(.GetAttribute("wWidth"), TWWidth.Value)
                XMLLoadAttribute(.GetAttribute("wPrecision"), TWPrecision.Value)
            End With
        End If

        'Player
        Dim ePlayer As XmlElement = Root.Item("Player")
        If ePlayer IsNot Nothing Then
            With ePlayer
                XMLLoadAttribute(.GetAttribute("CurrentPlayer"), CurrentPlayer)

                Dim xCount As Integer = .GetAttribute("Count")
                If xCount > 0 Then ReDim Preserve pArgs(xCount - 1)
            End With

            For Each eePlayer As XmlElement In ePlayer.ChildNodes
                Me.XMLLoadPlayer(eePlayer)
            Next
        End If

        'Columns
        Dim eColumns As XmlElement = Root.Item("Columns")
        If eColumns IsNot Nothing Then
            For Each eeCol As XmlElement In eColumns.ChildNodes
                Me.XMLLoadColumn(eeCol)
            Next
        End If

        'VisualSettings
        Dim eVisualSettings As XmlElement = Root.Item("VisualSettings")
        If eVisualSettings IsNot Nothing Then
            With eVisualSettings
                XMLLoadElementValue(.Item("ColumnTitle"), _theme.ColumnTitle.Color)
                XMLLoadElementValue(.Item("ColumnTitleFont"), _theme.ColumnTitleFont)
                XMLLoadElementValue(.Item("Bg"), _theme.Bg.Color)
                XMLLoadElementValue(.Item("Grid"), _theme.pGrid.Color)
                XMLLoadElementValue(.Item("Sub"), _theme.pSub.Color)
                XMLLoadElementValue(.Item("VLine"), _theme.pVLine.Color)
                XMLLoadElementValue(.Item("MLine"), _theme.pMLine.Color)

                XMLLoadElementValue(.Item("BGMWav"), _theme.pBGMWav.Color)
                TWTransparency.Value = _theme.pBGMWav.Color.A
                TWTransparency2.Value = _theme.pBGMWav.Color.A
                TWSaturation.Value = _theme.pBGMWav.Color.GetSaturation*1000
                TWSaturation2.Value = _theme.pBGMWav.Color.GetSaturation*1000

                XMLLoadElementValue(.Item("SelBox"), _theme.SelBox.Color)
                XMLLoadElementValue(.Item("TSCursor"), _theme.PECursor.Color)
                XMLLoadElementValue(.Item("TSHalf"), _theme.PEHalf.Color)
                XMLLoadElementValue(.Item("TSDeltaMouseOver"), _theme.PEDeltaMouseOver)
                XMLLoadElementValue(.Item("TSMouseOver"), _theme.PEMouseOver.Color)
                XMLLoadElementValue(.Item("TSSel"), _theme.PESel.Color)
                XMLLoadElementValue(.Item("TSBPM"), _theme.PEBPM.Color)
                XMLLoadElementValue(.Item("TSBPMFont"), _theme.PEBPMFont)
                XMLLoadElementValue(.Item("MiddleDeltaRelease"), _theme.MiddleDeltaRelease)
                XMLLoadElementValue(.Item("kHeight"), _theme.NoteHeight)
                XMLLoadElementValue(.Item("kFont"), _theme.kFont)
                XMLLoadElementValue(.Item("kMFont"), _theme.kMFont)
                XMLLoadElementValue(.Item("kLabelVShift"), _theme.kLabelVShift)
                XMLLoadElementValue(.Item("kLabelHShift"), _theme.kLabelHShift)
                XMLLoadElementValue(.Item("kLabelHShiftL"), _theme.kLabelHShiftL)
                XMLLoadElementValue(.Item("kMouseOver"), _theme.kMouseOver.Color)
                XMLLoadElementValue(.Item("kMouseOverE"), _theme.kMouseOverE.Color)
                XMLLoadElementValue(.Item("kSelected"), _theme.kSelected.Color)
                XMLLoadElementValue(.Item("kOpacity"), _theme.kOpacity)
            End With
        End If

        EndOfSub:
        UpdateColumnsX()
        FileStream.Close()
        'File.Delete(xTempFileName)
    End Sub

    Private Sub XMLLoadLocaleMenu(n As XmlElement, ByRef target As String)
        If n Is Nothing Then Exit Sub
        If n.HasAttribute("amp") Then
            target = n.InnerText.Insert(Integer.Parse(n.GetAttribute("amp")), "&")
        Else
            target = n.InnerText
        End If
    End Sub

    Private Sub XMLLoadLocale(n As XmlElement, ByRef target As String)
        If n IsNot Nothing Then target = n.InnerText
    End Sub

    Private Sub XMLLoadLocaleToolTipUniversal(n As XmlElement, target As Control)
        If n Is Nothing Then Exit Sub
        ToolTipUniversal.SetToolTip(target, n.InnerText)
    End Sub

    Private Sub LoadLocale(Path As String)
        If Not My.Computer.FileSystem.FileExists(Path) Then Return

        Dim Doc As XmlDocument = Nothing
        Dim FileStream As FileStream = Nothing

        Dim xPOHeaderPart2 As Boolean = POHeaderPart2.Visible
        Dim xPOGridPart2 As Boolean = POGridPart2.Visible
        Dim xPOWaveFormPart2 As Boolean = POWaveFormPart2.Visible
        POHeaderPart2.Visible = True
        POGridPart2.Visible = True
        POWaveFormPart2.Visible = True

        Try
            Doc = New XmlDocument
            FileStream = New FileStream(Path, FileMode.Open, FileAccess.Read)
            Doc.Load(FileStream)

            Dim Root As XmlElement = Doc.Item("iBMSC.Locale")
            If Root Is Nothing Then Throw New NullReferenceException

            XMLLoadLocale(Root.Item("OK"), Strings.OK)
            XMLLoadLocale(Root.Item("Cancel"), Strings.Cancel)
            XMLLoadLocale(Root.Item("None"), Strings.None)

            Dim eFont As XmlElement = Root.Item("Font")
            If eFont IsNot Nothing Then
                Dim xSize = 9
                If eFont.HasAttribute("Size") Then xSize = Val(eFont.GetAttribute("Size"))

                Dim fRegular As New Font(Me.Font.FontFamily, xSize, FontStyle.Regular)
                Dim xChildNode As XmlNode = eFont.LastChild
                Do While xChildNode IsNot Nothing
                    If xChildNode.LocalName <> "Family" Then Continue Do
                    If isFontInstalled(xChildNode.InnerText) Then
                        fRegular = New Font(xChildNode.InnerText, xSize)
                    End If
                    xChildNode = xChildNode.PreviousSibling
                Loop

                Dim rList() As Object =
                        {Me, mnSys, Menu1, mnMain, cmnLanguage, cmnTheme, cmnConversion, TBMain, FStatus, FStatus2}
                For Each c As Object In rList
                    Try
                        c.Font = fRegular
                    Catch ex As Exception
                    End Try
                Next

                Dim fBold As New Font(fRegular, FontStyle.Bold)

                Dim bList() As Object =
                        {TBStatistics, FSSS, FSSL, FSSH, TVCM, TVCD, TVCBPM, FSP1, FSP3, FSP2}
                For Each c As Object In bList
                    Try
                        c.Font = fBold
                    Catch ex As Exception
                    End Try
                Next
            End If

            Dim eMonoFont As XmlElement = Root.Item("MonoFont")
            If eMonoFont IsNot Nothing Then
                Dim xSize = 9
                If eMonoFont.HasAttribute("Size") Then xSize = Val(eMonoFont.GetAttribute("Size"))

                Dim fMono As New Font(POWAVInner.Font.FontFamily, xSize)
                Dim xChildNode As XmlNode = eMonoFont.LastChild
                Do While xChildNode IsNot Nothing
                    If xChildNode.LocalName <> "Family" Then Continue Do
                    If isFontInstalled(xChildNode.InnerText) Then
                        fMono = New Font(xChildNode.InnerText, xSize)
                    End If
                    xChildNode = xChildNode.PreviousSibling
                Loop

                Dim mList() As Object = {LWAV, LBMP, LBeat, TExpansion}
                For Each c As Object In mList
                    Try
                        c.font = fMono
                    Catch ex As Exception
                    End Try
                Next
            End If

            Dim eMenu As XmlElement = Root.Item("Menu")
            If eMenu IsNot Nothing Then

                Dim eFile As XmlElement = eMenu.Item("File")
                If eFile IsNot Nothing Then
                    XMLLoadLocaleMenu(eFile.Item("Title"), mnFile.Text)
                    XMLLoadLocaleMenu(eFile.Item("New"), mnNew.Text)
                    XMLLoadLocaleMenu(eFile.Item("Open"), mnOpen.Text)
                    XMLLoadLocaleMenu(eFile.Item("ImportSM"), mnImportSM.Text)
                    XMLLoadLocaleMenu(eFile.Item("ImportIBMSC"), mnImportIBMSC.Text)
                    XMLLoadLocaleMenu(eFile.Item("Save"), mnSave.Text)
                    XMLLoadLocaleMenu(eFile.Item("SaveAs"), mnSaveAs.Text)
                    XMLLoadLocaleMenu(eFile.Item("ExportIBMSC"), mnExport.Text)
                    If _recent(0) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent0"), mnOpenR0.Text)
                    If _recent(1) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent1"), mnOpenR1.Text)
                    If _recent(2) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent2"), mnOpenR2.Text)
                    If _recent(3) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent3"), mnOpenR3.Text)
                    If _recent(4) = "" Then XMLLoadLocaleMenu(eFile.Item("Recent4"), mnOpenR4.Text)
                    XMLLoadLocaleMenu(eFile.Item("Quit"), mnQuit.Text)
                End If

                Dim eEdit As XmlElement = eMenu.Item("Edit")
                If eEdit IsNot Nothing Then
                    XMLLoadLocaleMenu(eEdit.Item("Title"), mnEdit.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Undo"), mnUndo.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Redo"), mnRedo.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Cut"), mnCut.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Copy"), mnCopy.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Paste"), mnPaste.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Delete"), mnDelete.Text)
                    XMLLoadLocaleMenu(eEdit.Item("SelectAll"), mnSelectAll.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Find"), mnFind.Text)
                    XMLLoadLocaleMenu(eEdit.Item("Stat"), mnStatistics.Text)
                    XMLLoadLocaleMenu(eEdit.Item("TimeSelectionTool"), mnTimeSelect.Text)
                    XMLLoadLocaleMenu(eEdit.Item("SelectTool"), mnSelect.Text)
                    XMLLoadLocaleMenu(eEdit.Item("WriteTool"), mnWrite.Text)
                    XMLLoadLocaleMenu(eEdit.Item("MyO2"), mnMyO2.Text)
                End If

                Dim eView As XmlElement = eMenu.Item("View")
                If eView IsNot Nothing Then
                    XMLLoadLocaleMenu(eView.Item("Title"), mnSys.Text)
                End If

                Dim eOptions As XmlElement = eMenu.Item("Options")
                If eOptions IsNot Nothing Then
                    XMLLoadLocaleMenu(eOptions.Item("Title"), mnOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("NT"), mnNTInput.Text)
                    XMLLoadLocaleMenu(eOptions.Item("ErrorCheck"), mnErrorCheck.Text)
                    XMLLoadLocaleMenu(eOptions.Item("PreviewOnClick"), mnPreviewOnClick.Text)
                    XMLLoadLocaleMenu(eOptions.Item("ShowFileName"), mnShowFileName.Text)
                    XMLLoadLocaleMenu(eOptions.Item("GeneralOptions"), mnGOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("VisualOptions"), mnVOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("PlayerOptions"), mnPOptions.Text)
                    XMLLoadLocaleMenu(eOptions.Item("Language"), mnLanguage.Text)
                    XMLLoadLocaleMenu(eOptions.Item("Theme"), mnTheme.Text)
                End If

                XMLLoadLocaleMenu(eMenu.Item("Conversion"), mnConversion.Text)

                Dim ePreview As XmlElement = eMenu.Item("Preview")
                If ePreview IsNot Nothing Then
                    XMLLoadLocaleMenu(ePreview.Item("Title"), mnPreview.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayBegin"), mnPlayB.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayHere"), mnPlay.Text)
                    XMLLoadLocaleMenu(ePreview.Item("PlayStop"), mnStop.Text)
                End If

                Dim eAbout As XmlElement = eMenu.Item("About")
                If eAbout IsNot Nothing Then
                    'XMLLoadLocaleMenu(eAbout.Item("Title"), mnAbout.Text)
                    'XMLLoadLocaleMenu(eAbout.Item("About"), mnAbout1.Text)
                    'XMLLoadLocaleMenu(eAbout.Item("CheckUpdates"), mnUpdate.Text)
                    'XMLLoadLocaleMenu(eAbout.Item("CheckUpdatesC"), mnUpdateC.Text)
                End If
            End If

            Dim eToolBar As XmlElement = Root.Item("ToolBar")
            If eToolBar IsNot Nothing Then
                XMLLoadLocale(eToolBar.Item("New"), TBNew.Text)
                XMLLoadLocale(eToolBar.Item("Open"), TBOpen.Text)
                XMLLoadLocale(eToolBar.Item("Save"), TBSave.Text)
                XMLLoadLocale(eToolBar.Item("Cut"), TBCut.Text)
                XMLLoadLocale(eToolBar.Item("Copy"), TBCopy.Text)
                XMLLoadLocale(eToolBar.Item("Paste"), TBPaste.Text)
                XMLLoadLocale(eToolBar.Item("Find"), TBFind.Text)
                XMLLoadLocale(eToolBar.Item("Stat"), TBStatistics.ToolTipText)
                XMLLoadLocale(eToolBar.Item("Conversion"), POConvert.Text)
                XMLLoadLocale(eToolBar.Item("MyO2"), TBMyO2.Text)
                XMLLoadLocale(eToolBar.Item("ErrorCheck"), TBErrorCheck.Text)
                XMLLoadLocale(eToolBar.Item("PreviewOnClick"), TBPreviewOnClick.Text)
                XMLLoadLocale(eToolBar.Item("ShowFileName"), TBShowFileName.Text)
                XMLLoadLocale(eToolBar.Item("Undo"), TBUndo.Text)
                XMLLoadLocale(eToolBar.Item("Redo"), TBRedo.Text)
                XMLLoadLocale(eToolBar.Item("NT"), TBNTInput.Text)
                XMLLoadLocale(eToolBar.Item("TimeSelectionTool"), TBTimeSelect.Text)
                XMLLoadLocale(eToolBar.Item("SelectTool"), TBSelect.Text)
                XMLLoadLocale(eToolBar.Item("WriteTool"), TBWrite.Text)
                XMLLoadLocale(eToolBar.Item("PlayBegin"), TBPlayB.Text)
                XMLLoadLocale(eToolBar.Item("PlayHere"), TBPlay.Text)
                XMLLoadLocale(eToolBar.Item("PlayStop"), TBStop.Text)
                XMLLoadLocale(eToolBar.Item("PlayerOptions"), TBPOptions.Text)
                XMLLoadLocale(eToolBar.Item("VisualOptions"), TBVOptions.Text)
                XMLLoadLocale(eToolBar.Item("GeneralOptions"), TBGOptions.Text)
                XMLLoadLocale(eToolBar.Item("Language"), TBLanguage.Text)
                XMLLoadLocale(eToolBar.Item("Theme"), TBTheme.Text)
                ' XMLLoadLocale(eToolBar.Item("About"), TBAbout.Text)
            End If

            Dim eStatusBar As XmlElement = Root.Item("StatusBar")
            If eStatusBar IsNot Nothing Then
                XMLLoadLocale(eStatusBar.Item("ColumnCaption"), FSC.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("NoteIndex"), FSW.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("MeasureIndex"), FSM.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("GridResolution"), FSP1.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("ReducedResolution"), FSP3.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("MeasureResolution"), FSP2.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("AbsolutePosition"), FSP4.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("Length"), Strings.StatusBar.Length)
                XMLLoadLocale(eStatusBar.Item("LongNote"), Strings.StatusBar.LongNote)
                XMLLoadLocale(eStatusBar.Item("Hidden"), Strings.StatusBar.Hidden)
                XMLLoadLocale(eStatusBar.Item("Error"), Strings.StatusBar.Err)
                XMLLoadLocale(eStatusBar.Item("SelStart"), FSSS.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("SelLength"), FSSL.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("SelSplit"), FSSH.ToolTipText)
                XMLLoadLocale(eStatusBar.Item("Reverse"), BVCReverse.Text)
                XMLLoadLocale(eStatusBar.Item("ByMultiple"), BVCApply.Text)
                XMLLoadLocale(eStatusBar.Item("ByValue"), BVCCalculate.Text)
            End If

            Dim eSubMenu As XmlElement = Root.Item("SubMenu")
            If eSubMenu IsNot Nothing Then

                Dim eShowHide As XmlElement = eSubMenu.Item("ShowHide")
                If eShowHide IsNot Nothing Then
                    'Dim xToolTip As String = ToolTipUniversal.GetToolTip(ttlIcon)
                    'XMLLoadLocaleMenu(eShowHide.Item("ToolTip"), xToolTip)
                    'ToolTipUniversal.SetToolTip(ttlIcon, xToolTip)

                    XMLLoadLocaleMenu(eShowHide.Item("Menu"), mnSMenu.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("ToolBar"), mnSTB.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("OptionsPanel"), mnSOP.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("StatusBar"), mnSStatus.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("LSplit"), mnSLSplitter.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("RSplit"), mnSRSplitter.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Grid"), CGShow.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Sub"), CGShowS.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BG"), CGShowBG.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("MeasureIndex"), CGShowM.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("MeasureLine"), CGShowMB.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("Vertical"), CGShowV.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("ColumnCaption"), CGShowC.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BPM"), CGBPM.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("STOP"), CGSTOP.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("SCROLL"), CGSCROLL.Text)
                    XMLLoadLocaleMenu(eShowHide.Item("BLP"), CGBLP.Text)
                End If

                Dim eInsertMeasure As XmlElement = eSubMenu.Item("InsertMeasure")
                If eInsertMeasure IsNot Nothing Then
                    XMLLoadLocaleMenu(eInsertMeasure.Item("Insert"), MInsert.Text)
                    XMLLoadLocaleMenu(eInsertMeasure.Item("Remove"), MRemove.Text)
                End If

                Dim eLanguage As XmlElement = eSubMenu.Item("Language")
                If eLanguage IsNot Nothing Then
                    XMLLoadLocaleMenu(eLanguage.Item("Default"), TBLangDef.Text)
                    XMLLoadLocaleMenu(eLanguage.Item("Refresh"), TBLangRefresh.Text)
                End If

                Dim eTheme As XmlElement = eSubMenu.Item("Theme")
                If eTheme IsNot Nothing Then
                    XMLLoadLocaleMenu(eTheme.Item("Default"), TBThemeDef.Text)
                    XMLLoadLocaleMenu(eTheme.Item("Save"), TBThemeSave.Text)
                    XMLLoadLocaleMenu(eTheme.Item("Refresh"), TBThemeRefresh.Text)
                    XMLLoadLocaleMenu(eTheme.Item("LoadComptability"), TBThemeLoadComptability.Text)
                End If

                Dim eConvert As XmlElement = eSubMenu.Item("Convert")
                If eConvert IsNot Nothing Then
                    XMLLoadLocaleMenu(eConvert.Item("Long"), POBLong.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Short"), POBShort.Text)
                    XMLLoadLocaleMenu(eConvert.Item("LongShort"), POBLongShort.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Hidden"), POBHidden.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Visible"), POBVisible.Text)
                    XMLLoadLocaleMenu(eConvert.Item("HiddenVisible"), POBHiddenVisible.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Relabel"), POBModify.Text)
                    XMLLoadLocaleMenu(eConvert.Item("Mirror"), POBMirror.Text)
                End If

                Dim eWAV As XmlElement = eSubMenu.Item("WAV")
                If eWAV IsNot Nothing Then
                    XMLLoadLocaleMenu(eWAV.Item("MultiSelection"), CWAVMultiSelect.Text)
                    XMLLoadLocaleMenu(eWAV.Item("Synchronize"), CWAVChangeLabel.Text)
                End If

                Dim eBeat As XmlElement = eSubMenu.Item("Beat")
                If eBeat IsNot Nothing Then
                    XMLLoadLocaleMenu(eBeat.Item("Absolute"), CBeatPreserve.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Measure"), CBeatMeasure.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Cut"), CBeatCut.Text)
                    XMLLoadLocaleMenu(eBeat.Item("Scale"), CBeatScale.Text)
                End If
            End If

            Dim eOptionsPanel As XmlElement = Root.Item("OptionsPanel")
            If eOptionsPanel IsNot Nothing Then

                Dim eHeader As XmlElement = eOptionsPanel.Item("Header")
                If eHeader IsNot Nothing Then
                    XMLLoadLocale(eHeader.Item("Header"), POHeaderSwitch.Text)
                    XMLLoadLocale(eHeader.Item("Title"), Label3.Text)
                    XMLLoadLocale(eHeader.Item("Artist"), Label4.Text)
                    XMLLoadLocale(eHeader.Item("Genre"), Label2.Text)
                    XMLLoadLocale(eHeader.Item("BPM"), Label9.Text)
                    XMLLoadLocale(eHeader.Item("Player"), Label8.Text)
                    XMLLoadLocale(eHeader.Item("Rank"), Label10.Text)
                    XMLLoadLocale(eHeader.Item("PlayLevel"), Label6.Text)
                    XMLLoadLocale(eHeader.Item("SubTitle"), Label15.Text)
                    XMLLoadLocale(eHeader.Item("SubArtist"), Label17.Text)
                    'XMLLoadLocale(eHeader.Item("Maker"), Label14.Text)
                    XMLLoadLocale(eHeader.Item("StageFile"), Label16.Text)
                    XMLLoadLocale(eHeader.Item("Banner"), Label12.Text)
                    XMLLoadLocale(eHeader.Item("BackBMP"), Label11.Text)
                    'XMLLoadLocale(eHeader.Item("MidiFile"), Label18.Text)
                    XMLLoadLocale(eHeader.Item("Difficulty"), Label21.Text)
                    XMLLoadLocale(eHeader.Item("ExRank"), Label23.Text)
                    XMLLoadLocale(eHeader.Item("Total"), Label20.Text)
                    'XMLLoadLocale(eHeader.Item("VolWav"), Label22.Text)
                    XMLLoadLocale(eHeader.Item("Comment"), Label19.Text)
                    'XMLLoadLocale(eHeader.Item("LnType"), Label13.Text)
                    XMLLoadLocale(eHeader.Item("LnObj"), Label24.Text)
                    XMLLoadLocale(eHeader.Item("LandMine"), Label26.Text)
                    XMLLoadLocale(eHeader.Item("MissBMP"), Label27.Text)

                    RemoveHandler CHPlayer.SelectedIndexChanged, AddressOf CHPlayer_SelectedIndexChanged
                    XMLLoadLocale(eHeader.Item("Player1"), CHPlayer.Items.Item(0))
                    XMLLoadLocale(eHeader.Item("Player2"), CHPlayer.Items.Item(1))
                    XMLLoadLocale(eHeader.Item("Player3"), CHPlayer.Items.Item(2))
                    AddHandler CHPlayer.SelectedIndexChanged, AddressOf CHPlayer_SelectedIndexChanged

                    RemoveHandler CHRank.SelectedIndexChanged, AddressOf THGenre_TextChanged
                    XMLLoadLocale(eHeader.Item("Rank0"), CHRank.Items.Item(0))
                    XMLLoadLocale(eHeader.Item("Rank1"), CHRank.Items.Item(1))
                    XMLLoadLocale(eHeader.Item("Rank2"), CHRank.Items.Item(2))
                    XMLLoadLocale(eHeader.Item("Rank3"), CHRank.Items.Item(3))
                    XMLLoadLocale(eHeader.Item("Rank4"), CHRank.Items.Item(4))
                    AddHandler CHRank.SelectedIndexChanged, AddressOf THGenre_TextChanged

                    RemoveHandler CHDifficulty.SelectedIndexChanged, AddressOf THGenre_TextChanged
                    XMLLoadLocale(eHeader.Item("Difficulty0"), CHDifficulty.Items.Item(0))
                    XMLLoadLocale(eHeader.Item("Difficulty1"), CHDifficulty.Items.Item(1))
                    XMLLoadLocale(eHeader.Item("Difficulty2"), CHDifficulty.Items.Item(2))
                    XMLLoadLocale(eHeader.Item("Difficulty3"), CHDifficulty.Items.Item(3))
                    XMLLoadLocale(eHeader.Item("Difficulty4"), CHDifficulty.Items.Item(4))
                    XMLLoadLocale(eHeader.Item("Difficulty5"), CHDifficulty.Items.Item(5))
                    AddHandler CHDifficulty.SelectedIndexChanged, AddressOf THGenre_TextChanged
                End If

                Dim eGrid As XmlElement = eOptionsPanel.Item("Grid")
                If eGrid IsNot Nothing Then
                    XMLLoadLocale(eGrid.Item("Title"), POGridSwitch.Text)
                    XMLLoadLocale(eGrid.Item("Snap"), CGSnap.Text)
                    XMLLoadLocale(eGrid.Item("BCols"), Label1.Text)
                    XMLLoadLocale(eGrid.Item("DisableVertical"), CGDisableVertical.Text)
                    XMLLoadLocale(eGrid.Item("Scroll"), Label5.Text)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockLeft"), cVSLockL)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockMiddle"), cVSLock)
                    XMLLoadLocaleToolTipUniversal(eGrid.Item("LockRight"), cVSLockR)
                End If

                Dim eWaveForm As XmlElement = eOptionsPanel.Item("WaveForm")
                If eWaveForm IsNot Nothing Then
                    XMLLoadLocale(eWaveForm.Item("Title"), POWaveFormSwitch.Text)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Load"), BWLoad)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Clear"), BWClear)
                    XMLLoadLocaleToolTipUniversal(eWaveForm.Item("Lock"), BWLock)
                End If

                Dim eWAV As XmlElement = eOptionsPanel.Item("WAV")
                If eWAV IsNot Nothing Then
                    XMLLoadLocale(eWAV.Item("Title"), POWAVSwitch.Text)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveUp"), BBMPUp)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveDown"), BBMPDown)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Browse"), BBMPBrowse)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Remove"), BBMPRemove)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveUp"), BWAVUp)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("MoveDown"), BWAVDown)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Browse"), BWAVBrowse)
                    XMLLoadLocaleToolTipUniversal(eWAV.Item("Remove"), BWAVRemove)
                End If

                Dim eBMP As XmlElement = eOptionsPanel.Item("BMP")
                If eBMP IsNot Nothing Then
                    XMLLoadLocale(eBMP.Item("Title"), POBMPSwitch.Text)
                End If

                XMLLoadLocale(eOptionsPanel.Item("Beat"), POBeatSwitch.Text)
                XMLLoadLocale(eOptionsPanel.Item("Beat.Apply"), BBeatApply.Text)
                XMLLoadLocale(eOptionsPanel.Item("Beat.Apply"), BBeatApplyV.Text)
                XMLLoadLocale(eOptionsPanel.Item("Expansion"), POExpansionSwitch.Text)
            End If

            Dim eMessages As XmlElement = Root.Item("Messages")
            If eMessages IsNot Nothing Then
                XMLLoadLocale(eMessages.Item("Err"), Strings.Messages.Err)
                XMLLoadLocale(eMessages.Item("SaveOnExit"), Strings.Messages.SaveOnExit)
                XMLLoadLocale(eMessages.Item("SaveOnExit1"), Strings.Messages.SaveOnExit1)
                XMLLoadLocale(eMessages.Item("SaveOnExit2"), Strings.Messages.SaveOnExit2)
                XMLLoadLocale(eMessages.Item("PromptEnter"), Strings.Messages.PromptEnter)
                XMLLoadLocale(eMessages.Item("PromptEnterNumeric"), Strings.Messages.PromptEnterNumeric)
                XMLLoadLocale(eMessages.Item("PromptEnterBPM"), Strings.Messages.PromptEnterBPM)
                XMLLoadLocale(eMessages.Item("PromptEnterSTOP"), Strings.Messages.PromptEnterSTOP)
                XMLLoadLocale(eMessages.Item("PromptEnterSCROLL"), Strings.Messages.PromptEnterSCROLL)
                XMLLoadLocale(eMessages.Item("PromptSlashValue"), Strings.Messages.PromptSlashValue)
                XMLLoadLocale(eMessages.Item("InvalidLabel"), Strings.Messages.InvalidLabel)
                XMLLoadLocale(eMessages.Item("CannotFind"), Strings.Messages.CannotFind)
                XMLLoadLocale(eMessages.Item("PleaseRespecifyPath"), Strings.Messages.PleaseRespecifyPath)
                XMLLoadLocale(eMessages.Item("PlayerNotFound"), Strings.Messages.PlayerNotFound)
                XMLLoadLocale(eMessages.Item("PreviewDelError"), Strings.Messages.PreviewDelError)
                XMLLoadLocale(eMessages.Item("NegativeFactorError"), Strings.Messages.NegativeFactorError)
                XMLLoadLocale(eMessages.Item("NegativeDivisorError"), Strings.Messages.NegativeDivisorError)
                XMLLoadLocale(eMessages.Item("PreferencePostpone"), Strings.Messages.PreferencePostpone)
                XMLLoadLocale(eMessages.Item("EraserObsolete"), Strings.Messages.EraserObsolete)
                XMLLoadLocale(eMessages.Item("SaveWarning"), Strings.Messages.SaveWarning)
                XMLLoadLocale(eMessages.Item("NoteOverlapError"), Strings.Messages.NoteOverlapError)
                XMLLoadLocale(eMessages.Item("BPMOverflowError"), Strings.Messages.BPMOverflowError)
                XMLLoadLocale(eMessages.Item("STOPOverflowError"), Strings.Messages.STOPOverflowError)
                XMLLoadLocale(eMessages.Item("SCROLLOverflowError"), Strings.Messages.SCROLLOverflowError)
                XMLLoadLocale(eMessages.Item("SavedFileWillContainErrors"), Strings.Messages.SavedFileWillContainErrors)
                XMLLoadLocale(eMessages.Item("FileAssociationPrompt"), Strings.Messages.FileAssociationPrompt)
                XMLLoadLocale(eMessages.Item("FileAssociationError"), Strings.Messages.FileAssociationError)
                XMLLoadLocale(eMessages.Item("RestoreDefaultSettings"), Strings.Messages.RestoreDefaultSettings)
                XMLLoadLocale(eMessages.Item("RestoreAutosavedFile"), Strings.Messages.RestoreAutosavedFile)
            End If

            Dim eFileType As XmlElement = Root.Item("FileType")
            If eFileType IsNot Nothing Then
                XMLLoadLocale(eFileType.Item("_all"), Strings.FileType._all)
                XMLLoadLocale(eFileType.Item("_bms"), Strings.FileType._bms)
                XMLLoadLocale(eFileType.Item("BMS"), Strings.FileType.BMS)
                XMLLoadLocale(eFileType.Item("BME"), Strings.FileType.BME)
                XMLLoadLocale(eFileType.Item("BML"), Strings.FileType.BML)
                XMLLoadLocale(eFileType.Item("PMS"), Strings.FileType.PMS)
                XMLLoadLocale(eFileType.Item("TXT"), Strings.FileType.TXT)
                XMLLoadLocale(eFileType.Item("SM"), Strings.FileType.SM)
                XMLLoadLocale(eFileType.Item("IBMSC"), Strings.FileType.IBMSC)
                XMLLoadLocale(eFileType.Item("XML"), Strings.FileType.XML)
                XMLLoadLocale(eFileType.Item("THEME_XML"), Strings.FileType.THEME_XML)
                XMLLoadLocale(eFileType.Item("TH"), Strings.FileType.TH)
                XMLLoadLocale(eFileType.Item("_audio"), Strings.FileType._audio)
                XMLLoadLocale(eFileType.Item("_wave"), Strings.FileType._wave)
                XMLLoadLocale(eFileType.Item("WAV"), Strings.FileType.WAV)
                XMLLoadLocale(eFileType.Item("OGG"), Strings.FileType.OGG)
                XMLLoadLocale(eFileType.Item("MP3"), Strings.FileType.MP3)
                XMLLoadLocale(eFileType.Item("MID"), Strings.FileType.MID)
                XMLLoadLocale(eFileType.Item("_image"), Strings.FileType._image)
                XMLLoadLocale(eFileType.Item("_movie"), Strings.FileType._movie)
                XMLLoadLocale(eFileType.Item("BMP"), Strings.FileType.BMP)
                XMLLoadLocale(eFileType.Item("PNG"), Strings.FileType.PNG)
                XMLLoadLocale(eFileType.Item("JPG"), Strings.FileType.JPG)
                XMLLoadLocale(eFileType.Item("GIF"), Strings.FileType.GIF)
                XMLLoadLocale(eFileType.Item("MPG"), Strings.FileType.MPG)
                XMLLoadLocale(eFileType.Item("AVI"), Strings.FileType.AVI)
                XMLLoadLocale(eFileType.Item("MP4"), Strings.FileType.MP4)
                XMLLoadLocale(eFileType.Item("WMV"), Strings.FileType.WMV)
                XMLLoadLocale(eFileType.Item("WEBM"), Strings.FileType.WEBM)
                XMLLoadLocale(eFileType.Item("EXE"), Strings.FileType.EXE)
            End If

            Dim eStatistics As XmlElement = Root.Item("Statistics")
            If eStatistics IsNot Nothing Then
                XMLLoadLocale(eStatistics.Item("Title"), Strings.fStatistics.Title)
                XMLLoadLocale(eStatistics.Item("lBPM"), Strings.fStatistics.lBPM)
                XMLLoadLocale(eStatistics.Item("lSTOP"), Strings.fStatistics.lSTOP)
                XMLLoadLocale(eStatistics.Item("lSCROLL"), Strings.fStatistics.lSCROLL)
                XMLLoadLocale(eStatistics.Item("lA"), Strings.fStatistics.lA)
                XMLLoadLocale(eStatistics.Item("lD"), Strings.fStatistics.lD)
                XMLLoadLocale(eStatistics.Item("lBGM"), Strings.fStatistics.lBGM)
                XMLLoadLocale(eStatistics.Item("lTotal"), Strings.fStatistics.lTotal)
                XMLLoadLocale(eStatistics.Item("lShort"), Strings.fStatistics.lShort)
                XMLLoadLocale(eStatistics.Item("lLong"), Strings.fStatistics.lLong)
                XMLLoadLocale(eStatistics.Item("lLnObj"), Strings.fStatistics.lLnObj)
                XMLLoadLocale(eStatistics.Item("lHidden"), Strings.fStatistics.lHidden)
                XMLLoadLocale(eStatistics.Item("lErrors"), Strings.fStatistics.lErrors)
            End If

            Dim ePlayerOptions As XmlElement = Root.Item("PlayerOptions")
            If ePlayerOptions IsNot Nothing Then
                XMLLoadLocale(ePlayerOptions.Item("Title"), Strings.fopPlayer.Title)
                XMLLoadLocale(ePlayerOptions.Item("Add"), Strings.fopPlayer.Add)
                XMLLoadLocale(ePlayerOptions.Item("Remove"), Strings.fopPlayer.Remove)
                XMLLoadLocale(ePlayerOptions.Item("Path"), Strings.fopPlayer.Path)
                XMLLoadLocale(ePlayerOptions.Item("PlayFromBeginning"), Strings.fopPlayer.PlayFromBeginning)
                XMLLoadLocale(ePlayerOptions.Item("PlayFromHere"), Strings.fopPlayer.PlayFromHere)
                XMLLoadLocale(ePlayerOptions.Item("StopPlaying"), Strings.fopPlayer.StopPlaying)
                XMLLoadLocale(ePlayerOptions.Item("References"), Strings.fopPlayer.References)
                XMLLoadLocale(ePlayerOptions.Item("DirectoryOfApp"), Strings.fopPlayer.DirectoryOfApp)
                XMLLoadLocale(ePlayerOptions.Item("CurrMeasure"), Strings.fopPlayer.CurrMeasure)
                XMLLoadLocale(ePlayerOptions.Item("FileName"), Strings.fopPlayer.FileName)
                XMLLoadLocale(ePlayerOptions.Item("RestoreDefault"), Strings.fopPlayer.RestoreDefault)
            End If

            Dim eVisualOptions As XmlElement = Root.Item("VisualOptions")
            If eVisualOptions IsNot Nothing Then
                XMLLoadLocale(eVisualOptions.Item("Title"), Strings.fopVisual.Title)
                XMLLoadLocale(eVisualOptions.Item("Width"), Strings.fopVisual.Width)
                XMLLoadLocale(eVisualOptions.Item("Caption"), Strings.fopVisual.Caption)
                XMLLoadLocale(eVisualOptions.Item("Note"), Strings.fopVisual.Note)
                XMLLoadLocale(eVisualOptions.Item("Label"), Strings.fopVisual.Label)
                XMLLoadLocale(eVisualOptions.Item("LongNote"), Strings.fopVisual.LongNote)
                XMLLoadLocale(eVisualOptions.Item("LongNoteLabel"), Strings.fopVisual.LongNoteLabel)
                XMLLoadLocale(eVisualOptions.Item("Bg"), Strings.fopVisual.Bg)
                XMLLoadLocale(eVisualOptions.Item("ColumnCaption"), Strings.fopVisual.ColumnCaption)
                XMLLoadLocale(eVisualOptions.Item("ColumnCaptionFont"), Strings.fopVisual.ColumnCaptionFont)
                XMLLoadLocale(eVisualOptions.Item("Background"), Strings.fopVisual.Background)
                XMLLoadLocale(eVisualOptions.Item("Grid"), Strings.fopVisual.Grid)
                XMLLoadLocale(eVisualOptions.Item("SubGrid"), Strings.fopVisual.SubGrid)
                XMLLoadLocale(eVisualOptions.Item("VerticalLine"), Strings.fopVisual.VerticalLine)
                XMLLoadLocale(eVisualOptions.Item("MeasureBarLine"), Strings.fopVisual.MeasureBarLine)
                XMLLoadLocale(eVisualOptions.Item("BGMWaveform"), Strings.fopVisual.BGMWaveform)
                XMLLoadLocale(eVisualOptions.Item("NoteHeight"), Strings.fopVisual.NoteHeight)
                XMLLoadLocale(eVisualOptions.Item("NoteLabel"), Strings.fopVisual.NoteLabel)
                XMLLoadLocale(eVisualOptions.Item("MeasureLabel"), Strings.fopVisual.MeasureLabel)
                XMLLoadLocale(eVisualOptions.Item("LabelVerticalShift"), Strings.fopVisual.LabelVerticalShift)
                XMLLoadLocale(eVisualOptions.Item("LabelHorizontalShift"), Strings.fopVisual.LabelHorizontalShift)
                XMLLoadLocale(eVisualOptions.Item("LongNoteLabelHorizontalShift"),
                              Strings.fopVisual.LongNoteLabelHorizontalShift)
                XMLLoadLocale(eVisualOptions.Item("HiddenNoteOpacity"), Strings.fopVisual.HiddenNoteOpacity)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnMouseOver"), Strings.fopVisual.NoteBorderOnMouseOver)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnSelection"), Strings.fopVisual.NoteBorderOnSelection)
                XMLLoadLocale(eVisualOptions.Item("NoteBorderOnAdjustingLength"),
                              Strings.fopVisual.NoteBorderOnAdjustingLength)
                XMLLoadLocale(eVisualOptions.Item("SelectionBoxBorder"), Strings.fopVisual.SelectionBoxBorder)
                XMLLoadLocale(eVisualOptions.Item("TSCursor"), Strings.fopVisual.TSCursor)
                XMLLoadLocale(eVisualOptions.Item("TSSplitter"), Strings.fopVisual.TSSplitter)
                XMLLoadLocale(eVisualOptions.Item("TSCursorSensitivity"), Strings.fopVisual.TSCursorSensitivity)
                XMLLoadLocale(eVisualOptions.Item("TSMouseOverBorder"), Strings.fopVisual.TSMouseOverBorder)
                XMLLoadLocale(eVisualOptions.Item("TSFill"), Strings.fopVisual.TSFill)
                XMLLoadLocale(eVisualOptions.Item("TSBPM"), Strings.fopVisual.TSBPM)
                XMLLoadLocale(eVisualOptions.Item("TSBPMFont"), Strings.fopVisual.TSBPMFont)
                XMLLoadLocale(eVisualOptions.Item("MiddleSensitivity"), Strings.fopVisual.MiddleSensitivity)
            End If

            Dim eGeneralOptions As XmlElement = Root.Item("GeneralOptions")
            If eGeneralOptions IsNot Nothing Then
                XMLLoadLocale(eGeneralOptions.Item("Title"), Strings.fopGeneral.Title)
                XMLLoadLocale(eGeneralOptions.Item("MouseWheel"), Strings.fopGeneral.MouseWheel)
                XMLLoadLocale(eGeneralOptions.Item("TextEncoding"), Strings.fopGeneral.TextEncoding)
                XMLLoadLocale(eGeneralOptions.Item("PageUpDown"), Strings.fopGeneral.PageUpDown)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButton"), Strings.fopGeneral.MiddleButton)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButtonAuto"), Strings.fopGeneral.MiddleButtonAuto)
                XMLLoadLocale(eGeneralOptions.Item("MiddleButtonDrag"), Strings.fopGeneral.MiddleButtonDrag)
                XMLLoadLocale(eGeneralOptions.Item("AssociateFileType"), Strings.fopGeneral.AssociateFileType)
                XMLLoadLocale(eGeneralOptions.Item("MaxGridPartition"), Strings.fopGeneral.MaxGridPartition)
                XMLLoadLocale(eGeneralOptions.Item("BeepWhileSaved"), Strings.fopGeneral.BeepWhileSaved)
                XMLLoadLocale(eGeneralOptions.Item("ExtendBPM"), Strings.fopGeneral.ExtendBPM)
                XMLLoadLocale(eGeneralOptions.Item("ExtendSTOP"), Strings.fopGeneral.ExtendSTOP)
                XMLLoadLocale(eGeneralOptions.Item("AutoFocusOnMouseEnter"), Strings.fopGeneral.AutoFocusOnMouseEnter)
                XMLLoadLocale(eGeneralOptions.Item("DisableFirstClick"), Strings.fopGeneral.DisableFirstClick)
                XMLLoadLocale(eGeneralOptions.Item("AutoSave"), Strings.fopGeneral.AutoSave)
                XMLLoadLocale(eGeneralOptions.Item("minutes"), Strings.fopGeneral.minutes)
                XMLLoadLocale(eGeneralOptions.Item("StopPreviewOnClick"), Strings.fopGeneral.StopPreviewOnClick)
            End If

            Dim eFind As XmlElement = Root.Item("Find")
            If eFind IsNot Nothing Then
                XMLLoadLocale(eFind.Item("NoteRange"), Strings.fFind.NoteRange)
                XMLLoadLocale(eFind.Item("MeasureRange"), Strings.fFind.MeasureRange)
                XMLLoadLocale(eFind.Item("LabelRange"), Strings.fFind.LabelRange)
                XMLLoadLocale(eFind.Item("ValueRange"), Strings.fFind.ValueRange)
                XMLLoadLocale(eFind.Item("to"), Strings.fFind.to_)
                XMLLoadLocale(eFind.Item("Selected"), Strings.fFind.Selected)
                XMLLoadLocale(eFind.Item("UnSelected"), Strings.fFind.UnSelected)
                XMLLoadLocale(eFind.Item("ShortNote"), Strings.fFind.ShortNote)
                XMLLoadLocale(eFind.Item("LongNote"), Strings.fFind.LongNote)
                XMLLoadLocale(eFind.Item("Hidden"), Strings.fFind.Hidden)
                XMLLoadLocale(eFind.Item("Visible"), Strings.fFind.Visible)
                XMLLoadLocale(eFind.Item("Column"), Strings.fFind.Column)
                XMLLoadLocale(eFind.Item("SelectAll"), Strings.fFind.SelectAll)
                XMLLoadLocale(eFind.Item("SelectInverse"), Strings.fFind.SelectInverse)
                XMLLoadLocale(eFind.Item("UnselectAll"), Strings.fFind.UnselectAll)
                XMLLoadLocale(eFind.Item("Operation"), Strings.fFind.Operation)
                XMLLoadLocale(eFind.Item("ReplaceWithLabel"), Strings.fFind.ReplaceWithLabel)
                XMLLoadLocale(eFind.Item("ReplaceWithValue"), Strings.fFind.ReplaceWithValue)
                XMLLoadLocale(eFind.Item("Select"), Strings.fFind.Select_)
                XMLLoadLocale(eFind.Item("Unselect"), Strings.fFind.Unselect_)
                XMLLoadLocale(eFind.Item("Delete"), Strings.fFind.Delete_)
                XMLLoadLocale(eFind.Item("Close"), Strings.fFind.Close_)
            End If

            Dim eImportSM As XmlElement = Root.Item("ImportSM")
            If eImportSM IsNot Nothing Then
                XMLLoadLocale(eImportSM.Item("Title"), Strings.fImportSM.Title)
                XMLLoadLocale(eImportSM.Item("Difficulty"), Strings.fImportSM.Difficulty)
                XMLLoadLocale(eImportSM.Item("Note"), Strings.fImportSM.Note)
            End If

            Dim eFileAssociation As XmlElement = Root.Item("FileAssociation")
            If eFileAssociation IsNot Nothing Then
                XMLLoadLocale(eFileAssociation.Item("BMS"), Strings.FileAssociation.BMS)
                XMLLoadLocale(eFileAssociation.Item("BME"), Strings.FileAssociation.BME)
                XMLLoadLocale(eFileAssociation.Item("BML"), Strings.FileAssociation.BML)
                XMLLoadLocale(eFileAssociation.Item("PMS"), Strings.FileAssociation.PMS)
                XMLLoadLocale(eFileAssociation.Item("IBMSC"), Strings.FileAssociation.IBMSC)
                XMLLoadLocale(eFileAssociation.Item("Open"), Strings.FileAssociation.Open)
                XMLLoadLocale(eFileAssociation.Item("Preview"), Strings.FileAssociation.Preview)
                XMLLoadLocale(eFileAssociation.Item("ViewCode"), Strings.FileAssociation.ViewCode)
            End If

            _dispLang = Path.Replace(My.Application.Info.DirectoryPath & "\", "")

        Catch ex As Exception
            MsgBox(Path & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Exclamation)

        Finally
            If FileStream IsNot Nothing Then FileStream.Close()

            POHeaderPart2.Visible = xPOHeaderPart2
            POGridPart2.Visible = xPOGridPart2
            POWaveFormPart2.Visible = xPOWaveFormPart2
        End Try

        'File.Delete(xTempFileName)
    End Sub

    Private Sub LoadThemeComptability(xPath As String)
        Try
            Dim xStrLine() As String = Split(My.Computer.FileSystem.ReadAllText(xPath), vbCrLf)
            If xStrLine(0).Trim <> "iBMSC Configuration Settings Format" And xStrLine(0).Trim <> "iBMSC Theme Format" _
                Then Exit Sub

            Dim xW1 = ""
            Dim xW2 = ""
            Dim column = Columns.column

            For Each xLine As String In xStrLine
                xW1 = UCase(Mid(xLine, 1, InStr(xLine, "=") - 1))
                xW2 = Mid(xLine, InStr(xLine, "=") + 1)

                Select Case xW1
                    Case "VOTITLE" : _theme.ColumnTitle.Color = Color.FromArgb(Val(xW2))
                    Case "VOTITLEFONT" : _theme.ColumnTitleFont = StringToFont(xW2, Me.Font)
                    Case "VOBG" : _theme.Bg.Color = Color.FromArgb(Val(xW2))
                    Case "VOGRID" : _theme.pGrid.Color = Color.FromArgb(Val(xW2))
                    Case "VOSUB" : _theme.pSub.Color = Color.FromArgb(Val(xW2))
                    Case "VOVLINE" : _theme.pVLine.Color = Color.FromArgb(Val(xW2))
                    Case "VOMLINE" : _theme.pMLine.Color = Color.FromArgb(Val(xW2))
                    Case "VOBGMWAV"
                        _theme.pBGMWav.Color = Color.FromArgb(Val(xW2))
                        TWTransparency.Value = _theme.pBGMWav.Color.A
                        TWTransparency2.Value = _theme.pBGMWav.Color.A
                        TWSaturation.Value = _theme.pBGMWav.Color.GetSaturation*1000
                        TWSaturation2.Value = _theme.pBGMWav.Color.GetSaturation*1000
                    Case "VOSELBOX" : _theme.SelBox.Color = Color.FromArgb(Val(xW2))
                    Case "VOPECURSOR" : _theme.PECursor.Color = Color.FromArgb(Val(xW2))
                    Case "VOPEHALF" : _theme.PEHalf.Color = Color.FromArgb(Val(xW2))
                    Case "VOPEDELTAMOUSEOVER" : _theme.PEDeltaMouseOver = Val(xW2)
                    Case "VOPEMOUSEOVER" : _theme.PEMouseOver.Color = Color.FromArgb(Val(xW2))
                    Case "VOPESEL" : _theme.PESel.Color = Color.FromArgb(Val(xW2))
                    Case "VOPEBPM" : _theme.PEBPM.Color = Color.FromArgb(Val(xW2))
                    Case "VOPEBPMFONT" : _theme.PEBPMFont = StringToFont(xW2, Me.Font)
                    Case "VKHEIGHT" : _theme.NoteHeight = Val(xW2)
                    Case "VKFONT" : _theme.kFont = StringToFont(xW2, Me.Font)
                    Case "VKMFONT" : _theme.kMFont = StringToFont(xW2, Me.Font)
                    Case "VKLABELVSHIFT" : _theme.kLabelVShift = Val(xW2)
                    Case "VKLABELHSHIFT" : _theme.kLabelHShift = Val(xW2)
                    Case "VKLABELHSHIFTL" : _theme.kLabelHShiftL = Val(xW2)
                    Case "VKMOUSEOVER" : _theme.kMouseOver.Color = Color.FromArgb(Val(xW2))
                    Case "VKMOUSEOVERE " : _theme.kMouseOverE.Color = Color.FromArgb(Val(xW2))
                    Case "VKSELECTED" : _theme.kSelected.Color = Color.FromArgb(Val(xW2))
                        'Case "VKHIDTRANSPARENCY" : vo.kOpacity = Val(xW2)

                    Case "KLENGTH"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).Width = Val(xE(i))
                        Next

                    Case "KTITLE"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).Title = xE(i)
                        Next

                    Case "KCOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).setNoteColor(Val(xE(i)))
                        Next

                    Case "KCOLORL"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).setLNoteColor(Val(xE(i)))
                        Next

                    Case "KFORECOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).cText = Color.FromArgb(Val(xE(i)))
                        Next

                    Case "KFORECOLORL"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).cLText = Color.FromArgb(Val(xE(i)))
                        Next

                    Case "KBGCOLOR"
                        Dim xE() As String = LoadThemeComptability_SplitStringInto26Parts(xW2)
                        For i = 0 To 26
                            column(i).cBG = Color.FromArgb(Val(xE(i)))
                        Next

                End Select

            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, Strings.Messages.Err)

        Finally
            UpdateColumnsX()

        End Try
    End Sub

    Private Function LoadThemeComptability_SplitStringInto26Parts(xLine As String) As String()
        Dim xE() As String = Split(xLine, ",")
        ReDim Preserve xE(26)
        Return xE
    End Function

    Private Sub LoadLang(sender As Object, e As EventArgs)
        Dim xFN2 As String = sender.ToolTipText
        'ReadLanguagePack(xFN2)
        LoadLocale(xFN2)
    End Sub

    Private Sub LoadLocaleXML(xStr As FileInfo)
        Dim d As New XmlDocument
        Dim fs As New FileStream(xStr.FullName, FileMode.Open, FileAccess.Read)

        Try
            d.Load(fs)
            Dim xName As String = d.Item("iBMSC.Locale").GetAttribute("Name")
            If xName = "" Then xName = xStr.Name

            cmnLanguage.Items.Add(xName, Nothing, AddressOf LoadLang)
            cmnLanguage.Items(cmnLanguage.Items.Count - 1).ToolTipText = xStr.FullName

        Catch ex As Exception
            MsgBox(xStr.FullName & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Exclamation)

        End Try

        fs.Close()
    End Sub

    Private Sub LoadTheme(sender As Object, e As EventArgs)
        'If Not File.Exists(My.Application.Info.DirectoryPath & "\Data\" & sender.Text) Then Exit Sub
        'SaveTheme = True
        'LoadCFF(My.Computer.FileSystem.ReadAllText(My.Application.Info.DirectoryPath & "\Theme\" & sender.Text, System.Text.Encoding.Unicode))
        LoadSettings(My.Application.Info.DirectoryPath & "\Data\" & sender.Text)
        RefreshPanelAll()
    End Sub
End Class
