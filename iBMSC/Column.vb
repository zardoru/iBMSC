Namespace Editor
    Public Structure Column
        Private _Width As Integer
        Private _isNoteCol As Boolean
        Private _isVisible As Boolean
        Private _isEnabledAfterAll As Boolean

        Public Property Width() As Integer
            Get
                Return _Width
            End Get
            Set(ByVal value As Integer)
                _Width = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public Property isVisible() As Boolean
            Get
                Return _isVisible
            End Get
            Set(ByVal value As Boolean)
                _isVisible = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public Property isNoteCol() As Boolean
            Get
                Return _isNoteCol
            End Get
            Set(ByVal value As Boolean)
                _isNoteCol = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public ReadOnly Property isEnabledAfterAll() As Boolean
            Get
                Return _isEnabledAfterAll
            End Get
        End Property

        'Private Visibility As ColumnVisibility
        'Public Property Visible() As Boolean
        '    Get
        '        Return Not (Visibility And ColumnVisibility.Invisible)
        '    End Get
        '    Set(ByVal value As Boolean)
        '        If value Then
        '            Visibility = Visibility Or ColumnVisibility.Invisible
        '        Else
        '            Visibility = Visibility And (ColumnVisibility.Decorative Or ColumnVisibility.Zero_Width)
        '        End If
        '    End Set
        'End Property

        Public Left As Integer
        Public Title As String
        Public isNumeric As Boolean
        Public isSound As Boolean
        Public Identifier As Integer

        Public cNote As Integer
        Public cText As Color
        Public cLNote As Integer
        Public cLText As Color
        Public cBG As Color

        Private cCacheB As Integer
        Private cCacheD As Integer
        Private cCacheLB As Integer
        Private cCacheLD As Integer

        Public Function getBright(ByVal opacity As Single) As Color
            Return Color.FromArgb((CInt(((cCacheB >> 24) And &HFF) * opacity) << 24) Or (cCacheB And &HFFFFFF))
        End Function
        Public Function getDark(ByVal opacity As Single) As Color
            Return Color.FromArgb((CInt(((cCacheD >> 24) And &HFF) * opacity) << 24) Or (cCacheD And &HFFFFFF))
        End Function
        Public Function getLongBright(ByVal opacity As Single) As Color
            Return Color.FromArgb((CInt(((cCacheLB >> 24) And &HFF) * opacity) << 24) Or (cCacheLB And &HFFFFFF))
        End Function
        Public Function getLongDark(ByVal opacity As Single) As Color
            Return Color.FromArgb((CInt(((cCacheLD >> 24) And &HFF) * opacity) << 24) Or (cCacheLD And &HFFFFFF))
        End Function

        Public Sub setNoteColor(ByVal c As Integer)
            cNote = c
            'cCacheB = (c And &HFF000000) Or &H808080 Or ((c And &HFFFFFF) >> 1)
            'cCacheD = (c And &HFF000000) Or ((c And &HFEFEFE) >> 1)
            cCacheB = AdjustBrightness(Color.FromArgb(c), 50, ((c >> 24) And &HFF) / 255).ToArgb
            cCacheD = AdjustBrightness(Color.FromArgb(c), -25, ((c >> 24) And &HFF) / 255).ToArgb
        End Sub
        Public Sub setLNoteColor(ByVal c As Integer)
            cLNote = c
            'cCacheLB = (c And &HFF000000) Or &H808080 Or ((c And &HFFFFFF) >> 1)
            'cCacheLD = (c And &HFF000000) Or ((c And &HFEFEFE) >> 1)
            cCacheLB = AdjustBrightness(Color.FromArgb(c), 50, ((c >> 24) And &HFF) / 255).ToArgb
            cCacheLD = AdjustBrightness(Color.FromArgb(c), -25, ((c >> 24) And &HFF) / 255).ToArgb
        End Sub

        Public Sub New(ByVal xLeft As Integer, ByVal xWidth As Integer, ByVal xTitle As String,
        ByVal xNoteCol As Boolean, ByVal xisNumeric As Boolean, ByVal xisSound As Boolean, ByVal xVisible As Boolean, ByVal xIdentifier As Integer,
        ByVal xcNote As Integer, ByVal xcText As Integer, ByVal xcLNote As Integer, ByVal xcLText As Integer, ByVal xcBG As Integer)
            Left = xLeft
            Title = xTitle
            isNumeric = xisNumeric
            isSound = xisSound
            Identifier = xIdentifier

            _Width = xWidth
            _isVisible = xVisible
            _isNoteCol = xNoteCol
            _isEnabledAfterAll = xVisible And xNoteCol And (xWidth <> 0)

            setNoteColor(xcNote)
            cText = Color.FromArgb(xcText)
            setLNoteColor(xcLNote)
            cLText = Color.FromArgb(xcLText)
            cBG = Color.FromArgb(xcBG)
        End Sub
    End Structure
End Namespace
