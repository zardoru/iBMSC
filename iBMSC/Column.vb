Namespace Editor
    Public Class Column
        Private _width As Integer
        Private _isNoteCol As Boolean
        Private _isVisible As Boolean
        Private _isEnabledAfterAll As Boolean

        Public Property Width As Integer
            Get
                Return _Width
            End Get
            Set
                _Width = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public Property IsVisible As Boolean
            Get
                Return _isVisible
            End Get
            Set
                _isVisible = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public Property IsNoteCol As Boolean
            Get
                Return _isNoteCol
            End Get
            Set
                _isNoteCol = value
                _isEnabledAfterAll = _isVisible And _isNoteCol And (_Width <> 0)
            End Set
        End Property

        Public ReadOnly Property IsEnabledAfterAll As Boolean
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
        Public IsNumeric As Boolean
        Public IsSound As Boolean
        Public BmsChannel As Integer

        Public CNote As Integer
        Public CText As Color
        Public CLNote As Integer
        Public CLText As Color
        Public CBg As Color

        Private _cCacheB As Integer
        Private _cCacheD As Integer
        Private _cCacheLb As Integer
        Private _cCacheLd As Integer

        Public Function GetBright(opacity As Single) As Color
            Return Color.FromArgb((CInt(((_cCacheB >> 24) And &HFF)*opacity) << 24) Or (_cCacheB And &HFFFFFF))
        End Function

        Public Function GetDark(opacity As Single) As Color
            Return Color.FromArgb((CInt(((_cCacheD >> 24) And &HFF)*opacity) << 24) Or (_cCacheD And &HFFFFFF))
        End Function

        Public Function GetLongBright(opacity As Single) As Color
            Return Color.FromArgb((CInt(((_cCacheLb >> 24) And &HFF)*opacity) << 24) Or (_cCacheLb And &HFFFFFF))
        End Function

        Public Function GetLongDark(opacity As Single) As Color
            Return Color.FromArgb((CInt(((_cCacheLd >> 24) And &HFF)*opacity) << 24) Or (_cCacheLd And &HFFFFFF))
        End Function

        Public Sub SetNoteColor(c As Integer)
            cNote = c
            'cCacheB = (c And &HFF000000) Or &H808080 Or ((c And &HFFFFFF) >> 1)
            'cCacheD = (c And &HFF000000) Or ((c And &HFEFEFE) >> 1)
            _cCacheB = AdjustBrightness(Color.FromArgb(c), 50, ((c >> 24) And &HFF)/255).ToArgb
            _cCacheD = AdjustBrightness(Color.FromArgb(c), - 25, ((c >> 24) And &HFF)/255).ToArgb
        End Sub

        Public Sub SetLNoteColor(c As Integer)
            cLNote = c
            'cCacheLB = (c And &HFF000000) Or &H808080 Or ((c And &HFFFFFF) >> 1)
            'cCacheLD = (c And &HFF000000) Or ((c And &HFEFEFE) >> 1)
            _cCacheLb = AdjustBrightness(Color.FromArgb(c), 50, ((c >> 24) And &HFF)/255).ToArgb
            _cCacheLd = AdjustBrightness(Color.FromArgb(c), - 25, ((c >> 24) And &HFF)/255).ToArgb
        End Sub

        Public Sub New(xLeft As Integer, xWidth As Integer, xTitle As String,
                       xNoteCol As Boolean, xisNumeric As Boolean, xisSound As Boolean, xVisible As Boolean,
                       xIdentifier As Integer,
                       xcNote As Integer, xcText As Integer, xcLNote As Integer, xcLText As Integer, xcBg As Integer)
            Left = xLeft
            Title = xTitle
            isNumeric = xisNumeric
            isSound = xisSound
            BmsChannel = xIdentifier

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
    End Class
End Namespace
