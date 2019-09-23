Imports iBMSC.Editor

Public Enum TimeSelectLine
    None
    StartLine
    HalfLine
    EndLine
End Enum

Public Class TimeSelectState
    ' same, excpt without Editor.State.TimeSelect.
    Public StartPoint As Double = 192.0#
    Public EndPointLength As Double = 0.0#
    Public HalfPointLength As Double = 0.0#
    Public MouseOverLine As TimeSelectLine = TimeSelectLine.None
    Public Adjust As Boolean = False
    Public Notes() As Note = {} ' Editor.State.TimeSelect.K -> Notes
    Public PStart As Double = 192.0# ' az TODO: what's the difference? What's the P for?
    Public PLength As Double = 0.0#
    Public PHalf As Double = 0.0#

    Public Property EndPoint As Double
        Get
            Return StartPoint + EndPointLength
        End Get
        Set
            EndPointLength = value - StartPoint
        End Set
    End Property

    Public Property HalfPoint As Double
        Get
            Return StartPoint + HalfPointLength
        End Get
        Set
            HalfPointLength = value - StartPoint
        End Set
    End Property

    Public Sub ValidateSelection(maxvpos As Double)
        If StartPoint < 0 Then EndPointLength += StartPoint : HalfPointLength += StartPoint : StartPoint = 0

        If StartPoint > maxvpos - 1 Then
            EndPointLength += StartPoint - maxvpos + 1
            HalfPointLength += StartPoint - maxvpos + 1
            StartPoint = maxvpos - 1
        End If

        If StartPoint + EndPointLength < 0 Then
            EndPointLength = - StartPoint
        End If

        If StartPoint + EndPointLength > maxvpos - 1 Then
            EndPointLength = maxvpos - 1 - StartPoint
        End If

        If Math.Sign(HalfPointLength) <> Math.Sign(EndPointLength) Then HalfPointLength = 0
        If Math.Abs(HalfPointLength) > Math.Abs(EndPointLength) Then HalfPointLength = EndPointLength
    End Sub
End Class

Public Class MouseState
    Public CurrentMouseColumn As Integer = - 1
    Public CurrentMouseRow As Double = - 1.0#
    
    'Mouse is clicked on which point (location for display) (for selection box)
    Public LastMouseDownLocation As PointF = New Point(- 1, - 1) 
    
    'Mouse is moved to which point   (location for display) (for selection box)
    Public pMouseMove As PointF = New Point(- 1, - 1) 
    
    Public CurrentHoveredNoteIndex As Integer = - 1              'Mouse is on which note (for drawing green outline)
    Public MiddleButtonLocation As New Point(0, 0)
    Public MiddleButtonClicked As Boolean = False
    Public MouseMoveStatus As Point = New Point(0, 0)  'mouse is moved to which point (For Status Panel)
    Public PanX As Integer
    Public PanY As Integer
    Public PanVerticalScroll As Integer ' vscroll when middle mouse was clicked
    Public PanHorizontalScroll As Integer ' hscroll when middle mouse was clicked
End Class

Public Class NtState
    Public IsAdjustingNoteLength As Boolean     'If adjusting note length instead of moving it
    Public IsAdjustingUpperEnd As Boolean      'true = Adjusting upper end, false = adjusting lower end
    Public IsAdjustingSingleNote As Boolean     'true if there is only one note to be adjusted
End Class

Public Class EditorState
    Public Mouse As MouseState = New MouseState
    Public NT As NtState = New NtState
    Public TimeSelect As TimeSelectState = New TimeSelectState
    Public IsDuplicatingSelectedNotes As Boolean = False 'Indicates if the CTRL key is pressed while mousedown
    Public SelectedNotesWereDuplicated As Boolean = False     'Indicates if duplicate notes of select/unselect note
    Public OverwriteLastUndoRedoCommand As Boolean       'temp variables for undo, if undo command is added
End Class