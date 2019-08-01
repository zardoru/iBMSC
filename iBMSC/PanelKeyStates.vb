Module PanelKeyStates
    Public Function ModifierLongNoteActive() As Boolean
        Return My.Computer.Keyboard.ShiftKeyDown And
               Not My.Computer.Keyboard.CtrlKeyDown
    End Function

    Public Function ModifierHiddenActive() As Boolean
        Return My.Computer.Keyboard.CtrlKeyDown And
               Not My.Computer.Keyboard.ShiftKeyDown
    End Function

    Public Function ModifierLandmineActive() As Boolean
        Return My.Computer.Keyboard.CtrlKeyDown And
               My.Computer.Keyboard.ShiftKeyDown
    End Function

    Public Function ModifierMultiselectActive() As Boolean
        Return My.Computer.Keyboard.ShiftKeyDown And My.Computer.Keyboard.CtrlKeyDown
    End Function
End Module
