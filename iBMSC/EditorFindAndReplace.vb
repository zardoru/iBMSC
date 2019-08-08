Imports System.Linq
Imports iBMSC.Editor

Public Partial Class MainWindow
    '----Find Delete Replace Options
    Dim fdriMesL As Integer
    Dim fdriMesU As Integer
    Dim fdriLblL As Integer
    Dim fdriLblU As Integer
    Dim fdriValL As Integer
    Dim fdriValU As Integer
    Dim fdriCol() As Integer
    
    Private Function fdrCheck(xNote As Note) As Boolean
        Return _
            xNote.VPosition >= MeasureBottom(fdriMesL) And
            xNote.VPosition < MeasureBottom(fdriMesU) + MeasureLength(fdriMesU) AndAlso
            IIf(Columns.IsColumnNumeric(xNote.ColumnIndex),
                xNote.Value >= fdriValL And xNote.Value <= fdriValU,
                xNote.Value >= fdriLblL And xNote.Value <= fdriLblU) AndAlso
            Array.IndexOf(fdriCol, xNote.ColumnIndex) <> - 1
    End Function

    Private Function fdrRangeS(xbLim1 As Boolean, xbLim2 As Boolean, xVal As Boolean) As Boolean
        Return (Not xbLim1 And xbLim2 And xVal) Or (xbLim1 And Not xbLim2 And Not xVal) Or (xbLim1 And xbLim2)
    End Function

    Public Sub fdrSelect(iRange As Integer,
                         xMesL As Integer, xMesU As Integer,
                         xLblL As String, xLblU As String,
                         xValL As Integer, xValU As Integer,
                         iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL)*10000
        fdriLblU = C36to10(xLblU)*10000
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
        For i = 1 To UBound(Notes)
            xSel(i) = Notes(i).Selected
        Next

        'Main process
        For i = 1 To UBound(Notes)
            Dim bbba As Boolean = xbSel And xSel(i)
            Dim bbbb As Boolean = xbUnsel And Not xSel(i)
            Dim bbbc As Boolean = Columns.IsEnabled(Notes(i).ColumnIndex)
            Dim bbbd As Boolean = fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote))
            Dim bbbe As Boolean = fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden)
            Dim bbbf As Boolean = fdrCheck(Notes(i))

            If ((xbSel And xSel(i)) Or (xbUnsel And Not xSel(i))) AndAlso
               Columns.IsEnabled(Notes(i).ColumnIndex) AndAlso
               fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And
               fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                Notes(i).Selected = fdrCheck(Notes(i))
            End If
        Next

        RefreshPanelAll()
        Beep()
    End Sub

    Public Sub fdrUnselect(iRange As Integer,
                           xMesL As Integer, xMesU As Integer,
                           xLblL As String, xLblU As String,
                           xValL As Integer, xValU As Integer,
                           iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL)*10000
        fdriLblU = C36to10(xLblU)*10000
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
        For i = 1 To UBound(Notes)
            xSel(i) = Notes(i).Selected
        Next

        'Main process
        For i = 1 To UBound(Notes)
            If ((xbSel And xSel(i)) Or (xbUnsel And Not xSel(i))) AndAlso
               Columns.IsEnabled(Notes(i).ColumnIndex) AndAlso
               fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And
               fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                Notes(i).Selected = Not fdrCheck(Notes(i))
            End If
        Next

        RefreshPanelAll()
        Beep()
    End Sub

    Public Sub fdrDelete(iRange As Integer,
                         xMesL As Integer, xMesU As Integer,
                         xLblL As String, xLblU As String,
                         xValL As Integer, xValU As Integer,
                         iCol() As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL)*10000
        fdriLblU = C36to10(xLblU)*10000
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
        Dim i = 1
        Do While i <= UBound(Notes)
            If ((xbSel And Notes(i).Selected) Or (xbUnsel And Not Notes(i).Selected)) AndAlso
               fdrCheck(Notes(i)) AndAlso
               Columns.IsEnabled(Notes(i).ColumnIndex) AndAlso
               fdrRangeS(xbShort, xbLong, IIf(NTInput, Notes(i).Length, Notes(i).LongNote)) And
               fdrRangeS(xbVisible, xbHidden, Notes(i).Hidden) Then
                RedoRemoveNote(Notes(i), xUndo, xRedo)
                RemoveNote(i, False)
            Else
                i += 1
            End If
        Loop

        AddUndoChain(xUndo, xBaseRedo.Next)
        ValidateNotesArray()
        RefreshPanelAll()

        Beep()
    End Sub

    Public Sub fdrReplaceL(iRange As Integer,
                           xMesL As Integer, xMesU As Integer,
                           xLblL As String, xLblU As String,
                           xValL As Integer, xValU As Integer,
                           iCol() As Integer, xReplaceLbl As String)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL)*10000
        fdriLblU = C36to10(xLblU)*10000
        fdriValL = xValL
        fdriValU = xValU
        fdriCol = iCol

        Dim xbSel As Boolean = iRange Mod 2 = 0
        Dim xbUnsel As Boolean = iRange Mod 3 = 0
        Dim xbShort As Boolean = iRange Mod 5 = 0
        Dim xbLong As Boolean = iRange Mod 7 = 0
        Dim xbHidden As Boolean = iRange Mod 11 = 0
        Dim xbVisible As Boolean = iRange Mod 13 = 0

        Dim xxLbl As Integer = C36to10(xReplaceLbl)*10000

        Dim xUndo As UndoRedo.LinkedURCmd = Nothing
        Dim xRedo As UndoRedo.LinkedURCmd = New UndoRedo.Void
        Dim xBaseRedo As UndoRedo.LinkedURCmd = xRedo

        'Main process
        For Each note In Notes
            If ((xbSel And note.Selected) Or (xbUnsel And Not note.Selected)) AndAlso
               fdrCheck(note) AndAlso
               Columns.IsEnabled(note.ColumnIndex) And
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

    Public Sub fdrReplaceV(iRange As Integer,
                           xMesL As Integer, xMesU As Integer,
                           xLblL As String, xLblU As String,
                           xValL As Integer, xValU As Integer,
                           iCol() As Integer, xReplaceVal As Integer)

        fdriMesL = xMesL
        fdriMesU = xMesU
        fdriLblL = C36to10(xLblL)*10000
        fdriLblU = C36to10(xLblU)*10000
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
               Columns.IsEnabled(note.ColumnIndex) And
               Columns.IsColumnNumeric(note.ColumnIndex) AndAlso
               fdrRangeS(xbShort, xbLong, IIf(NTInput, note.Length, note.LongNote)) And
               fdrRangeS(xbVisible, xbHidden, note.Hidden) Then
                RedoRelabelNote(note, xReplaceVal, xUndo, xRedo)
                note.Value = xReplaceVal
            End If
        Next

        AddUndoChain(xUndo, xBaseRedo.Next)
        RefreshPanelAll()
        Beep()
    End Sub
End Class