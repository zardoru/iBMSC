Public Class FrmCompareMethod

    Dim K() As Long

    Public Sub SortByVPositionQuick(ByVal xMin As Integer, ByVal xMax As Integer)
        Dim xNote As Long
        Dim iHi As Integer
        Dim iLo As Integer
        Dim xI1 As Integer

        ' If min >= max, the list contains 0 or 1 items so it is sorted.
        If xMin >= xMax Then Exit Sub

        ' Pick the dividing value.
        xI1 = CInt((xMax - xMin) / 2) + xMin
        xNote = K(xI1)

        ' Swap it to the front.
        K(xI1) = K(xMin)

        iLo = xMin
        iHi = xMax
        Do
            ' Look down from hi for a value < med_value.
            Do While K(iHi) >= xNote
                iHi = iHi - 1
                If iHi <= iLo Then Exit Do
            Loop
            If iHi <= iLo Then
                K(iLo) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            K(iLo) = K(iHi)

            ' Look up from lo for a value >= med_value.
            iLo = iLo + 1
            Do While K(iLo) < xNote
                iLo = iLo + 1
                If iLo >= iHi Then Exit Do
            Loop
            If iLo >= iHi Then
                iLo = iHi
                K(iHi) = xNote
                Exit Do
            End If

            ' Swap the lo and hi values.
            K(iHi) = K(iLo)

            '-----------------Draw
            Dim currentContext As BufferedGraphicsContext = BufferedGraphicsManager.Current
            Dim e1 As BufferedGraphics = currentContext.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

            Dim xI2 As Integer
            For xI2 = 0 To UBound(K)
                e1.Graphics.DrawLine(Pens.Red, xI2, 0, xI2, CInt(K(xI2) / 20))
            Next

            e1.Render(PMain.CreateGraphics)
            e1.Dispose()
            currentContext.Dispose()
            Application.DoEvents()
            '-----------------End Draw


        Loop

        ' Sort the two sublists.
        SortByVPositionQuick(xMin, iLo - 1)
        SortByVPositionQuick(iLo + 1, xMax)


        '-----------------Draw
        Dim currentContext2 As BufferedGraphicsContext = BufferedGraphicsManager.Current
        Dim e2 As BufferedGraphics = currentContext2.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

        Dim xI3 As Integer
        For xI3 = 0 To UBound(K)
            e2.Graphics.DrawLine(Pens.Red, xI3, 0, xI3, CInt(K(xI3) / 20))
        Next

        e2.Render(PMain.CreateGraphics)
        e2.Dispose()
        currentContext2.Dispose()
        Application.DoEvents()
        '-----------------End Draw


    End Sub

    Public Sub SortByVPositionBubble() 'One-directional Bubble Sort
        Dim xI1 As Integer = 1
        Dim xNote As Long
        Do While xI1 <= UBound(K)
            If K(xI1 - 1) > K(xI1) Then
                'Swap
                xNote = K(xI1 - 1)
                K(xI1 - 1) = K(xI1)
                K(xI1) = xNote
                'To previous
                If xI1 > 1 Then xI1 -= 2
            End If
            xI1 += 1
            '-----------------Draw
            Dim currentContext As BufferedGraphicsContext = BufferedGraphicsManager.Current
            Dim e1 As BufferedGraphics = currentContext.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

            Dim xI2 As Integer
            For xI2 = 0 To UBound(K)
                e1.Graphics.DrawLine(Pens.Red, xI2, 0, xI2, CInt(K(xI2) / 20))
            Next

            e1.Render(PMain.CreateGraphics)
            e1.Dispose()
            currentContext.Dispose()
            Application.DoEvents()
            '-----------------End Draw
        Loop
        '-----------------Draw
        Dim currentContext2 As BufferedGraphicsContext = BufferedGraphicsManager.Current
        Dim e2 As BufferedGraphics = currentContext2.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

        Dim xI3 As Integer
        For xI3 = 0 To UBound(K)
            e2.Graphics.DrawLine(Pens.Red, xI3, 0, xI3, CInt(K(xI3) / 20))
        Next

        e2.Render(PMain.CreateGraphics)
        e2.Dispose()
        currentContext2.Dispose()
        Application.DoEvents()
        '-----------------End Draw

    End Sub

    Public Sub SortByVPositionInsertion() 'Insertion Sort
        If UBound(K) <= 0 Then Exit Sub
        Dim xNote As Long
        Dim xI1 As Integer
        Dim xI2 As Integer
        For xI1 = 1 To UBound(K)
            xNote = K(xI1)
            For xI2 = xI1 - 1 To 0 Step -1
                If K(xI2) > xNote Then
                    K(xI2 + 1) = K(xI2)
                    If xI2 = 0 Then
                        K(xI2) = xNote
                        Exit For
                    End If
                Else
                    K(xI2 + 1) = xNote
                    Exit For
                End If
                '-----------------Draw
                Dim currentContext As BufferedGraphicsContext = BufferedGraphicsManager.Current
                Dim e1 As BufferedGraphics = currentContext.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

                Dim xI4 As Integer
                For xI4 = 0 To UBound(K)
                    e1.Graphics.DrawLine(Pens.Red, xI4, 0, xI4, CInt(K(xI4) / 20))
                Next

                e1.Render(PMain.CreateGraphics)
                e1.Dispose()
                currentContext.Dispose()
                Application.DoEvents()
                '-----------------End Draw
            Next
        Next
        '-----------------Draw
        Dim currentContext2 As BufferedGraphicsContext = BufferedGraphicsManager.Current
        Dim e2 As BufferedGraphics = currentContext2.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

        Dim xI3 As Integer
        For xI3 = 0 To UBound(K)
            e2.Graphics.DrawLine(Pens.Red, xI3, 0, xI3, CInt(K(xI3) / 20))
        Next

        e2.Render(PMain.CreateGraphics)
        e2.Dispose()
        currentContext2.Dispose()
        Application.DoEvents()
        '-----------------End Draw


    End Sub

    Private Sub FrmSetCompareMethod_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub FrmSetCompareMethod_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
    End Sub

    Private Sub PMain_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PMain.Paint
        If K Is Nothing Then Exit Sub

    End Sub

    Private Sub bQuickSort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bQuickSort.Click
        If K Is Nothing Then Exit Sub
        SortByVPositionQuick(0, UBound(K))
    End Sub

    Private Sub bBubbleSort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bBubbleSort.Click
        If K Is Nothing Then Exit Sub
        SortByVPositionBubble()
    End Sub

    Private Sub bInsertionSort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bInsertionSort.Click
        If K Is Nothing Then Exit Sub
        SortByVPositionInsertion()
    End Sub

    Private Sub bImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bImport.Click
        Erase K
        K = Form1.ReturnK
        '-----------------Draw
        Dim currentContext2 As BufferedGraphicsContext = BufferedGraphicsManager.Current
        Dim e2 As BufferedGraphics = currentContext2.Allocate(PMain.CreateGraphics, PMain.DisplayRectangle)

        Dim xI3 As Integer
        For xI3 = 0 To UBound(K)
            e2.Graphics.DrawLine(Pens.Red, xI3, 0, xI3, CInt(K(xI3) / 20))
        Next

        e2.Render(PMain.CreateGraphics)
        e2.Dispose()
        currentContext2.Dispose()
        Application.DoEvents()
        '-----------------End Draw
    End Sub
End Class