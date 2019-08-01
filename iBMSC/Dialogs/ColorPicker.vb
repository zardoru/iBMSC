Imports System.Drawing.Drawing2D

Public Class ColorPicker
    Public OrigColor As Color = Color.Black
    Public NewColor As Color = Color.Black
    Dim DrawingIndex As Integer = 1
    Dim PassiveValueChange As Boolean = False

    Dim mMain As Point = New Point(0, 255)
    Dim mAlpha As Integer = 255
    Dim m1 As Integer = 255

    Public Sub SetOrigColor(xColor As Color)
        OrigColor = xColor
        SetNewColor(xColor)
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub SetNewColor(xColor As Color, Optional ByVal xSetText As Boolean = True)
        NewColor = xColor
        RefreshPrev(pPrev.DisplayRectangle)
        If xSetText Then tStr.Text = Mid("0000000" & Hex(xColor.ToArgb), Len(Hex(xColor.ToArgb)))
    End Sub

    Private Sub SetCursor()
        Select Case DrawingIndex
            Case 0
                mMain = New Point(inS.Value/1000*255, (1000 - inL.Value)/1000*255)
                m1 = (360 - inH.Value)*255/360
            Case 1
                mMain = New Point(inH.Value*255/360, (1000 - inL.Value)/1000*255)
                m1 = (1000 - inS.Value)*255/1000
            Case 2
                mMain = New Point(inH.Value*255/360, (1000 - inS.Value)*255/1000)
                m1 = (1000 - inL.Value)*255/1000
            Case 3
                mMain = New Point(inB.Value, 255 - inG.Value)
                m1 = 255 - inR.Value
            Case 4
                mMain = New Point(inB.Value, 255 - inR.Value)
                m1 = 255 - inG.Value
            Case 5
                mMain = New Point(inG.Value, 255 - inR.Value)
                m1 = 255 - inB.Value
        End Select
        mAlpha = inA.Value
    End Sub

    ''' <param name="xH">Hue (0-359)</param>
    ''' <param name="xS">Saturation (0-1000)</param>
    ''' <param name="xL">Lightness (0-1000)</param>
    ''' <param name="xA">Alpha (0-255)</param>
    Private Function HSL2RGB(xH As Integer, xS As Integer, xL As Integer, Optional ByVal xA As Integer = 255) As Color
        If xH > 360 Or xS > 1000 Or xL > 1000 Or xA > 255 Then Return Color.Black

        'Dim xxH As Double = xH
        Dim xxS As Double = xS/1000
        Dim xxB As Double = (xL - 500)/500
        Dim xR As Double
        Dim xG As Double
        Dim xB As Double

        If xH < 60 Then
            xB = - 1
            xR = 1
            xG = (xH - 30)/30
        ElseIf xH < 120 Then
            xB = - 1
            xG = 1
            xR = (90 - xH)/30
        ElseIf xH < 180 Then
            xR = - 1
            xG = 1
            xB = (xH - 150)/30
        ElseIf xH < 240 Then
            xR = - 1
            xB = 1
            xG = (210 - xH)/30
        ElseIf xH < 300 Then
            xG = - 1
            xB = 1
            xR = (xH - 270)/30
        Else
            xG = - 1
            xR = 1
            xB = (330 - xH)/30
        End If

        xR = (xR*xxS*(1 - Math.Abs(xxB)) + xxB + 1)*255/2
        xG = (xG*xxS*(1 - Math.Abs(xxB)) + xxB + 1)*255/2
        xB = (xB*xxS*(1 - Math.Abs(xxB)) + xxB + 1)*255/2

        Return Color.FromArgb(xA, xR, xG, xB)
    End Function

    Private Sub PCMain_MouseDown(sender As Object, e As MouseEventArgs) Handles PCMain.MouseDown
        PCMain_MouseMove(sender, e)
    End Sub

    Private Sub PCMain_MouseMove(sender As Object, e As MouseEventArgs) Handles PCMain.MouseMove
        If Not e.Button = MouseButtons.Left Then Exit Sub
        mMain = e.Location
        If mMain.X < 0 Then mMain.X = 0
        If mMain.X > 255 Then mMain.X = 255
        If mMain.Y < 0 Then mMain.Y = 0
        If mMain.Y > 255 Then mMain.Y = 255

        Select Case DrawingIndex
            Case 0
                inL.Value = 1000 - mMain.Y/255*1000
                inS.Value = mMain.X/255*1000
            Case 1
                inL.Value = 1000 - mMain.Y/255*1000
                inH.Value = mMain.X/255*360
            Case 2
                inS.Value = 1000 - mMain.Y/255*1000
                inH.Value = mMain.X/255*360
            Case 3
                inG.Value = 255 - mMain.Y
                inB.Value = mMain.X
            Case 4
                inR.Value = 255 - mMain.Y
                inB.Value = mMain.X
            Case 5
                inR.Value = 255 - mMain.Y
                inG.Value = mMain.X
        End Select
    End Sub

    Private Sub PCMain_Paint(sender As Object, e As PaintEventArgs) Handles PCMain.Paint
        RefreshMain(e.ClipRectangle)
    End Sub

    Private Sub RefreshMain(xRegion As Rectangle)
        Dim e1 As BufferedGraphics = BufferedGraphicsManager.Current.Allocate(PCMain.CreateGraphics, xRegion)
        Dim i As Integer
        'Dim xWidth As Integer = PCMain.DisplayRectangle.Width
        'Dim xHeight As Integer = PCMain.DisplayRectangle.Height
        Dim xPrecision As Integer = tbPrecision.Value

        Select Case DrawingIndex
            Case 0
                Dim xHue As Integer = inH.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 128),
                        Color.White,
                        HSL2RGB(xHue, i/255*1000, 500)),
                                              i,
                                              0,
                                              xPrecision,
                                              128)
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 128),
                        New Point(0, 256),
                        HSL2RGB(xHue, i/255*1000, 500),
                        Color.Black),
                                              i,
                                              128,
                                              xPrecision,
                                              128)
                Next
            Case 1
                Dim xSaturation As Integer = inS.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 128),
                        Color.White,
                        HSL2RGB(i/255*360, xSaturation, 500)),
                                              i,
                                              0,
                                              xPrecision,
                                              128)
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 128),
                        New Point(0, 256),
                        HSL2RGB(i/255*360, xSaturation, 500),
                        Color.Black),
                                              i,
                                              128,
                                              xPrecision,
                                              128)
                Next
            Case 2
                Dim xLightness As Integer = inL.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 256),
                        HSL2RGB(i/255*360, 1000, xLightness),
                        HSL2RGB(i/255*360, 0, xLightness)),
                                              i,
                                              0,
                                              xPrecision,
                                              256)
                Next
            Case 3
                Dim xRed As Integer = inR.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 256),
                        Color.FromArgb(xRed, 255, i),
                        Color.FromArgb(xRed, 0, i)),
                                              i,
                                              0,
                                              xPrecision,
                                              256)
                Next
            Case 4
                Dim xGreen As Integer = inG.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 256),
                        Color.FromArgb(255, xGreen, i),
                        Color.FromArgb(0, xGreen, i)),
                                              i,
                                              0,
                                              xPrecision,
                                              256)
                Next
            Case 5
                Dim xBlue As Integer = inB.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New LinearGradientBrush(
                        New Point(0, 0),
                        New Point(0, 256),
                        Color.FromArgb(255, i, xBlue),
                        Color.FromArgb(0, i, xBlue)),
                                              i,
                                              0,
                                              xPrecision,
                                              256)
                Next
        End Select

        e1.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        e1.Graphics.DrawEllipse(IIf(inL.Value > 500, Pens.Black, Pens.White), mMain.X - 4, mMain.Y - 4, 8, 8)

        e1.Render(PCMain.CreateGraphics)
        e1.Dispose()
    End Sub

    Private Sub rbH_CheckedChanged(sender As Object, e As EventArgs) Handles rbH.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 0 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub rbS_CheckedChanged(sender As Object, e As EventArgs) Handles rbS.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 1 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub rbL_CheckedChanged(sender As Object, e As EventArgs) Handles rbL.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 2 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub rbR_CheckedChanged(sender As Object, e As EventArgs) Handles rbR.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 3 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub rbG_CheckedChanged(sender As Object, e As EventArgs) Handles rbG.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 4 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub rbB_CheckedChanged(sender As Object, e As EventArgs) Handles rbB.CheckedChanged
        If sender.Checked Then _
            DrawingIndex = 5 : SetCursor() : RefreshMain(PCMain.DisplayRectangle) : Refresh1(PC1.DisplayRectangle) : _
                RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub inH_ValueChanged(sender As Object, e As EventArgs) Handles inH.ValueChanged
        If inH.Value = 360 Then inH.Value = 0

        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(HSL2RGB(inH.Value, inS.Value, inL.Value, inA.Value))
        inR.Value = NewColor.R
        inG.Value = NewColor.G
        inB.Value = NewColor.B

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inS_ValueChanged(sender As Object, e As EventArgs) Handles inS.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(HSL2RGB(inH.Value, inS.Value, inL.Value, inA.Value))
        inR.Value = NewColor.R
        inG.Value = NewColor.G
        inB.Value = NewColor.B

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inL_ValueChanged(sender As Object, e As EventArgs) Handles inL.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(HSL2RGB(inH.Value, inS.Value, inL.Value, inA.Value))
        inR.Value = NewColor.R
        inG.Value = NewColor.G
        inB.Value = NewColor.B

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inR_ValueChanged(sender As Object, e As EventArgs) Handles inR.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(Color.FromArgb(inA.Value, inR.Value, inG.Value, inB.Value))
        inH.Value = NewColor.GetHue
        inS.Value = NewColor.GetSaturation*1000
        inL.Value = NewColor.GetBrightness*1000

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inG_ValueChanged(sender As Object, e As EventArgs) Handles inG.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(Color.FromArgb(inA.Value, inR.Value, inG.Value, inB.Value))
        inH.Value = NewColor.GetHue
        inS.Value = NewColor.GetSaturation*1000
        inL.Value = NewColor.GetBrightness*1000

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inB_ValueChanged(sender As Object, e As EventArgs) Handles inB.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(Color.FromArgb(inA.Value, inR.Value, inG.Value, inB.Value))
        inH.Value = NewColor.GetHue
        inS.Value = NewColor.GetSaturation*1000
        inL.Value = NewColor.GetBrightness*1000

        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub inA_ValueChanged(sender As Object, e As EventArgs) Handles inA.ValueChanged
        If PassiveValueChange Then Exit Sub
        PassiveValueChange = True
        SetNewColor(Color.FromArgb(inA.Value, inR.Value, inG.Value, inB.Value))

        'RefreshMain(PCMain.DisplayRectangle)
        'Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
        PassiveValueChange = False
    End Sub

    Private Sub PC1_MouseDown(sender As Object, e As MouseEventArgs) Handles PC1.MouseDown
        PC1_MouseMove(sender, e)
    End Sub

    Private Sub PC1_MouseMove(sender As Object, e As MouseEventArgs) Handles PC1.MouseMove
        If Not e.Button = MouseButtons.Left Then Exit Sub
        m1 = e.Y
        If m1 < 0 Then m1 = 0
        If m1 > 255 Then m1 = 255

        Select Case DrawingIndex
            Case 0
                inH.Value = 360 - m1/255*360
            Case 1
                inS.Value = 1000 - m1/255*1000
            Case 2
                inL.Value = 1000 - m1/255*1000
            Case 3
                inR.Value = 255 - m1
            Case 4
                inG.Value = 255 - m1
            Case 5
                inB.Value = 255 - m1
        End Select
    End Sub

    Private Sub PC1_Paint(sender As Object, e As PaintEventArgs) Handles PC1.Paint
        Refresh1(e.ClipRectangle)
    End Sub

    Private Sub Refresh1(xRegion As Rectangle)
        Dim e1 As BufferedGraphics = BufferedGraphicsManager.Current.Allocate(PC1.CreateGraphics, xRegion)
        Dim i As Integer
        Dim xWidth As Integer = PC1.DisplayRectangle.Width
        Dim xPrecision As Integer = tbPrecision.Value

        Select Case DrawingIndex
            Case 0
                Dim xxS As Integer = inS.Value
                Dim xxL As Integer = inL.Value
                For i = 0 To 255 Step xPrecision
                    e1.Graphics.FillRectangle(New SolidBrush(HSL2RGB((255 - i)/255*360, xxS, xxL)), 0, i, xWidth,
                                              xPrecision)
                Next
            Case 1
                Dim xxH As Integer = inH.Value
                Dim xxL As Integer = inL.Value
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 0),
                    New Point(0, 256),
                    HSL2RGB(xxH, 1000, xxL),
                    HSL2RGB(xxH, 0, xxL)),
                                          0, 0, xWidth, 256)
            Case 2
                Dim xxH As Integer = inH.Value
                Dim xxS As Integer = inS.Value
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 0),
                    New Point(0, 128),
                    HSL2RGB(xxH, xxS, 1000),
                    HSL2RGB(xxH, xxS, 500)),
                                          0, 0, xWidth, 128)
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 128),
                    New Point(0, 256),
                    HSL2RGB(xxH, xxS, 500),
                    HSL2RGB(xxH, xxS, 0)),
                                          0, 128, xWidth, 128)
            Case 3
                Dim xxG As Integer = inG.Value
                Dim xxB As Integer = inB.Value
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 0),
                    New Point(0, 256),
                    Color.FromArgb(255, xxG, xxB),
                    Color.FromArgb(0, xxG, xxB)),
                                          0, 0, xWidth, 256)
            Case 4
                Dim xxR As Integer = inR.Value
                Dim xxB As Integer = inB.Value
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 0),
                    New Point(0, 256),
                    Color.FromArgb(xxR, 255, xxB),
                    Color.FromArgb(xxR, 0, xxB)),
                                          0, 0, xWidth, 256)
            Case 5
                Dim xxR As Integer = inR.Value
                Dim xxG As Integer = inG.Value
                e1.Graphics.FillRectangle(New LinearGradientBrush(
                    New Point(0, 0),
                    New Point(0, 256),
                    Color.FromArgb(xxR, xxG, 255),
                    Color.FromArgb(xxR, xxG, 0)),
                                          0, 0, xWidth, 256)
        End Select

        Dim xTri1() As Point = {New Point(4, m1), New Point(- 1, m1 - 2), New Point(- 1, m1 + 2)}
        Dim xTri2() As Point = {New Point(xWidth - 5, m1), New Point(xWidth, m1 - 2), New Point(xWidth, m1 + 2)}

        e1.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        e1.Graphics.FillPolygon(IIf(inL.Value > 500, Brushes.Black, Brushes.White), xTri1)
        e1.Graphics.FillPolygon(IIf(inL.Value > 500, Brushes.Black, Brushes.White), xTri2)

        e1.Render(PC1.CreateGraphics)
        e1.Dispose()
    End Sub

    Private Sub PCA_MouseDown(sender As Object, e As MouseEventArgs) Handles PCA.MouseDown
        PCA_MouseMove(sender, e)
    End Sub

    Private Sub PCA_MouseMove(sender As Object, e As MouseEventArgs) Handles PCA.MouseMove
        If Not e.Button = MouseButtons.Left Then Exit Sub
        mAlpha = e.X
        If mAlpha < 0 Then mAlpha = 0
        If mAlpha > 255 Then mAlpha = 255

        inA.Value = mAlpha
    End Sub

    Private Sub PCA_Paint(sender As Object, e As PaintEventArgs) Handles PCA.Paint
        RefreshA(e.ClipRectangle)
    End Sub

    Private Sub RefreshA(xRegion As Rectangle)
        Dim e1 As BufferedGraphics = BufferedGraphicsManager.Current.Allocate(PCA.CreateGraphics, xRegion)
        Dim i As Integer
        Dim xHeight As Integer = PCA.DisplayRectangle.Height
        Dim xPrecision As Integer = tbPrecision.Value
        Dim xColor As Color = Color.FromArgb(inR.Value, inG.Value, inB.Value)
        'Dim xAlphaRect(314) As Rectangle

        'e1.Graphics.FillRectangle(Brushes.White, PCA.DisplayRectangle)
        'For i = 0 To 62
        '    xAlphaRect(i * 5) = New Rectangle(i * 8, 0, 4, 4)
        '    xAlphaRect(i * 5 + 1) = New Rectangle(i * 8 + 4, 4, 4, 4)
        '    xAlphaRect(i * 5 + 2) = New Rectangle(i * 8, 8, 4, 4)
        '    xAlphaRect(i * 5 + 3) = New Rectangle(i * 8 + 4, 12, 4, 4)
        '    xAlphaRect(i * 5 + 4) = New Rectangle(i * 8, 16, 4, 4)
        'Next
        'e1.Graphics.FillRectangles(New Drawing.SolidBrush(Color.FromArgb(235, 235, 235)), xAlphaRect)
        e1.Graphics.DrawImageUnscaledAndClipped(My.Resources.TransparentBG, xRegion)

        For i = 0 To 255 Step xPrecision
            e1.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(i, xColor)), i, 0, xPrecision, xHeight)
        Next

        Dim xTri1() As Point = {New Point(mAlpha, 4), New Point(mAlpha - 2, - 1), New Point(mAlpha + 2, - 1)}
        Dim xTri2() As Point =
                {New Point(mAlpha, xHeight - 5), New Point(mAlpha - 2, xHeight), New Point(mAlpha + 2, xHeight)}
        i = inL.Value + (255 - inA.Value)*1000/255

        e1.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        e1.Graphics.FillPolygon(IIf(i > 500, Brushes.Black, Brushes.White), xTri1)
        e1.Graphics.FillPolygon(IIf(i > 500, Brushes.Black, Brushes.White), xTri2)

        e1.Render(PCA.CreateGraphics)
        e1.Dispose()
    End Sub

    Private Sub tbPrecision_ValueChanged(sender As Object, e As EventArgs) Handles tbPrecision.ValueChanged
        RefreshMain(PCMain.DisplayRectangle)
        Refresh1(PC1.DisplayRectangle)
        RefreshA(PCA.DisplayRectangle)
    End Sub

    Private Sub pPrev_Paint(sender As Object, e As PaintEventArgs) Handles pPrev.Paint
        RefreshPrev(e.ClipRectangle)
    End Sub

    Private Sub RefreshPrev(xRegion As Rectangle)
        Dim e1 As BufferedGraphics = BufferedGraphicsManager.Current.Allocate(pPrev.CreateGraphics, xRegion)
        Dim xFont As Font = pPrev.Font
        e1.Graphics.DrawImageUnscaledAndClipped(My.Resources.TransparentBG, xRegion)

        If xRegion.X < 62 Then
            e1.Graphics.FillRectangle(New SolidBrush(OrigColor), 0, 0, 61, 28)
            e1.Graphics.DrawLine(New Pen(Color.FromKnownColor(KnownColor.WindowFrame)), 61, 0, 61, 28)
            e1.Graphics.DrawString("Orig", xFont,
                                   IIf(OrigColor.GetBrightness + (255 - OrigColor.A)/255 > 0.5, Brushes.Black,
                                       Brushes.White),
                                   31 - e1.Graphics.MeasureString("Orig", xFont).Width/2,
                                   14 - e1.Graphics.MeasureString("Orig", xFont).Height/2)
        End If

        e1.Graphics.FillRectangle(New SolidBrush(NewColor), 62, 0, 61, 28)
        e1.Graphics.DrawString("New", xFont,
                               IIf(NewColor.GetBrightness + (255 - NewColor.A)/255 > 0.5, Brushes.Black, Brushes.White),
                               93 - e1.Graphics.MeasureString("New", xFont).Width/2,
                               14 - e1.Graphics.MeasureString("New", xFont).Height/2)

        e1.Render(pPrev.CreateGraphics)
        e1.Dispose()
    End Sub

    Private Sub tStr_GotFocus(sender As Object, e As EventArgs) Handles tStr.GotFocus
        tStr.SelectAll()
    End Sub

    Private Sub tStr_LostFocus(sender As Object, e As EventArgs) Handles tStr.LostFocus
        Dim i As Integer

        Try
            i = Convert.ToInt32(tStr.Text, 16)
            SetNewColor(Color.FromArgb(i))

            PassiveValueChange = True
            inR.Value = NewColor.R
            inG.Value = NewColor.G
            inB.Value = NewColor.B
            inH.Value = NewColor.GetHue
            inS.Value = NewColor.GetSaturation*1000
            inL.Value = NewColor.GetBrightness*1000
            inA.Value = NewColor.A
            SetCursor()

            RefreshMain(PCMain.DisplayRectangle)
            Refresh1(PC1.DisplayRectangle)
            RefreshA(PCA.DisplayRectangle)
            PassiveValueChange = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            tStr.Text = Mid("0000000" & Hex(NewColor.ToArgb), Len(Hex(NewColor.ToArgb)))
        End Try
    End Sub

    Private Sub tStr_TextChanged(sender As Object, e As EventArgs) Handles tStr.TextChanged
        If PassiveValueChange Then Exit Sub
        Dim i As Integer

        Try
            i = Convert.ToInt32(tStr.Text, 16)
            SetNewColor(Color.FromArgb(i), False)

            PassiveValueChange = True
            inR.Value = NewColor.R
            inG.Value = NewColor.G
            inB.Value = NewColor.B
            inH.Value = NewColor.GetHue
            inS.Value = NewColor.GetSaturation*1000
            inL.Value = NewColor.GetBrightness*1000
            inA.Value = NewColor.A
            SetCursor()

            RefreshMain(PCMain.DisplayRectangle)
            Refresh1(PC1.DisplayRectangle)
            RefreshA(PCA.DisplayRectangle)
            PassiveValueChange = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ColorPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font
    End Sub
End Class
