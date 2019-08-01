Imports System.Text
Imports Microsoft.Win32

Public Class OpGeneral
    Public zWheel As Integer
    Public zPgUpDn As Integer
    Public zEncoding As Encoding
    Public zMiddle As Integer
    'Public zSort As Integer
    Public zAutoSave As Integer
    Public zGridPartition As Integer

    'Dim lpfa() As String

    <DllImport("shell32.dll")>
    Shared Sub SHChangeNotify(wEventId As Integer, uFlags As Integer, dwItem1 As Integer, dwItem2 As Integer)
    End Sub

    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        Select Case CWheel.SelectedIndex
            Case 0 : zWheel = 192
            Case 1 : zWheel = 96
            Case 2 : zWheel = 64
            Case 3 : zWheel = 48
        End Select
        Select Case CPgUpDn.SelectedIndex
            Case 0 : zPgUpDn = 1536
            Case 1 : zPgUpDn = 1152
            Case 2 : zPgUpDn = 768
            Case 3 : zPgUpDn = 576
            Case 4 : zPgUpDn = 384
            Case 5 : zPgUpDn = 192
            Case 6 : zPgUpDn = 96
        End Select
        Select Case CTextEncoding.SelectedIndex
            Case 0 : zEncoding = Encoding.Default
            Case 1 : zEncoding = Encoding.Unicode
            Case 2 : zEncoding = Encoding.ASCII
            Case 3 : zEncoding = Encoding.BigEndianUnicode
            Case 4 : zEncoding = Encoding.UTF32
            Case 5 : zEncoding = Encoding.UTF7
            Case 6 : zEncoding = Encoding.UTF8
            Case 7 : zEncoding = Encoding.GetEncoding(932)
            Case 8 : zEncoding = Encoding.GetEncoding(51949)
        End Select
        'zSort = CSortingMethod.SelectedIndex
        zMiddle = IIf(rMiddleDrag.Checked, 1, 0)
        zAutoSave = IIf(cAutoSave.Checked, 1, 0)*NAutoSave.Value*60000
        zGridPartition = nGridPartition.Value
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(xMsWheel As Integer, xPgUpDn As Integer, xMiddleButton As Integer, xTextEncoding As Integer,
                   xGridPartition As Integer,
                   xAutoSave As Integer, xBeep As Boolean, xBPMx As Boolean, xSTOPx As Boolean,
                   xMFEnter As Boolean, xMFClick As Boolean, xMStopPreview As Boolean)
        InitializeComponent()

        On Error Resume Next
        Select Case xMsWheel
            Case 192 : CWheel.SelectedIndex = 0
            Case 96 : CWheel.SelectedIndex = 1
            Case 64 : CWheel.SelectedIndex = 2
            Case 48 : CWheel.SelectedIndex = 3
        End Select

        Select Case xPgUpDn
            Case 1536 : CPgUpDn.SelectedIndex = 0
            Case 1152 : CPgUpDn.SelectedIndex = 1
            Case 768 : CPgUpDn.SelectedIndex = 2
            Case 576 : CPgUpDn.SelectedIndex = 3
            Case 384 : CPgUpDn.SelectedIndex = 4
            Case 192 : CPgUpDn.SelectedIndex = 5
            Case 96 : CPgUpDn.SelectedIndex = 6
        End Select

        CTextEncoding.SelectedIndex = xTextEncoding
        'CSortingMethod.SelectedIndex = xSort
        nGridPartition.Value = xGridPartition

        If xMiddleButton = 0 Then rMiddleAuto.Checked = True _
            Else rMiddleDrag.Checked = True

        If xAutoSave/60000 > NAutoSave.Maximum Or xAutoSave/60000 < NAutoSave.Minimum Then
            cAutoSave.Checked = False
        Else
            NAutoSave.Value = xAutoSave/60000
        End If

        cBeep.Checked = xBeep
        cBpm1296.Checked = xBPMx
        cStop1296.Checked = xSTOPx
        cMEnterFocus.Checked = xMFEnter
        cMClickFocus.Checked = xMFClick
        cMStopPreview.Checked = xMStopPreview
    End Sub

    Private Sub OpGeneral_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = MainWindow.Font

        'lpfa = Form1.lpfa
        'Dim xL() As String = Form1.lpgo
        Me.Text = Strings.fopGeneral.Title
        Label1.Text = Strings.fopGeneral.MouseWheel
        Label2.Text = Strings.fopGeneral.TextEncoding
        'Label3.Text = Locale.fopGeneral.SortingMethod
        Label5.Text = Strings.fopGeneral.PageUpDown
        Label3.Text = Strings.fopGeneral.MiddleButton
        Label4.Text = Strings.fopGeneral.AssociateFileType
        'CSortingMethod.Items.Item(0) = Locale.fopGeneral.sortBubble
        'CSortingMethod.Items.Item(1) = Locale.fopGeneral.sortInsertion
        'CSortingMethod.Items.Item(2) = Locale.fopGeneral.sortQuick
        'CSortingMethod.Items.Item(3) = Locale.fopGeneral.sortQuickD3
        'CSortingMethod.Items.Item(4) = Locale.fopGeneral.sortHeap
        rMiddleAuto.Text = Strings.fopGeneral.MiddleButtonAuto
        rMiddleDrag.Text = Strings.fopGeneral.MiddleButtonDrag
        Label6.Text = Strings.fopGeneral.MaxGridPartition

        cBeep.Text = Strings.fopGeneral.BeepWhileSaved
        cBpm1296.Text = Strings.fopGeneral.ExtendBPM
        cStop1296.Text = Strings.fopGeneral.ExtendSTOP
        cMEnterFocus.Text = Strings.fopGeneral.AutoFocusOnMouseEnter
        cMClickFocus.Text = Strings.fopGeneral.DisableFirstClick
        cAutoSave.Text = Strings.fopGeneral.AutoSave
        Label7.Text = Strings.fopGeneral.minutes
        cMStopPreview.Text = Strings.fopGeneral.StopPreviewOnClick

        Dim enc = Encoding.Default
        CTextEncoding.Items(0) = "System ANSI (" & enc.EncodingName & ")"

        OK_Button.Text = Strings.OK
        Cancel_Button.Text = Strings.Cancel
    End Sub

    Private Sub TBAssociate_Click(sender As Object, e As EventArgs) Handles TBAssociate.Click
        Associate(".bms", "iBMSC.BMS", Strings.FileAssociation.BMS, False)
    End Sub

    Private Sub TBAssociateIBMSC_Click(sender As Object, e As EventArgs) Handles TBAssociateIBMSC.Click
        Associate(".ibmsc", "iBMSC.iBMSC", Strings.FileAssociation.IBMSC, True)
    End Sub

    Private Sub TBAssociateBME_Click(sender As Object, e As EventArgs) Handles TBAssociateBME.Click
        Associate(".bme", "iBMSC.BME", Strings.FileAssociation.BME, False)
    End Sub

    Private Sub TBAssociateBML_Click(sender As Object, e As EventArgs) Handles TBAssociateBML.Click
        Associate(".bml", "iBMSC.BML", Strings.FileAssociation.BML, False)
    End Sub

    Private Sub TBAssociatePMS_Click(sender As Object, e As EventArgs) Handles TBAssociatePMS.Click
        Associate(".pms", "iBMSC.PMS", Strings.FileAssociation.PMS, False)
    End Sub

    Private Sub Associate(xExt As String, xClass As String, xDescription As String, isIBMSC As Boolean)
        If _
            MsgBox(Replace(Strings.Messages.FileAssociationPrompt, "{}", "*" & xExt),
                   MsgBoxStyle.YesNo Or MsgBoxStyle.Question) <> MsgBoxResult.Yes Then Exit Sub

        Dim xReg As RegistryKey

        Try
            With Registry.ClassesRoot
                If Array.IndexOf(.GetSubKeyNames(), xExt) <> - 1 Then .DeleteSubKeyTree(xExt)
                .CreateSubKey(xExt)
                xReg = .OpenSubKey(xExt, True)
                xReg.SetValue("", xClass, RegistryValueKind.String)

                If Array.IndexOf(.GetSubKeyNames(), xClass) <> - 1 Then .DeleteSubKeyTree(xClass)
                .CreateSubKey(xClass)
                xReg = .OpenSubKey(xClass, True)
                xReg.SetValue("", xDescription, RegistryValueKind.String)

                'Default Icon
                xReg.CreateSubKey("DefaultIcon")
                xReg = .OpenSubKey(xClass & "\DefaultIcon", True)
                xReg.SetValue("", My.Application.Info.DirectoryPath & "\TypeBMS.ico", RegistryValueKind.String)

                xReg = .OpenSubKey(xClass, True)
                xReg.CreateSubKey("shell")
                xReg = .OpenSubKey(xClass & "\shell", True)
                xReg.SetValue("", "open")

                xReg = .OpenSubKey(xClass & "\shell", True)
                xReg.CreateSubKey("open\command")
                xReg = .OpenSubKey(xClass & "\shell\open", True)
                xReg.SetValue("", Strings.FileAssociation.Open)
                xReg = .OpenSubKey(xClass & "\shell\open\command", True)
                xReg.SetValue("", """" & Application.ExecutablePath & """ ""%1""")

                If Not isIBMSC Then
                    xReg = .OpenSubKey(xClass & "\shell", True)
                    xReg.CreateSubKey("preview\command")
                    xReg = .OpenSubKey(xClass & "\shell\preview", True)
                    xReg.SetValue("", Strings.FileAssociation.Preview)
                    xReg = .OpenSubKey(xClass & "\shell\preview\command", True)
                    xReg.SetValue("", """" & My.Application.Info.DirectoryPath & "\uBMplay.exe" & """ ""%1""")

                    xReg = .OpenSubKey(xClass & "\shell", True)
                    xReg.CreateSubKey("viewcode\command")
                    xReg = .OpenSubKey(xClass & "\shell\viewcode", True)
                    xReg.SetValue("", Strings.FileAssociation.ViewCode)
                    xReg = .OpenSubKey(xClass & "\shell\viewcode\command", True)
                    xReg.SetValue("", Environment.SystemDirectory & "\notepad.exe %1")
                End If
            End With

            With Registry.CurrentUser
                .CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt)

                xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt, True)
                'xReg.DeleteSubKey("UserChoice")
                xReg.CreateSubKey("UserChoice")
                xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt & "\UserChoice",
                                   True)
                xReg.SetValue("Progid", xClass)
            End With

        Catch ex As Exception
            MsgBox(Strings.Messages.FileAssociationError & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.Exclamation,
                   Strings.Messages.Err)
        End Try

        SHChangeNotify(&H8000000, 0, 0, 0)
        Beep()
    End Sub

    '    Private Sub TBAssociateE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAssociatePMS.Click
    '        If MsgBox(Replace(Locale.Messages.FileAssociationPrompt, "{}", "*.bme, *.bml"), MsgBoxStyle.YesNo + MsgBoxStyle.Question) <> MsgBoxResult.Yes Then Exit Sub
    '
    '        Dim xReg As Microsoft.Win32.RegistryKey
    '        On Error GoTo Jump1
    '
    '        Dim xExt As String = ".bme"
    '        Dim xCtg As String = "iBMSC.BME"
    '
    'Jump2:
    '
    '        With Microsoft.Win32.Registry.ClassesRoot
    '            If Array.IndexOf(.GetSubKeyNames(), xExt) <> -1 Then .DeleteSubKeyTree(xExt)
    '            .CreateSubKey(xExt)
    '            xReg = .OpenSubKey(xExt, True)
    '            xReg.SetValue("", xCtg, Microsoft.Win32.RegistryValueKind.String)
    '
    '            If Array.IndexOf(.GetSubKeyNames(), xCtg) <> -1 Then .DeleteSubKeyTree(xCtg)
    '            .CreateSubKey(xCtg)
    '            xReg = .OpenSubKey(xCtg, True)
    '            xReg.SetValue("", lpfa(0), Microsoft.Win32.RegistryValueKind.String)
    '
    '            xReg.CreateSubKey("DefaultIcon")
    '            xReg = .OpenSubKey(xCtg & "\DefaultIcon", True)
    '            'xReg.SetValue("", My.Application.Info.DirectoryPath & "\TypeBMS.ico", Microsoft.Win32.RegistryValueKind.String)
    '
    '            xReg = .OpenSubKey(xCtg, True)
    '            xReg.CreateSubKey("shell")
    '            xReg = .OpenSubKey(xCtg & "\shell", True)
    '            xReg.SetValue("", "open")
    '
    '            xReg.CreateSubKey("open\command")
    '            xReg.CreateSubKey("preview\command")
    '            xReg.CreateSubKey("viewcode\command")
    '
    '            xReg = .OpenSubKey(xCtg & "\shell\open", True)
    '            xReg.SetValue("", lpfa(1))
    '            xReg = .OpenSubKey(xCtg & "\shell\preview", True)
    '            xReg.SetValue("", lpfa(2))
    '            xReg = .OpenSubKey(xCtg & "\shell\viewcode", True)
    '            xReg.SetValue("", lpfa(3))
    '
    '            xReg = .OpenSubKey(xCtg & "\shell\open\command", True)
    '            xReg.SetValue("", """" & Application.ExecutablePath & """ ""%1""")
    '            xReg = .OpenSubKey(xCtg & "\shell\preview\command", True)
    '            xReg.SetValue("", """" & My.Application.Info.DirectoryPath & "\uBMplay.exe" & """ ""%1""")
    '            xReg = .OpenSubKey(xCtg & "\shell\viewcode\command", True)
    '            xReg.SetValue("", Environment.SystemDirectory & "\notepad.exe %1")
    '        End With
    '
    '        With Microsoft.Win32.Registry.CurrentUser
    '            .CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt)
    '
    '            xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt, True)
    '            xReg.CreateSubKey("UserChoice")
    '            xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt & "\UserChoice", True)
    '            xReg.SetValue("Progid", xCtg)
    '        End With
    '
    '        If xExt <> ".bml" Or xCtg <> "iBMSC.BML" Then
    '            xExt = ".bml"
    '            xCtg = "iBMSC.BML"
    '            GoTo Jump2
    '        End If
    '
    'Jump1:
    '        Beep()
    '    End Sub

    '    Private Sub TBAssociateIBMSC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TBAssociateIBMSC.Click
    '        If MsgBox(Replace(Locale.Messages.FileAssociationPrompt, "{}", "*.ibmsc"), MsgBoxStyle.YesNo + MsgBoxStyle.Question) <> MsgBoxResult.Yes Then Exit Sub
    '
    '        Dim xReg As Microsoft.Win32.RegistryKey
    '        On Error GoTo Jump1
    '
    '        Dim xExt As String = ".ibmsc"
    '        Dim xCtg As String = "iBMSC.iBMSC"
    '
    '        With Microsoft.Win32.Registry.ClassesRoot
    '            If Array.IndexOf(.GetSubKeyNames(), xExt) <> -1 Then .DeleteSubKeyTree(xExt)
    '            .CreateSubKey(xExt)
    '            xReg = .OpenSubKey(xExt, True)
    '            xReg.SetValue("", xCtg, Microsoft.Win32.RegistryValueKind.String)
    '
    '            If Array.IndexOf(.GetSubKeyNames(), xCtg) <> -1 Then .DeleteSubKeyTree(xCtg)
    '            .CreateSubKey(xCtg)
    '            xReg = .OpenSubKey(xCtg, True)
    '            xReg.SetValue("", lpfa(0), Microsoft.Win32.RegistryValueKind.String)
    '
    '            xReg.CreateSubKey("DefaultIcon")
    '            xReg = .OpenSubKey(xCtg & "\DefaultIcon", True)
    '            'xReg.SetValue("", My.Application.Info.DirectoryPath & "\TypeBMS.ico", Microsoft.Win32.RegistryValueKind.String)
    '
    '            xReg = .OpenSubKey(xCtg, True)
    '            xReg.CreateSubKey("shell")
    '            xReg = .OpenSubKey(xCtg & "\shell", True)
    '            xReg.SetValue("", "open")
    '
    '            xReg.CreateSubKey("open\command")
    '            'xReg.CreateSubKey("preview\command")
    '            xReg.CreateSubKey("viewcode\command")
    '
    '            xReg = .OpenSubKey(xCtg & "\shell\open", True)
    '            xReg.SetValue("", lpfa(1))
    '            'xReg = .OpenSubKey(xCtg & "\shell\preview", True)
    '            'xReg.SetValue("", lpfa(2))
    '            xReg = .OpenSubKey(xCtg & "\shell\viewcode", True)
    '            xReg.SetValue("", lpfa(3))
    '
    '            xReg = .OpenSubKey(xCtg & "\shell\open\command", True)
    '            xReg.SetValue("", """" & Application.ExecutablePath & """ ""%1""")
    '            'xReg = .OpenSubKey(xCtg & "\shell\preview\command", True)
    '            'xReg.SetValue("", """" & My.Application.Info.DirectoryPath & "\uBMplay.exe" & """ ""%1""")
    '            xReg = .OpenSubKey(xCtg & "\shell\viewcode\command", True)
    '            xReg.SetValue("", Environment.SystemDirectory & "\notepad.exe %1")
    '        End With
    '
    '        With Microsoft.Win32.Registry.CurrentUser
    '            .CreateSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt)
    '
    '            xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt, True)
    '            xReg.CreateSubKey("UserChoice")
    '            xReg = .OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & xExt & "\UserChoice", True)
    '            xReg.SetValue("Progid", xCtg)
    '        End With
    '
    'Jump1:
    '        Beep()
    '    End Sub

    Private Sub cAutoSave_CheckedChanged(sender As Object, e As EventArgs) Handles cAutoSave.CheckedChanged
        NAutoSave.Enabled = cAutoSave.Checked
    End Sub
End Class
