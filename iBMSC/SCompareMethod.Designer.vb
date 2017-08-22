<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmCompareMethod
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.bQuickSort = New System.Windows.Forms.Button
        Me.PTitle = New System.Windows.Forms.Panel
        Me.PMain = New System.Windows.Forms.Panel
        Me.bBubbleSort = New System.Windows.Forms.Button
        Me.bInsertionSort = New System.Windows.Forms.Button
        Me.bImport = New System.Windows.Forms.Button
        Me.PTitle.SuspendLayout()
        Me.SuspendLayout()
        '
        'bQuickSort
        '
        Me.bQuickSort.Location = New System.Drawing.Point(203, 12)
        Me.bQuickSort.Name = "bQuickSort"
        Me.bQuickSort.Size = New System.Drawing.Size(108, 23)
        Me.bQuickSort.TabIndex = 0
        Me.bQuickSort.Text = "Quick Sort"
        Me.bQuickSort.UseVisualStyleBackColor = True
        '
        'PTitle
        '
        Me.PTitle.Controls.Add(Me.bImport)
        Me.PTitle.Controls.Add(Me.bInsertionSort)
        Me.PTitle.Controls.Add(Me.bBubbleSort)
        Me.PTitle.Controls.Add(Me.bQuickSort)
        Me.PTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.PTitle.Location = New System.Drawing.Point(0, 0)
        Me.PTitle.Name = "PTitle"
        Me.PTitle.Size = New System.Drawing.Size(1033, 50)
        Me.PTitle.TabIndex = 1
        '
        'PMain
        '
        Me.PMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PMain.Location = New System.Drawing.Point(0, 50)
        Me.PMain.Name = "PMain"
        Me.PMain.Size = New System.Drawing.Size(1033, 540)
        Me.PMain.TabIndex = 1
        '
        'bBubbleSort
        '
        Me.bBubbleSort.Location = New System.Drawing.Point(317, 12)
        Me.bBubbleSort.Name = "bBubbleSort"
        Me.bBubbleSort.Size = New System.Drawing.Size(108, 23)
        Me.bBubbleSort.TabIndex = 1
        Me.bBubbleSort.Text = "Bubble Sort"
        Me.bBubbleSort.UseVisualStyleBackColor = True
        '
        'bInsertionSort
        '
        Me.bInsertionSort.Location = New System.Drawing.Point(431, 12)
        Me.bInsertionSort.Name = "bInsertionSort"
        Me.bInsertionSort.Size = New System.Drawing.Size(108, 23)
        Me.bInsertionSort.TabIndex = 2
        Me.bInsertionSort.Text = "Insertion Sort"
        Me.bInsertionSort.UseVisualStyleBackColor = True
        '
        'bImport
        '
        Me.bImport.Location = New System.Drawing.Point(12, 12)
        Me.bImport.Name = "bImport"
        Me.bImport.Size = New System.Drawing.Size(75, 23)
        Me.bImport.TabIndex = 3
        Me.bImport.Text = "Import K"
        Me.bImport.UseVisualStyleBackColor = True
        '
        'FrmCompareMethod
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1033, 590)
        Me.Controls.Add(Me.PMain)
        Me.Controls.Add(Me.PTitle)
        Me.Name = "FrmCompareMethod"
        Me.Text = "Form2"
        Me.PTitle.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents bQuickSort As System.Windows.Forms.Button
    Friend WithEvents PTitle As System.Windows.Forms.Panel
    Friend WithEvents PMain As System.Windows.Forms.Panel
    Friend WithEvents bInsertionSort As System.Windows.Forms.Button
    Friend WithEvents bBubbleSort As System.Windows.Forms.Button
    Friend WithEvents bImport As System.Windows.Forms.Button
End Class
