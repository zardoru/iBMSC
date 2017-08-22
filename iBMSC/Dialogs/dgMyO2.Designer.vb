<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dgMyO2
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(dgMyO2))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.vBPM = New System.Windows.Forms.NumericUpDown
        Me.bApply1 = New System.Windows.Forms.Button
        Me.bApply2 = New System.Windows.Forms.Button
        Me.bApply3 = New System.Windows.Forms.Button
        Me.lResult = New System.Windows.Forms.DataGridView
        Me.Index = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Measure = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.ColumnName = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.Grid = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.LongNote = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.Hidden = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.AdjTo64 = New System.Windows.Forms.DataGridViewCheckBoxColumn
        Me.D64 = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.D48 = New System.Windows.Forms.DataGridViewTextBoxColumn
        CType(Me.vBPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'vBPM
        '
        Me.vBPM.DecimalPlaces = 4
        Me.vBPM.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        resources.ApplyResources(Me.vBPM, "vBPM")
        Me.vBPM.Maximum = New Decimal(New Integer() {655359999, 0, 0, 262144})
        Me.vBPM.Minimum = New Decimal(New Integer() {1, 0, 0, 262144})
        Me.vBPM.Name = "vBPM"
        Me.vBPM.Value = New Decimal(New Integer() {120, 0, 0, 0})
        '
        'bApply1
        '
        resources.ApplyResources(Me.bApply1, "bApply1")
        Me.bApply1.Name = "bApply1"
        Me.bApply1.UseVisualStyleBackColor = True
        '
        'bApply2
        '
        resources.ApplyResources(Me.bApply2, "bApply2")
        Me.bApply2.Name = "bApply2"
        Me.bApply2.UseVisualStyleBackColor = True
        '
        'bApply3
        '
        resources.ApplyResources(Me.bApply3, "bApply3")
        Me.bApply3.Name = "bApply3"
        Me.bApply3.UseVisualStyleBackColor = True
        '
        'lResult
        '
        Me.lResult.AllowUserToAddRows = False
        Me.lResult.AllowUserToDeleteRows = False
        Me.lResult.AllowUserToResizeColumns = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer), CType(CType(248, Byte), Integer))
        Me.lResult.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.lResult.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.lResult.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.lResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.lResult.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Index, Me.Measure, Me.ColumnName, Me.Grid, Me.LongNote, Me.Hidden, Me.AdjTo64, Me.D64, Me.D48})
        resources.ApplyResources(Me.lResult, "lResult")
        Me.lResult.Name = "lResult"
        '
        'Index
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Index.DefaultCellStyle = DataGridViewCellStyle3
        resources.ApplyResources(Me.Index, "Index")
        Me.Index.Name = "Index"
        Me.Index.ReadOnly = True
        '
        'Measure
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Measure.DefaultCellStyle = DataGridViewCellStyle4
        resources.ApplyResources(Me.Measure, "Measure")
        Me.Measure.Name = "Measure"
        Me.Measure.ReadOnly = True
        '
        'ColumnName
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ColumnName.DefaultCellStyle = DataGridViewCellStyle5
        resources.ApplyResources(Me.ColumnName, "ColumnName")
        Me.ColumnName.Name = "ColumnName"
        Me.ColumnName.ReadOnly = True
        '
        'Grid
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Grid.DefaultCellStyle = DataGridViewCellStyle6
        resources.ApplyResources(Me.Grid, "Grid")
        Me.Grid.Name = "Grid"
        Me.Grid.ReadOnly = True
        '
        'LongNote
        '
        Me.LongNote.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        resources.ApplyResources(Me.LongNote, "LongNote")
        Me.LongNote.Name = "LongNote"
        Me.LongNote.ReadOnly = True
        '
        'Hidden
        '
        Me.Hidden.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        resources.ApplyResources(Me.Hidden, "Hidden")
        Me.Hidden.Name = "Hidden"
        Me.Hidden.ReadOnly = True
        '
        'AdjTo64
        '
        Me.AdjTo64.FlatStyle = System.Windows.Forms.FlatStyle.System
        resources.ApplyResources(Me.AdjTo64, "AdjTo64")
        Me.AdjTo64.Name = "AdjTo64"
        '
        'D64
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.Color.Maroon
        Me.D64.DefaultCellStyle = DataGridViewCellStyle7
        resources.ApplyResources(Me.D64, "D64")
        Me.D64.Name = "D64"
        Me.D64.ReadOnly = True
        '
        'D48
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.Color.Navy
        Me.D48.DefaultCellStyle = DataGridViewCellStyle8
        resources.ApplyResources(Me.D48, "D48")
        Me.D48.Name = "D48"
        Me.D48.ReadOnly = True
        '
        'dgMyO2
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.lResult)
        Me.Controls.Add(Me.bApply3)
        Me.Controls.Add(Me.bApply2)
        Me.Controls.Add(Me.bApply1)
        Me.Controls.Add(Me.vBPM)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "dgMyO2"
        Me.TopMost = True
        CType(Me.vBPM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents vBPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents bApply1 As System.Windows.Forms.Button
    Friend WithEvents bApply2 As System.Windows.Forms.Button
    Friend WithEvents bApply3 As System.Windows.Forms.Button
    Friend WithEvents lResult As System.Windows.Forms.DataGridView
    Friend WithEvents Index As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Measure As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColumnName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Grid As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LongNote As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Hidden As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents AdjTo64 As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents D64 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents D48 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
