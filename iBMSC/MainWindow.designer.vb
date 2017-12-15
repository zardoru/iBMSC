<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainWindow
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainWindow))
        Me.cmnLanguage = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TBLangDef = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBLangRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator9 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnLanguage = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBLanguage = New System.Windows.Forms.ToolStripDropDownButton()
        Me.cmnTheme = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TBThemeDef = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBThemeSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBThemeRefresh = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBThemeLoadComptability = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnTheme = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBTheme = New System.Windows.Forms.ToolStripDropDownButton()
        Me.POptionsScroll = New System.Windows.Forms.Panel()
        Me.POptions = New System.Windows.Forms.Panel()
        Me.POExpansion = New System.Windows.Forms.Panel()
        Me.POExpansionInner = New System.Windows.Forms.Panel()
        Me.TExpansion = New System.Windows.Forms.TextBox()
        Me.POExpansionResizer = New System.Windows.Forms.Button()
        Me.POExpansionSwitch = New System.Windows.Forms.CheckBox()
        Me.POBeat = New System.Windows.Forms.Panel()
        Me.POBeatInner = New System.Windows.Forms.TableLayoutPanel()
        Me.POBeatExpander = New System.Windows.Forms.CheckBox()
        Me.POBeatResizer = New System.Windows.Forms.Button()
        Me.TableLayoutPanel7 = New System.Windows.Forms.TableLayoutPanel()
        Me.nBeatD = New System.Windows.Forms.NumericUpDown()
        Me.BBeatApplyV = New System.Windows.Forms.Button()
        Me.nBeatN = New System.Windows.Forms.NumericUpDown()
        Me.BBeatApply = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.tBeatValue = New System.Windows.Forms.TextBox()
        Me.LBeat = New System.Windows.Forms.ListBox()
        Me.POBeatPart2 = New System.Windows.Forms.TableLayoutPanel()
        Me.CBeatScale = New System.Windows.Forms.RadioButton()
        Me.CBeatCut = New System.Windows.Forms.RadioButton()
        Me.CBeatMeasure = New System.Windows.Forms.RadioButton()
        Me.CBeatPreserve = New System.Windows.Forms.RadioButton()
        Me.POBeatSwitch = New System.Windows.Forms.CheckBox()
        Me.POWAV = New System.Windows.Forms.Panel()
        Me.POWAVInner = New System.Windows.Forms.TableLayoutPanel()
        Me.POWAVExpander = New System.Windows.Forms.CheckBox()
        Me.LWAV = New System.Windows.Forms.ListBox()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.BWAVUp = New System.Windows.Forms.Button()
        Me.BWAVDown = New System.Windows.Forms.Button()
        Me.BWAVBrowse = New System.Windows.Forms.Button()
        Me.BWAVRemove = New System.Windows.Forms.Button()
        Me.POWAVResizer = New System.Windows.Forms.Button()
        Me.POWAVPart2 = New System.Windows.Forms.TableLayoutPanel()
        Me.CWAVMultiSelect = New System.Windows.Forms.CheckBox()
        Me.CWAVChangeLabel = New System.Windows.Forms.CheckBox()
        Me.POWAVSwitch = New System.Windows.Forms.CheckBox()
        Me.POWaveForm = New System.Windows.Forms.Panel()
        Me.POWaveFormInner = New System.Windows.Forms.Panel()
        Me.POWaveFormPart2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TWSaturation = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.TWTransparency = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.TWPrecision = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.TWWidth = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.TWLeft = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.TWSaturation2 = New System.Windows.Forms.TrackBar()
        Me.TWLeft2 = New System.Windows.Forms.TrackBar()
        Me.TWTransparency2 = New System.Windows.Forms.TrackBar()
        Me.TWWidth2 = New System.Windows.Forms.TrackBar()
        Me.TWPrecision2 = New System.Windows.Forms.TrackBar()
        Me.POWaveFormExpander = New System.Windows.Forms.CheckBox()
        Me.POWaveFormPart1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.BWLoad = New System.Windows.Forms.Button()
        Me.BWClear = New System.Windows.Forms.Button()
        Me.BWLock = New System.Windows.Forms.CheckBox()
        Me.TWFileName = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TWPosition2 = New System.Windows.Forms.TrackBar()
        Me.TWPosition = New System.Windows.Forms.NumericUpDown()
        Me.POWaveFormSwitch = New System.Windows.Forms.CheckBox()
        Me.POGrid = New System.Windows.Forms.Panel()
        Me.POGridInner = New System.Windows.Forms.Panel()
        Me.POGridPart2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.cVSLockL = New System.Windows.Forms.CheckBox()
        Me.cVSLock = New System.Windows.Forms.CheckBox()
        Me.cVSLockR = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CGB = New System.Windows.Forms.NumericUpDown()
        Me.POGridExpander = New System.Windows.Forms.CheckBox()
        Me.POGridPart1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.PictureBox9 = New System.Windows.Forms.PictureBox()
        Me.CGHeight2 = New System.Windows.Forms.TrackBar()
        Me.CGHeight = New System.Windows.Forms.NumericUpDown()
        Me.PictureBox10 = New System.Windows.Forms.PictureBox()
        Me.CGWidth2 = New System.Windows.Forms.TrackBar()
        Me.CGWidth = New System.Windows.Forms.NumericUpDown()
        Me.CGDisableVertical = New System.Windows.Forms.CheckBox()
        Me.CGSnap = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.CGDivide = New System.Windows.Forms.NumericUpDown()
        Me.CGSub = New System.Windows.Forms.NumericUpDown()
        Me.BGSlash = New System.Windows.Forms.Button()
        Me.POGridSwitch = New System.Windows.Forms.CheckBox()
        Me.POHeader = New System.Windows.Forms.Panel()
        Me.POHeaderInner = New System.Windows.Forms.Panel()
        Me.POHeaderPart2 = New System.Windows.Forms.TableLayoutPanel()
        Me.CHDifficulty = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.THExRank = New System.Windows.Forms.TextBox()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.CHLnObj = New System.Windows.Forms.ComboBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.THComment = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.THTotal = New System.Windows.Forms.TextBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.BHStageFile = New System.Windows.Forms.Button()
        Me.BHBanner = New System.Windows.Forms.Button()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.BHBackBMP = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.THBackBMP = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.THBanner = New System.Windows.Forms.TextBox()
        Me.THStageFile = New System.Windows.Forms.TextBox()
        Me.THSubTitle = New System.Windows.Forms.TextBox()
        Me.THSubArtist = New System.Windows.Forms.TextBox()
        Me.POHeaderExpander = New System.Windows.Forms.CheckBox()
        Me.POHeaderPart1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.THPlayLevel = New System.Windows.Forms.TextBox()
        Me.CHRank = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.CHPlayer = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.THGenre = New System.Windows.Forms.TextBox()
        Me.THBPM = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.THArtist = New System.Windows.Forms.TextBox()
        Me.THTitle = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.POHeaderSwitch = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Menu1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.MInsert = New System.Windows.Forms.ToolStripMenuItem()
        Me.MRemove = New System.Windows.Forms.ToolStripMenuItem()
        Me.AutoSaveTimer = New System.Windows.Forms.Timer(Me.components)
        Me.mnMain = New System.Windows.Forms.MenuStrip()
        Me.mnFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOpen = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnImportSM = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnImportIBMSC = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator14 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnSave = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnExport = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator15 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnOpenR0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOpenR1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOpenR2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOpenR3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOpenR4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator16 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnQuit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnUndo = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnRedo = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator17 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnCut = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnPaste = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnDelete = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSelectAll = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnGotoMeasure = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator18 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnFind = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnStatistics = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator19 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnTimeSelect = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSelect = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnWrite = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator23 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnMyO2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSys = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSMenu = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSTB = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSOP = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSStatus = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSLSplitter = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnSRSplitter = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator21 = New System.Windows.Forms.ToolStripSeparator()
        Me.CGShow = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowS = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowBG = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowM = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowMB = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowV = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGShowC = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator22 = New System.Windows.Forms.ToolStripSeparator()
        Me.CGBPM = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGSTOP = New System.Windows.Forms.ToolStripMenuItem()
        Me.CGBLP = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnNTInput = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnErrorCheck = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnPreviewOnClick = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnShowFileName = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator20 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnGOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnVOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnPOptions = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnConversion = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmnConversion = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.POBLong = New System.Windows.Forms.ToolStripMenuItem()
        Me.POBShort = New System.Windows.Forms.ToolStripMenuItem()
        Me.POBLongShort = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator10 = New System.Windows.Forms.ToolStripSeparator()
        Me.POBHidden = New System.Windows.Forms.ToolStripMenuItem()
        Me.POBVisible = New System.Windows.Forms.ToolStripMenuItem()
        Me.POBHiddenVisible = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator11 = New System.Windows.Forms.ToolStripSeparator()
        Me.POBModify = New System.Windows.Forms.ToolStripMenuItem()
        Me.POBMirror = New System.Windows.Forms.ToolStripMenuItem()
        Me.POConvert = New System.Windows.Forms.ToolStripDropDownButton()
        Me.mnPreview = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnPlayB = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnPlay = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnStop = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBMain = New System.Windows.Forms.ToolStrip()
        Me.TBNew = New System.Windows.Forms.ToolStripButton()
        Me.TBOpen = New System.Windows.Forms.ToolStripSplitButton()
        Me.TBOpenR0 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBOpenR1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBOpenR2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBOpenR3 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBOpenR4 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator12 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBImportSM = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBImportIBMSC = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBSave = New System.Windows.Forms.ToolStripSplitButton()
        Me.TBSaveAs = New System.Windows.Forms.ToolStripMenuItem()
        Me.TBExport = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBCut = New System.Windows.Forms.ToolStripButton()
        Me.TBCopy = New System.Windows.Forms.ToolStripButton()
        Me.TBPaste = New System.Windows.Forms.ToolStripButton()
        Me.TBFind = New System.Windows.Forms.ToolStripButton()
        Me.TBStatistics = New System.Windows.Forms.ToolStripButton()
        Me.TBMyO2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBErrorCheck = New System.Windows.Forms.ToolStripButton()
        Me.TBPreviewOnClick = New System.Windows.Forms.ToolStripButton()
        Me.TBShowFileName = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBUndo = New System.Windows.Forms.ToolStripButton()
        Me.TBRedo = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBNTInput = New System.Windows.Forms.ToolStripButton()
        Me.TBTimeSelect = New System.Windows.Forms.ToolStripButton()
        Me.TBSelect = New System.Windows.Forms.ToolStripButton()
        Me.TBWrite = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBPlayB = New System.Windows.Forms.ToolStripButton()
        Me.TBPlay = New System.Windows.Forms.ToolStripButton()
        Me.TBStop = New System.Windows.Forms.ToolStripButton()
        Me.TBPOptions = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.TBVOptions = New System.Windows.Forms.ToolStripButton()
        Me.TBGOptions = New System.Windows.Forms.ToolStripButton()
        Me.TBAbout = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator13 = New System.Windows.Forms.ToolStripSeparator()
        Me.POBStorm = New System.Windows.Forms.ToolStripButton()
        Me.pStatus = New System.Windows.Forms.Panel()
        Me.FStatus2 = New System.Windows.Forms.StatusStrip()
        Me.FSSS = New System.Windows.Forms.ToolStripButton()
        Me.FSSL = New System.Windows.Forms.ToolStripButton()
        Me.FSSH = New System.Windows.Forms.ToolStripButton()
        Me.BVCReverse = New System.Windows.Forms.ToolStripButton()
        Me.LblMultiply = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TVCM = New System.Windows.Forms.ToolStripTextBox()
        Me.LblDivide = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TVCD = New System.Windows.Forms.ToolStripTextBox()
        Me.BVCApply = New System.Windows.Forms.ToolStripButton()
        Me.TVCBPM = New System.Windows.Forms.ToolStripTextBox()
        Me.BVCCalculate = New System.Windows.Forms.ToolStripButton()
        Me.FStatus = New System.Windows.Forms.StatusStrip()
        Me.FSC = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSW = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSM = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSP1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSP3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSP2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSP4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FST = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSH = New System.Windows.Forms.ToolStripStatusLabel()
        Me.FSE = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TimerMiddle = New System.Windows.Forms.Timer(Me.components)
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.PMain = New System.Windows.Forms.Panel()
        Me.PMainIn = New System.Windows.Forms.Panel()
        Me.ttlIcon = New System.Windows.Forms.PictureBox()
        Me.MainPanelScroll = New System.Windows.Forms.VScrollBar()
        Me.HS = New System.Windows.Forms.HScrollBar()
        Me.SpR = New System.Windows.Forms.Button()
        Me.SpL = New System.Windows.Forms.Button()
        Me.PMainR = New System.Windows.Forms.Panel()
        Me.PMainInR = New System.Windows.Forms.Panel()
        Me.RightPanelScroll = New System.Windows.Forms.VScrollBar()
        Me.HSR = New System.Windows.Forms.HScrollBar()
        Me.PMainL = New System.Windows.Forms.Panel()
        Me.PMainInL = New System.Windows.Forms.Panel()
        Me.LeftPanelScroll = New System.Windows.Forms.VScrollBar()
        Me.HSL = New System.Windows.Forms.HScrollBar()
        Me.POptionsResizer = New System.Windows.Forms.Button()
        Me.ToolTipUniversal = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmnLanguage.SuspendLayout()
        Me.cmnTheme.SuspendLayout()
        Me.POptionsScroll.SuspendLayout()
        Me.POptions.SuspendLayout()
        Me.POExpansion.SuspendLayout()
        Me.POExpansionInner.SuspendLayout()
        Me.POBeat.SuspendLayout()
        Me.POBeatInner.SuspendLayout()
        Me.TableLayoutPanel7.SuspendLayout()
        CType(Me.nBeatD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nBeatN, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.POBeatPart2.SuspendLayout()
        Me.POWAV.SuspendLayout()
        Me.POWAVInner.SuspendLayout()
        Me.FlowLayoutPanel3.SuspendLayout()
        Me.POWAVPart2.SuspendLayout()
        Me.POWaveForm.SuspendLayout()
        Me.POWaveFormInner.SuspendLayout()
        Me.POWaveFormPart2.SuspendLayout()
        CType(Me.TWSaturation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWTransparency, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWPrecision, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWWidth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWLeft, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWSaturation2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWLeft2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWTransparency2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWWidth2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWPrecision2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.POWaveFormPart1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWPosition2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TWPosition, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.POGrid.SuspendLayout()
        Me.POGridInner.SuspendLayout()
        Me.POGridPart2.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        CType(Me.CGB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.POGridPart1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGHeight2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGHeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGWidth2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGWidth, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGDivide, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CGSub, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.POHeader.SuspendLayout()
        Me.POHeaderInner.SuspendLayout()
        Me.POHeaderPart2.SuspendLayout()
        Me.POHeaderPart1.SuspendLayout()
        CType(Me.THBPM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Menu1.SuspendLayout()
        Me.mnMain.SuspendLayout()
        Me.cmnConversion.SuspendLayout()
        Me.TBMain.SuspendLayout()
        Me.pStatus.SuspendLayout()
        Me.FStatus2.SuspendLayout()
        Me.FStatus.SuspendLayout()
        Me.ToolStripContainer1.ContentPanel.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.PMain.SuspendLayout()
        Me.PMainIn.SuspendLayout()
        CType(Me.ttlIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PMainR.SuspendLayout()
        Me.PMainL.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmnLanguage
        '
        Me.cmnLanguage.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TBLangDef, Me.TBLangRefresh, Me.ToolStripSeparator9})
        Me.cmnLanguage.Name = "cmnLanguage"
        Me.cmnLanguage.OwnerItem = Me.TBLanguage
        Me.cmnLanguage.Size = New System.Drawing.Size(121, 54)
        '
        'TBLangDef
        '
        Me.TBLangDef.Name = "TBLangDef"
        Me.TBLangDef.Size = New System.Drawing.Size(120, 22)
        Me.TBLangDef.Text = "(Default)"
        '
        'TBLangRefresh
        '
        Me.TBLangRefresh.Image = Global.iBMSC.My.Resources.Resources.x16Refresh
        Me.TBLangRefresh.Name = "TBLangRefresh"
        Me.TBLangRefresh.Size = New System.Drawing.Size(120, 22)
        Me.TBLangRefresh.Text = "Refresh"
        '
        'ToolStripSeparator9
        '
        Me.ToolStripSeparator9.Name = "ToolStripSeparator9"
        Me.ToolStripSeparator9.Size = New System.Drawing.Size(117, 6)
        '
        'mnLanguage
        '
        Me.mnLanguage.DropDown = Me.cmnLanguage
        Me.mnLanguage.Image = Global.iBMSC.My.Resources.Resources.x16Language
        Me.mnLanguage.Name = "mnLanguage"
        Me.mnLanguage.Size = New System.Drawing.Size(229, 22)
        Me.mnLanguage.Text = "&Language"
        '
        'TBLanguage
        '
        Me.TBLanguage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBLanguage.DropDown = Me.cmnLanguage
        Me.TBLanguage.Image = Global.iBMSC.My.Resources.Resources.x16Language
        Me.TBLanguage.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBLanguage.Name = "TBLanguage"
        Me.TBLanguage.Size = New System.Drawing.Size(29, 22)
        Me.TBLanguage.Text = "Language"
        '
        'cmnTheme
        '
        Me.cmnTheme.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TBThemeDef, Me.TBThemeSave, Me.TBThemeRefresh, Me.TBThemeLoadComptability, Me.ToolStripSeparator6})
        Me.cmnTheme.Name = "cmnLanguage"
        Me.cmnTheme.OwnerItem = Me.TBTheme
        Me.cmnTheme.Size = New System.Drawing.Size(246, 98)
        '
        'TBThemeDef
        '
        Me.TBThemeDef.Name = "TBThemeDef"
        Me.TBThemeDef.Size = New System.Drawing.Size(245, 22)
        Me.TBThemeDef.Text = "(Default)"
        '
        'TBThemeSave
        '
        Me.TBThemeSave.Image = Global.iBMSC.My.Resources.Resources.x16SaveAs
        Me.TBThemeSave.Name = "TBThemeSave"
        Me.TBThemeSave.Size = New System.Drawing.Size(245, 22)
        Me.TBThemeSave.Text = "Save Theme"
        '
        'TBThemeRefresh
        '
        Me.TBThemeRefresh.Image = Global.iBMSC.My.Resources.Resources.x16Refresh
        Me.TBThemeRefresh.Name = "TBThemeRefresh"
        Me.TBThemeRefresh.Size = New System.Drawing.Size(245, 22)
        Me.TBThemeRefresh.Text = "Refresh"
        '
        'TBThemeLoadComptability
        '
        Me.TBThemeLoadComptability.Name = "TBThemeLoadComptability"
        Me.TBThemeLoadComptability.Size = New System.Drawing.Size(245, 22)
        Me.TBThemeLoadComptability.Text = "Load Theme File from iBMSC 2.x"
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(242, 6)
        '
        'mnTheme
        '
        Me.mnTheme.DropDown = Me.cmnTheme
        Me.mnTheme.Image = Global.iBMSC.My.Resources.Resources.x16Theme
        Me.mnTheme.Name = "mnTheme"
        Me.mnTheme.Size = New System.Drawing.Size(229, 22)
        Me.mnTheme.Text = "&Theme"
        '
        'TBTheme
        '
        Me.TBTheme.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBTheme.DropDown = Me.cmnTheme
        Me.TBTheme.Image = Global.iBMSC.My.Resources.Resources.x16Theme
        Me.TBTheme.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBTheme.Name = "TBTheme"
        Me.TBTheme.Size = New System.Drawing.Size(29, 22)
        Me.TBTheme.Text = "Theme"
        '
        'POptionsScroll
        '
        Me.POptionsScroll.AutoScroll = True
        Me.POptionsScroll.Controls.Add(Me.POptions)
        Me.POptionsScroll.Dock = System.Windows.Forms.DockStyle.Right
        Me.POptionsScroll.Location = New System.Drawing.Point(882, 0)
        Me.POptionsScroll.Name = "POptionsScroll"
        Me.POptionsScroll.Size = New System.Drawing.Size(200, 730)
        Me.POptionsScroll.TabIndex = 28
        '
        'POptions
        '
        Me.POptions.AutoSize = True
        Me.POptions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POptions.Controls.Add(Me.POExpansion)
        Me.POptions.Controls.Add(Me.POBeat)
        Me.POptions.Controls.Add(Me.POWAV)
        Me.POptions.Controls.Add(Me.POWaveForm)
        Me.POptions.Controls.Add(Me.POGrid)
        Me.POptions.Controls.Add(Me.POHeader)
        Me.POptions.Dock = System.Windows.Forms.DockStyle.Top
        Me.POptions.Location = New System.Drawing.Point(0, 0)
        Me.POptions.Name = "POptions"
        Me.POptions.Size = New System.Drawing.Size(183, 1722)
        Me.POptions.TabIndex = 29
        '
        'POExpansion
        '
        Me.POExpansion.AutoSize = True
        Me.POExpansion.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POExpansion.Controls.Add(Me.POExpansionInner)
        Me.POExpansion.Controls.Add(Me.POExpansionSwitch)
        Me.POExpansion.Dock = System.Windows.Forms.DockStyle.Top
        Me.POExpansion.Location = New System.Drawing.Point(0, 1452)
        Me.POExpansion.Name = "POExpansion"
        Me.POExpansion.Size = New System.Drawing.Size(183, 270)
        Me.POExpansion.TabIndex = 6
        '
        'POExpansionInner
        '
        Me.POExpansionInner.Controls.Add(Me.TExpansion)
        Me.POExpansionInner.Controls.Add(Me.POExpansionResizer)
        Me.POExpansionInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POExpansionInner.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POExpansionInner.Location = New System.Drawing.Point(0, 20)
        Me.POExpansionInner.Name = "POExpansionInner"
        Me.POExpansionInner.Size = New System.Drawing.Size(183, 250)
        Me.POExpansionInner.TabIndex = 7
        Me.POExpansionInner.Visible = False
        '
        'TExpansion
        '
        Me.TExpansion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TExpansion.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TExpansion.HideSelection = False
        Me.TExpansion.Location = New System.Drawing.Point(0, 0)
        Me.TExpansion.Multiline = True
        Me.TExpansion.Name = "TExpansion"
        Me.TExpansion.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TExpansion.Size = New System.Drawing.Size(183, 245)
        Me.TExpansion.TabIndex = 0
        Me.TExpansion.WordWrap = False
        '
        'POExpansionResizer
        '
        Me.POExpansionResizer.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.POExpansionResizer.FlatAppearance.BorderSize = 0
        Me.POExpansionResizer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.POExpansionResizer.Location = New System.Drawing.Point(0, 245)
        Me.POExpansionResizer.Margin = New System.Windows.Forms.Padding(0)
        Me.POExpansionResizer.Name = "POExpansionResizer"
        Me.POExpansionResizer.Size = New System.Drawing.Size(183, 5)
        Me.POExpansionResizer.TabIndex = 65
        Me.POExpansionResizer.TabStop = False
        Me.POExpansionResizer.UseVisualStyleBackColor = True
        '
        'POExpansionSwitch
        '
        Me.POExpansionSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POExpansionSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POExpansionSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POExpansionSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POExpansionSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POExpansionSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POExpansionSwitch.Name = "POExpansionSwitch"
        Me.POExpansionSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POExpansionSwitch.TabIndex = 6
        Me.POExpansionSwitch.TabStop = False
        Me.POExpansionSwitch.Text = "Expansion Code"
        Me.POExpansionSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POExpansionSwitch.UseCompatibleTextRendering = True
        Me.POExpansionSwitch.UseVisualStyleBackColor = False
        '
        'POBeat
        '
        Me.POBeat.AutoSize = True
        Me.POBeat.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POBeat.Controls.Add(Me.POBeatInner)
        Me.POBeat.Controls.Add(Me.POBeatSwitch)
        Me.POBeat.Dock = System.Windows.Forms.DockStyle.Top
        Me.POBeat.Location = New System.Drawing.Point(0, 1182)
        Me.POBeat.Name = "POBeat"
        Me.POBeat.Size = New System.Drawing.Size(183, 270)
        Me.POBeat.TabIndex = 5
        '
        'POBeatInner
        '
        Me.POBeatInner.ColumnCount = 1
        Me.POBeatInner.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POBeatInner.Controls.Add(Me.POBeatExpander, 0, 1)
        Me.POBeatInner.Controls.Add(Me.POBeatResizer, 0, 4)
        Me.POBeatInner.Controls.Add(Me.TableLayoutPanel7, 0, 0)
        Me.POBeatInner.Controls.Add(Me.LBeat, 0, 3)
        Me.POBeatInner.Controls.Add(Me.POBeatPart2, 0, 2)
        Me.POBeatInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POBeatInner.Location = New System.Drawing.Point(0, 20)
        Me.POBeatInner.Name = "POBeatInner"
        Me.POBeatInner.RowCount = 5
        Me.POBeatInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatInner.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POBeatInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatInner.Size = New System.Drawing.Size(183, 250)
        Me.POBeatInner.TabIndex = 6
        Me.POBeatInner.Visible = False
        '
        'POBeatExpander
        '
        Me.POBeatExpander.Appearance = System.Windows.Forms.Appearance.Button
        Me.POBeatExpander.AutoSize = True
        Me.POBeatExpander.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POBeatExpander.Dock = System.Windows.Forms.DockStyle.Top
        Me.POBeatExpander.FlatAppearance.BorderSize = 0
        Me.POBeatExpander.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.POBeatExpander.Location = New System.Drawing.Point(0, 54)
        Me.POBeatExpander.Margin = New System.Windows.Forms.Padding(0)
        Me.POBeatExpander.Name = "POBeatExpander"
        Me.POBeatExpander.Size = New System.Drawing.Size(183, 25)
        Me.POBeatExpander.TabIndex = 65
        Me.POBeatExpander.TabStop = False
        Me.POBeatExpander.Text = "Expand..."
        Me.POBeatExpander.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POBeatExpander.UseVisualStyleBackColor = False
        '
        'POBeatResizer
        '
        Me.POBeatInner.SetColumnSpan(Me.POBeatResizer, 2)
        Me.POBeatResizer.Dock = System.Windows.Forms.DockStyle.Top
        Me.POBeatResizer.FlatAppearance.BorderSize = 0
        Me.POBeatResizer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.POBeatResizer.Location = New System.Drawing.Point(0, 245)
        Me.POBeatResizer.Margin = New System.Windows.Forms.Padding(0)
        Me.POBeatResizer.Name = "POBeatResizer"
        Me.POBeatResizer.Size = New System.Drawing.Size(183, 5)
        Me.POBeatResizer.TabIndex = 64
        Me.POBeatResizer.TabStop = False
        Me.POBeatResizer.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel7
        '
        Me.TableLayoutPanel7.AutoSize = True
        Me.TableLayoutPanel7.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel7.ColumnCount = 4
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel7.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel7.Controls.Add(Me.nBeatD, 2, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.BBeatApplyV, 3, 1)
        Me.TableLayoutPanel7.Controls.Add(Me.nBeatN, 0, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.BBeatApply, 3, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.Label7, 1, 0)
        Me.TableLayoutPanel7.Controls.Add(Me.tBeatValue, 0, 1)
        Me.TableLayoutPanel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel7.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel7.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel7.Name = "TableLayoutPanel7"
        Me.TableLayoutPanel7.RowCount = 2
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel7.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel7.Size = New System.Drawing.Size(183, 54)
        Me.TableLayoutPanel7.TabIndex = 63
        '
        'nBeatD
        '
        Me.nBeatD.Dock = System.Windows.Forms.DockStyle.Fill
        Me.nBeatD.Location = New System.Drawing.Point(66, 3)
        Me.nBeatD.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nBeatD.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nBeatD.Name = "nBeatD"
        Me.nBeatD.Size = New System.Drawing.Size(45, 23)
        Me.nBeatD.TabIndex = 37
        Me.nBeatD.Value = New Decimal(New Integer() {4, 0, 0, 0})
        '
        'BBeatApplyV
        '
        Me.BBeatApplyV.AutoSize = True
        Me.BBeatApplyV.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BBeatApplyV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BBeatApplyV.Location = New System.Drawing.Point(114, 29)
        Me.BBeatApplyV.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.BBeatApplyV.Name = "BBeatApplyV"
        Me.BBeatApplyV.Size = New System.Drawing.Size(66, 25)
        Me.BBeatApplyV.TabIndex = 35
        Me.BBeatApplyV.Text = "Apply"
        Me.BBeatApplyV.UseVisualStyleBackColor = True
        '
        'nBeatN
        '
        Me.nBeatN.Dock = System.Windows.Forms.DockStyle.Fill
        Me.nBeatN.Location = New System.Drawing.Point(3, 3)
        Me.nBeatN.Maximum = New Decimal(New Integer() {99999, 0, 0, 0})
        Me.nBeatN.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nBeatN.Name = "nBeatN"
        Me.nBeatN.Size = New System.Drawing.Size(45, 23)
        Me.nBeatN.TabIndex = 27
        Me.nBeatN.Value = New Decimal(New Integer() {4, 0, 0, 0})
        '
        'BBeatApply
        '
        Me.BBeatApply.AutoSize = True
        Me.BBeatApply.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BBeatApply.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BBeatApply.Location = New System.Drawing.Point(114, 2)
        Me.BBeatApply.Margin = New System.Windows.Forms.Padding(0, 2, 3, 2)
        Me.BBeatApply.Name = "BBeatApply"
        Me.BBeatApply.Size = New System.Drawing.Size(66, 25)
        Me.BBeatApply.TabIndex = 30
        Me.BBeatApply.Text = "Apply"
        Me.BBeatApply.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(51, 7)
        Me.Label7.Margin = New System.Windows.Forms.Padding(0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(12, 15)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "/"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tBeatValue
        '
        Me.TableLayoutPanel7.SetColumnSpan(Me.tBeatValue, 3)
        Me.tBeatValue.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tBeatValue.Location = New System.Drawing.Point(3, 30)
        Me.tBeatValue.Margin = New System.Windows.Forms.Padding(3, 1, 3, 0)
        Me.tBeatValue.Name = "tBeatValue"
        Me.tBeatValue.Size = New System.Drawing.Size(108, 23)
        Me.tBeatValue.TabIndex = 36
        Me.tBeatValue.Text = "1"
        '
        'LBeat
        '
        Me.LBeat.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LBeat.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBeat.IntegralHeight = False
        Me.LBeat.ItemHeight = 14
        Me.LBeat.Items.AddRange(New Object() {"000: 15.984375 (1023/64)"})
        Me.LBeat.Location = New System.Drawing.Point(3, 155)
        Me.LBeat.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.LBeat.Name = "LBeat"
        Me.LBeat.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.LBeat.Size = New System.Drawing.Size(177, 90)
        Me.LBeat.TabIndex = 26
        '
        'POBeatPart2
        '
        Me.POBeatPart2.AutoSize = True
        Me.POBeatPart2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POBeatPart2.ColumnCount = 1
        Me.POBeatPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POBeatPart2.Controls.Add(Me.CBeatScale, 0, 3)
        Me.POBeatPart2.Controls.Add(Me.CBeatCut, 0, 2)
        Me.POBeatPart2.Controls.Add(Me.CBeatMeasure, 0, 1)
        Me.POBeatPart2.Controls.Add(Me.CBeatPreserve, 0, 0)
        Me.POBeatPart2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.POBeatPart2.Location = New System.Drawing.Point(0, 79)
        Me.POBeatPart2.Margin = New System.Windows.Forms.Padding(0)
        Me.POBeatPart2.Name = "POBeatPart2"
        Me.POBeatPart2.RowCount = 4
        Me.POBeatPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POBeatPart2.Size = New System.Drawing.Size(183, 76)
        Me.POBeatPart2.TabIndex = 66
        Me.POBeatPart2.Visible = False
        '
        'CBeatScale
        '
        Me.CBeatScale.AutoSize = True
        Me.CBeatScale.Location = New System.Drawing.Point(3, 57)
        Me.CBeatScale.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CBeatScale.Name = "CBeatScale"
        Me.CBeatScale.Size = New System.Drawing.Size(151, 19)
        Me.CBeatScale.TabIndex = 3
        Me.CBeatScale.Text = "Scale to measure length"
        Me.CBeatScale.UseVisualStyleBackColor = True
        '
        'CBeatCut
        '
        Me.CBeatCut.AutoEllipsis = True
        Me.CBeatCut.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CBeatCut.Location = New System.Drawing.Point(3, 38)
        Me.CBeatCut.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CBeatCut.Name = "CBeatCut"
        Me.CBeatCut.Size = New System.Drawing.Size(177, 19)
        Me.CBeatCut.TabIndex = 2
        Me.CBeatCut.Text = "Keep measure position and cut overflow"
        Me.CBeatCut.UseVisualStyleBackColor = True
        '
        'CBeatMeasure
        '
        Me.CBeatMeasure.AutoSize = True
        Me.CBeatMeasure.Location = New System.Drawing.Point(3, 19)
        Me.CBeatMeasure.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CBeatMeasure.Name = "CBeatMeasure"
        Me.CBeatMeasure.Size = New System.Drawing.Size(145, 19)
        Me.CBeatMeasure.TabIndex = 1
        Me.CBeatMeasure.Text = "Keep measure position"
        Me.CBeatMeasure.UseVisualStyleBackColor = True
        '
        'CBeatPreserve
        '
        Me.CBeatPreserve.AutoSize = True
        Me.CBeatPreserve.Checked = True
        Me.CBeatPreserve.Location = New System.Drawing.Point(3, 0)
        Me.CBeatPreserve.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CBeatPreserve.Name = "CBeatPreserve"
        Me.CBeatPreserve.Size = New System.Drawing.Size(145, 19)
        Me.CBeatPreserve.TabIndex = 0
        Me.CBeatPreserve.TabStop = True
        Me.CBeatPreserve.Text = "Keep absolute position"
        Me.CBeatPreserve.UseVisualStyleBackColor = True
        '
        'POBeatSwitch
        '
        Me.POBeatSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POBeatSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POBeatSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POBeatSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POBeatSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POBeatSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POBeatSwitch.Name = "POBeatSwitch"
        Me.POBeatSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POBeatSwitch.TabIndex = 5
        Me.POBeatSwitch.TabStop = False
        Me.POBeatSwitch.Text = "Beat"
        Me.POBeatSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POBeatSwitch.UseCompatibleTextRendering = True
        Me.POBeatSwitch.UseVisualStyleBackColor = False
        '
        'POWAV
        '
        Me.POWAV.AllowDrop = True
        Me.POWAV.AutoSize = True
        Me.POWAV.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWAV.Controls.Add(Me.POWAVInner)
        Me.POWAV.Controls.Add(Me.POWAVSwitch)
        Me.POWAV.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWAV.Location = New System.Drawing.Point(0, 912)
        Me.POWAV.Name = "POWAV"
        Me.POWAV.Size = New System.Drawing.Size(183, 270)
        Me.POWAV.TabIndex = 4
        '
        'POWAVInner
        '
        Me.POWAVInner.ColumnCount = 1
        Me.POWAVInner.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POWAVInner.Controls.Add(Me.POWAVExpander, 0, 1)
        Me.POWAVInner.Controls.Add(Me.LWAV, 0, 3)
        Me.POWAVInner.Controls.Add(Me.FlowLayoutPanel3, 0, 0)
        Me.POWAVInner.Controls.Add(Me.POWAVResizer, 0, 4)
        Me.POWAVInner.Controls.Add(Me.POWAVPart2, 0, 2)
        Me.POWAVInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWAVInner.Location = New System.Drawing.Point(0, 20)
        Me.POWAVInner.Name = "POWAVInner"
        Me.POWAVInner.RowCount = 5
        Me.POWAVInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVInner.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POWAVInner.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVInner.Size = New System.Drawing.Size(183, 250)
        Me.POWAVInner.TabIndex = 5
        '
        'POWAVExpander
        '
        Me.POWAVExpander.Appearance = System.Windows.Forms.Appearance.Button
        Me.POWAVExpander.AutoSize = True
        Me.POWAVExpander.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POWAVExpander.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWAVExpander.FlatAppearance.BorderSize = 0
        Me.POWAVExpander.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.POWAVExpander.Location = New System.Drawing.Point(0, 30)
        Me.POWAVExpander.Margin = New System.Windows.Forms.Padding(0)
        Me.POWAVExpander.Name = "POWAVExpander"
        Me.POWAVExpander.Size = New System.Drawing.Size(183, 25)
        Me.POWAVExpander.TabIndex = 34
        Me.POWAVExpander.TabStop = False
        Me.POWAVExpander.Text = "Expand..."
        Me.POWAVExpander.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POWAVExpander.UseVisualStyleBackColor = False
        '
        'LWAV
        '
        Me.LWAV.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LWAV.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LWAV.IntegralHeight = False
        Me.LWAV.ItemHeight = 14
        Me.LWAV.Location = New System.Drawing.Point(3, 93)
        Me.LWAV.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.LWAV.Name = "LWAV"
        Me.LWAV.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.LWAV.Size = New System.Drawing.Size(177, 152)
        Me.LWAV.TabIndex = 25
        '
        'FlowLayoutPanel3
        '
        Me.FlowLayoutPanel3.AutoSize = True
        Me.FlowLayoutPanel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.FlowLayoutPanel3.Controls.Add(Me.BWAVUp)
        Me.FlowLayoutPanel3.Controls.Add(Me.BWAVDown)
        Me.FlowLayoutPanel3.Controls.Add(Me.BWAVBrowse)
        Me.FlowLayoutPanel3.Controls.Add(Me.BWAVRemove)
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(3, 3)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(96, 24)
        Me.FlowLayoutPanel3.TabIndex = 26
        Me.FlowLayoutPanel3.WrapContents = False
        '
        'BWAVUp
        '
        Me.BWAVUp.Image = Global.iBMSC.My.Resources.Resources.x16Up
        Me.BWAVUp.Location = New System.Drawing.Point(0, 0)
        Me.BWAVUp.Margin = New System.Windows.Forms.Padding(0)
        Me.BWAVUp.Name = "BWAVUp"
        Me.BWAVUp.Size = New System.Drawing.Size(24, 24)
        Me.BWAVUp.TabIndex = 26
        Me.ToolTipUniversal.SetToolTip(Me.BWAVUp, "Move Up")
        Me.BWAVUp.UseVisualStyleBackColor = True
        '
        'BWAVDown
        '
        Me.BWAVDown.Image = Global.iBMSC.My.Resources.Resources.x16Down
        Me.BWAVDown.Location = New System.Drawing.Point(24, 0)
        Me.BWAVDown.Margin = New System.Windows.Forms.Padding(0)
        Me.BWAVDown.Name = "BWAVDown"
        Me.BWAVDown.Size = New System.Drawing.Size(24, 24)
        Me.BWAVDown.TabIndex = 27
        Me.ToolTipUniversal.SetToolTip(Me.BWAVDown, "Move Down")
        Me.BWAVDown.UseVisualStyleBackColor = True
        '
        'BWAVBrowse
        '
        Me.BWAVBrowse.Image = Global.iBMSC.My.Resources.Resources.x16PlayerBrowse
        Me.BWAVBrowse.Location = New System.Drawing.Point(48, 0)
        Me.BWAVBrowse.Margin = New System.Windows.Forms.Padding(0)
        Me.BWAVBrowse.Name = "BWAVBrowse"
        Me.BWAVBrowse.Size = New System.Drawing.Size(24, 24)
        Me.BWAVBrowse.TabIndex = 30
        Me.ToolTipUniversal.SetToolTip(Me.BWAVBrowse, "Browse")
        Me.BWAVBrowse.UseVisualStyleBackColor = True
        '
        'BWAVRemove
        '
        Me.BWAVRemove.Image = Global.iBMSC.My.Resources.Resources.x16Remove
        Me.BWAVRemove.Location = New System.Drawing.Point(72, 0)
        Me.BWAVRemove.Margin = New System.Windows.Forms.Padding(0)
        Me.BWAVRemove.Name = "BWAVRemove"
        Me.BWAVRemove.Size = New System.Drawing.Size(24, 24)
        Me.BWAVRemove.TabIndex = 31
        Me.ToolTipUniversal.SetToolTip(Me.BWAVRemove, "Remove")
        Me.BWAVRemove.UseVisualStyleBackColor = True
        '
        'POWAVResizer
        '
        Me.POWAVResizer.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWAVResizer.FlatAppearance.BorderSize = 0
        Me.POWAVResizer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.POWAVResizer.Location = New System.Drawing.Point(0, 245)
        Me.POWAVResizer.Margin = New System.Windows.Forms.Padding(0)
        Me.POWAVResizer.Name = "POWAVResizer"
        Me.POWAVResizer.Size = New System.Drawing.Size(183, 5)
        Me.POWAVResizer.TabIndex = 33
        Me.POWAVResizer.TabStop = False
        Me.POWAVResizer.UseVisualStyleBackColor = True
        '
        'POWAVPart2
        '
        Me.POWAVPart2.AutoSize = True
        Me.POWAVPart2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWAVPart2.ColumnCount = 1
        Me.POWAVPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POWAVPart2.Controls.Add(Me.CWAVMultiSelect, 0, 0)
        Me.POWAVPart2.Controls.Add(Me.CWAVChangeLabel, 0, 1)
        Me.POWAVPart2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.POWAVPart2.Location = New System.Drawing.Point(0, 55)
        Me.POWAVPart2.Margin = New System.Windows.Forms.Padding(0)
        Me.POWAVPart2.Name = "POWAVPart2"
        Me.POWAVPart2.RowCount = 2
        Me.POWAVPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWAVPart2.Size = New System.Drawing.Size(183, 38)
        Me.POWAVPart2.TabIndex = 35
        Me.POWAVPart2.Visible = False
        '
        'CWAVMultiSelect
        '
        Me.CWAVMultiSelect.AutoSize = True
        Me.CWAVMultiSelect.Checked = True
        Me.CWAVMultiSelect.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CWAVMultiSelect.Location = New System.Drawing.Point(3, 0)
        Me.CWAVMultiSelect.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CWAVMultiSelect.Name = "CWAVMultiSelect"
        Me.CWAVMultiSelect.Size = New System.Drawing.Size(154, 19)
        Me.CWAVMultiSelect.TabIndex = 0
        Me.CWAVMultiSelect.Text = "Allow Multiple Selection"
        Me.CWAVMultiSelect.UseVisualStyleBackColor = True
        '
        'CWAVChangeLabel
        '
        Me.CWAVChangeLabel.AutoSize = True
        Me.CWAVChangeLabel.Checked = True
        Me.CWAVChangeLabel.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CWAVChangeLabel.Location = New System.Drawing.Point(3, 19)
        Me.CWAVChangeLabel.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CWAVChangeLabel.Name = "CWAVChangeLabel"
        Me.CWAVChangeLabel.Size = New System.Drawing.Size(155, 19)
        Me.CWAVChangeLabel.TabIndex = 1
        Me.CWAVChangeLabel.Text = "Synchronize Note Labels"
        Me.CWAVChangeLabel.UseVisualStyleBackColor = True
        '
        'POWAVSwitch
        '
        Me.POWAVSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POWAVSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POWAVSwitch.Checked = True
        Me.POWAVSwitch.CheckState = System.Windows.Forms.CheckState.Checked
        Me.POWAVSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POWAVSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWAVSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POWAVSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POWAVSwitch.Name = "POWAVSwitch"
        Me.POWAVSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POWAVSwitch.TabIndex = 4
        Me.POWAVSwitch.TabStop = False
        Me.POWAVSwitch.Text = "#WAV (Sounds List)"
        Me.POWAVSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POWAVSwitch.UseCompatibleTextRendering = True
        Me.POWAVSwitch.UseVisualStyleBackColor = False
        '
        'POWaveForm
        '
        Me.POWaveForm.AutoSize = True
        Me.POWaveForm.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWaveForm.Controls.Add(Me.POWaveFormInner)
        Me.POWaveForm.Controls.Add(Me.POWaveFormSwitch)
        Me.POWaveForm.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveForm.Location = New System.Drawing.Point(0, 669)
        Me.POWaveForm.Name = "POWaveForm"
        Me.POWaveForm.Size = New System.Drawing.Size(183, 243)
        Me.POWaveForm.TabIndex = 3
        '
        'POWaveFormInner
        '
        Me.POWaveFormInner.AutoSize = True
        Me.POWaveFormInner.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWaveFormInner.Controls.Add(Me.POWaveFormPart2)
        Me.POWaveFormInner.Controls.Add(Me.POWaveFormExpander)
        Me.POWaveFormInner.Controls.Add(Me.POWaveFormPart1)
        Me.POWaveFormInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveFormInner.Location = New System.Drawing.Point(0, 20)
        Me.POWaveFormInner.Name = "POWaveFormInner"
        Me.POWaveFormInner.Size = New System.Drawing.Size(183, 223)
        Me.POWaveFormInner.TabIndex = 29
        Me.POWaveFormInner.Visible = False
        '
        'POWaveFormPart2
        '
        Me.POWaveFormPart2.AutoSize = True
        Me.POWaveFormPart2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWaveFormPart2.ColumnCount = 3
        Me.POWaveFormPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.POWaveFormPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.0!))
        Me.POWaveFormPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.POWaveFormPart2.Controls.Add(Me.TWSaturation, 2, 4)
        Me.POWaveFormPart2.Controls.Add(Me.PictureBox2, 0, 0)
        Me.POWaveFormPart2.Controls.Add(Me.TWTransparency, 2, 3)
        Me.POWaveFormPart2.Controls.Add(Me.PictureBox3, 0, 1)
        Me.POWaveFormPart2.Controls.Add(Me.TWPrecision, 2, 2)
        Me.POWaveFormPart2.Controls.Add(Me.PictureBox4, 0, 2)
        Me.POWaveFormPart2.Controls.Add(Me.TWWidth, 2, 1)
        Me.POWaveFormPart2.Controls.Add(Me.PictureBox5, 0, 3)
        Me.POWaveFormPart2.Controls.Add(Me.TWLeft, 2, 0)
        Me.POWaveFormPart2.Controls.Add(Me.PictureBox6, 0, 4)
        Me.POWaveFormPart2.Controls.Add(Me.TWSaturation2, 1, 4)
        Me.POWaveFormPart2.Controls.Add(Me.TWLeft2, 1, 0)
        Me.POWaveFormPart2.Controls.Add(Me.TWTransparency2, 1, 3)
        Me.POWaveFormPart2.Controls.Add(Me.TWWidth2, 1, 1)
        Me.POWaveFormPart2.Controls.Add(Me.TWPrecision2, 1, 2)
        Me.POWaveFormPart2.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveFormPart2.Location = New System.Drawing.Point(0, 83)
        Me.POWaveFormPart2.Name = "POWaveFormPart2"
        Me.POWaveFormPart2.RowCount = 5
        Me.POWaveFormPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.POWaveFormPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.POWaveFormPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.POWaveFormPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.POWaveFormPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.POWaveFormPart2.Size = New System.Drawing.Size(183, 140)
        Me.POWaveFormPart2.TabIndex = 5
        '
        'TWSaturation
        '
        Me.TWSaturation.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWSaturation.Increment = New Decimal(New Integer() {50, 0, 0, 0})
        Me.TWSaturation.Location = New System.Drawing.Point(137, 112)
        Me.TWSaturation.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWSaturation.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.TWSaturation.Name = "TWSaturation"
        Me.TWSaturation.Size = New System.Drawing.Size(43, 23)
        Me.TWSaturation.TabIndex = 68
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = Global.iBMSC.My.Resources.Resources.WAVLeft
        Me.PictureBox2.Location = New System.Drawing.Point(3, 0)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox2.TabIndex = 60
        Me.PictureBox2.TabStop = False
        '
        'TWTransparency
        '
        Me.TWTransparency.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWTransparency.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        Me.TWTransparency.Location = New System.Drawing.Point(137, 84)
        Me.TWTransparency.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWTransparency.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.TWTransparency.Name = "TWTransparency"
        Me.TWTransparency.Size = New System.Drawing.Size(43, 23)
        Me.TWTransparency.TabIndex = 69
        Me.TWTransparency.Value = New Decimal(New Integer() {80, 0, 0, 0})
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = Global.iBMSC.My.Resources.Resources.WAVWidth
        Me.PictureBox3.Location = New System.Drawing.Point(3, 28)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox3.TabIndex = 61
        Me.PictureBox3.TabStop = False
        '
        'TWPrecision
        '
        Me.TWPrecision.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWPrecision.Location = New System.Drawing.Point(137, 56)
        Me.TWPrecision.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWPrecision.Maximum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.TWPrecision.Name = "TWPrecision"
        Me.TWPrecision.Size = New System.Drawing.Size(43, 23)
        Me.TWPrecision.TabIndex = 46
        Me.TWPrecision.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = Global.iBMSC.My.Resources.Resources.WAVPrecision
        Me.PictureBox4.Location = New System.Drawing.Point(3, 56)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox4.TabIndex = 62
        Me.PictureBox4.TabStop = False
        '
        'TWWidth
        '
        Me.TWWidth.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWWidth.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        Me.TWWidth.Location = New System.Drawing.Point(137, 28)
        Me.TWWidth.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWWidth.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.TWWidth.Name = "TWWidth"
        Me.TWWidth.Size = New System.Drawing.Size(43, 23)
        Me.TWWidth.TabIndex = 45
        Me.TWWidth.Value = New Decimal(New Integer() {200, 0, 0, 0})
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = Global.iBMSC.My.Resources.Resources.WAVTransparency
        Me.PictureBox5.Location = New System.Drawing.Point(3, 84)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox5.TabIndex = 67
        Me.PictureBox5.TabStop = False
        '
        'TWLeft
        '
        Me.TWLeft.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWLeft.Increment = New Decimal(New Integer() {5, 0, 0, 0})
        Me.TWLeft.Location = New System.Drawing.Point(137, 0)
        Me.TWLeft.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWLeft.Maximum = New Decimal(New Integer() {800, 0, 0, 0})
        Me.TWLeft.Name = "TWLeft"
        Me.TWLeft.Size = New System.Drawing.Size(43, 23)
        Me.TWLeft.TabIndex = 44
        Me.TWLeft.Value = New Decimal(New Integer() {50, 0, 0, 0})
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = Global.iBMSC.My.Resources.Resources.WAVSaturation
        Me.PictureBox6.Location = New System.Drawing.Point(3, 112)
        Me.PictureBox6.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox6.TabIndex = 66
        Me.PictureBox6.TabStop = False
        '
        'TWSaturation2
        '
        Me.TWSaturation2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWSaturation2.LargeChange = 200
        Me.TWSaturation2.Location = New System.Drawing.Point(30, 112)
        Me.TWSaturation2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWSaturation2.Maximum = 1000
        Me.TWSaturation2.Name = "TWSaturation2"
        Me.TWSaturation2.Size = New System.Drawing.Size(107, 28)
        Me.TWSaturation2.SmallChange = 50
        Me.TWSaturation2.TabIndex = 70
        Me.TWSaturation2.TickFrequency = 200
        '
        'TWLeft2
        '
        Me.TWLeft2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWLeft2.LargeChange = 50
        Me.TWLeft2.Location = New System.Drawing.Point(30, 0)
        Me.TWLeft2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWLeft2.Maximum = 800
        Me.TWLeft2.Name = "TWLeft2"
        Me.TWLeft2.Size = New System.Drawing.Size(107, 28)
        Me.TWLeft2.SmallChange = 10
        Me.TWLeft2.TabIndex = 63
        Me.TWLeft2.TickFrequency = 100
        Me.TWLeft2.Value = 50
        '
        'TWTransparency2
        '
        Me.TWTransparency2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWTransparency2.LargeChange = 64
        Me.TWTransparency2.Location = New System.Drawing.Point(30, 84)
        Me.TWTransparency2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWTransparency2.Maximum = 255
        Me.TWTransparency2.Name = "TWTransparency2"
        Me.TWTransparency2.Size = New System.Drawing.Size(107, 28)
        Me.TWTransparency2.SmallChange = 8
        Me.TWTransparency2.TabIndex = 71
        Me.TWTransparency2.TickFrequency = 64
        Me.TWTransparency2.Value = 80
        '
        'TWWidth2
        '
        Me.TWWidth2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWWidth2.LargeChange = 50
        Me.TWWidth2.Location = New System.Drawing.Point(30, 28)
        Me.TWWidth2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWWidth2.Maximum = 1000
        Me.TWWidth2.Name = "TWWidth2"
        Me.TWWidth2.Size = New System.Drawing.Size(107, 28)
        Me.TWWidth2.SmallChange = 10
        Me.TWWidth2.TabIndex = 64
        Me.TWWidth2.TickFrequency = 100
        Me.TWWidth2.Value = 200
        '
        'TWPrecision2
        '
        Me.TWPrecision2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWPrecision2.LargeChange = 4
        Me.TWPrecision2.Location = New System.Drawing.Point(30, 56)
        Me.TWPrecision2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWPrecision2.Maximum = 50
        Me.TWPrecision2.Name = "TWPrecision2"
        Me.TWPrecision2.Size = New System.Drawing.Size(107, 28)
        Me.TWPrecision2.TabIndex = 65
        Me.TWPrecision2.TickFrequency = 5
        Me.TWPrecision2.Value = 5
        '
        'POWaveFormExpander
        '
        Me.POWaveFormExpander.Appearance = System.Windows.Forms.Appearance.Button
        Me.POWaveFormExpander.AutoSize = True
        Me.POWaveFormExpander.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POWaveFormExpander.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveFormExpander.FlatAppearance.BorderSize = 0
        Me.POWaveFormExpander.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.POWaveFormExpander.Location = New System.Drawing.Point(0, 58)
        Me.POWaveFormExpander.Margin = New System.Windows.Forms.Padding(0)
        Me.POWaveFormExpander.Name = "POWaveFormExpander"
        Me.POWaveFormExpander.Size = New System.Drawing.Size(183, 25)
        Me.POWaveFormExpander.TabIndex = 29
        Me.POWaveFormExpander.TabStop = False
        Me.POWaveFormExpander.Text = "Expand..."
        Me.POWaveFormExpander.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POWaveFormExpander.UseVisualStyleBackColor = False
        '
        'POWaveFormPart1
        '
        Me.POWaveFormPart1.AutoSize = True
        Me.POWaveFormPart1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POWaveFormPart1.ColumnCount = 1
        Me.POWaveFormPart1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POWaveFormPart1.Controls.Add(Me.TableLayoutPanel1, 0, 0)
        Me.POWaveFormPart1.Controls.Add(Me.TableLayoutPanel6, 0, 1)
        Me.POWaveFormPart1.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveFormPart1.Location = New System.Drawing.Point(0, 0)
        Me.POWaveFormPart1.Name = "POWaveFormPart1"
        Me.POWaveFormPart1.RowCount = 2
        Me.POWaveFormPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWaveFormPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POWaveFormPart1.Size = New System.Drawing.Size(183, 58)
        Me.POWaveFormPart1.TabIndex = 4
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TWFileName, 1, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(183, 30)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoSize = True
        Me.FlowLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.FlowLayoutPanel1.Controls.Add(Me.BWLoad)
        Me.FlowLayoutPanel1.Controls.Add(Me.BWClear)
        Me.FlowLayoutPanel1.Controls.Add(Me.BWLock)
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 3)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(72, 24)
        Me.FlowLayoutPanel1.TabIndex = 72
        Me.FlowLayoutPanel1.WrapContents = False
        '
        'BWLoad
        '
        Me.BWLoad.Image = Global.iBMSC.My.Resources.Resources.x16Add
        Me.BWLoad.Location = New System.Drawing.Point(0, 0)
        Me.BWLoad.Margin = New System.Windows.Forms.Padding(0)
        Me.BWLoad.Name = "BWLoad"
        Me.BWLoad.Size = New System.Drawing.Size(24, 24)
        Me.BWLoad.TabIndex = 40
        Me.ToolTipUniversal.SetToolTip(Me.BWLoad, "Load WaveForm")
        Me.BWLoad.UseVisualStyleBackColor = True
        '
        'BWClear
        '
        Me.BWClear.Image = Global.iBMSC.My.Resources.Resources.x16Remove
        Me.BWClear.Location = New System.Drawing.Point(24, 0)
        Me.BWClear.Margin = New System.Windows.Forms.Padding(0)
        Me.BWClear.Name = "BWClear"
        Me.BWClear.Size = New System.Drawing.Size(24, 24)
        Me.BWClear.TabIndex = 41
        Me.ToolTipUniversal.SetToolTip(Me.BWClear, "Clear WaveForm")
        Me.BWClear.UseVisualStyleBackColor = True
        '
        'BWLock
        '
        Me.BWLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.BWLock.Checked = True
        Me.BWLock.CheckState = System.Windows.Forms.CheckState.Checked
        Me.BWLock.Image = Global.iBMSC.My.Resources.Resources.x16Lock
        Me.BWLock.Location = New System.Drawing.Point(48, 0)
        Me.BWLock.Margin = New System.Windows.Forms.Padding(0)
        Me.BWLock.Name = "BWLock"
        Me.BWLock.Size = New System.Drawing.Size(24, 24)
        Me.BWLock.TabIndex = 0
        Me.ToolTipUniversal.SetToolTip(Me.BWLock, "Lock to BGM")
        Me.BWLock.UseVisualStyleBackColor = True
        '
        'TWFileName
        '
        Me.TWFileName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWFileName.Location = New System.Drawing.Point(81, 3)
        Me.TWFileName.Name = "TWFileName"
        Me.TWFileName.ReadOnly = True
        Me.TWFileName.Size = New System.Drawing.Size(99, 23)
        Me.TWFileName.TabIndex = 42
        Me.TWFileName.Text = "(None)"
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.AutoSize = True
        Me.TableLayoutPanel6.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel6.ColumnCount = 3
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel6.Controls.Add(Me.PictureBox1, 0, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.TWPosition2, 1, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.TWPosition, 2, 0)
        Me.TableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(0, 30)
        Me.TableLayoutPanel6.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 1
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(183, 28)
        Me.TableLayoutPanel6.TabIndex = 1
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.iBMSC.My.Resources.Resources.WAVOffset
        Me.PictureBox1.Location = New System.Drawing.Point(3, 0)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 59
        Me.PictureBox1.TabStop = False
        '
        'TWPosition2
        '
        Me.TWPosition2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWPosition2.Enabled = False
        Me.TWPosition2.LargeChange = 24
        Me.TWPosition2.Location = New System.Drawing.Point(30, 0)
        Me.TWPosition2.Margin = New System.Windows.Forms.Padding(0)
        Me.TWPosition2.Maximum = 960
        Me.TWPosition2.Name = "TWPosition2"
        Me.TWPosition2.Size = New System.Drawing.Size(91, 28)
        Me.TWPosition2.TabIndex = 58
        Me.TWPosition2.TickFrequency = 192
        '
        'TWPosition
        '
        Me.TWPosition.DecimalPlaces = 2
        Me.TWPosition.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TWPosition.Enabled = False
        Me.TWPosition.Location = New System.Drawing.Point(121, 0)
        Me.TWPosition.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.TWPosition.Maximum = New Decimal(New Integer() {192000, 0, 0, 0})
        Me.TWPosition.Name = "TWPosition"
        Me.TWPosition.Size = New System.Drawing.Size(59, 23)
        Me.TWPosition.TabIndex = 43
        '
        'POWaveFormSwitch
        '
        Me.POWaveFormSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POWaveFormSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POWaveFormSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POWaveFormSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POWaveFormSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POWaveFormSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POWaveFormSwitch.Name = "POWaveFormSwitch"
        Me.POWaveFormSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POWaveFormSwitch.TabIndex = 3
        Me.POWaveFormSwitch.TabStop = False
        Me.POWaveFormSwitch.Text = "WaveForm"
        Me.POWaveFormSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POWaveFormSwitch.UseCompatibleTextRendering = True
        Me.POWaveFormSwitch.UseVisualStyleBackColor = False
        '
        'POGrid
        '
        Me.POGrid.AutoSize = True
        Me.POGrid.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POGrid.Controls.Add(Me.POGridInner)
        Me.POGrid.Controls.Add(Me.POGridSwitch)
        Me.POGrid.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGrid.Location = New System.Drawing.Point(0, 439)
        Me.POGrid.Name = "POGrid"
        Me.POGrid.Size = New System.Drawing.Size(183, 230)
        Me.POGrid.TabIndex = 2
        '
        'POGridInner
        '
        Me.POGridInner.AutoSize = True
        Me.POGridInner.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POGridInner.Controls.Add(Me.POGridPart2)
        Me.POGridInner.Controls.Add(Me.POGridExpander)
        Me.POGridInner.Controls.Add(Me.POGridPart1)
        Me.POGridInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGridInner.Location = New System.Drawing.Point(0, 20)
        Me.POGridInner.Name = "POGridInner"
        Me.POGridInner.Size = New System.Drawing.Size(183, 210)
        Me.POGridInner.TabIndex = 3
        '
        'POGridPart2
        '
        Me.POGridPart2.AutoSize = True
        Me.POGridPart2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POGridPart2.ColumnCount = 1
        Me.POGridPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POGridPart2.Controls.Add(Me.TableLayoutPanel5, 0, 1)
        Me.POGridPart2.Controls.Add(Me.TableLayoutPanel4, 0, 0)
        Me.POGridPart2.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGridPart2.Location = New System.Drawing.Point(0, 152)
        Me.POGridPart2.Name = "POGridPart2"
        Me.POGridPart2.RowCount = 2
        Me.POGridPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.POGridPart2.Size = New System.Drawing.Size(183, 58)
        Me.POGridPart2.TabIndex = 0
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.AutoSize = True
        Me.TableLayoutPanel5.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel5.ColumnCount = 2
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.FlowLayoutPanel2, 1, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Label5, 0, 0)
        Me.TableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(0, 29)
        Me.TableLayoutPanel5.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 1
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(183, 29)
        Me.TableLayoutPanel5.TabIndex = 46
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.AutoSize = True
        Me.FlowLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.FlowLayoutPanel2.Controls.Add(Me.cVSLockL)
        Me.FlowLayoutPanel2.Controls.Add(Me.cVSLock)
        Me.FlowLayoutPanel2.Controls.Add(Me.cVSLockR)
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(115, 3)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(65, 23)
        Me.FlowLayoutPanel2.TabIndex = 72
        Me.FlowLayoutPanel2.WrapContents = False
        '
        'cVSLockL
        '
        Me.cVSLockL.Appearance = System.Windows.Forms.Appearance.Button
        Me.cVSLockL.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cVSLockL.Image = Global.iBMSC.My.Resources.Resources.x16Lock
        Me.cVSLockL.Location = New System.Drawing.Point(0, 0)
        Me.cVSLockL.Margin = New System.Windows.Forms.Padding(0)
        Me.cVSLockL.Name = "cVSLockL"
        Me.cVSLockL.Size = New System.Drawing.Size(23, 23)
        Me.cVSLockL.TabIndex = 38
        Me.cVSLockL.Tag = "0"
        Me.ToolTipUniversal.SetToolTip(Me.cVSLockL, "Lock Left Editing Panel")
        '
        'cVSLock
        '
        Me.cVSLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.cVSLock.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cVSLock.Image = Global.iBMSC.My.Resources.Resources.x16Lock
        Me.cVSLock.Location = New System.Drawing.Point(23, 0)
        Me.cVSLock.Margin = New System.Windows.Forms.Padding(0)
        Me.cVSLock.Name = "cVSLock"
        Me.cVSLock.Size = New System.Drawing.Size(23, 23)
        Me.cVSLock.TabIndex = 40
        Me.cVSLock.Tag = "1"
        Me.ToolTipUniversal.SetToolTip(Me.cVSLock, "Lock Middle Editing Panel")
        '
        'cVSLockR
        '
        Me.cVSLockR.Appearance = System.Windows.Forms.Appearance.Button
        Me.cVSLockR.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cVSLockR.Image = Global.iBMSC.My.Resources.Resources.x16Lock
        Me.cVSLockR.Location = New System.Drawing.Point(46, 0)
        Me.cVSLockR.Margin = New System.Windows.Forms.Padding(0)
        Me.cVSLockR.Name = "cVSLockR"
        Me.cVSLockR.Size = New System.Drawing.Size(23, 23)
        Me.cVSLockR.TabIndex = 41
        Me.cVSLockR.Tag = "2"
        Me.ToolTipUniversal.SetToolTip(Me.cVSLockR, "Lock Right Editing Panel")
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label5.Location = New System.Drawing.Point(3, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(106, 29)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Vertical Scroll Lock"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.AutoSize = True
        Me.TableLayoutPanel4.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel4.Controls.Add(Me.Label1, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.CGB, 1, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel4.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(183, 29)
        Me.TableLayoutPanel4.TabIndex = 44
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 29)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "Number of B Columns"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CGB
        '
        Me.CGB.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGB.Location = New System.Drawing.Point(135, 3)
        Me.CGB.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.CGB.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.CGB.Name = "CGB"
        Me.CGB.Size = New System.Drawing.Size(45, 23)
        Me.CGB.TabIndex = 35
        Me.CGB.Value = New Decimal(New Integer() {15, 0, 0, 0})
        '
        'POGridExpander
        '
        Me.POGridExpander.Appearance = System.Windows.Forms.Appearance.Button
        Me.POGridExpander.AutoSize = True
        Me.POGridExpander.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POGridExpander.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGridExpander.FlatAppearance.BorderSize = 0
        Me.POGridExpander.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.POGridExpander.Location = New System.Drawing.Point(0, 127)
        Me.POGridExpander.Margin = New System.Windows.Forms.Padding(0)
        Me.POGridExpander.Name = "POGridExpander"
        Me.POGridExpander.Size = New System.Drawing.Size(183, 25)
        Me.POGridExpander.TabIndex = 27
        Me.POGridExpander.TabStop = False
        Me.POGridExpander.Text = "Expand..."
        Me.POGridExpander.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POGridExpander.UseVisualStyleBackColor = False
        '
        'POGridPart1
        '
        Me.POGridPart1.AutoSize = True
        Me.POGridPart1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POGridPart1.ColumnCount = 1
        Me.POGridPart1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POGridPart1.Controls.Add(Me.TableLayoutPanel3, 0, 1)
        Me.POGridPart1.Controls.Add(Me.CGDisableVertical, 0, 3)
        Me.POGridPart1.Controls.Add(Me.CGSnap, 0, 2)
        Me.POGridPart1.Controls.Add(Me.TableLayoutPanel2, 0, 0)
        Me.POGridPart1.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGridPart1.Location = New System.Drawing.Point(0, 0)
        Me.POGridPart1.Name = "POGridPart1"
        Me.POGridPart1.RowCount = 4
        Me.POGridPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POGridPart1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.POGridPart1.Size = New System.Drawing.Size(183, 127)
        Me.POGridPart1.TabIndex = 11
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.AutoSize = True
        Me.TableLayoutPanel3.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.PictureBox9, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.CGHeight2, 1, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.CGHeight, 2, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.PictureBox10, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.CGWidth2, 1, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.CGWidth, 2, 1)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(0, 30)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 2
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 28.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(183, 56)
        Me.TableLayoutPanel3.TabIndex = 12
        '
        'PictureBox9
        '
        Me.PictureBox9.Image = Global.iBMSC.My.Resources.Resources.lgheight
        Me.PictureBox9.Location = New System.Drawing.Point(3, 0)
        Me.PictureBox9.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox9.TabIndex = 27
        Me.PictureBox9.TabStop = False
        '
        'CGHeight2
        '
        Me.CGHeight2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGHeight2.LargeChange = 4
        Me.CGHeight2.Location = New System.Drawing.Point(30, 0)
        Me.CGHeight2.Margin = New System.Windows.Forms.Padding(0)
        Me.CGHeight2.Maximum = 20
        Me.CGHeight2.Minimum = 1
        Me.CGHeight2.Name = "CGHeight2"
        Me.CGHeight2.Size = New System.Drawing.Size(107, 28)
        Me.CGHeight2.TabIndex = 29
        Me.CGHeight2.TickFrequency = 2
        Me.CGHeight2.Value = 4
        '
        'CGHeight
        '
        Me.CGHeight.DecimalPlaces = 2
        Me.CGHeight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGHeight.Increment = New Decimal(New Integer() {25, 0, 0, 131072})
        Me.CGHeight.Location = New System.Drawing.Point(137, 0)
        Me.CGHeight.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.CGHeight.Maximum = New Decimal(New Integer() {99, 0, 0, 0})
        Me.CGHeight.Minimum = New Decimal(New Integer() {25, 0, 0, 131072})
        Me.CGHeight.Name = "CGHeight"
        Me.CGHeight.Size = New System.Drawing.Size(43, 23)
        Me.CGHeight.TabIndex = 23
        Me.CGHeight.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'PictureBox10
        '
        Me.PictureBox10.Image = Global.iBMSC.My.Resources.Resources.lgwidth
        Me.PictureBox10.Location = New System.Drawing.Point(3, 28)
        Me.PictureBox10.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox10.TabIndex = 28
        Me.PictureBox10.TabStop = False
        '
        'CGWidth2
        '
        Me.CGWidth2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGWidth2.LargeChange = 4
        Me.CGWidth2.Location = New System.Drawing.Point(30, 28)
        Me.CGWidth2.Margin = New System.Windows.Forms.Padding(0)
        Me.CGWidth2.Maximum = 20
        Me.CGWidth2.Minimum = 1
        Me.CGWidth2.Name = "CGWidth2"
        Me.CGWidth2.Size = New System.Drawing.Size(107, 28)
        Me.CGWidth2.TabIndex = 30
        Me.CGWidth2.TickFrequency = 2
        Me.CGWidth2.Value = 4
        '
        'CGWidth
        '
        Me.CGWidth.DecimalPlaces = 2
        Me.CGWidth.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGWidth.Increment = New Decimal(New Integer() {25, 0, 0, 131072})
        Me.CGWidth.Location = New System.Drawing.Point(137, 28)
        Me.CGWidth.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.CGWidth.Maximum = New Decimal(New Integer() {99, 0, 0, 0})
        Me.CGWidth.Minimum = New Decimal(New Integer() {25, 0, 0, 131072})
        Me.CGWidth.Name = "CGWidth"
        Me.CGWidth.Size = New System.Drawing.Size(43, 23)
        Me.CGWidth.TabIndex = 24
        Me.CGWidth.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'CGDisableVertical
        '
        Me.CGDisableVertical.AutoSize = True
        Me.CGDisableVertical.Location = New System.Drawing.Point(3, 108)
        Me.CGDisableVertical.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.CGDisableVertical.Name = "CGDisableVertical"
        Me.CGDisableVertical.Size = New System.Drawing.Size(162, 19)
        Me.CGDisableVertical.TabIndex = 45
        Me.CGDisableVertical.Text = "Disable vertical moves (D)"
        Me.CGDisableVertical.UseVisualStyleBackColor = True
        '
        'CGSnap
        '
        Me.CGSnap.AutoSize = True
        Me.CGSnap.Checked = True
        Me.CGSnap.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGSnap.Location = New System.Drawing.Point(3, 89)
        Me.CGSnap.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.CGSnap.Name = "CGSnap"
        Me.CGSnap.Size = New System.Drawing.Size(109, 19)
        Me.CGSnap.TabIndex = 10
        Me.CGSnap.Text = "Snap to grid (G)"
        Me.CGSnap.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AutoSize = True
        Me.TableLayoutPanel2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.Controls.Add(Me.PictureBox7, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.CGDivide, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.CGSub, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.BGSlash, 3, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Margin = New System.Windows.Forms.Padding(0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(183, 30)
        Me.TableLayoutPanel2.TabIndex = 11
        '
        'PictureBox7
        '
        Me.PictureBox7.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.PictureBox7.Image = Global.iBMSC.My.Resources.Resources.lgpartition
        Me.PictureBox7.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(24, 24)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox7.TabIndex = 25
        Me.PictureBox7.TabStop = False
        '
        'CGDivide
        '
        Me.CGDivide.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGDivide.Location = New System.Drawing.Point(33, 3)
        Me.CGDivide.Maximum = New Decimal(New Integer() {1920, 0, 0, 0})
        Me.CGDivide.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.CGDivide.Name = "CGDivide"
        Me.CGDivide.Size = New System.Drawing.Size(58, 23)
        Me.CGDivide.TabIndex = 36
        Me.CGDivide.Value = New Decimal(New Integer() {16, 0, 0, 0})
        '
        'CGSub
        '
        Me.CGSub.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CGSub.Location = New System.Drawing.Point(97, 3)
        Me.CGSub.Maximum = New Decimal(New Integer() {1920, 0, 0, 0})
        Me.CGSub.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.CGSub.Name = "CGSub"
        Me.CGSub.Size = New System.Drawing.Size(58, 23)
        Me.CGSub.TabIndex = 37
        Me.CGSub.Value = New Decimal(New Integer() {4, 0, 0, 0})
        '
        'BGSlash
        '
        Me.BGSlash.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.BGSlash.Image = Global.iBMSC.My.Resources.Resources.Shortcut
        Me.BGSlash.Location = New System.Drawing.Point(158, 4)
        Me.BGSlash.Margin = New System.Windows.Forms.Padding(0, 0, 3, 0)
        Me.BGSlash.Name = "BGSlash"
        Me.BGSlash.Size = New System.Drawing.Size(22, 22)
        Me.BGSlash.TabIndex = 38
        Me.BGSlash.UseVisualStyleBackColor = True
        '
        'POGridSwitch
        '
        Me.POGridSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POGridSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POGridSwitch.Checked = True
        Me.POGridSwitch.CheckState = System.Windows.Forms.CheckState.Checked
        Me.POGridSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POGridSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POGridSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POGridSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POGridSwitch.Name = "POGridSwitch"
        Me.POGridSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POGridSwitch.TabIndex = 2
        Me.POGridSwitch.TabStop = False
        Me.POGridSwitch.Text = "Grid"
        Me.POGridSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POGridSwitch.UseCompatibleTextRendering = True
        Me.POGridSwitch.UseVisualStyleBackColor = False
        '
        'POHeader
        '
        Me.POHeader.AutoSize = True
        Me.POHeader.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POHeader.Controls.Add(Me.POHeaderInner)
        Me.POHeader.Controls.Add(Me.POHeaderSwitch)
        Me.POHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeader.Location = New System.Drawing.Point(0, 0)
        Me.POHeader.Name = "POHeader"
        Me.POHeader.Size = New System.Drawing.Size(183, 439)
        Me.POHeader.TabIndex = 1
        '
        'POHeaderInner
        '
        Me.POHeaderInner.AutoSize = True
        Me.POHeaderInner.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POHeaderInner.Controls.Add(Me.POHeaderPart2)
        Me.POHeaderInner.Controls.Add(Me.POHeaderExpander)
        Me.POHeaderInner.Controls.Add(Me.POHeaderPart1)
        Me.POHeaderInner.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeaderInner.Location = New System.Drawing.Point(0, 20)
        Me.POHeaderInner.Name = "POHeaderInner"
        Me.POHeaderInner.Size = New System.Drawing.Size(183, 419)
        Me.POHeaderInner.TabIndex = 2
        '
        'POHeaderPart2
        '
        Me.POHeaderPart2.AutoSize = True
        Me.POHeaderPart2.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POHeaderPart2.ColumnCount = 3
        Me.POHeaderPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.POHeaderPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POHeaderPart2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.POHeaderPart2.Controls.Add(Me.CHDifficulty, 1, 5)
        Me.POHeaderPart2.Controls.Add(Me.Label13, 2, 6)
        Me.POHeaderPart2.Controls.Add(Me.THExRank, 1, 6)
        Me.POHeaderPart2.Controls.Add(Me.Label25, 2, 7)
        Me.POHeaderPart2.Controls.Add(Me.CHLnObj, 1, 9)
        Me.POHeaderPart2.Controls.Add(Me.Label23, 0, 6)
        Me.POHeaderPart2.Controls.Add(Me.Label21, 0, 5)
        Me.POHeaderPart2.Controls.Add(Me.THComment, 1, 8)
        Me.POHeaderPart2.Controls.Add(Me.Label24, 0, 9)
        Me.POHeaderPart2.Controls.Add(Me.Label15, 0, 0)
        Me.POHeaderPart2.Controls.Add(Me.THTotal, 1, 7)
        Me.POHeaderPart2.Controls.Add(Me.Label20, 0, 7)
        Me.POHeaderPart2.Controls.Add(Me.BHStageFile, 2, 2)
        Me.POHeaderPart2.Controls.Add(Me.BHBanner, 2, 3)
        Me.POHeaderPart2.Controls.Add(Me.Label19, 0, 8)
        Me.POHeaderPart2.Controls.Add(Me.BHBackBMP, 2, 4)
        Me.POHeaderPart2.Controls.Add(Me.Label17, 0, 1)
        Me.POHeaderPart2.Controls.Add(Me.Label16, 0, 2)
        Me.POHeaderPart2.Controls.Add(Me.Label12, 0, 3)
        Me.POHeaderPart2.Controls.Add(Me.THBackBMP, 1, 4)
        Me.POHeaderPart2.Controls.Add(Me.Label11, 0, 4)
        Me.POHeaderPart2.Controls.Add(Me.THBanner, 1, 3)
        Me.POHeaderPart2.Controls.Add(Me.THStageFile, 1, 2)
        Me.POHeaderPart2.Controls.Add(Me.THSubTitle, 1, 0)
        Me.POHeaderPart2.Controls.Add(Me.THSubArtist, 1, 1)
        Me.POHeaderPart2.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeaderPart2.Location = New System.Drawing.Point(0, 186)
        Me.POHeaderPart2.Name = "POHeaderPart2"
        Me.POHeaderPart2.RowCount = 10
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart2.Size = New System.Drawing.Size(183, 233)
        Me.POHeaderPart2.TabIndex = 27
        '
        'CHDifficulty
        '
        Me.POHeaderPart2.SetColumnSpan(Me.CHDifficulty, 2)
        Me.CHDifficulty.Dock = System.Windows.Forms.DockStyle.Top
        Me.CHDifficulty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CHDifficulty.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CHDifficulty.Items.AddRange(New Object() {"None", "1 - EZ / Beginner", "2 - NM / Normal", "3 - HD / Hyper", "4 - MX / Another", "5 - SC / Insane"})
        Me.CHDifficulty.Location = New System.Drawing.Point(67, 118)
        Me.CHDifficulty.Margin = New System.Windows.Forms.Padding(0)
        Me.CHDifficulty.Name = "CHDifficulty"
        Me.CHDifficulty.Size = New System.Drawing.Size(116, 23)
        Me.CHDifficulty.TabIndex = 63
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label13.Location = New System.Drawing.Point(158, 141)
        Me.Label13.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(25, 23)
        Me.Label13.TabIndex = 63
        Me.Label13.Text = "(%)"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'THExRank
        '
        Me.THExRank.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THExRank.Location = New System.Drawing.Point(67, 141)
        Me.THExRank.Margin = New System.Windows.Forms.Padding(0)
        Me.THExRank.Name = "THExRank"
        Me.THExRank.Size = New System.Drawing.Size(88, 23)
        Me.THExRank.TabIndex = 27
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label25.Location = New System.Drawing.Point(158, 164)
        Me.Label25.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(25, 23)
        Me.Label25.TabIndex = 23
        Me.Label25.Text = "(%)"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CHLnObj
        '
        Me.POHeaderPart2.SetColumnSpan(Me.CHLnObj, 2)
        Me.CHLnObj.Dock = System.Windows.Forms.DockStyle.Top
        Me.CHLnObj.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CHLnObj.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CHLnObj.Items.AddRange(New Object() {"None (#LnType 1)", "01", "02", "03", "04", "05", "06", "07", "08", "09", "0A", "0B", "0C", "0D", "0E", "0F", "0G", "0H", "0I", "0J", "0K", "0L", "0M", "0N", "0O", "0P", "0Q", "0R", "0S", "0T", "0U", "0V", "0W", "0X", "0Y", "0Z", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "1A", "1B", "1C", "1D", "1E", "1F", "1G", "1H", "1I", "1J", "1K", "1L", "1M", "1N", "1O", "1P", "1Q", "1R", "1S", "1T", "1U", "1V", "1W", "1X", "1Y", "1Z", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "2A", "2B", "2C", "2D", "2E", "2F", "2G", "2H", "2I", "2J", "2K", "2L", "2M", "2N", "2O", "2P", "2Q", "2R", "2S", "2T", "2U", "2V", "2W", "2X", "2Y", "2Z", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "3A", "3B", "3C", "3D", "3E", "3F", "3G", "3H", "3I", "3J", "3K", "3L", "3M", "3N", "3O", "3P", "3Q", "3R", "3S", "3T", "3U", "3V", "3W", "3X", "3Y", "3Z", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "4A", "4B", "4C", "4D", "4E", "4F", "4G", "4H", "4I", "4J", "4K", "4L", "4M", "4N", "4O", "4P", "4Q", "4R", "4S", "4T", "4U", "4V", "4W", "4X", "4Y", "4Z", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "5A", "5B", "5C", "5D", "5E", "5F", "5G", "5H", "5I", "5J", "5K", "5L", "5M", "5N", "5O", "5P", "5Q", "5R", "5S", "5T", "5U", "5V", "5W", "5X", "5Y", "5Z", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "6A", "6B", "6C", "6D", "6E", "6F", "6G", "6H", "6I", "6J", "6K", "6L", "6M", "6N", "6O", "6P", "6Q", "6R", "6S", "6T", "6U", "6V", "6W", "6X", "6Y", "6Z", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "7A", "7B", "7C", "7D", "7E", "7F", "7G", "7H", "7I", "7J", "7K", "7L", "7M", "7N", "7O", "7P", "7Q", "7R", "7S", "7T", "7U", "7V", "7W", "7X", "7Y", "7Z", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "8A", "8B", "8C", "8D", "8E", "8F", "8G", "8H", "8I", "8J", "8K", "8L", "8M", "8N", "8O", "8P", "8Q", "8R", "8S", "8T", "8U", "8V", "8W", "8X", "8Y", "8Z", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "9A", "9B", "9C", "9D", "9E", "9F", "9G", "9H", "9I", "9J", "9K", "9L", "9M", "9N", "9O", "9P", "9Q", "9R", "9S", "9T", "9U", "9V", "9W", "9X", "9Y", "9Z", "A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "B0", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "C0", "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", "E0", "E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", "F0", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ", "G0", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ", "H0", "H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ", "I0", "I1", "I2", "I3", "I4", "I5", "I6", "I7", "I8", "I9", "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ", "J0", "J1", "J2", "J3", "J4", "J5", "J6", "J7", "J8", "J9", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ", "K0", "K1", "K2", "K3", "K4", "K5", "K6", "K7", "K8", "K9", "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ", "L0", "L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9", "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", "LY", "LZ", "M0", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "N0", "N1", "N2", "N3", "N4", "N5", "N6", "N7", "N8", "N9", "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", "NU", "NV", "NW", "NX", "NY", "NZ", "O0", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OI", "OJ", "OK", "OL", "OM", "ON", "OO", "OP", "OQ", "OR", "OS", "OT", "OU", "OV", "OW", "OX", "OY", "OZ", "P0", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ", "Q0", "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ", "R0", "R1", "R2", "R3", "R4", "R5", "R6", "R7", "R8", "R9", "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL", "RM", "RN", "RO", "RP", "RQ", "RR", "RS", "RT", "RU", "RV", "RW", "RX", "RY", "RZ", "S0", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9", "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SP", "SQ", "SR", "SS", "ST", "SU", "SV", "SW", "SX", "SY", "SZ", "T0", "T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", "TI", "TJ", "TK", "TL", "TM", "TN", "TO", "TP", "TQ", "TR", "TS", "TT", "TU", "TV", "TW", "TX", "TY", "TZ", "U0", "U1", "U2", "U3", "U4", "U5", "U6", "U7", "U8", "U9", "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UI", "UJ", "UK", "UL", "UM", "UN", "UO", "UP", "UQ", "UR", "US", "UT", "UU", "UV", "UW", "UX", "UY", "UZ", "V0", "V1", "V2", "V3", "V4", "V5", "V6", "V7", "V8", "V9", "VA", "VB", "VC", "VD", "VE", "VF", "VG", "VH", "VI", "VJ", "VK", "VL", "VM", "VN", "VO", "VP", "VQ", "VR", "VS", "VT", "VU", "VV", "VW", "VX", "VY", "VZ", "W0", "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8", "W9", "WA", "WB", "WC", "WD", "WE", "WF", "WG", "WH", "WI", "WJ", "WK", "WL", "WM", "WN", "WO", "WP", "WQ", "WR", "WS", "WT", "WU", "WV", "WW", "WX", "WY", "WZ", "X0", "X1", "X2", "X3", "X4", "X5", "X6", "X7", "X8", "X9", "XA", "XB", "XC", "XD", "XE", "XF", "XG", "XH", "XI", "XJ", "XK", "XL", "XM", "XN", "XO", "XP", "XQ", "XR", "XS", "XT", "XU", "XV", "XW", "XX", "XY", "XZ", "Y0", "Y1", "Y2", "Y3", "Y4", "Y5", "Y6", "Y7", "Y8", "Y9", "YA", "YB", "YC", "YD", "YE", "YF", "YG", "YH", "YI", "YJ", "YK", "YL", "YM", "YN", "YO", "YP", "YQ", "YR", "YS", "YT", "YU", "YV", "YW", "YX", "YY", "YZ", "Z0", "Z1", "Z2", "Z3", "Z4", "Z5", "Z6", "Z7", "Z8", "Z9", "ZA", "ZB", "ZC", "ZD", "ZE", "ZF", "ZG", "ZH", "ZI", "ZJ", "ZK", "ZL", "ZM", "ZN", "ZO", "ZP", "ZQ", "ZR", "ZS", "ZT", "ZU", "ZV", "ZW", "ZX", "ZY", "ZZ"})
        Me.CHLnObj.Location = New System.Drawing.Point(67, 210)
        Me.CHLnObj.Margin = New System.Windows.Forms.Padding(0)
        Me.CHLnObj.Name = "CHLnObj"
        Me.CHLnObj.Size = New System.Drawing.Size(116, 23)
        Me.CHLnObj.TabIndex = 28
        '
        'Label23
        '
        Me.Label23.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(20, 145)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(44, 15)
        Me.Label23.TabIndex = 26
        Me.Label23.Text = "ExRank"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label21
        '
        Me.Label21.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(9, 122)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(55, 15)
        Me.Label21.TabIndex = 25
        Me.Label21.Text = "Difficulty"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THComment
        '
        Me.POHeaderPart2.SetColumnSpan(Me.THComment, 2)
        Me.THComment.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THComment.Location = New System.Drawing.Point(67, 187)
        Me.THComment.Margin = New System.Windows.Forms.Padding(0)
        Me.THComment.Name = "THComment"
        Me.THComment.Size = New System.Drawing.Size(116, 23)
        Me.THComment.TabIndex = 19
        '
        'Label24
        '
        Me.Label24.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label24.AutoSize = True
        Me.Label24.Location = New System.Drawing.Point(25, 214)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(39, 15)
        Me.Label24.TabIndex = 27
        Me.Label24.Text = "LnObj"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label15
        '
        Me.Label15.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(14, 4)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(50, 15)
        Me.Label15.TabIndex = 6
        Me.Label15.Text = "SubTitle"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THTotal
        '
        Me.THTotal.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THTotal.Location = New System.Drawing.Point(67, 164)
        Me.THTotal.Margin = New System.Windows.Forms.Padding(0)
        Me.THTotal.Name = "THTotal"
        Me.THTotal.Size = New System.Drawing.Size(88, 23)
        Me.THTotal.TabIndex = 5
        '
        'Label20
        '
        Me.Label20.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(30, 168)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(34, 15)
        Me.Label20.TabIndex = 5
        Me.Label20.Text = "Total"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BHStageFile
        '
        Me.BHStageFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BHStageFile.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BHStageFile.Location = New System.Drawing.Point(158, 46)
        Me.BHStageFile.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.BHStageFile.Name = "BHStageFile"
        Me.BHStageFile.Size = New System.Drawing.Size(25, 24)
        Me.BHStageFile.TabIndex = 20
        Me.BHStageFile.Text = "..."
        Me.BHStageFile.UseVisualStyleBackColor = True
        '
        'BHBanner
        '
        Me.BHBanner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BHBanner.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BHBanner.Location = New System.Drawing.Point(158, 70)
        Me.BHBanner.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.BHBanner.Name = "BHBanner"
        Me.BHBanner.Size = New System.Drawing.Size(25, 24)
        Me.BHBanner.TabIndex = 21
        Me.BHBanner.Text = "..."
        Me.BHBanner.UseVisualStyleBackColor = True
        '
        'Label19
        '
        Me.Label19.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(3, 191)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(61, 15)
        Me.Label19.TabIndex = 13
        Me.Label19.Text = "Comment"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'BHBackBMP
        '
        Me.BHBackBMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BHBackBMP.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BHBackBMP.Location = New System.Drawing.Point(158, 94)
        Me.BHBackBMP.Margin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.BHBackBMP.Name = "BHBackBMP"
        Me.BHBackBMP.Size = New System.Drawing.Size(25, 24)
        Me.BHBackBMP.TabIndex = 22
        Me.BHBackBMP.Text = "..."
        Me.BHBackBMP.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(9, 27)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(55, 15)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "SubArtist"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label16
        '
        Me.Label16.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(7, 50)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(57, 15)
        Me.Label16.TabIndex = 9
        Me.Label16.Text = "Stage File"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(20, 74)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 15)
        Me.Label12.TabIndex = 13
        Me.Label12.Text = "Banner"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THBackBMP
        '
        Me.THBackBMP.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THBackBMP.Location = New System.Drawing.Point(67, 94)
        Me.THBackBMP.Margin = New System.Windows.Forms.Padding(0)
        Me.THBackBMP.Name = "THBackBMP"
        Me.THBackBMP.Size = New System.Drawing.Size(88, 23)
        Me.THBackBMP.TabIndex = 17
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(4, 98)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(60, 15)
        Me.Label11.TabIndex = 16
        Me.Label11.Text = "Back BMP"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THBanner
        '
        Me.THBanner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THBanner.Location = New System.Drawing.Point(67, 70)
        Me.THBanner.Margin = New System.Windows.Forms.Padding(0)
        Me.THBanner.Name = "THBanner"
        Me.THBanner.Size = New System.Drawing.Size(88, 23)
        Me.THBanner.TabIndex = 19
        '
        'THStageFile
        '
        Me.THStageFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THStageFile.Location = New System.Drawing.Point(67, 46)
        Me.THStageFile.Margin = New System.Windows.Forms.Padding(0)
        Me.THStageFile.Name = "THStageFile"
        Me.THStageFile.Size = New System.Drawing.Size(88, 23)
        Me.THStageFile.TabIndex = 18
        '
        'THSubTitle
        '
        Me.POHeaderPart2.SetColumnSpan(Me.THSubTitle, 2)
        Me.THSubTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THSubTitle.Location = New System.Drawing.Point(67, 0)
        Me.THSubTitle.Margin = New System.Windows.Forms.Padding(0)
        Me.THSubTitle.Name = "THSubTitle"
        Me.THSubTitle.Size = New System.Drawing.Size(116, 23)
        Me.THSubTitle.TabIndex = 6
        '
        'THSubArtist
        '
        Me.POHeaderPart2.SetColumnSpan(Me.THSubArtist, 2)
        Me.THSubArtist.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THSubArtist.Location = New System.Drawing.Point(67, 23)
        Me.THSubArtist.Margin = New System.Windows.Forms.Padding(0)
        Me.THSubArtist.Name = "THSubArtist"
        Me.THSubArtist.Size = New System.Drawing.Size(116, 23)
        Me.THSubArtist.TabIndex = 7
        '
        'POHeaderExpander
        '
        Me.POHeaderExpander.Appearance = System.Windows.Forms.Appearance.Button
        Me.POHeaderExpander.AutoSize = True
        Me.POHeaderExpander.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POHeaderExpander.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeaderExpander.FlatAppearance.BorderSize = 0
        Me.POHeaderExpander.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.POHeaderExpander.Location = New System.Drawing.Point(0, 161)
        Me.POHeaderExpander.Margin = New System.Windows.Forms.Padding(0)
        Me.POHeaderExpander.Name = "POHeaderExpander"
        Me.POHeaderExpander.Size = New System.Drawing.Size(183, 25)
        Me.POHeaderExpander.TabIndex = 26
        Me.POHeaderExpander.TabStop = False
        Me.POHeaderExpander.Text = "Expand..."
        Me.POHeaderExpander.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POHeaderExpander.UseVisualStyleBackColor = False
        '
        'POHeaderPart1
        '
        Me.POHeaderPart1.AutoSize = True
        Me.POHeaderPart1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.POHeaderPart1.ColumnCount = 2
        Me.POHeaderPart1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.POHeaderPart1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.POHeaderPart1.Controls.Add(Me.Label3, 0, 0)
        Me.POHeaderPart1.Controls.Add(Me.THPlayLevel, 1, 6)
        Me.POHeaderPart1.Controls.Add(Me.CHRank, 1, 5)
        Me.POHeaderPart1.Controls.Add(Me.Label10, 0, 5)
        Me.POHeaderPart1.Controls.Add(Me.CHPlayer, 1, 4)
        Me.POHeaderPart1.Controls.Add(Me.Label4, 0, 1)
        Me.POHeaderPart1.Controls.Add(Me.THGenre, 1, 2)
        Me.POHeaderPart1.Controls.Add(Me.THBPM, 1, 3)
        Me.POHeaderPart1.Controls.Add(Me.Label2, 0, 2)
        Me.POHeaderPart1.Controls.Add(Me.THArtist, 1, 1)
        Me.POHeaderPart1.Controls.Add(Me.THTitle, 1, 0)
        Me.POHeaderPart1.Controls.Add(Me.Label9, 0, 3)
        Me.POHeaderPart1.Controls.Add(Me.Label8, 0, 4)
        Me.POHeaderPart1.Controls.Add(Me.Label6, 0, 6)
        Me.POHeaderPart1.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeaderPart1.Location = New System.Drawing.Point(0, 0)
        Me.POHeaderPart1.Name = "POHeaderPart1"
        Me.POHeaderPart1.RowCount = 7
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.POHeaderPart1.Size = New System.Drawing.Size(183, 161)
        Me.POHeaderPart1.TabIndex = 25
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(32, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 15)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Title"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THPlayLevel
        '
        Me.THPlayLevel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THPlayLevel.Location = New System.Drawing.Point(65, 138)
        Me.THPlayLevel.Margin = New System.Windows.Forms.Padding(0)
        Me.THPlayLevel.Name = "THPlayLevel"
        Me.THPlayLevel.Size = New System.Drawing.Size(118, 23)
        Me.THPlayLevel.TabIndex = 8
        '
        'CHRank
        '
        Me.CHRank.Dock = System.Windows.Forms.DockStyle.Top
        Me.CHRank.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CHRank.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CHRank.Items.AddRange(New Object() {"0 - Very Hard", "1 - Hard", "2 - Normal", "3 - Easy", "4 - Very Easy"})
        Me.CHRank.Location = New System.Drawing.Point(65, 115)
        Me.CHRank.Margin = New System.Windows.Forms.Padding(0)
        Me.CHRank.Name = "CHRank"
        Me.CHRank.Size = New System.Drawing.Size(118, 23)
        Me.CHRank.TabIndex = 15
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(29, 119)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(33, 15)
        Me.Label10.TabIndex = 16
        Me.Label10.Text = "Rank"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CHPlayer
        '
        Me.CHPlayer.Dock = System.Windows.Forms.DockStyle.Top
        Me.CHPlayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CHPlayer.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CHPlayer.Items.AddRange(New Object() {"1 - Single Play", "2 - Couple Play", "3 - Double Play"})
        Me.CHPlayer.Location = New System.Drawing.Point(65, 92)
        Me.CHPlayer.Margin = New System.Windows.Forms.Padding(0)
        Me.CHPlayer.Name = "CHPlayer"
        Me.CHPlayer.Size = New System.Drawing.Size(118, 23)
        Me.CHPlayer.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(27, 27)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(35, 15)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Artist"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THGenre
        '
        Me.THGenre.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THGenre.Location = New System.Drawing.Point(65, 46)
        Me.THGenre.Margin = New System.Windows.Forms.Padding(0)
        Me.THGenre.Name = "THGenre"
        Me.THGenre.Size = New System.Drawing.Size(118, 23)
        Me.THGenre.TabIndex = 5
        '
        'THBPM
        '
        Me.THBPM.DecimalPlaces = 4
        Me.THBPM.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THBPM.Increment = New Decimal(New Integer() {10, 0, 0, 0})
        Me.THBPM.Location = New System.Drawing.Point(65, 69)
        Me.THBPM.Margin = New System.Windows.Forms.Padding(0)
        Me.THBPM.Maximum = New Decimal(New Integer() {655359999, 0, 0, 262144})
        Me.THBPM.Minimum = New Decimal(New Integer() {1, 0, 0, 262144})
        Me.THBPM.Name = "THBPM"
        Me.THBPM.Size = New System.Drawing.Size(118, 23)
        Me.THBPM.TabIndex = 10
        Me.THBPM.Value = New Decimal(New Integer() {120, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 50)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(38, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Genre"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'THArtist
        '
        Me.THArtist.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THArtist.Location = New System.Drawing.Point(65, 23)
        Me.THArtist.Margin = New System.Windows.Forms.Padding(0)
        Me.THArtist.Name = "THArtist"
        Me.THArtist.Size = New System.Drawing.Size(118, 23)
        Me.THArtist.TabIndex = 7
        '
        'THTitle
        '
        Me.THTitle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.THTitle.Location = New System.Drawing.Point(65, 0)
        Me.THTitle.Margin = New System.Windows.Forms.Padding(0)
        Me.THTitle.Name = "THTitle"
        Me.THTitle.Size = New System.Drawing.Size(118, 23)
        Me.THTitle.TabIndex = 6
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(30, 73)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 15)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "BPM"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(23, 96)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(39, 15)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Player"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 142)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 15)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Play Level"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'POHeaderSwitch
        '
        Me.POHeaderSwitch.Appearance = System.Windows.Forms.Appearance.Button
        Me.POHeaderSwitch.BackColor = System.Drawing.SystemColors.Control
        Me.POHeaderSwitch.Checked = True
        Me.POHeaderSwitch.CheckState = System.Windows.Forms.CheckState.Checked
        Me.POHeaderSwitch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.POHeaderSwitch.Dock = System.Windows.Forms.DockStyle.Top
        Me.POHeaderSwitch.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.POHeaderSwitch.Location = New System.Drawing.Point(0, 0)
        Me.POHeaderSwitch.Name = "POHeaderSwitch"
        Me.POHeaderSwitch.Size = New System.Drawing.Size(183, 20)
        Me.POHeaderSwitch.TabIndex = 1
        Me.POHeaderSwitch.TabStop = False
        Me.POHeaderSwitch.Text = "Header"
        Me.POHeaderSwitch.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.POHeaderSwitch.UseCompatibleTextRendering = True
        Me.POHeaderSwitch.UseVisualStyleBackColor = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 15
        '
        'Menu1
        '
        Me.Menu1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MInsert, Me.MRemove})
        Me.Menu1.Name = "Menu1"
        Me.Menu1.Size = New System.Drawing.Size(166, 48)
        '
        'MInsert
        '
        Me.MInsert.Image = Global.iBMSC.My.Resources.Resources.x16Add
        Me.MInsert.Name = "MInsert"
        Me.MInsert.Size = New System.Drawing.Size(165, 22)
        Me.MInsert.Text = "Insert Measure"
        '
        'MRemove
        '
        Me.MRemove.Image = Global.iBMSC.My.Resources.Resources.x16Remove
        Me.MRemove.Name = "MRemove"
        Me.MRemove.Size = New System.Drawing.Size(165, 22)
        Me.MRemove.Text = "Remove Measure"
        '
        'AutoSaveTimer
        '
        Me.AutoSaveTimer.Enabled = True
        Me.AutoSaveTimer.Interval = 300000
        '
        'mnMain
        '
        Me.mnMain.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuBar
        Me.mnMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.mnMain.Dock = System.Windows.Forms.DockStyle.None
        Me.mnMain.GripMargin = New System.Windows.Forms.Padding(2)
        Me.mnMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnFile, Me.mnEdit, Me.mnSys, Me.mnOptions, Me.mnConversion, Me.mnPreview})
        Me.mnMain.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.mnMain.Location = New System.Drawing.Point(0, 0)
        Me.mnMain.Name = "mnMain"
        Me.mnMain.Size = New System.Drawing.Size(877, 23)
        Me.mnMain.TabIndex = 57
        '
        'mnFile
        '
        Me.mnFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnNew, Me.mnOpen, Me.mnImportSM, Me.mnImportIBMSC, Me.ToolStripSeparator14, Me.mnSave, Me.mnSaveAs, Me.mnExport, Me.ToolStripSeparator15, Me.mnOpenR0, Me.mnOpenR1, Me.mnOpenR2, Me.mnOpenR3, Me.mnOpenR4, Me.ToolStripSeparator16, Me.mnQuit})
        Me.mnFile.Name = "mnFile"
        Me.mnFile.Size = New System.Drawing.Size(37, 19)
        Me.mnFile.Text = "&File"
        '
        'mnNew
        '
        Me.mnNew.Image = Global.iBMSC.My.Resources.Resources.x16New
        Me.mnNew.Name = "mnNew"
        Me.mnNew.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.N), System.Windows.Forms.Keys)
        Me.mnNew.Size = New System.Drawing.Size(199, 22)
        Me.mnNew.Text = "&New"
        '
        'mnOpen
        '
        Me.mnOpen.Image = Global.iBMSC.My.Resources.Resources.x16Open
        Me.mnOpen.Name = "mnOpen"
        Me.mnOpen.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.O), System.Windows.Forms.Keys)
        Me.mnOpen.Size = New System.Drawing.Size(199, 22)
        Me.mnOpen.Text = "&Open"
        '
        'mnImportSM
        '
        Me.mnImportSM.Image = Global.iBMSC.My.Resources.Resources.x16Import2
        Me.mnImportSM.Name = "mnImportSM"
        Me.mnImportSM.Size = New System.Drawing.Size(199, 22)
        Me.mnImportSM.Text = "Import from .S&M file"
        '
        'mnImportIBMSC
        '
        Me.mnImportIBMSC.Image = Global.iBMSC.My.Resources.Resources.x16Import2
        Me.mnImportIBMSC.Name = "mnImportIBMSC"
        Me.mnImportIBMSC.Size = New System.Drawing.Size(199, 22)
        Me.mnImportIBMSC.Text = "Import from .&IBMSC file"
        '
        'ToolStripSeparator14
        '
        Me.ToolStripSeparator14.Name = "ToolStripSeparator14"
        Me.ToolStripSeparator14.Size = New System.Drawing.Size(196, 6)
        '
        'mnSave
        '
        Me.mnSave.Image = Global.iBMSC.My.Resources.Resources.x16Save
        Me.mnSave.Name = "mnSave"
        Me.mnSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
        Me.mnSave.Size = New System.Drawing.Size(199, 22)
        Me.mnSave.Text = "&Save"
        '
        'mnSaveAs
        '
        Me.mnSaveAs.Image = Global.iBMSC.My.Resources.Resources.x16SaveAs
        Me.mnSaveAs.Name = "mnSaveAs"
        Me.mnSaveAs.Size = New System.Drawing.Size(199, 22)
        Me.mnSaveAs.Text = "Save &As..."
        '
        'mnExport
        '
        Me.mnExport.Image = Global.iBMSC.My.Resources.Resources.x16Export
        Me.mnExport.Name = "mnExport"
        Me.mnExport.Size = New System.Drawing.Size(199, 22)
        Me.mnExport.Text = "&Export .IBMSC file"
        '
        'ToolStripSeparator15
        '
        Me.ToolStripSeparator15.Name = "ToolStripSeparator15"
        Me.ToolStripSeparator15.Size = New System.Drawing.Size(196, 6)
        '
        'mnOpenR0
        '
        Me.mnOpenR0.Enabled = False
        Me.mnOpenR0.Name = "mnOpenR0"
        Me.mnOpenR0.Size = New System.Drawing.Size(199, 22)
        Me.mnOpenR0.Tag = "0"
        Me.mnOpenR0.Text = "Recent #0"
        '
        'mnOpenR1
        '
        Me.mnOpenR1.Enabled = False
        Me.mnOpenR1.Name = "mnOpenR1"
        Me.mnOpenR1.Size = New System.Drawing.Size(199, 22)
        Me.mnOpenR1.Tag = "1"
        Me.mnOpenR1.Text = "Recent #1"
        '
        'mnOpenR2
        '
        Me.mnOpenR2.Enabled = False
        Me.mnOpenR2.Name = "mnOpenR2"
        Me.mnOpenR2.Size = New System.Drawing.Size(199, 22)
        Me.mnOpenR2.Tag = "2"
        Me.mnOpenR2.Text = "Recent #2"
        '
        'mnOpenR3
        '
        Me.mnOpenR3.Enabled = False
        Me.mnOpenR3.Name = "mnOpenR3"
        Me.mnOpenR3.Size = New System.Drawing.Size(199, 22)
        Me.mnOpenR3.Tag = "3"
        Me.mnOpenR3.Text = "Recent #3"
        '
        'mnOpenR4
        '
        Me.mnOpenR4.Enabled = False
        Me.mnOpenR4.Name = "mnOpenR4"
        Me.mnOpenR4.Size = New System.Drawing.Size(199, 22)
        Me.mnOpenR4.Tag = "4"
        Me.mnOpenR4.Text = "Recent #4"
        '
        'ToolStripSeparator16
        '
        Me.ToolStripSeparator16.Name = "ToolStripSeparator16"
        Me.ToolStripSeparator16.Size = New System.Drawing.Size(196, 6)
        '
        'mnQuit
        '
        Me.mnQuit.Name = "mnQuit"
        Me.mnQuit.Size = New System.Drawing.Size(199, 22)
        Me.mnQuit.Text = "&Quit"
        '
        'mnEdit
        '
        Me.mnEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnUndo, Me.mnRedo, Me.ToolStripSeparator17, Me.mnCut, Me.mnCopy, Me.mnPaste, Me.mnDelete, Me.mnSelectAll, Me.mnGotoMeasure, Me.ToolStripSeparator18, Me.mnFind, Me.mnStatistics, Me.ToolStripSeparator19, Me.mnTimeSelect, Me.mnSelect, Me.mnWrite, Me.ToolStripSeparator23, Me.mnMyO2})
        Me.mnEdit.Name = "mnEdit"
        Me.mnEdit.Size = New System.Drawing.Size(39, 19)
        Me.mnEdit.Text = "&Edit"
        '
        'mnUndo
        '
        Me.mnUndo.Enabled = False
        Me.mnUndo.Image = Global.iBMSC.My.Resources.Resources.x16Undo
        Me.mnUndo.Name = "mnUndo"
        Me.mnUndo.ShortcutKeyDisplayString = "Ctrl+Z"
        Me.mnUndo.Size = New System.Drawing.Size(233, 22)
        Me.mnUndo.Text = "&Undo"
        '
        'mnRedo
        '
        Me.mnRedo.Enabled = False
        Me.mnRedo.Image = Global.iBMSC.My.Resources.Resources.x16Redo
        Me.mnRedo.Name = "mnRedo"
        Me.mnRedo.ShortcutKeyDisplayString = "Ctrl+Y"
        Me.mnRedo.Size = New System.Drawing.Size(233, 22)
        Me.mnRedo.Text = "&Redo"
        '
        'ToolStripSeparator17
        '
        Me.ToolStripSeparator17.Name = "ToolStripSeparator17"
        Me.ToolStripSeparator17.Size = New System.Drawing.Size(230, 6)
        '
        'mnCut
        '
        Me.mnCut.Image = Global.iBMSC.My.Resources.Resources.x16Cut
        Me.mnCut.Name = "mnCut"
        Me.mnCut.ShortcutKeyDisplayString = "Ctrl+X"
        Me.mnCut.Size = New System.Drawing.Size(233, 22)
        Me.mnCut.Text = "Cu&t"
        '
        'mnCopy
        '
        Me.mnCopy.Image = Global.iBMSC.My.Resources.Resources.x16Copy
        Me.mnCopy.Name = "mnCopy"
        Me.mnCopy.ShortcutKeyDisplayString = "Ctrl+C"
        Me.mnCopy.Size = New System.Drawing.Size(233, 22)
        Me.mnCopy.Text = "&Copy"
        '
        'mnPaste
        '
        Me.mnPaste.Image = Global.iBMSC.My.Resources.Resources.x16Paste
        Me.mnPaste.Name = "mnPaste"
        Me.mnPaste.ShortcutKeyDisplayString = "Ctrl+V"
        Me.mnPaste.Size = New System.Drawing.Size(233, 22)
        Me.mnPaste.Text = "&Paste"
        '
        'mnDelete
        '
        Me.mnDelete.Image = Global.iBMSC.My.Resources.Resources.x16Remove
        Me.mnDelete.Name = "mnDelete"
        Me.mnDelete.ShortcutKeyDisplayString = "Del"
        Me.mnDelete.Size = New System.Drawing.Size(233, 22)
        Me.mnDelete.Text = "De&lete"
        '
        'mnSelectAll
        '
        Me.mnSelectAll.Name = "mnSelectAll"
        Me.mnSelectAll.ShortcutKeyDisplayString = "Ctrl+A"
        Me.mnSelectAll.Size = New System.Drawing.Size(233, 22)
        Me.mnSelectAll.Text = "Select &All"
        '
        'mnGotoMeasure
        '
        Me.mnGotoMeasure.Name = "mnGotoMeasure"
        Me.mnGotoMeasure.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.G), System.Windows.Forms.Keys)
        Me.mnGotoMeasure.Size = New System.Drawing.Size(233, 22)
        Me.mnGotoMeasure.Text = "Go To Measure..."
        '
        'ToolStripSeparator18
        '
        Me.ToolStripSeparator18.Name = "ToolStripSeparator18"
        Me.ToolStripSeparator18.Size = New System.Drawing.Size(230, 6)
        '
        'mnFind
        '
        Me.mnFind.Image = Global.iBMSC.My.Resources.Resources.x16Find
        Me.mnFind.Name = "mnFind"
        Me.mnFind.ShortcutKeyDisplayString = "Ctrl+F"
        Me.mnFind.Size = New System.Drawing.Size(233, 22)
        Me.mnFind.Text = "&Find / Delete / Replace"
        '
        'mnStatistics
        '
        Me.mnStatistics.Image = Global.iBMSC.My.Resources.Resources.x16Statistics
        Me.mnStatistics.Name = "mnStatistics"
        Me.mnStatistics.ShortcutKeyDisplayString = "Ctrl+T"
        Me.mnStatistics.Size = New System.Drawing.Size(233, 22)
        Me.mnStatistics.Text = "St&atistics"
        '
        'ToolStripSeparator19
        '
        Me.ToolStripSeparator19.Name = "ToolStripSeparator19"
        Me.ToolStripSeparator19.Size = New System.Drawing.Size(230, 6)
        '
        'mnTimeSelect
        '
        Me.mnTimeSelect.CheckOnClick = True
        Me.mnTimeSelect.Image = Global.iBMSC.My.Resources.Resources.x16TimeSelection
        Me.mnTimeSelect.Name = "mnTimeSelect"
        Me.mnTimeSelect.ShortcutKeys = System.Windows.Forms.Keys.F1
        Me.mnTimeSelect.Size = New System.Drawing.Size(233, 22)
        Me.mnTimeSelect.Text = "T&ime Selection Tool"
        '
        'mnSelect
        '
        Me.mnSelect.Checked = True
        Me.mnSelect.CheckOnClick = True
        Me.mnSelect.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnSelect.Image = Global.iBMSC.My.Resources.Resources.x16Select
        Me.mnSelect.Name = "mnSelect"
        Me.mnSelect.ShortcutKeys = System.Windows.Forms.Keys.F2
        Me.mnSelect.Size = New System.Drawing.Size(233, 22)
        Me.mnSelect.Text = "&Select Tool"
        '
        'mnWrite
        '
        Me.mnWrite.CheckOnClick = True
        Me.mnWrite.Image = Global.iBMSC.My.Resources.Resources.x16Pen
        Me.mnWrite.Name = "mnWrite"
        Me.mnWrite.ShortcutKeys = System.Windows.Forms.Keys.F3
        Me.mnWrite.Size = New System.Drawing.Size(233, 22)
        Me.mnWrite.Text = "&Write Tool"
        '
        'ToolStripSeparator23
        '
        Me.ToolStripSeparator23.Name = "ToolStripSeparator23"
        Me.ToolStripSeparator23.Size = New System.Drawing.Size(230, 6)
        '
        'mnMyO2
        '
        Me.mnMyO2.Image = Global.iBMSC.My.Resources.Resources.x16MyO2
        Me.mnMyO2.Name = "mnMyO2"
        Me.mnMyO2.Size = New System.Drawing.Size(233, 22)
        Me.mnMyO2.Text = "MyO2 ToolBox (Chinese Only)"
        '
        'mnSys
        '
        Me.mnSys.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnSMenu, Me.mnSTB, Me.mnSOP, Me.mnSStatus, Me.mnSLSplitter, Me.mnSRSplitter, Me.ToolStripSeparator21, Me.CGShow, Me.CGShowS, Me.CGShowBG, Me.CGShowM, Me.CGShowMB, Me.CGShowV, Me.CGShowC, Me.ToolStripSeparator22, Me.CGBPM, Me.CGSTOP, Me.CGBLP})
        Me.mnSys.Name = "mnSys"
        Me.mnSys.Size = New System.Drawing.Size(44, 19)
        Me.mnSys.Text = "&View"
        '
        'mnSMenu
        '
        Me.mnSMenu.Checked = True
        Me.mnSMenu.CheckOnClick = True
        Me.mnSMenu.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnSMenu.Name = "mnSMenu"
        Me.mnSMenu.Size = New System.Drawing.Size(172, 22)
        Me.mnSMenu.Text = "&Main Menu"
        Me.mnSMenu.Visible = False
        '
        'mnSTB
        '
        Me.mnSTB.Checked = True
        Me.mnSTB.CheckOnClick = True
        Me.mnSTB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnSTB.Name = "mnSTB"
        Me.mnSTB.Size = New System.Drawing.Size(172, 22)
        Me.mnSTB.Text = "&ToolBar"
        '
        'mnSOP
        '
        Me.mnSOP.Checked = True
        Me.mnSOP.CheckOnClick = True
        Me.mnSOP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnSOP.Name = "mnSOP"
        Me.mnSOP.Size = New System.Drawing.Size(172, 22)
        Me.mnSOP.Text = "&Options Panel"
        '
        'mnSStatus
        '
        Me.mnSStatus.Checked = True
        Me.mnSStatus.CheckOnClick = True
        Me.mnSStatus.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnSStatus.Name = "mnSStatus"
        Me.mnSStatus.Size = New System.Drawing.Size(172, 22)
        Me.mnSStatus.Text = "&Status Bar"
        '
        'mnSLSplitter
        '
        Me.mnSLSplitter.CheckOnClick = True
        Me.mnSLSplitter.Name = "mnSLSplitter"
        Me.mnSLSplitter.Size = New System.Drawing.Size(172, 22)
        Me.mnSLSplitter.Text = "&Left Splitter"
        '
        'mnSRSplitter
        '
        Me.mnSRSplitter.CheckOnClick = True
        Me.mnSRSplitter.Name = "mnSRSplitter"
        Me.mnSRSplitter.Size = New System.Drawing.Size(172, 22)
        Me.mnSRSplitter.Text = "&Right Splitter"
        '
        'ToolStripSeparator21
        '
        Me.ToolStripSeparator21.Name = "ToolStripSeparator21"
        Me.ToolStripSeparator21.Size = New System.Drawing.Size(169, 6)
        '
        'CGShow
        '
        Me.CGShow.Checked = True
        Me.CGShow.CheckOnClick = True
        Me.CGShow.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShow.Name = "CGShow"
        Me.CGShow.Size = New System.Drawing.Size(172, 22)
        Me.CGShow.Text = "Grid"
        '
        'CGShowS
        '
        Me.CGShowS.Checked = True
        Me.CGShowS.CheckOnClick = True
        Me.CGShowS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowS.Name = "CGShowS"
        Me.CGShowS.Size = New System.Drawing.Size(172, 22)
        Me.CGShowS.Text = "Sub"
        '
        'CGShowBG
        '
        Me.CGShowBG.Checked = True
        Me.CGShowBG.CheckOnClick = True
        Me.CGShowBG.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowBG.Name = "CGShowBG"
        Me.CGShowBG.Size = New System.Drawing.Size(172, 22)
        Me.CGShowBG.Text = "BackGround"
        '
        'CGShowM
        '
        Me.CGShowM.Checked = True
        Me.CGShowM.CheckOnClick = True
        Me.CGShowM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowM.Name = "CGShowM"
        Me.CGShowM.Size = New System.Drawing.Size(172, 22)
        Me.CGShowM.Text = "Measure Index"
        '
        'CGShowMB
        '
        Me.CGShowMB.Checked = True
        Me.CGShowMB.CheckOnClick = True
        Me.CGShowMB.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowMB.Name = "CGShowMB"
        Me.CGShowMB.Size = New System.Drawing.Size(172, 22)
        Me.CGShowMB.Text = "Measure Line"
        '
        'CGShowV
        '
        Me.CGShowV.Checked = True
        Me.CGShowV.CheckOnClick = True
        Me.CGShowV.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowV.Name = "CGShowV"
        Me.CGShowV.Size = New System.Drawing.Size(172, 22)
        Me.CGShowV.Text = "Vertical Line"
        '
        'CGShowC
        '
        Me.CGShowC.Checked = True
        Me.CGShowC.CheckOnClick = True
        Me.CGShowC.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGShowC.Name = "CGShowC"
        Me.CGShowC.Size = New System.Drawing.Size(172, 22)
        Me.CGShowC.Text = "Column Caption"
        '
        'ToolStripSeparator22
        '
        Me.ToolStripSeparator22.Name = "ToolStripSeparator22"
        Me.ToolStripSeparator22.Size = New System.Drawing.Size(169, 6)
        '
        'CGBPM
        '
        Me.CGBPM.Checked = True
        Me.CGBPM.CheckOnClick = True
        Me.CGBPM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGBPM.Name = "CGBPM"
        Me.CGBPM.Size = New System.Drawing.Size(172, 22)
        Me.CGBPM.Text = "BPM"
        '
        'CGSTOP
        '
        Me.CGSTOP.Checked = True
        Me.CGSTOP.CheckOnClick = True
        Me.CGSTOP.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CGSTOP.Name = "CGSTOP"
        Me.CGSTOP.Size = New System.Drawing.Size(172, 22)
        Me.CGSTOP.Text = "STOP"
        '
        'CGBLP
        '
        Me.CGBLP.CheckOnClick = True
        Me.CGBLP.Name = "CGBLP"
        Me.CGBLP.Size = New System.Drawing.Size(172, 22)
        Me.CGBLP.Text = "BGA / Layer / Poor"
        '
        'mnOptions
        '
        Me.mnOptions.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnNTInput, Me.mnErrorCheck, Me.mnPreviewOnClick, Me.mnShowFileName, Me.ToolStripSeparator20, Me.mnGOptions, Me.mnVOptions, Me.mnPOptions, Me.mnLanguage, Me.mnTheme})
        Me.mnOptions.Name = "mnOptions"
        Me.mnOptions.Size = New System.Drawing.Size(61, 19)
        Me.mnOptions.Text = "&Options"
        '
        'mnNTInput
        '
        Me.mnNTInput.Checked = True
        Me.mnNTInput.CheckOnClick = True
        Me.mnNTInput.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnNTInput.Image = Global.iBMSC.My.Resources.Resources.x16NTInput
        Me.mnNTInput.Name = "mnNTInput"
        Me.mnNTInput.ShortcutKeys = System.Windows.Forms.Keys.F8
        Me.mnNTInput.Size = New System.Drawing.Size(229, 22)
        Me.mnNTInput.Text = "L&N Input Style - NT/BMSE"
        '
        'mnErrorCheck
        '
        Me.mnErrorCheck.Checked = True
        Me.mnErrorCheck.CheckOnClick = True
        Me.mnErrorCheck.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnErrorCheck.Image = Global.iBMSC.My.Resources.Resources.x16CheckError
        Me.mnErrorCheck.Name = "mnErrorCheck"
        Me.mnErrorCheck.Size = New System.Drawing.Size(229, 22)
        Me.mnErrorCheck.Text = "&Error Check"
        '
        'mnPreviewOnClick
        '
        Me.mnPreviewOnClick.Checked = True
        Me.mnPreviewOnClick.CheckOnClick = True
        Me.mnPreviewOnClick.CheckState = System.Windows.Forms.CheckState.Checked
        Me.mnPreviewOnClick.Image = Global.iBMSC.My.Resources.Resources.x16PreviewOnClick
        Me.mnPreviewOnClick.Name = "mnPreviewOnClick"
        Me.mnPreviewOnClick.Size = New System.Drawing.Size(229, 22)
        Me.mnPreviewOnClick.Text = "Preview on &Click"
        '
        'mnShowFileName
        '
        Me.mnShowFileName.CheckOnClick = True
        Me.mnShowFileName.Image = Global.iBMSC.My.Resources.Resources.x16ShowFileNameN
        Me.mnShowFileName.Name = "mnShowFileName"
        Me.mnShowFileName.Size = New System.Drawing.Size(229, 22)
        Me.mnShowFileName.Text = "Show &File Name on Notes"
        '
        'ToolStripSeparator20
        '
        Me.ToolStripSeparator20.Name = "ToolStripSeparator20"
        Me.ToolStripSeparator20.Size = New System.Drawing.Size(226, 6)
        '
        'mnGOptions
        '
        Me.mnGOptions.Image = Global.iBMSC.My.Resources.Resources.x16GeneralOptions
        Me.mnGOptions.Name = "mnGOptions"
        Me.mnGOptions.Size = New System.Drawing.Size(229, 22)
        Me.mnGOptions.Text = "&General Options"
        '
        'mnVOptions
        '
        Me.mnVOptions.Image = Global.iBMSC.My.Resources.Resources.x16VisualOptions
        Me.mnVOptions.Name = "mnVOptions"
        Me.mnVOptions.Size = New System.Drawing.Size(229, 22)
        Me.mnVOptions.Text = "&Visual Options"
        '
        'mnPOptions
        '
        Me.mnPOptions.Image = Global.iBMSC.My.Resources.Resources.x16PlayerOptions
        Me.mnPOptions.Name = "mnPOptions"
        Me.mnPOptions.Size = New System.Drawing.Size(229, 22)
        Me.mnPOptions.Text = "&Player Options"
        '
        'mnConversion
        '
        Me.mnConversion.DropDown = Me.cmnConversion
        Me.mnConversion.Name = "mnConversion"
        Me.mnConversion.Size = New System.Drawing.Size(84, 19)
        Me.mnConversion.Text = "&Conversions"
        '
        'cmnConversion
        '
        Me.cmnConversion.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.POBLong, Me.POBShort, Me.POBLongShort, Me.ToolStripSeparator10, Me.POBHidden, Me.POBVisible, Me.POBHiddenVisible, Me.ToolStripSeparator11, Me.POBModify, Me.POBMirror})
        Me.cmnConversion.Name = "cmnLanguage"
        Me.cmnConversion.OwnerItem = Me.mnConversion
        Me.cmnConversion.Size = New System.Drawing.Size(222, 192)
        '
        'POBLong
        '
        Me.POBLong.Enabled = False
        Me.POBLong.Image = Global.iBMSC.My.Resources.Resources.ConvertNotesL
        Me.POBLong.Name = "POBLong"
        Me.POBLong.Size = New System.Drawing.Size(221, 22)
        Me.POBLong.Text = "→ &Long Note"
        '
        'POBShort
        '
        Me.POBShort.Image = Global.iBMSC.My.Resources.Resources.ConvertNotesN
        Me.POBShort.Name = "POBShort"
        Me.POBShort.Size = New System.Drawing.Size(221, 22)
        Me.POBShort.Text = "→ &Short Note"
        '
        'POBLongShort
        '
        Me.POBLongShort.Enabled = False
        Me.POBLongShort.Image = Global.iBMSC.My.Resources.Resources.ConvertNotes
        Me.POBLongShort.Name = "POBLongShort"
        Me.POBLongShort.Size = New System.Drawing.Size(221, 22)
        Me.POBLongShort.Text = "Long Note ↔ Short Note"
        '
        'ToolStripSeparator10
        '
        Me.ToolStripSeparator10.Name = "ToolStripSeparator10"
        Me.ToolStripSeparator10.Size = New System.Drawing.Size(218, 6)
        '
        'POBHidden
        '
        Me.POBHidden.Image = Global.iBMSC.My.Resources.Resources.ConvertNotesH
        Me.POBHidden.Name = "POBHidden"
        Me.POBHidden.Size = New System.Drawing.Size(221, 22)
        Me.POBHidden.Text = "→ &Hidden Note"
        '
        'POBVisible
        '
        Me.POBVisible.Image = Global.iBMSC.My.Resources.Resources.ConvertNotesV
        Me.POBVisible.Name = "POBVisible"
        Me.POBVisible.Size = New System.Drawing.Size(221, 22)
        Me.POBVisible.Text = "→ &Visible Note"
        '
        'POBHiddenVisible
        '
        Me.POBHiddenVisible.Image = Global.iBMSC.My.Resources.Resources.ConvertNotesHV
        Me.POBHiddenVisible.Name = "POBHiddenVisible"
        Me.POBHiddenVisible.Size = New System.Drawing.Size(221, 22)
        Me.POBHiddenVisible.Text = "Hidden Note ↔ Visible Note"
        '
        'ToolStripSeparator11
        '
        Me.ToolStripSeparator11.Name = "ToolStripSeparator11"
        Me.ToolStripSeparator11.Size = New System.Drawing.Size(218, 6)
        '
        'POBModify
        '
        Me.POBModify.Image = Global.iBMSC.My.Resources.Resources.x16ModifyLabel
        Me.POBModify.Name = "POBModify"
        Me.POBModify.Size = New System.Drawing.Size(221, 22)
        Me.POBModify.Text = "&Modify Labels"
        '
        'POBMirror
        '
        Me.POBMirror.Image = Global.iBMSC.My.Resources.Resources.x16Mirror
        Me.POBMirror.Name = "POBMirror"
        Me.POBMirror.Size = New System.Drawing.Size(221, 22)
        Me.POBMirror.Text = "Mi&rror"
        '
        'POConvert
        '
        Me.POConvert.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.POConvert.DropDown = Me.cmnConversion
        Me.POConvert.Image = Global.iBMSC.My.Resources.Resources.ConvertNotes
        Me.POConvert.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.POConvert.Name = "POConvert"
        Me.POConvert.Size = New System.Drawing.Size(29, 22)
        Me.POConvert.Text = "Convert Notes"
        '
        'mnPreview
        '
        Me.mnPreview.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnPlayB, Me.mnPlay, Me.mnStop})
        Me.mnPreview.Name = "mnPreview"
        Me.mnPreview.Size = New System.Drawing.Size(60, 19)
        Me.mnPreview.Text = "&Preview"
        '
        'mnPlayB
        '
        Me.mnPlayB.Image = Global.iBMSC.My.Resources.Resources.x16PlayB
        Me.mnPlayB.Name = "mnPlayB"
        Me.mnPlayB.ShortcutKeys = System.Windows.Forms.Keys.F5
        Me.mnPlayB.Size = New System.Drawing.Size(201, 22)
        Me.mnPlayB.Text = "Play from &beginning"
        '
        'mnPlay
        '
        Me.mnPlay.Image = Global.iBMSC.My.Resources.Resources.x16Play
        Me.mnPlay.Name = "mnPlay"
        Me.mnPlay.ShortcutKeys = System.Windows.Forms.Keys.F6
        Me.mnPlay.Size = New System.Drawing.Size(201, 22)
        Me.mnPlay.Text = "&Play from here"
        '
        'mnStop
        '
        Me.mnStop.Image = Global.iBMSC.My.Resources.Resources.x16Stop
        Me.mnStop.Name = "mnStop"
        Me.mnStop.ShortcutKeys = System.Windows.Forms.Keys.F7
        Me.mnStop.Size = New System.Drawing.Size(201, 22)
        Me.mnStop.Text = "&Stop"
        '
        'TBMain
        '
        Me.TBMain.AccessibleRole = System.Windows.Forms.AccessibleRole.ToolBar
        Me.TBMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.TBMain.Dock = System.Windows.Forms.DockStyle.None
        Me.TBMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TBNew, Me.TBOpen, Me.TBSave, Me.ToolStripSeparator1, Me.TBCut, Me.TBCopy, Me.TBPaste, Me.TBFind, Me.TBStatistics, Me.POConvert, Me.TBMyO2, Me.ToolStripSeparator4, Me.TBErrorCheck, Me.TBPreviewOnClick, Me.TBShowFileName, Me.ToolStripSeparator2, Me.TBUndo, Me.TBRedo, Me.ToolStripSeparator5, Me.TBNTInput, Me.TBTimeSelect, Me.TBSelect, Me.TBWrite, Me.ToolStripSeparator3, Me.TBPlayB, Me.TBPlay, Me.TBStop, Me.TBPOptions, Me.ToolStripSeparator7, Me.TBVOptions, Me.TBGOptions, Me.TBLanguage, Me.TBTheme, Me.TBAbout, Me.ToolStripSeparator13, Me.POBStorm})
        Me.TBMain.Location = New System.Drawing.Point(3, 23)
        Me.TBMain.Name = "TBMain"
        Me.TBMain.Size = New System.Drawing.Size(739, 25)
        Me.TBMain.TabIndex = 64
        Me.TBMain.Text = "Main Toolbar"
        '
        'TBNew
        '
        Me.TBNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBNew.Image = Global.iBMSC.My.Resources.Resources.x16New
        Me.TBNew.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.TBNew.Name = "TBNew"
        Me.TBNew.Size = New System.Drawing.Size(23, 22)
        Me.TBNew.Text = "New (Ctrl+N)"
        '
        'TBOpen
        '
        Me.TBOpen.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBOpen.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TBOpenR0, Me.TBOpenR1, Me.TBOpenR2, Me.TBOpenR3, Me.TBOpenR4, Me.ToolStripSeparator12, Me.TBImportSM, Me.TBImportIBMSC})
        Me.TBOpen.Image = Global.iBMSC.My.Resources.Resources.x16Open
        Me.TBOpen.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBOpen.Name = "TBOpen"
        Me.TBOpen.Size = New System.Drawing.Size(32, 22)
        Me.TBOpen.Text = "Open (Ctrl+O)"
        '
        'TBOpenR0
        '
        Me.TBOpenR0.Enabled = False
        Me.TBOpenR0.Name = "TBOpenR0"
        Me.TBOpenR0.Size = New System.Drawing.Size(199, 22)
        Me.TBOpenR0.Tag = "0"
        Me.TBOpenR0.Text = "Recent #0"
        '
        'TBOpenR1
        '
        Me.TBOpenR1.Enabled = False
        Me.TBOpenR1.Name = "TBOpenR1"
        Me.TBOpenR1.Size = New System.Drawing.Size(199, 22)
        Me.TBOpenR1.Tag = "1"
        Me.TBOpenR1.Text = "Recent #1"
        '
        'TBOpenR2
        '
        Me.TBOpenR2.Enabled = False
        Me.TBOpenR2.Name = "TBOpenR2"
        Me.TBOpenR2.Size = New System.Drawing.Size(199, 22)
        Me.TBOpenR2.Tag = "2"
        Me.TBOpenR2.Text = "Recent #2"
        '
        'TBOpenR3
        '
        Me.TBOpenR3.Enabled = False
        Me.TBOpenR3.Name = "TBOpenR3"
        Me.TBOpenR3.Size = New System.Drawing.Size(199, 22)
        Me.TBOpenR3.Tag = "3"
        Me.TBOpenR3.Text = "Recent #3"
        '
        'TBOpenR4
        '
        Me.TBOpenR4.Enabled = False
        Me.TBOpenR4.Name = "TBOpenR4"
        Me.TBOpenR4.Size = New System.Drawing.Size(199, 22)
        Me.TBOpenR4.Tag = "4"
        Me.TBOpenR4.Text = "Recent #4"
        '
        'ToolStripSeparator12
        '
        Me.ToolStripSeparator12.Name = "ToolStripSeparator12"
        Me.ToolStripSeparator12.Size = New System.Drawing.Size(196, 6)
        '
        'TBImportSM
        '
        Me.TBImportSM.Image = Global.iBMSC.My.Resources.Resources.x16Import2
        Me.TBImportSM.Name = "TBImportSM"
        Me.TBImportSM.Size = New System.Drawing.Size(199, 22)
        Me.TBImportSM.Text = "Import from .SM file"
        '
        'TBImportIBMSC
        '
        Me.TBImportIBMSC.Image = Global.iBMSC.My.Resources.Resources.x16Import2
        Me.TBImportIBMSC.Name = "TBImportIBMSC"
        Me.TBImportIBMSC.Size = New System.Drawing.Size(199, 22)
        Me.TBImportIBMSC.Text = "Import from .IBMSC file"
        '
        'TBSave
        '
        Me.TBSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBSave.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TBSaveAs, Me.TBExport})
        Me.TBSave.Image = Global.iBMSC.My.Resources.Resources.x16Save
        Me.TBSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBSave.Name = "TBSave"
        Me.TBSave.Size = New System.Drawing.Size(32, 22)
        Me.TBSave.Text = "Save (Ctrl+S)"
        '
        'TBSaveAs
        '
        Me.TBSaveAs.Image = Global.iBMSC.My.Resources.Resources.x16SaveAs
        Me.TBSaveAs.Name = "TBSaveAs"
        Me.TBSaveAs.Size = New System.Drawing.Size(167, 22)
        Me.TBSaveAs.Text = "Save As..."
        '
        'TBExport
        '
        Me.TBExport.Image = Global.iBMSC.My.Resources.Resources.x16Export
        Me.TBExport.Name = "TBExport"
        Me.TBExport.Size = New System.Drawing.Size(167, 22)
        Me.TBExport.Text = "Export .IBMSC file"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'TBCut
        '
        Me.TBCut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBCut.Image = Global.iBMSC.My.Resources.Resources.x16Cut
        Me.TBCut.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBCut.Name = "TBCut"
        Me.TBCut.Size = New System.Drawing.Size(23, 22)
        Me.TBCut.Text = "Cut (Ctrl+X)"
        '
        'TBCopy
        '
        Me.TBCopy.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBCopy.Image = Global.iBMSC.My.Resources.Resources.x16Copy
        Me.TBCopy.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBCopy.Name = "TBCopy"
        Me.TBCopy.Size = New System.Drawing.Size(23, 22)
        Me.TBCopy.Text = "Copy (Ctrl+C)"
        '
        'TBPaste
        '
        Me.TBPaste.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBPaste.Image = Global.iBMSC.My.Resources.Resources.x16Paste
        Me.TBPaste.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBPaste.Name = "TBPaste"
        Me.TBPaste.Size = New System.Drawing.Size(23, 22)
        Me.TBPaste.Text = "Paste (Ctrl+V)"
        '
        'TBFind
        '
        Me.TBFind.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBFind.Image = Global.iBMSC.My.Resources.Resources.x16Find
        Me.TBFind.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBFind.Name = "TBFind"
        Me.TBFind.Size = New System.Drawing.Size(23, 22)
        Me.TBFind.Text = "Find / Delete / Replace (Ctrl+F)"
        '
        'TBStatistics
        '
        Me.TBStatistics.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TBStatistics.Image = Global.iBMSC.My.Resources.Resources.x16Statistics
        Me.TBStatistics.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBStatistics.Name = "TBStatistics"
        Me.TBStatistics.Size = New System.Drawing.Size(34, 22)
        Me.TBStatistics.Text = "0"
        Me.TBStatistics.ToolTipText = "Statistics (Ctrl+T)"
        '
        'TBMyO2
        '
        Me.TBMyO2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBMyO2.Image = Global.iBMSC.My.Resources.Resources.x16MyO2
        Me.TBMyO2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBMyO2.Name = "TBMyO2"
        Me.TBMyO2.Size = New System.Drawing.Size(23, 22)
        Me.TBMyO2.Text = "MyO2 ToolBox (Chinese Only)"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'TBErrorCheck
        '
        Me.TBErrorCheck.Checked = True
        Me.TBErrorCheck.CheckOnClick = True
        Me.TBErrorCheck.CheckState = System.Windows.Forms.CheckState.Checked
        Me.TBErrorCheck.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBErrorCheck.Image = Global.iBMSC.My.Resources.Resources.x16CheckError
        Me.TBErrorCheck.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBErrorCheck.Name = "TBErrorCheck"
        Me.TBErrorCheck.Size = New System.Drawing.Size(23, 22)
        Me.TBErrorCheck.Text = "Error Check"
        '
        'TBPreviewOnClick
        '
        Me.TBPreviewOnClick.Checked = True
        Me.TBPreviewOnClick.CheckOnClick = True
        Me.TBPreviewOnClick.CheckState = System.Windows.Forms.CheckState.Checked
        Me.TBPreviewOnClick.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBPreviewOnClick.Image = Global.iBMSC.My.Resources.Resources.x16PreviewOnClick
        Me.TBPreviewOnClick.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBPreviewOnClick.Name = "TBPreviewOnClick"
        Me.TBPreviewOnClick.Size = New System.Drawing.Size(23, 22)
        Me.TBPreviewOnClick.Text = "Preview On Click"
        '
        'TBShowFileName
        '
        Me.TBShowFileName.CheckOnClick = True
        Me.TBShowFileName.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBShowFileName.Image = Global.iBMSC.My.Resources.Resources.x16ShowFileNameN
        Me.TBShowFileName.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBShowFileName.Name = "TBShowFileName"
        Me.TBShowFileName.Size = New System.Drawing.Size(23, 22)
        Me.TBShowFileName.Text = "Show File Name on Notes"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'TBUndo
        '
        Me.TBUndo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBUndo.Enabled = False
        Me.TBUndo.Image = Global.iBMSC.My.Resources.Resources.x16Undo
        Me.TBUndo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBUndo.Name = "TBUndo"
        Me.TBUndo.Size = New System.Drawing.Size(23, 22)
        Me.TBUndo.Text = "Undo (Ctrl+Z)"
        '
        'TBRedo
        '
        Me.TBRedo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBRedo.Enabled = False
        Me.TBRedo.Image = Global.iBMSC.My.Resources.Resources.x16Redo
        Me.TBRedo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBRedo.Name = "TBRedo"
        Me.TBRedo.Size = New System.Drawing.Size(23, 22)
        Me.TBRedo.Text = "Redo (Ctrl+Y)"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 25)
        '
        'TBNTInput
        '
        Me.TBNTInput.Checked = True
        Me.TBNTInput.CheckOnClick = True
        Me.TBNTInput.CheckState = System.Windows.Forms.CheckState.Checked
        Me.TBNTInput.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBNTInput.Image = Global.iBMSC.My.Resources.Resources.x16NTInput
        Me.TBNTInput.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBNTInput.Name = "TBNTInput"
        Me.TBNTInput.Size = New System.Drawing.Size(23, 22)
        Me.TBNTInput.Text = "LongNote Input Style - NoteTool/BMSE"
        '
        'TBTimeSelect
        '
        Me.TBTimeSelect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBTimeSelect.Image = Global.iBMSC.My.Resources.Resources.x16TimeSelection
        Me.TBTimeSelect.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBTimeSelect.Name = "TBTimeSelect"
        Me.TBTimeSelect.Size = New System.Drawing.Size(23, 22)
        Me.TBTimeSelect.Text = "Time Selection Tool (F1)"
        '
        'TBSelect
        '
        Me.TBSelect.Checked = True
        Me.TBSelect.CheckState = System.Windows.Forms.CheckState.Checked
        Me.TBSelect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBSelect.Image = Global.iBMSC.My.Resources.Resources.x16Select
        Me.TBSelect.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBSelect.Name = "TBSelect"
        Me.TBSelect.Size = New System.Drawing.Size(23, 22)
        Me.TBSelect.Text = "Select Tool (F2)"
        '
        'TBWrite
        '
        Me.TBWrite.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBWrite.Image = Global.iBMSC.My.Resources.Resources.x16Pen
        Me.TBWrite.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBWrite.Name = "TBWrite"
        Me.TBWrite.Size = New System.Drawing.Size(23, 22)
        Me.TBWrite.Text = "Write Tool (F3)"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'TBPlayB
        '
        Me.TBPlayB.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBPlayB.Image = Global.iBMSC.My.Resources.Resources.x16PlayB
        Me.TBPlayB.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBPlayB.Name = "TBPlayB"
        Me.TBPlayB.Size = New System.Drawing.Size(23, 22)
        Me.TBPlayB.Text = "Play from beginning (F5)"
        '
        'TBPlay
        '
        Me.TBPlay.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBPlay.Image = Global.iBMSC.My.Resources.Resources.x16Play
        Me.TBPlay.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBPlay.Name = "TBPlay"
        Me.TBPlay.Size = New System.Drawing.Size(23, 22)
        Me.TBPlay.Text = "Play from here (F6)"
        '
        'TBStop
        '
        Me.TBStop.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBStop.Image = Global.iBMSC.My.Resources.Resources.x16Stop
        Me.TBStop.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBStop.Name = "TBStop"
        Me.TBStop.Size = New System.Drawing.Size(23, 22)
        Me.TBStop.Text = "Stop (F7)"
        '
        'TBPOptions
        '
        Me.TBPOptions.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBPOptions.Image = Global.iBMSC.My.Resources.Resources.x16PlayerOptions
        Me.TBPOptions.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBPOptions.Name = "TBPOptions"
        Me.TBPOptions.Size = New System.Drawing.Size(23, 22)
        Me.TBPOptions.Text = "Player Options"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(6, 25)
        '
        'TBVOptions
        '
        Me.TBVOptions.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBVOptions.Image = Global.iBMSC.My.Resources.Resources.x16VisualOptions
        Me.TBVOptions.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBVOptions.Name = "TBVOptions"
        Me.TBVOptions.Size = New System.Drawing.Size(23, 22)
        Me.TBVOptions.Text = "Visual Options"
        '
        'TBGOptions
        '
        Me.TBGOptions.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBGOptions.Image = Global.iBMSC.My.Resources.Resources.x16GeneralOptions
        Me.TBGOptions.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBGOptions.Name = "TBGOptions"
        Me.TBGOptions.Size = New System.Drawing.Size(23, 22)
        Me.TBGOptions.Text = "General Options"
        '
        'TBAbout
        '
        Me.TBAbout.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TBAbout.Image = Global.iBMSC.My.Resources.Resources.x16About
        Me.TBAbout.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TBAbout.Name = "TBAbout"
        Me.TBAbout.Size = New System.Drawing.Size(23, 22)
        Me.TBAbout.Text = "About"
        '
        'ToolStripSeparator13
        '
        Me.ToolStripSeparator13.Name = "ToolStripSeparator13"
        Me.ToolStripSeparator13.Size = New System.Drawing.Size(6, 25)
        Me.ToolStripSeparator13.Visible = False
        '
        'POBStorm
        '
        Me.POBStorm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.POBStorm.Image = Global.iBMSC.My.Resources.Resources.x16Storm
        Me.POBStorm.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.POBStorm.Name = "POBStorm"
        Me.POBStorm.Size = New System.Drawing.Size(23, 22)
        Me.POBStorm.Text = "Storm"
        Me.POBStorm.Visible = False
        '
        'pStatus
        '
        Me.pStatus.AutoSize = True
        Me.pStatus.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.pStatus.Controls.Add(Me.FStatus2)
        Me.pStatus.Controls.Add(Me.FStatus)
        Me.pStatus.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pStatus.Location = New System.Drawing.Point(0, 708)
        Me.pStatus.Name = "pStatus"
        Me.pStatus.Size = New System.Drawing.Size(877, 22)
        Me.pStatus.TabIndex = 62
        '
        'FStatus2
        '
        Me.FStatus2.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar
        Me.FStatus2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FSSS, Me.FSSL, Me.FSSH, Me.BVCReverse, Me.LblMultiply, Me.TVCM, Me.LblDivide, Me.TVCD, Me.BVCApply, Me.TVCBPM, Me.BVCCalculate})
        Me.FStatus2.Location = New System.Drawing.Point(0, 0)
        Me.FStatus2.Name = "FStatus2"
        Me.FStatus2.ShowItemToolTips = True
        Me.FStatus2.Size = New System.Drawing.Size(803, 22)
        Me.FStatus2.TabIndex = 0
        Me.FStatus2.Text = "Status"
        Me.FStatus2.Visible = False
        '
        'FSSS
        '
        Me.FSSS.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.FSSS.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSSS.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.FSSS.Name = "FSSS"
        Me.FSSS.Size = New System.Drawing.Size(23, 20)
        Me.FSSS.Text = "0"
        Me.FSSS.ToolTipText = "Selection Start Position"
        '
        'FSSL
        '
        Me.FSSL.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.FSSL.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSSL.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.FSSL.Name = "FSSL"
        Me.FSSL.Size = New System.Drawing.Size(23, 20)
        Me.FSSL.Text = "0"
        Me.FSSL.ToolTipText = "Selection Length"
        '
        'FSSH
        '
        Me.FSSH.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.FSSH.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSSH.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.FSSH.Name = "FSSH"
        Me.FSSH.Size = New System.Drawing.Size(23, 20)
        Me.FSSH.Text = "0"
        Me.FSSH.ToolTipText = "Selection Split Position"
        '
        'BVCReverse
        '
        Me.BVCReverse.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BVCReverse.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.BVCReverse.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BVCReverse.Name = "BVCReverse"
        Me.BVCReverse.Size = New System.Drawing.Size(102, 20)
        Me.BVCReverse.Text = "Reverse Selection"
        '
        'LblMultiply
        '
        Me.LblMultiply.Name = "LblMultiply"
        Me.LblMultiply.Size = New System.Drawing.Size(15, 17)
        Me.LblMultiply.Text = "×"
        '
        'TVCM
        '
        Me.TVCM.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TVCM.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TVCM.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TVCM.Name = "TVCM"
        Me.TVCM.Size = New System.Drawing.Size(60, 22)
        Me.TVCM.Text = "2"
        Me.TVCM.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'LblDivide
        '
        Me.LblDivide.Name = "LblDivide"
        Me.LblDivide.Size = New System.Drawing.Size(15, 17)
        Me.LblDivide.Text = "÷"
        '
        'TVCD
        '
        Me.TVCD.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TVCD.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TVCD.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TVCD.Name = "TVCD"
        Me.TVCD.Size = New System.Drawing.Size(60, 22)
        Me.TVCD.Text = "1"
        Me.TVCD.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BVCApply
        '
        Me.BVCApply.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BVCApply.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.BVCApply.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BVCApply.Name = "BVCApply"
        Me.BVCApply.Size = New System.Drawing.Size(71, 20)
        Me.BVCApply.Text = "By Multiple"
        '
        'TVCBPM
        '
        Me.TVCBPM.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TVCBPM.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TVCBPM.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TVCBPM.Name = "TVCBPM"
        Me.TVCBPM.Size = New System.Drawing.Size(80, 22)
        Me.TVCBPM.Text = "120"
        Me.TVCBPM.TextBoxTextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'BVCCalculate
        '
        Me.BVCCalculate.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.BVCCalculate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.BVCCalculate.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BVCCalculate.Name = "BVCCalculate"
        Me.BVCCalculate.Size = New System.Drawing.Size(56, 20)
        Me.BVCCalculate.Text = "By Value"
        '
        'FStatus
        '
        Me.FStatus.AccessibleRole = System.Windows.Forms.AccessibleRole.StatusBar
        Me.FStatus.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FSC, Me.FSW, Me.FSM, Me.FSP1, Me.FSP3, Me.FSP2, Me.FSP4, Me.FST, Me.FSH, Me.FSE})
        Me.FStatus.Location = New System.Drawing.Point(0, 0)
        Me.FStatus.Name = "FStatus"
        Me.FStatus.ShowItemToolTips = True
        Me.FStatus.Size = New System.Drawing.Size(877, 22)
        Me.FStatus.SizingGrip = False
        Me.FStatus.TabIndex = 62
        Me.FStatus.Text = "Status"
        '
        'FSC
        '
        Me.FSC.AutoSize = False
        Me.FSC.Name = "FSC"
        Me.FSC.Size = New System.Drawing.Size(70, 17)
        Me.FSC.Text = "BPM"
        Me.FSC.ToolTipText = "Column Caption"
        '
        'FSW
        '
        Me.FSW.AutoSize = False
        Me.FSW.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.FSW.Name = "FSW"
        Me.FSW.Size = New System.Drawing.Size(40, 17)
        Me.FSW.Text = "01"
        Me.FSW.ToolTipText = "Note Index"
        '
        'FSM
        '
        Me.FSM.AutoSize = False
        Me.FSM.ForeColor = System.Drawing.Color.Teal
        Me.FSM.Name = "FSM"
        Me.FSM.Size = New System.Drawing.Size(40, 17)
        Me.FSM.Text = "000"
        Me.FSM.ToolTipText = "Measure Index"
        '
        'FSP1
        '
        Me.FSP1.AutoSize = False
        Me.FSP1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSP1.Name = "FSP1"
        Me.FSP1.Size = New System.Drawing.Size(170, 17)
        Me.FSP1.Text = "9.41176470588235 / 9999"
        Me.FSP1.ToolTipText = "Grid Resolution"
        '
        'FSP3
        '
        Me.FSP3.AutoSize = False
        Me.FSP3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSP3.ForeColor = System.Drawing.Color.Maroon
        Me.FSP3.Name = "FSP3"
        Me.FSP3.Size = New System.Drawing.Size(85, 17)
        Me.FSP3.Text = "9999 / 9999"
        Me.FSP3.ToolTipText = "Reduced Resolution"
        '
        'FSP2
        '
        Me.FSP2.AutoSize = False
        Me.FSP2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.FSP2.ForeColor = System.Drawing.Color.Green
        Me.FSP2.Name = "FSP2"
        Me.FSP2.Size = New System.Drawing.Size(170, 17)
        Me.FSP2.Text = "112.941176470588 / 9999"
        Me.FSP2.ToolTipText = "Measure Resolution"
        '
        'FSP4
        '
        Me.FSP4.AutoSize = False
        Me.FSP4.Name = "FSP4"
        Me.FSP4.Size = New System.Drawing.Size(115, 17)
        Me.FSP4.Text = "112.941176470588"
        Me.FSP4.ToolTipText = "Absolute Position"
        '
        'FST
        '
        Me.FST.ForeColor = System.Drawing.Color.Olive
        Me.FST.LinkColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.FST.Name = "FST"
        Me.FST.Size = New System.Drawing.Size(64, 17)
        Me.FST.Text = "Length = 0"
        '
        'FSH
        '
        Me.FSH.ForeColor = System.Drawing.Color.Blue
        Me.FSH.Name = "FSH"
        Me.FSH.Size = New System.Drawing.Size(46, 17)
        Me.FSH.Text = "Hidden"
        '
        'FSE
        '
        Me.FSE.ForeColor = System.Drawing.Color.Red
        Me.FSE.Name = "FSE"
        Me.FSE.Size = New System.Drawing.Size(32, 17)
        Me.FSE.Text = "Error"
        '
        'TimerMiddle
        '
        Me.TimerMiddle.Interval = 15
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.PMain)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.SpR)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.SpL)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.PMainR)
        Me.ToolStripContainer1.ContentPanel.Controls.Add(Me.PMainL)
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(877, 660)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(877, 708)
        Me.ToolStripContainer1.TabIndex = 65
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.mnMain)
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.TBMain)
        '
        'PMain
        '
        Me.PMain.BackColor = System.Drawing.Color.Black
        Me.PMain.Controls.Add(Me.PMainIn)
        Me.PMain.Controls.Add(Me.MainPanelScroll)
        Me.PMain.Controls.Add(Me.HS)
        Me.PMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PMain.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMain.ForeColor = System.Drawing.Color.White
        Me.PMain.Location = New System.Drawing.Point(5, 0)
        Me.PMain.Name = "PMain"
        Me.PMain.Size = New System.Drawing.Size(867, 660)
        Me.PMain.TabIndex = 58
        Me.PMain.Tag = "1"
        '
        'PMainIn
        '
        Me.PMainIn.BackColor = System.Drawing.Color.Black
        Me.PMainIn.Controls.Add(Me.ttlIcon)
        Me.PMainIn.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PMainIn.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMainIn.ForeColor = System.Drawing.Color.White
        Me.PMainIn.Location = New System.Drawing.Point(0, 0)
        Me.PMainIn.Name = "PMainIn"
        Me.PMainIn.Size = New System.Drawing.Size(850, 643)
        Me.PMainIn.TabIndex = 0
        Me.PMainIn.TabStop = True
        Me.PMainIn.Tag = "1"
        '
        'ttlIcon
        '
        Me.ttlIcon.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ttlIcon.Image = Global.iBMSC.My.Resources.Resources.icon2_16
        Me.ttlIcon.Location = New System.Drawing.Point(824, 10)
        Me.ttlIcon.Name = "ttlIcon"
        Me.ttlIcon.Size = New System.Drawing.Size(16, 16)
        Me.ttlIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.ttlIcon.TabIndex = 62
        Me.ttlIcon.TabStop = False
        Me.ToolTipUniversal.SetToolTip(Me.ttlIcon, "Show / Hide")
        '
        'MainPanelScroll
        '
        Me.MainPanelScroll.AccessibleRole = System.Windows.Forms.AccessibleRole.ScrollBar
        Me.MainPanelScroll.Dock = System.Windows.Forms.DockStyle.Right
        Me.MainPanelScroll.LargeChange = 592
        Me.MainPanelScroll.Location = New System.Drawing.Point(850, 0)
        Me.MainPanelScroll.Maximum = 591
        Me.MainPanelScroll.Minimum = -10000
        Me.MainPanelScroll.Name = "MainPanelScroll"
        Me.MainPanelScroll.Size = New System.Drawing.Size(17, 643)
        Me.MainPanelScroll.SmallChange = 12
        Me.MainPanelScroll.TabIndex = 2
        Me.MainPanelScroll.Tag = "1"
        '
        'HS
        '
        Me.HS.AccessibleRole = System.Windows.Forms.AccessibleRole.ScrollBar
        Me.HS.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.HS.LargeChange = 777
        Me.HS.Location = New System.Drawing.Point(0, 643)
        Me.HS.Maximum = 1233
        Me.HS.Name = "HS"
        Me.HS.Size = New System.Drawing.Size(867, 17)
        Me.HS.TabIndex = 3
        Me.HS.Tag = "1"
        '
        'SpR
        '
        Me.SpR.Dock = System.Windows.Forms.DockStyle.Right
        Me.SpR.FlatAppearance.BorderSize = 0
        Me.SpR.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SpR.Location = New System.Drawing.Point(872, 0)
        Me.SpR.Name = "SpR"
        Me.SpR.Size = New System.Drawing.Size(5, 660)
        Me.SpR.TabIndex = 59
        Me.SpR.TabStop = False
        Me.SpR.UseVisualStyleBackColor = True
        Me.SpR.Visible = False
        '
        'SpL
        '
        Me.SpL.Dock = System.Windows.Forms.DockStyle.Left
        Me.SpL.FlatAppearance.BorderSize = 0
        Me.SpL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.SpL.Location = New System.Drawing.Point(0, 0)
        Me.SpL.Name = "SpL"
        Me.SpL.Size = New System.Drawing.Size(5, 660)
        Me.SpL.TabIndex = 60
        Me.SpL.TabStop = False
        Me.SpL.UseVisualStyleBackColor = True
        Me.SpL.Visible = False
        '
        'PMainR
        '
        Me.PMainR.BackColor = System.Drawing.Color.Black
        Me.PMainR.Controls.Add(Me.PMainInR)
        Me.PMainR.Controls.Add(Me.RightPanelScroll)
        Me.PMainR.Controls.Add(Me.HSR)
        Me.PMainR.Dock = System.Windows.Forms.DockStyle.Right
        Me.PMainR.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMainR.ForeColor = System.Drawing.Color.White
        Me.PMainR.Location = New System.Drawing.Point(877, 0)
        Me.PMainR.Name = "PMainR"
        Me.PMainR.Size = New System.Drawing.Size(0, 660)
        Me.PMainR.TabIndex = 56
        Me.PMainR.Tag = "2"
        '
        'PMainInR
        '
        Me.PMainInR.BackColor = System.Drawing.Color.Black
        Me.PMainInR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PMainInR.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMainInR.ForeColor = System.Drawing.Color.White
        Me.PMainInR.Location = New System.Drawing.Point(0, 0)
        Me.PMainInR.Name = "PMainInR"
        Me.PMainInR.Size = New System.Drawing.Size(0, 643)
        Me.PMainInR.TabIndex = 0
        Me.PMainInR.TabStop = True
        Me.PMainInR.Tag = "2"
        '
        'RightPanelScroll
        '
        Me.RightPanelScroll.Dock = System.Windows.Forms.DockStyle.Right
        Me.RightPanelScroll.LargeChange = 592
        Me.RightPanelScroll.Location = New System.Drawing.Point(-17, 0)
        Me.RightPanelScroll.Maximum = 591
        Me.RightPanelScroll.Minimum = -10000
        Me.RightPanelScroll.Name = "RightPanelScroll"
        Me.RightPanelScroll.Size = New System.Drawing.Size(17, 643)
        Me.RightPanelScroll.SmallChange = 12
        Me.RightPanelScroll.TabIndex = 2
        Me.RightPanelScroll.Tag = "2"
        '
        'HSR
        '
        Me.HSR.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.HSR.LargeChange = 777
        Me.HSR.Location = New System.Drawing.Point(0, 643)
        Me.HSR.Maximum = 1233
        Me.HSR.Name = "HSR"
        Me.HSR.Size = New System.Drawing.Size(0, 17)
        Me.HSR.TabIndex = 3
        Me.HSR.Tag = "2"
        '
        'PMainL
        '
        Me.PMainL.BackColor = System.Drawing.Color.Black
        Me.PMainL.Controls.Add(Me.PMainInL)
        Me.PMainL.Controls.Add(Me.LeftPanelScroll)
        Me.PMainL.Controls.Add(Me.HSL)
        Me.PMainL.Dock = System.Windows.Forms.DockStyle.Left
        Me.PMainL.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMainL.ForeColor = System.Drawing.Color.White
        Me.PMainL.Location = New System.Drawing.Point(0, 0)
        Me.PMainL.Name = "PMainL"
        Me.PMainL.Size = New System.Drawing.Size(0, 660)
        Me.PMainL.TabIndex = 54
        Me.PMainL.Tag = "0"
        '
        'PMainInL
        '
        Me.PMainInL.BackColor = System.Drawing.Color.Black
        Me.PMainInL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PMainInL.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PMainInL.ForeColor = System.Drawing.Color.White
        Me.PMainInL.Location = New System.Drawing.Point(0, 0)
        Me.PMainInL.Name = "PMainInL"
        Me.PMainInL.Size = New System.Drawing.Size(0, 643)
        Me.PMainInL.TabIndex = 0
        Me.PMainInL.TabStop = True
        Me.PMainInL.Tag = "0"
        '
        'LeftPanelScroll
        '
        Me.LeftPanelScroll.Dock = System.Windows.Forms.DockStyle.Right
        Me.LeftPanelScroll.LargeChange = 592
        Me.LeftPanelScroll.Location = New System.Drawing.Point(-17, 0)
        Me.LeftPanelScroll.Maximum = 591
        Me.LeftPanelScroll.Minimum = -10000
        Me.LeftPanelScroll.Name = "LeftPanelScroll"
        Me.LeftPanelScroll.Size = New System.Drawing.Size(17, 643)
        Me.LeftPanelScroll.SmallChange = 12
        Me.LeftPanelScroll.TabIndex = 2
        Me.LeftPanelScroll.Tag = "0"
        '
        'HSL
        '
        Me.HSL.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.HSL.LargeChange = 777
        Me.HSL.Location = New System.Drawing.Point(0, 643)
        Me.HSL.Maximum = 1233
        Me.HSL.Name = "HSL"
        Me.HSL.Size = New System.Drawing.Size(0, 17)
        Me.HSL.TabIndex = 3
        Me.HSL.Tag = "0"
        '
        'POptionsResizer
        '
        Me.POptionsResizer.Dock = System.Windows.Forms.DockStyle.Right
        Me.POptionsResizer.FlatAppearance.BorderSize = 0
        Me.POptionsResizer.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.POptionsResizer.Location = New System.Drawing.Point(877, 0)
        Me.POptionsResizer.Name = "POptionsResizer"
        Me.POptionsResizer.Size = New System.Drawing.Size(5, 730)
        Me.POptionsResizer.TabIndex = 67
        Me.POptionsResizer.TabStop = False
        Me.POptionsResizer.UseVisualStyleBackColor = True
        '
        'MainWindow
        '
        Me.AllowDrop = True
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(1082, 730)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Controls.Add(Me.pStatus)
        Me.Controls.Add(Me.POptionsResizer)
        Me.Controls.Add(Me.POptionsScroll)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MainMenuStrip = Me.mnMain
        Me.Name = "MainWindow"
        Me.cmnLanguage.ResumeLayout(False)
        Me.cmnTheme.ResumeLayout(False)
        Me.POptionsScroll.ResumeLayout(False)
        Me.POptionsScroll.PerformLayout()
        Me.POptions.ResumeLayout(False)
        Me.POptions.PerformLayout()
        Me.POExpansion.ResumeLayout(False)
        Me.POExpansionInner.ResumeLayout(False)
        Me.POExpansionInner.PerformLayout()
        Me.POBeat.ResumeLayout(False)
        Me.POBeatInner.ResumeLayout(False)
        Me.POBeatInner.PerformLayout()
        Me.TableLayoutPanel7.ResumeLayout(False)
        Me.TableLayoutPanel7.PerformLayout()
        CType(Me.nBeatD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nBeatN, System.ComponentModel.ISupportInitialize).EndInit()
        Me.POBeatPart2.ResumeLayout(False)
        Me.POBeatPart2.PerformLayout()
        Me.POWAV.ResumeLayout(False)
        Me.POWAVInner.ResumeLayout(False)
        Me.POWAVInner.PerformLayout()
        Me.FlowLayoutPanel3.ResumeLayout(False)
        Me.POWAVPart2.ResumeLayout(False)
        Me.POWAVPart2.PerformLayout()
        Me.POWaveForm.ResumeLayout(False)
        Me.POWaveForm.PerformLayout()
        Me.POWaveFormInner.ResumeLayout(False)
        Me.POWaveFormInner.PerformLayout()
        Me.POWaveFormPart2.ResumeLayout(False)
        Me.POWaveFormPart2.PerformLayout()
        CType(Me.TWSaturation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWTransparency, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWPrecision, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWWidth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWLeft, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWSaturation2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWLeft2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWTransparency2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWWidth2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWPrecision2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.POWaveFormPart1.ResumeLayout(False)
        Me.POWaveFormPart1.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.TableLayoutPanel6.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWPosition2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TWPosition, System.ComponentModel.ISupportInitialize).EndInit()
        Me.POGrid.ResumeLayout(False)
        Me.POGrid.PerformLayout()
        Me.POGridInner.ResumeLayout(False)
        Me.POGridInner.PerformLayout()
        Me.POGridPart2.ResumeLayout(False)
        Me.POGridPart2.PerformLayout()
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        CType(Me.CGB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.POGridPart1.ResumeLayout(False)
        Me.POGridPart1.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGHeight2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGHeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGWidth2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGWidth, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGDivide, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CGSub, System.ComponentModel.ISupportInitialize).EndInit()
        Me.POHeader.ResumeLayout(False)
        Me.POHeader.PerformLayout()
        Me.POHeaderInner.ResumeLayout(False)
        Me.POHeaderInner.PerformLayout()
        Me.POHeaderPart2.ResumeLayout(False)
        Me.POHeaderPart2.PerformLayout()
        Me.POHeaderPart1.ResumeLayout(False)
        Me.POHeaderPart1.PerformLayout()
        CType(Me.THBPM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Menu1.ResumeLayout(False)
        Me.mnMain.ResumeLayout(False)
        Me.mnMain.PerformLayout()
        Me.cmnConversion.ResumeLayout(False)
        Me.TBMain.ResumeLayout(False)
        Me.TBMain.PerformLayout()
        Me.pStatus.ResumeLayout(False)
        Me.pStatus.PerformLayout()
        Me.FStatus2.ResumeLayout(False)
        Me.FStatus2.PerformLayout()
        Me.FStatus.ResumeLayout(False)
        Me.FStatus.PerformLayout()
        Me.ToolStripContainer1.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.PMain.ResumeLayout(False)
        Me.PMainIn.ResumeLayout(False)
        Me.PMainIn.PerformLayout()
        CType(Me.ttlIcon, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PMainR.ResumeLayout(False)
        Me.PMainL.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents POptionsScroll As System.Windows.Forms.Panel
    Friend WithEvents LWAV As System.Windows.Forms.ListBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MInsert As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MRemove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Menu1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AutoSaveTimer As System.Windows.Forms.Timer
    Friend WithEvents mnMain As System.Windows.Forms.MenuStrip
    Friend WithEvents mnFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnOpen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnImportSM As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnImportIBMSC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator14 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSaveAs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnExport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator15 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnOpenR0 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnOpenR1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnOpenR2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnOpenR3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnOpenR4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator16 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnQuit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnUndo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnRedo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator17 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnCut As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnPaste As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator18 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnFind As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnStatistics As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnDelete As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSelectAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnMyO2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator19 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnPreview As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnPlayB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnPlay As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnStop As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnTimeSelect As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSelect As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnWrite As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator23 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnConversion As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnErrorCheck As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnPreviewOnClick As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnShowFileName As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator20 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnGOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnVOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnLanguage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnTheme As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmnLanguage As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents TBLangDef As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator9 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents cmnTheme As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents TBThemeDef As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBThemeSave As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBThemeRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents cmnConversion As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents POBLong As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POBShort As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POBLongShort As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator10 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents POBHidden As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POBVisible As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POBHiddenVisible As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator11 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents POBMirror As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POBModify As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSys As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSMenu As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSTB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSOP As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSStatus As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSLSplitter As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnSRSplitter As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnNTInput As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBLangRefresh As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGSub As System.Windows.Forms.NumericUpDown
    Friend WithEvents CGDivide As System.Windows.Forms.NumericUpDown
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents CGSnap As System.Windows.Forms.CheckBox
    Friend WithEvents BWLock As System.Windows.Forms.CheckBox
    Friend WithEvents TWSaturation As System.Windows.Forms.NumericUpDown
    Friend WithEvents TWTransparency As System.Windows.Forms.NumericUpDown
    Friend WithEvents TWPrecision As System.Windows.Forms.NumericUpDown
    Friend WithEvents TWWidth As System.Windows.Forms.NumericUpDown
    Friend WithEvents TWLeft As System.Windows.Forms.NumericUpDown
    Friend WithEvents TWPosition As System.Windows.Forms.NumericUpDown
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents TWSaturation2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWTransparency2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWPrecision2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWWidth2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWLeft2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWPosition2 As System.Windows.Forms.TrackBar
    Friend WithEvents TWFileName As System.Windows.Forms.TextBox
    Friend WithEvents BWClear As System.Windows.Forms.Button
    Friend WithEvents BWLoad As System.Windows.Forms.Button
    Friend WithEvents TBMain As System.Windows.Forms.ToolStrip
    Friend WithEvents TBNew As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBOpen As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents TBOpenR0 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBOpenR1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBOpenR2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBOpenR3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBOpenR4 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator12 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBImportSM As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBImportIBMSC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBSave As System.Windows.Forms.ToolStripSplitButton
    Friend WithEvents TBSaveAs As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TBExport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBCut As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBCopy As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBPaste As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBFind As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBStatistics As System.Windows.Forms.ToolStripButton
    Friend WithEvents pStatus As System.Windows.Forms.Panel
    Friend WithEvents FStatus As System.Windows.Forms.StatusStrip
    Friend WithEvents FSC As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSP1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSP2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSP3 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSP4 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FST As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSH As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSE As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FStatus2 As System.Windows.Forms.StatusStrip
    Friend WithEvents BVCReverse As System.Windows.Forms.ToolStripButton
    Friend WithEvents LblMultiply As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TVCBPM As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents TVCM As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents LblDivide As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TVCD As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BVCApply As System.Windows.Forms.ToolStripButton
    Friend WithEvents BVCCalculate As System.Windows.Forms.ToolStripButton
    Friend WithEvents FSW As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents FSSS As System.Windows.Forms.ToolStripButton
    Friend WithEvents FSSL As System.Windows.Forms.ToolStripButton
    Friend WithEvents FSSH As System.Windows.Forms.ToolStripButton
    Friend WithEvents TExpansion As System.Windows.Forms.TextBox
    Friend WithEvents TimerMiddle As System.Windows.Forms.Timer
    Friend WithEvents ToolStripSeparator21 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CGShow As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowS As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowBG As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowM As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowMB As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowV As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGShowC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator22 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CGBLP As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CGSTOP As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents nBeatN As System.Windows.Forms.NumericUpDown
    Friend WithEvents LBeat As System.Windows.Forms.ListBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents BBeatApply As System.Windows.Forms.Button
    Friend WithEvents mnPOptions As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents PMain As System.Windows.Forms.Panel
    Friend WithEvents PMainIn As System.Windows.Forms.Panel
    Friend WithEvents ttlIcon As System.Windows.Forms.PictureBox
    Friend WithEvents MainPanelScroll As System.Windows.Forms.VScrollBar
    Friend WithEvents HS As System.Windows.Forms.HScrollBar
    Friend WithEvents PMainR As System.Windows.Forms.Panel
    Friend WithEvents PMainInR As System.Windows.Forms.Panel
    Friend WithEvents RightPanelScroll As System.Windows.Forms.VScrollBar
    Friend WithEvents HSR As System.Windows.Forms.HScrollBar
    Friend WithEvents PMainL As System.Windows.Forms.Panel
    Friend WithEvents PMainInL As System.Windows.Forms.Panel
    Friend WithEvents LeftPanelScroll As System.Windows.Forms.VScrollBar
    Friend WithEvents HSL As System.Windows.Forms.HScrollBar
    Friend WithEvents POConvert As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents TBLanguage As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents TBTheme As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents TBMyO2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBErrorCheck As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBPreviewOnClick As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBShowFileName As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBUndo As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBRedo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBNTInput As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBTimeSelect As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBSelect As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBWrite As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBPlayB As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBPlay As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBStop As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBPOptions As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator7 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TBVOptions As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBGOptions As System.Windows.Forms.ToolStripButton
    Friend WithEvents TBAbout As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator13 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents POBStorm As System.Windows.Forms.ToolStripButton
    Friend WithEvents POptions As System.Windows.Forms.Panel
    Friend WithEvents POHeader As System.Windows.Forms.Panel
    Friend WithEvents POHeaderSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents POGrid As System.Windows.Forms.Panel
    Friend WithEvents POGridSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents POHeaderInner As System.Windows.Forms.Panel
    Friend WithEvents POHeaderPart2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents THExRank As System.Windows.Forms.TextBox
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents CHLnObj As System.Windows.Forms.ComboBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents THComment As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents THTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents BHStageFile As System.Windows.Forms.Button
    Friend WithEvents BHBanner As System.Windows.Forms.Button
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents BHBackBMP As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents THBackBMP As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents THBanner As System.Windows.Forms.TextBox
    Friend WithEvents THStageFile As System.Windows.Forms.TextBox
    Friend WithEvents THSubTitle As System.Windows.Forms.TextBox
    Friend WithEvents THSubArtist As System.Windows.Forms.TextBox
    Friend WithEvents POHeaderExpander As System.Windows.Forms.CheckBox
    Friend WithEvents POHeaderPart1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents THPlayLevel As System.Windows.Forms.TextBox
    Friend WithEvents CHRank As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CHPlayer As System.Windows.Forms.ComboBox
    Friend WithEvents CHDifficulty As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents THGenre As System.Windows.Forms.TextBox
    Friend WithEvents THBPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents THArtist As System.Windows.Forms.TextBox
    Friend WithEvents THTitle As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents POGridPart1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents POGridInner As System.Windows.Forms.Panel
    Friend WithEvents POGridPart2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel5 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents cVSLockR As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cVSLock As System.Windows.Forms.CheckBox
    Friend WithEvents cVSLockL As System.Windows.Forms.CheckBox
    Friend WithEvents CGDisableVertical As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel4 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CGB As System.Windows.Forms.NumericUpDown
    Friend WithEvents POGridExpander As System.Windows.Forms.CheckBox
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents CGHeight2 As System.Windows.Forms.TrackBar
    Friend WithEvents CGHeight As System.Windows.Forms.NumericUpDown
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents CGWidth2 As System.Windows.Forms.TrackBar
    Friend WithEvents CGWidth As System.Windows.Forms.NumericUpDown
    Friend WithEvents POWaveForm As System.Windows.Forms.Panel
    Friend WithEvents POWaveFormSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents POWaveFormInner As System.Windows.Forms.Panel
    Friend WithEvents POWaveFormPart2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents POWaveFormExpander As System.Windows.Forms.CheckBox
    Friend WithEvents POWaveFormPart1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel6 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents FlowLayoutPanel2 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents POWAV As System.Windows.Forms.Panel
    Friend WithEvents POWAVInner As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents POWAVSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents POBeat As System.Windows.Forms.Panel
    Friend WithEvents POBeatInner As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents POBeatSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents FlowLayoutPanel3 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents BWAVUp As System.Windows.Forms.Button
    Friend WithEvents BWAVDown As System.Windows.Forms.Button
    Friend WithEvents BWAVBrowse As System.Windows.Forms.Button
    Friend WithEvents BWAVRemove As System.Windows.Forms.Button
    Friend WithEvents POExpansion As System.Windows.Forms.Panel
    Friend WithEvents POExpansionInner As System.Windows.Forms.Panel
    Friend WithEvents POExpansionSwitch As System.Windows.Forms.CheckBox
    Friend WithEvents POWAVResizer As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel7 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents POExpansionResizer As System.Windows.Forms.Button
    Friend WithEvents POBeatResizer As System.Windows.Forms.Button
    Friend WithEvents POptionsResizer As System.Windows.Forms.Button
    Friend WithEvents SpR As System.Windows.Forms.Button
    Friend WithEvents SpL As System.Windows.Forms.Button
    Friend WithEvents TBThemeLoadComptability As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POWAVExpander As System.Windows.Forms.CheckBox
    Friend WithEvents POWAVPart2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents CWAVMultiSelect As System.Windows.Forms.CheckBox
    Friend WithEvents CWAVChangeLabel As System.Windows.Forms.CheckBox
    Friend WithEvents nBeatD As System.Windows.Forms.NumericUpDown
    Friend WithEvents BBeatApplyV As System.Windows.Forms.Button
    Friend WithEvents tBeatValue As System.Windows.Forms.TextBox
    Friend WithEvents POBeatExpander As System.Windows.Forms.CheckBox
    Friend WithEvents POBeatPart2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents CBeatScale As System.Windows.Forms.RadioButton
    Friend WithEvents CBeatCut As System.Windows.Forms.RadioButton
    Friend WithEvents CBeatMeasure As System.Windows.Forms.RadioButton
    Friend WithEvents CBeatPreserve As System.Windows.Forms.RadioButton
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents ToolTipUniversal As System.Windows.Forms.ToolTip
    Friend WithEvents CGBPM As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BGSlash As System.Windows.Forms.Button
    Friend WithEvents FSM As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mnGotoMeasure As ToolStripMenuItem
End Class
